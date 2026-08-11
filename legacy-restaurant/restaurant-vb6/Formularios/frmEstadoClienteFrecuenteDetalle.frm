VERSION 5.00
Begin VB.Form frmEstadoClienteFrecuenteDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10020
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2400
      Width           =   10080
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5190
         TabIndex        =   21
         Top             =   60
         Width           =   5250
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3750
            Picture         =   "frmEstadoClienteFrecuenteDetalle.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4230
            Picture         =   "frmEstadoClienteFrecuenteDetalle.frx":0542
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4710
            Picture         =   "frmEstadoClienteFrecuenteDetalle.frx":0A84
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmEstadoClienteFrecuenteDetalle.frx":0FC6
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmEstadoClienteFrecuenteDetalle.frx":1508
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmEstadoClienteFrecuenteDetalle.frx":1A4A
            Style           =   1  'Graphical
            TabIndex        =   7
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
            Left            =   1440
            TabIndex        =   22
            Top             =   180
            Width           =   2265
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
         Left            =   6510
         Picture         =   "frmEstadoClienteFrecuenteDetalle.frx":1F8C
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   5340
         Picture         =   "frmEstadoClienteFrecuenteDetalle.frx":24BE
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   7680
         Picture         =   "frmEstadoClienteFrecuenteDetalle.frx":29F0
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   8850
         Picture         =   "frmEstadoClienteFrecuenteDetalle.frx":2AF2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   1170
      End
   End
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
      Height          =   2340
      Left            =   2175
      TabIndex        =   16
      Top             =   45
      Width           =   7875
      Begin VB.CheckBox chkpermitido 
         Alignment       =   1  'Right Justify
         Caption         =   "No permite seleccionar al cliente"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1750
         Width           =   2895
      End
      Begin VB.CheckBox chkAlerta 
         Alignment       =   1  'Right Justify
         Caption         =   "Muestra Alerta al seleccionar Cliente :"
         Height          =   195
         Left            =   270
         TabIndex        =   3
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   2400
         TabIndex        =   5
         Top             =   2040
         Width           =   855
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
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   315
         Width           =   570
      End
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
         Left            =   3030
         MaxLength       =   20
         TabIndex        =   1
         Text            =   " "
         Top             =   720
         Width           =   4725
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
         Left            =   3030
         MaxLength       =   20
         TabIndex        =   2
         Text            =   " "
         Top             =   1050
         Width           =   1995
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   2340
         TabIndex        =   19
         Top             =   375
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción (Reportes) :"
         Height          =   195
         Index           =   1
         Left            =   1215
         TabIndex        =   18
         Top             =   765
         Width           =   1710
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descrip. Resumida (Botonera) :"
         Height          =   195
         Left            =   720
         TabIndex        =   17
         Top             =   1050
         Width           =   2205
      End
   End
   Begin VB.Image Image 
      Height          =   2385
      Left            =   45
      Picture         =   "frmEstadoClienteFrecuenteDetalle.frx":2BE4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2100
   End
End
Attribute VB_Name = "frmEstadoClienteFrecuenteDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Asignar()
    With frmEstadoClienteFrecuente.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
       If IsNull(!nBoton) Or (!nBoton = 0) Then
            chkAlerta.value = 0
        Else
            chkAlerta.value = 1
       End If
       
       If IsNull(!nValor) Or (!nValor = 0) Then
            chkpermitido.value = 0
        Else
            chkpermitido.value = 1
       End If
       
       
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
        
        
       
    End With
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmEstadoClienteFrecuente.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmEstadoClienteFrecuente.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmEstadoClienteFrecuente.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmEstadoClienteFrecuente.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmEstadoClienteFrecuente.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmEstadoClienteFrecuente.grdGrilla
    End Select
    Asignar
    cmdTexto.Caption = "Registro " & IIf(frmEstadoClienteFrecuente.RsCabecera.RecordCount = 0, 0, frmEstadoClienteFrecuente.RsCabecera.AbsolutePosition) & " de " & frmEstadoClienteFrecuente.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.value = 1
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
                    
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                                   
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='ESTADOFRECUENTE' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, nValor, tValor, tIcono,nboton, lActivo) " & _
                         "values ('ESTADOFRECUENTE', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & IIf(chkpermitido.value, 1, 0) & ", '', ''," & IIf(chkAlerta.value, 1, 0) & " , " & _
                                       chkActivo.value & ") "
                  Cn.Execute Isql
                  frmEstadoClienteFrecuente.RsCabecera.Sort = "Codigo ASC"
                  frmEstadoClienteFrecuente.RsCabecera.Requery
                  frmEstadoClienteFrecuente.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmEstadoClienteFrecuente.RsCabecera.RecordCount = 0, 0, frmEstadoClienteFrecuente.RsCabecera.AbsolutePosition) & " de " & frmEstadoClienteFrecuente.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "lActivo =" & chkActivo.value & " , nValor=" & IIf(chkpermitido.value, 1, 0) & " , NBOTON=" & IIf(chkAlerta.value, 1, 0) & ", lReplica=1 " & _
                         " where tTAbla = 'ESTADOFRECUENTE' and tCodigo = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                   nPos = frmEstadoClienteFrecuente.RsCabecera.Bookmark
                   frmEstadoClienteFrecuente.RsCabecera.Requery
                   If frmEstadoClienteFrecuente.RsCabecera.RecordCount = 0 Then
                      frmEstadoClienteFrecuente.RsCabecera.Filter = adFilterNone
                   End If
                   frmEstadoClienteFrecuente.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
               If frmEstadoClienteFrecuente.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Motivo " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'ESTADOFRECUENTE' and tCodigo = '" & txtCodigo & "'"
               frmEstadoClienteFrecuente.RsCabecera.Requery
               If frmEstadoClienteFrecuente.RsCabecera.RecordCount <> 0 Then
                  frmEstadoClienteFrecuente.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmEstadoClienteFrecuente.RsCabecera.RecordCount = 0, 0, frmEstadoClienteFrecuente.RsCabecera.AbsolutePosition) & " de " & frmEstadoClienteFrecuente.RsCabecera.RecordCount
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
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Estado de Cliente Frecuente "
    fraDetalle.Caption = Me.Caption
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmEstadoClienteFrecuente.RsCabecera.RecordCount = 0, 0, frmEstadoClienteFrecuente.RsCabecera.AbsolutePosition) & " de " & frmEstadoClienteFrecuente.RsCabecera.RecordCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmEstadoClienteFrecuenteDetalle = Nothing
End Sub

Private Sub cmdAgregar_AfterClick()
    Sw = True
    ActivarBotones (False)
    Blanquear Me
    chkActivo.value = 1
    chkAlerta.value = 0
    chkpermitido.value = 0
    'Cambia el Nombre del Primer Text
    txtDetallado.SetFocus
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


