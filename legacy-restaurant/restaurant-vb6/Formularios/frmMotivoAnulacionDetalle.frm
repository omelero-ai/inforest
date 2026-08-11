VERSION 5.00
Begin VB.Form frmMotivoAnulacionDetalle 
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10005
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2160
      Width           =   10065
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5190
         TabIndex        =   13
         Top             =   60
         Width           =   5250
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3750
            Picture         =   "frmMotivoAnulacionDetalle.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4230
            Picture         =   "frmMotivoAnulacionDetalle.frx":0542
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4710
            Picture         =   "frmMotivoAnulacionDetalle.frx":0A84
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmMotivoAnulacionDetalle.frx":0FC6
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmMotivoAnulacionDetalle.frx":1508
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmMotivoAnulacionDetalle.frx":1A4A
            Style           =   1  'Graphical
            TabIndex        =   14
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
            TabIndex        =   20
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
         Picture         =   "frmMotivoAnulacionDetalle.frx":1F8C
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmMotivoAnulacionDetalle.frx":24BE
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmMotivoAnulacionDetalle.frx":29F0
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmMotivoAnulacionDetalle.frx":2AF2
         Style           =   1  'Graphical
         TabIndex        =   9
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
      Height          =   2100
      Left            =   2160
      TabIndex        =   0
      Top             =   0
      Width           =   7875
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1755
         TabIndex        =   4
         Top             =   1530
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   315
         Width           =   1170
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
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   2
         Text            =   " "
         Top             =   720
         Width           =   5325
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
         Left            =   2430
         MaxLength       =   24
         TabIndex        =   1
         Text            =   " "
         Top             =   1125
         Width           =   2595
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   1740
         TabIndex        =   7
         Top             =   375
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción (Reportes) :"
         Height          =   195
         Index           =   1
         Left            =   615
         TabIndex        =   6
         Top             =   765
         Width           =   1710
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descrip. Resumida (Botonera) :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1140
         Width           =   2205
      End
   End
   Begin VB.Image Image 
      Height          =   2145
      Left            =   0
      Picture         =   "frmMotivoAnulacionDetalle.frx":2BE4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2100
   End
End
Attribute VB_Name = "frmMotivoAnulacionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmMotivoAnulacion.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmMotivoAnulacion.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmMotivoAnulacion.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmMotivoAnulacion.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmMotivoAnulacion.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmMotivoAnulacion.grdGrilla
    End Select
    Asignar
    cmdTexto.Caption = "Registro " & IIf(frmMotivoAnulacion.RsCabecera.RecordCount = 0, 0, frmMotivoAnulacion.RsCabecera.AbsolutePosition) & " de " & frmMotivoAnulacion.RsCabecera.RecordCount
End Sub
Sub Asignar()
    With frmMotivoAnulacion.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
       
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
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
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                                   
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='MOTIVOANULACION' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 3)
                  End If
                  Sw = False
                   
                  
                 sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "MOTIVO ELIMINACION", "01", sUsuario, sPasa, "", _
                   "TTABLA", "TABLA", "MOTIVOANULACION", "tCodigo", "Codigo Motivo", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumido", txtResumido.Text, "lActivo", "Flag Activo", IIf(chkActivo = 1, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                   
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, nValor, tValor, tIcono, lActivo) " & _
                         "values ('MOTIVOANULACION', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', 0, '', '', " & _
                                       chkActivo.Value & ") "
                  Cn.Execute Isql
                  
                     
                  
                  frmMotivoAnulacion.RsCabecera.Sort = "Codigo ASC"
                  frmMotivoAnulacion.RsCabecera.Requery
                  frmMotivoAnulacion.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmMotivoAnulacion.RsCabecera.RecordCount = 0, 0, frmMotivoAnulacion.RsCabecera.AbsolutePosition) & " de " & frmMotivoAnulacion.RsCabecera.RecordCount
               Else
               
                sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "MOTIVO ANULACION", "01", sUsuario, sPasa, "", _
                   "TTABLA", "TABLA", "MOTIVOANULACION", "tCodigo", "Codigo Motivo", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumido", txtResumido.Text, "lActivo", "Flag Activo", IIf(chkActivo = 1, "Verdadero", "Falso"))
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
               
               
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "lActivo =" & chkActivo.Value & ", lReplica=1 " & _
                         " where tTAbla = 'MOTIVOANULACION' and tCodigo = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                   nPos = frmMotivoAnulacion.RsCabecera.Bookmark
                   frmMotivoAnulacion.RsCabecera.Requery
                   If frmMotivoAnulacion.RsCabecera.RecordCount = 0 Then
                      frmMotivoAnulacion.RsCabecera.Filter = adFilterNone
                   End If
                   frmMotivoAnulacion.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
                Cn.Execute "update TTABLA set lNuevoModificado=1 where tTAbla = 'MOTIVOANULACION' and tCodigo = '" & txtCodigo & "'"
                
          Case Is = 2 ' Eliminar
               If frmMotivoAnulacion.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Motivo " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
                         
                 sPasa = txtCodigo.Text
                'Inserta Movimiento en este el codigo del registro viene del "txtcodigo.text"
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "MOTIVO ANULACION", "03", sUsuario, txtCodigo.Text, "", "TTABLA", "TABLA", "MOTIVOANULACION", "tCodigo", "Código Motivo", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text)
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.

               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'MOTIVOANULACION' and tCodigo = '" & txtCodigo & "'"
               frmMotivoAnulacion.RsCabecera.Requery
               If frmMotivoAnulacion.RsCabecera.RecordCount <> 0 Then
                  frmMotivoAnulacion.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmMotivoAnulacion.RsCabecera.RecordCount = 0, 0, frmMotivoAnulacion.RsCabecera.AbsolutePosition) & " de " & frmMotivoAnulacion.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If

          
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub

Private Sub txtDetallado_LostFocus()
   Call ValidaStr(txtDetallado)
End Sub

Private Sub txtResumido_LostFocus()
   Call ValidaStr(txtResumido)
End Sub

Private Sub Form_Load()
 Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Motivos de Anulación "
    fraDetalle.Caption = Me.Caption
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmMotivoAnulacion.RsCabecera.RecordCount = 0, 0, frmMotivoAnulacion.RsCabecera.AbsolutePosition) & " de " & frmMotivoAnulacion.RsCabecera.RecordCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Cambia el Nombre del Formulario
    Set frmMotivoAnulacionDetalle = Nothing
End Sub

