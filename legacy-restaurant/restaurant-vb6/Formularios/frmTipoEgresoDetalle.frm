VERSION 5.00
Begin VB.Form frmTipoEgresoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2820
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9450
   Icon            =   "frmTipoEgresoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   9450
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
      Height          =   2025
      Left            =   2100
      TabIndex        =   16
      Top             =   0
      Width           =   7305
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Top             =   1440
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   330
         Width           =   1170
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
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   1
         Text            =   " "
         Top             =   1038
         Width           =   2595
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
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   0
         Text            =   " "
         Top             =   684
         Width           =   5250
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   20
         Top             =   729
         Width           =   1650
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   19
         Top             =   1083
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   1290
         TabIndex        =   18
         Top             =   375
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9390
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2070
      Width           =   9450
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   14
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "frmTipoEgresoDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmTipoEgresoDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmTipoEgresoDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTipoEgresoDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTipoEgresoDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTipoEgresoDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   4
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
            TabIndex        =   15
            Top             =   150
            Width           =   1665
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
         Left            =   5880
         Picture         =   "frmTipoEgresoDetalle.frx":23CE
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   4710
         Picture         =   "frmTipoEgresoDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   7050
         Picture         =   "frmTipoEgresoDetalle.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   8220
         Picture         =   "frmTipoEgresoDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   2040
      Left            =   15
      Picture         =   "frmTipoEgresoDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2040
   End
End
Attribute VB_Name = "frmTipoEgresoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
 Sub Asignar()
    With frmTipoEgreso.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTipoEgreso.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTipoEgreso.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTipoEgreso.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTipoEgreso.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTipoEgreso.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTipoEgreso.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmTipoEgreso.RsCabecera.RecordCount = 0, 0, frmTipoEgreso.RsCabecera.AbsolutePosition) & " de " & frmTipoEgreso.RsCabecera.RecordCount
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
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                    
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='TIPOEGRESO' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 3)
                  End If
                  Sw = False
                                  
                sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "TIPO EGRESO", "01", sUsuario, sPasa, "", _
                   "TTABLA", "TABLA", "TIPOEGRESO", "tCodigo", "Codigo EGRESO", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumido", txtResumido.Text, "lActivo", "Flag Activo", IIf(chkActivo = 1, "Verdadero", "Falso"))
                If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                                                                    
                                  
                                  
                  'Cambiar el SQL 'KDS
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido,   lActivo) " & _
                         "values ('TIPOEGRESO', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       chkActivo.value & ") "
                  Cn.Execute Isql
                  ActivarBotones (True)
                  frmTipoEgreso.RsCabecera.Sort = "Codigo ASC"
                  frmTipoEgreso.RsCabecera.Requery
                  frmTipoEgreso.RsCabecera.MoveLast
                  cmdTexto.Caption = "Registro " & IIf(frmTipoEgreso.RsCabecera.RecordCount = 0, 0, frmTipoEgreso.RsCabecera.AbsolutePosition) & " de " & frmTipoEgreso.RsCabecera.RecordCount
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  
               Else
               
                 sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "TIPO EGRESO", "02", sUsuario, sPasa, "", _
                   "TTABLA", "TABLA", "TIPOEGRESO", "tCodigo", "Codigo EGRESO", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumido", txtResumido.Text, "lActivo", "Flag Activo", IIf(chkActivo = 1, "Verdadero", "Falso"))
                If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                                                                    
                       
               
               
               
                  'Cambiar el SQL 'KDS
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "lActivo =" & chkActivo.value & " " & _
                         " where tTAbla = 'TIPOEGRESO' and tCodigo = '" & txtCodigo & "'"
                   Cn.Execute Isql
                   nPos = frmTipoEgreso.RsCabecera.Bookmark
                   frmTipoEgreso.RsCabecera.Requery
                   If frmTipoEgreso.RsCabecera.RecordCount = 0 Then
                      frmTipoEgreso.RsCabecera.Filter = adFilterNone
                   End If
                   frmTipoEgreso.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               
          Case Is = 2 ' Eliminar
               If frmTipoEgreso.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If Calcular("select count(*)  as codigo from megreso where ttipoegreso='" & txtCodigo & "' ", Cn) > 0 Then
                MsgBox "Se tiene movimientos con este Tipo de Egreso " & vbCrLf & "No es posible Eliminar"
                Exit Sub
               
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el TipoEgreso " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
                sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "TIPO EGRESO", "03", sUsuario, sPasa, "", _
                   "TTABLA", "TABLA", "TIPOEGRESO", "tCodigo", "Codigo EGRESO", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text)
                If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                                                                    
                                      
               
               
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'TIPOEGRESO' and tCodigo = '" & txtCodigo & "'"
               frmTipoEgreso.RsCabecera.Requery
               If frmTipoEgreso.RsCabecera.RecordCount <> 0 Then
                  frmTipoEgreso.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmTipoEgreso.RsCabecera.RecordCount = 0, 0, frmTipoEgreso.RsCabecera.AbsolutePosition) & " de " & frmTipoEgreso.RsCabecera.RecordCount
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
    Me.Caption = " Tipos de Egresos "
    fraDetalle.Caption = Me.Caption
     
    'Ingresar la Tabla
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
      
    Else
       ActivarBotones (True)
       Asignar
    End If
    cmdTexto.Caption = "Registro " & IIf(frmTipoEgreso.RsCabecera.RecordCount = 0, 0, frmTipoEgreso.RsCabecera.AbsolutePosition) & " de " & frmTipoEgreso.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTipoEgresoDetalle = Nothing
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


