VERSION 5.00
Begin VB.Form frmTipoCancelacionDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3300
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   10095
   Icon            =   "frmTipoCancelacionDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   10095
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
      Height          =   2505
      Left            =   2175
      TabIndex        =   17
      Top             =   45
      Width           =   7875
      Begin VB.CheckBox ChkCreditoCorporativo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activa credito corporativo :"
         Height          =   270
         Left            =   330
         TabIndex        =   23
         Top             =   1920
         Width           =   2250
      End
      Begin VB.TextBox txtCuentaContable 
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
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   3
         Text            =   " "
         Top             =   1365
         Width           =   2595
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   270
         Left            =   1720
         TabIndex        =   4
         Top             =   1680
         Width           =   855
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   765
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
         Left            =   2400
         MaxLength       =   24
         TabIndex        =   2
         Text            =   " "
         Top             =   1065
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable :"
         Height          =   195
         Left            =   1050
         TabIndex        =   22
         Top             =   1410
         Width           =   1275
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción (Reportes) :"
         Height          =   195
         Index           =   1
         Left            =   615
         TabIndex        =   21
         Top             =   810
         Width           =   1710
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descrip. Resumida (Botonera) :"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1095
         Width           =   2205
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   1740
         TabIndex        =   18
         Top             =   405
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10035
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2550
      Width           =   10095
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
         Picture         =   "frmTipoCancelacionDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmTipoCancelacionDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmTipoCancelacionDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1170
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
         Picture         =   "frmTipoCancelacionDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5190
         TabIndex        =   10
         Top             =   60
         Width           =   5250
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTipoCancelacionDetalle.frx":109A
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
            Picture         =   "frmTipoCancelacionDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTipoCancelacionDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4710
            Picture         =   "frmTipoCancelacionDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4230
            Picture         =   "frmTipoCancelacionDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3750
            Picture         =   "frmTipoCancelacionDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   11
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
            TabIndex        =   19
            Top             =   180
            Width           =   2265
         End
      End
   End
   Begin VB.Image Image 
      Height          =   2550
      Left            =   45
      Picture         =   "frmTipoCancelacionDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2100
   End
End
Attribute VB_Name = "frmTipoCancelacionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub Asignar()
    With frmTipoCancelacion.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        txtCuentaContable = IIf(IsNull(!tvalor), "", !tvalor)
        ChkCreditoCorporativo.Value = IIf(IsNull(!nValor), 0, IIf(!nValor = "1", 1, 0))
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTipoCancelacion.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTipoCancelacion.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTipoCancelacion.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTipoCancelacion.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTipoCancelacion.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTipoCancelacion.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmTipoCancelacion.RsCabecera.RecordCount = 0, 0, frmTipoCancelacion.RsCabecera.AbsolutePosition) & " de " & frmTipoCancelacion.RsCabecera.RecordCount
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
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='TipoCancelacion' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 3)
                  End If
                  Sw = False
                   
                   
                   
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "TIPO CANCELACION", "01", sUsuario, sPasa, "", _
                   "TTABLA", "TABLA", "TIPOCANCELACION", "tCodigo", "Codigo Tipo", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumido", txtResumido.Text, _
                   "tValor", "Cuenta Contable", txtCuentaContable.Text, "lActivo", "Flag Activo", IIf(chkActivo = 1, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
               
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, nValor, tValor, tIcono, lActivo) " & _
                         "values ('TIPOCANCELACION', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', '" & IIf(ChkCreditoCorporativo.Value, "1", "0") & "','" & txtCuentaContable.Text & "',''," & _
                                       chkActivo.Value & ") "
                  Cn.Execute Isql
                  
                  
                                    
                 sPasa = txtCodigo.Text
                  

                  
                  frmTipoCancelacion.RsCabecera.Sort = "Codigo ASC"
                  frmTipoCancelacion.RsCabecera.Requery
                  frmTipoCancelacion.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmTipoCancelacion.RsCabecera.RecordCount = 0, 0, frmTipoCancelacion.RsCabecera.AbsolutePosition) & " de " & frmTipoCancelacion.RsCabecera.RecordCount
               Else
                                   
                 sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "TIPO CANCELACION", "02", sUsuario, sPasa, "", _
                   "TTABLA", "TABLA", "TIPOCANCELACION", "tCodigo", "Codigo Tipo", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumido", txtResumido.Text, _
                   "tValor", "Cuenta Contable", txtCuentaContable.Text, "lActivo", "Flag Activo", IIf(chkActivo = 1, "Verdadero", "Falso"))
                
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                
               
               
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "nValor ='" & IIf(ChkCreditoCorporativo.Value, "1", "0") & "', " & _
                         "tValor ='" & txtCuentaContable.Text & "', " & _
                         "lActivo =" & chkActivo.Value & _
                         ",lreplica=1 where tTabla = 'TIPOCANCELACION' and tCodigo = '" & txtCodigo & "'"
                   Cn.Execute Isql
                   nPos = frmTipoCancelacion.RsCabecera.Bookmark
                   frmTipoCancelacion.RsCabecera.Requery
                   If frmTipoCancelacion.RsCabecera.RecordCount = 0 Then
                      frmTipoCancelacion.RsCabecera.Filter = adFilterNone
                   End If
                   frmTipoCancelacion.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
            Cn.Execute "update TTABLA set lNuevoModificado=1 where tTabla = 'TIPOCANCELACION' and tCodigo = '" & txtCodigo & "'"
          Case Is = 2 ' Eliminar
               If frmTipoCancelacion.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Tipo de Cancelación " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               
                                   
                 sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "TIPO CANCELACION", "03", sUsuario, sPasa, "", _
                   "TTABLA", "TABLA", "TIPOCANCELACION", "tCodigo", "Codigo Tipo", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text)
                   
                   
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                   
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                
                
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'TIPOCANCELACION' and tCodigo = '" & txtCodigo & "'"
               frmTipoCancelacion.RsCabecera.Requery
               If frmTipoCancelacion.RsCabecera.RecordCount <> 0 Then
                  frmTipoCancelacion.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmTipoCancelacion.RsCabecera.RecordCount = 0, 0, frmTipoCancelacion.RsCabecera.AbsolutePosition) & " de " & frmTipoCancelacion.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Otros tipos de Cancelación"
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
    
    cmdTexto.Caption = "Registro " & IIf(frmTipoCancelacion.RsCabecera.RecordCount = 0, 0, frmTipoCancelacion.RsCabecera.AbsolutePosition) & " de " & frmTipoCancelacion.RsCabecera.RecordCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmTipoCancelacionDetalle = Nothing
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


Private Sub Text1_Change()

End Sub

