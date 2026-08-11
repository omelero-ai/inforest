VERSION 5.00
Begin VB.Form frmMotivoDescuentoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5955
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   10125
   Icon            =   "frmMotivoDescuentoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10125
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
      Height          =   5205
      Left            =   2175
      TabIndex        =   23
      Top             =   0
      Width           =   7935
      Begin VB.CheckBox chkAnticipoDes 
         Alignment       =   1  'Right Justify
         Caption         =   "Aplica Anticipo :"
         Height          =   195
         Left            =   1560
         TabIndex        =   34
         Top             =   4320
         Width           =   1440
      End
      Begin VB.CheckBox chkAplicaPedido 
         Alignment       =   1  'Right Justify
         Caption         =   "Descuento aplicable solo al Pedido :"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   4080
         Width           =   2880
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   4680
         TabIndex        =   31
         Top             =   2820
         Width           =   2715
         Begin VB.OptionButton optTopeP 
            Caption         =   "Por Pedido"
            Height          =   255
            Left            =   180
            TabIndex        =   7
            Top             =   300
            Width           =   2415
         End
         Begin VB.OptionButton optTopeM 
            Caption         =   "Acumulado Mensual"
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   660
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   4680
         TabIndex        =   30
         Top             =   1500
         Width           =   2715
         Begin VB.OptionButton optDescuentoM 
            Caption         =   "Monto"
            Height          =   195
            Left            =   180
            TabIndex        =   5
            Top             =   660
            Width           =   2295
         End
         Begin VB.OptionButton optDescuentoR 
            Caption         =   "Ratio Porcentual"
            Height          =   195
            Left            =   180
            TabIndex        =   4
            Top             =   300
            Width           =   2295
         End
      End
      Begin VB.TextBox txtDescuento 
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
         Height          =   285
         Left            =   2805
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1560
         Width           =   1740
      End
      Begin VB.TextBox txtTope 
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
         Height          =   285
         Left            =   2805
         MaxLength       =   15
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   2880
         Width           =   1740
      End
      Begin VB.CheckBox chkBloqueo 
         Alignment       =   1  'Right Justify
         Caption         =   "Permite Cambiar El Ratio o Monto :"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   3840
         Width           =   2760
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
         Left            =   2805
         MaxLength       =   24
         TabIndex        =   2
         Text            =   " "
         Top             =   1140
         Width           =   3135
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
         Left            =   2805
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   735
         Width           =   4980
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
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   330
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   2145
         TabIndex        =   10
         Top             =   4800
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "(No aplicable para Cuentas Corrientes)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   570
         TabIndex        =   33
         Top             =   3300
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descuento :"
         Height          =   195
         Left            =   1815
         TabIndex        =   29
         Top             =   1560
         Width           =   870
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descuento máximo permitido :"
         Height          =   195
         Left            =   570
         TabIndex        =   28
         Top             =   2925
         Width           =   2115
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Left            =   1005
         TabIndex        =   27
         Top             =   1185
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Index           =   1
         Left            =   1035
         TabIndex        =   25
         Top             =   780
         Width           =   1650
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   2100
         TabIndex        =   24
         Top             =   375
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10065
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5205
      Width           =   10125
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
         Picture         =   "frmMotivoDescuentoDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "frmMotivoDescuentoDetalle.frx":0534
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
         Left            =   5310
         Picture         =   "frmMotivoDescuentoDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmMotivoDescuentoDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5190
         TabIndex        =   22
         Top             =   60
         Width           =   5250
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmMotivoDescuentoDetalle.frx":109A
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
            Picture         =   "frmMotivoDescuentoDetalle.frx":15DC
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
            Picture         =   "frmMotivoDescuentoDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4710
            Picture         =   "frmMotivoDescuentoDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4230
            Picture         =   "frmMotivoDescuentoDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3750
            Picture         =   "frmMotivoDescuentoDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   18
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
            TabIndex        =   26
            Top             =   180
            Width           =   2265
         End
      End
   End
   Begin VB.Image Image 
      Height          =   5205
      Left            =   45
      Picture         =   "frmMotivoDescuentoDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2100
   End
End
Attribute VB_Name = "frmMotivoDescuentoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nRatio As Double
Sub Asignar()
    With frmMotivoDescuento.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        
        txtDescuento = Format(IIf(IsNull(!nRatio), 0, !nRatio), "##0.00")
        optDescuentoR.value = IIf(!lRatio, 1, 0)
        optDescuentoM.value = IIf(!lRatio, 0, 1)
        
        txtTope = Format(IIf(IsNull(!nTope), 0, !nTope), "###,##0.00")
        optTopeP.value = IIf(!lTopePedido, 1, 0)
        optTopeM.value = IIf(!lTopePedido, 0, 1)
        
        'Check Box
        chkBloqueo.value = IIf(!lBloqueo, 1, 0)
        chkAplicaPedido.value = IIf(!lAplicablePedido, 1, 0)
        chkActivo.value = IIf(!lActivo, 1, 0)
        Me.chkAnticipoDes.value = IIf(!AplicaAnticipo, 1, 0)
     End With
End Sub



Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmMotivoDescuento.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmMotivoDescuento.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmMotivoDescuento.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmMotivoDescuento.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmMotivoDescuento.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmMotivoDescuento.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmMotivoDescuento.RsCabecera.RecordCount = 0, 0, frmMotivoDescuento.RsCabecera.AbsolutePosition) & " de " & frmMotivoDescuento.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               txtDescuento.Text = "0.00"
               txtTope.Text = "0.00"
               chkActivo.value = 1
               optDescuentoR.value = True
               optTopeP.value = True
                                       'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
                    
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               
               If txtResumido.Text = "" Then MsgBox "Ingrese el Prefijo", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                    
               If optDescuentoR.value = True And val(Me.txtDescuento.Text) >= 100 Then MsgBox "El Ratio del descuento no puede ser mayor ó Igual a 100%": txtDescuento.SetFocus: Exit Sub
                    
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tDescuento) as Codigo from TMOTIVODESCUENTO", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 3)
                  End If
                  Sw = False
                    
                  sPasa = txtCodigo.Text
                  
                   'Inserta Movimiento auditoria
                   lAuditoria = RegistraMovimientoAuditoria(tMolduloSeg, sMDB, "TMOTIVODESCUENTO", "MOTIVO DESCUENTO", "01", sUsuario, sPasa, "", _
                   "TDESCUENTO", "CODIGO DESCUENTO", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, _
                   "nRatio", "Descuento", val(txtDescuento.Text), "lRatio", "Opcion Ratio o Monto", IIf(optDescuentoR.value, "Verdadero", "Falso"), _
                   "nTope", "Tope Máximo", val(txtTope.Text), "lTopePedido", "Opcion Por Pedido o Mensual", IIf(optTopeP.value, "Verdadero", "Falso"), _
                   "lBloqueo", "Flag Modificar Ratio o Monto", IIf(chkBloqueo.value, "Verdadero", "Falso"), "lAplicablePedido", "Flag Aplicable Solo Pedido", IIf(chkAplicaPedido.value, "Verdadero", "Falso"), "lActivo", "Flag Activo", IIf(chkActivo.value, "Verdadero", "Falso"), "lAplicaAnt", "Flag Anticipo", IIf(Me.chkAnticipoDes.value, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                 'Cambiar el SQL
                 Isql = "insert into tMotivoDescuento( " & _
                         " tDescuento, tDetallado, tResumido, nRatio, nTope, lRatio, lTopePedido, lBloqueo, lActivo, lAplicablepedido,lAplicaAnt ) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                CDbl(txtDescuento.Text) & ", " & _
                                CDbl(txtTope.Text) & ", " & _
                                IIf(optDescuentoR.value, 1, 0) & ", " & _
                                IIf(optTopeP.value, 1, 0) & ", " & _
                                chkBloqueo.value & " ," & _
                                chkActivo.value & ", " & chkAplicaPedido.value & "," & Me.chkAnticipoDes.value & ") "
                  Cn.Execute Isql
                
                  frmMotivoDescuento.RsCabecera.Sort = "Codigo ASC"
                  frmMotivoDescuento.RsCabecera.Requery
                  frmMotivoDescuento.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmMotivoDescuento.RsCabecera.RecordCount = 0, 0, frmMotivoDescuento.RsCabecera.AbsolutePosition) & " de " & frmMotivoDescuento.RsCabecera.RecordCount
               Else
               
               sPasa = txtCodigo.Text
                  'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tMolduloSeg, sMDB, "TMOTIVODESCUENTO", "MOTIVO DESCUENTO", "02", sUsuario, sPasa, "", _
                   "TDESCUENTO", "CODIGO DESCUENTO", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, _
                   "nRatio", "Descuento", val(txtDescuento.Text), "lRatio", "Opcion Ratio o Monto", IIf(optDescuentoR.value, "Verdadero", "Falso"), _
                   "nTope", "Tope Máximo", val(txtTope.Text), "lTopePedido", "Opcion Por Pedido o Mensual", IIf(optTopeP.value, "Verdadero", "Falso"), _
                   "lBloqueo", "Flag Modificar Ratio o Monto", IIf(chkBloqueo.value, "Verdadero", "Falso"), "lAplicablePedido", "Flag Aplicable Solo Pedido", IIf(chkAplicaPedido.value, "Verdadero", "Falso"), "lActivo", "Flag Activo", IIf(chkActivo.value, "Verdadero", "Falso"), "lAplicaAnt", "Flag Anticipo", IIf(Me.chkAnticipoDes.value, "Verdadero", "Falso"))
                   
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
               
                  'Cambiar el SQL
                  Isql = "update TMOTIVODESCUENTO set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "nRatio =" & CDbl(txtDescuento.Text) & ", " & _
                         "nTope =" & CDbl(txtTope.Text) & ", " & _
                         "lRatio =" & IIf(optDescuentoR.value, 1, 0) & ", " & _
                         "lTopePedido =" & IIf(optTopeP.value, 1, 0) & ", " & _
                         "lBloqueo =" & chkBloqueo.value & ", " & _
                         "lAplicablePedido =" & chkAplicaPedido.value & ", " & _
                         "lActivo =" & chkActivo.value & ",lAplicaAnt=" & Me.chkAnticipoDes.value & ", lReplica=1 " & _
                         " where tDescuento = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                   nPos = frmMotivoDescuento.RsCabecera.Bookmark
                   frmMotivoDescuento.RsCabecera.Requery
                   If frmMotivoDescuento.RsCabecera.RecordCount = 0 Then
                      frmMotivoDescuento.RsCabecera.Filter = adFilterNone
                   End If
                   frmMotivoDescuento.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               
               Cn.Execute "update TMOTIVODESCUENTO set lNuevoModificado=1 where tDescuento = '" & txtCodigo & "'"
            
          Case Is = 2 ' Eliminar
               If frmMotivoDescuento.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Motivo Descuento " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
                 
               sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tMolduloSeg, sMDB, "TMOTIVODESCUENTO", "MOTIVO DESCUENTO", "03", sUsuario, sPasa, "", _
                   "TDESCUENTO", "CODIGO DESCUENTO", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text)
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
               
               
               
               'Cambia el Delete
               Cn.Execute "delete from TMOTIVODESCUENTO where tDescuento = '" & txtCodigo & "'"
               frmMotivoDescuento.RsCabecera.Requery
               If frmMotivoDescuento.RsCabecera.RecordCount <> 0 Then
                  frmMotivoDescuento.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmMotivoDescuento.RsCabecera.RecordCount = 0, 0, frmMotivoDescuento.RsCabecera.AbsolutePosition) & " de " & frmMotivoDescuento.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Motivos de Descuento "
    fraDetalle.Caption = Me.Caption
    'Me.Label4.Caption = "( " & sMonedaN & " )"
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       txtDescuento.Text = "0.00"
       txtTope.Text = "0.00"
       optDescuentoR.value = True
       optTopeP.value = True
       chkActivo.value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmMotivoDescuento.RsCabecera.RecordCount = 0, 0, frmMotivoDescuento.RsCabecera.AbsolutePosition) & " de " & frmMotivoDescuento.RsCabecera.RecordCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmMotivoDescuentoDetalle = Nothing
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

Private Sub txtDescuento_GotFocus()
   txtDescuento.SelLength = Len(Trim(txtDescuento.Text))
End Sub

Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtDescuento
End Sub

Private Sub txtDescuento_LostFocus()
   txtDescuento.Text = Format(txtDescuento.Text, "##0.00")
    If val(txtDescuento.Text) < 0 Then
        MsgBox "Error en Descuento", vbInformation, sMensaje
        txtDescuento.SetFocus
        Exit Sub
    End If
    
    If optDescuentoR.value = True And val(Me.txtDescuento.Text) >= 100 Then
        MsgBox "Error en ratio de Descuento", vbInformation, sMensaje
        'txtDescuento.SetFocus
        Exit Sub
    End If
   
End Sub

Private Sub txtDetallado_LostFocus()
   Call ValidaStr(txtDetallado)
End Sub

Private Sub txtResumido_LostFocus()
   Call ValidaStr(txtResumido)
End Sub

Private Sub TxtTope_GotFocus()
   txtTope.SelLength = Len(Trim(txtTope.Text))
End Sub

Private Sub TxtTope_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtTope
End Sub

Private Sub TxtTope_LostFocus()
   txtTope.Text = Format(txtTope.Text, "###,##0.00")
End Sub




