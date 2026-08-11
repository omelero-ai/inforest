VERSION 5.00
Begin VB.Form frmCortesiaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2850
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9480
   Icon            =   "frmCortesiaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   9480
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
      Height          =   2070
      Left            =   1740
      TabIndex        =   17
      Top             =   0
      Width           =   7650
      Begin VB.CheckBox chkAplicaTransferencia 
         Alignment       =   1  'Right Justify
         Caption         =   "Aplicar a Transferencia Gratuita"
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   1695
         Width           =   2655
      End
      Begin VB.TextBox txtCortesiaTope 
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
         Left            =   1515
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1320
         Width           =   1740
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
         Left            =   1515
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   615
         Width           =   5850
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
         Left            =   1515
         MaxLength       =   24
         TabIndex        =   2
         Text            =   " "
         Top             =   960
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
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   1695
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "(Acumulado Por Mes; 0= Sin tope)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3360
         TabIndex        =   23
         Top             =   1350
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tope Permitido :"
         Height          =   195
         Left            =   285
         TabIndex        =   22
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Detallada :"
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   660
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Resumida :"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1005
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   855
         TabIndex        =   18
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9420
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2100
      Width           =   9480
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
         Picture         =   "frmCortesiaDetalle.frx":0442
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
         Left            =   7050
         Picture         =   "frmCortesiaDetalle.frx":0534
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
         Left            =   4710
         Picture         =   "frmCortesiaDetalle.frx":0636
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
         Left            =   5880
         Picture         =   "frmCortesiaDetalle.frx":0B68
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
         ScaleWidth      =   4590
         TabIndex        =   10
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmCortesiaDetalle.frx":109A
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
            Picture         =   "frmCortesiaDetalle.frx":15DC
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
            Picture         =   "frmCortesiaDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmCortesiaDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmCortesiaDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "frmCortesiaDetalle.frx":2AE4
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
            Left            =   1470
            TabIndex        =   21
            Top             =   180
            Width           =   1665
         End
      End
   End
   Begin VB.Image Image 
      Height          =   2025
      Left            =   0
      Picture         =   "frmCortesiaDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1710
   End
End
Attribute VB_Name = "frmCortesiaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim frmCortesia.RsCabecera As Recordset

Sub Asignar()
    With frmCortesia.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        txtCortesiaTope = !Tope
        'Check Box
        chkAplicaTransferencia = IIf(!Transferencia = "1", 1, 0)
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
    
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmCortesia.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmCortesia.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmCortesia.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmCortesia.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmCortesia.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmCortesia.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmCortesia.RsCabecera.AbsolutePosition & " de " & frmCortesia.RsCabecera.RecordCount
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
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(convert(int,tCodigo)) as Codigo from TTABLA where tTabla = 'CORTESIA' ", Cn)
                  'nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla = 'CORTESIA' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "0001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 4)
                  End If
                  Sw = False
                   
                sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "CORTESIA", "01", sUsuario, sPasa, "", _
                   "TTABLA", "TABLA", "CORTESIA", "tCodigo", "Codigo Cortesia", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumido", txtResumido.Text, "lActivo", "Flag Activo", IIf(chkActivo = 1, "Verdadero", "Falso"), "nvalor", "Tope Cortesia", val(txtCortesiaTope.Text))
                If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                                         
                   
                  'Cambiar el SQL
                  
                  'GCAA 02022021 - REQ REST-C043-2021 - Grupo Hersil
                  'tValor va guardar chkAplicaTransferencia
                  
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tValor, tResumido,nvalor, lActivo) " & _
                         "values ('CORTESIA', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " " & chkAplicaTransferencia.Value & ", " & _
                                " '" & txtResumido.Text & "'," & val(txtCortesiaTope.Text) & ", " & _
                                       chkActivo.Value & ") "
            
            
            
            
            
            
                  Cn.Execute Isql
                  
                   
                  frmCortesia.RsCabecera.Sort = "Codigo ASC"
                  frmCortesia.RsCabecera.Requery
                  frmCortesia.RsCabecera.MoveLast
                  ActivarBotones (True)
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  cmdTexto.Caption = "Registro " & IIf(frmCortesia.RsCabecera.RecordCount = 0, 0, frmCortesia.RsCabecera.AbsolutePosition) & " de " & frmCortesia.RsCabecera.RecordCount
               Else
               
               
               sPasa = txtCodigo.Text
               lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "CORTESIA", "02", sUsuario, sPasa, "", _
               "TTABLA", "TABLA", "CORTESIA", "tCodigo", "Codigo Cortesia", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumido", txtResumido.Text, "lActivo", "Flag Activo", IIf(chkActivo = 1, "Verdadero", "Falso"), "nvalor", "Tope Cortesia", val(txtCortesiaTope.Text))
                If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                End If

                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tValor =" & chkAplicaTransferencia.Value & ", " & _
                         "tResumido ='" & txtResumido.Text & "', nvalor=" & val(txtCortesiaTope.Text) & ", " & _
                         "lActivo =" & chkActivo.Value & _
                         ", lReplica=1 where tTabla = 'CORTESIA' and tCodigo = '" & txtCodigo & "'"
                   Cn.Execute Isql
                   nPos = frmCortesia.RsCabecera.Bookmark
                   frmCortesia.RsCabecera.Requery
                   If frmCortesia.RsCabecera.RecordCount = 0 Then
                      frmCortesia.RsCabecera.Filter = adFilterNone
                   End If
                   frmCortesia.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
                Cn.Execute "update TTABLA set lNuevoModificado=1 where tTabla = 'CORTESIA' and tCodigo = '" & txtCodigo & "'"
          Case Is = 2 ' Eliminar
               If frmCortesia.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar la Cortesía" & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                
                
                sPasa = txtCodigo.Text
                'Inserta Movimiento en este el codigo del registro viene del "txtcodigo.text"
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "CORTESIA", "03", sUsuario, txtCodigo.Text, "", "TTABLA", "TABLA", "CORTESIA", "tCodigo", "Código Cortesia", txtCodigo.Text, "tDetallado", "Descripcion Detallada", txtDetallado.Text)
                 If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.

                
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'CORTESIA' and tCodigo = '" & txtCodigo & "'"
               frmCortesia.RsCabecera.Requery
               If frmCortesia.RsCabecera.RecordCount <> 0 Then
                  frmCortesia.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmCortesia.RsCabecera.RecordCount = 0, 0, frmCortesia.RsCabecera.AbsolutePosition) & " de " & frmCortesia.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Cortesia "
    fraDetalle.Caption = Me.Caption
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       'frmCortesia.RsCabecera.Find ("Codigo = '" & frmCortesia.RsCabecera!Codigo & "'")
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & frmCortesia.RsCabecera.AbsolutePosition & " de " & frmCortesia.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmCortesiaDetalle = Nothing
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
Private Sub TxtCortesiaTope_GotFocus()
   txtCortesiaTope.SelLength = Len(Trim(txtCortesiaTope.Text))
End Sub

Private Sub TxtCortesiaTope_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtCortesiaTope
End Sub

Private Sub TxtCortesiaTope_LostFocus()
   txtCortesiaTope.Text = Format(txtCortesiaTope.Text, "########.00")
End Sub

