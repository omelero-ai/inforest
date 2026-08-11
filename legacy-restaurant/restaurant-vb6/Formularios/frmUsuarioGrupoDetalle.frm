VERSION 5.00
Begin VB.Form frmUsuarioGrupoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7335
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9930
   Icon            =   "frmUsuarioGrupoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9930
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
      Height          =   6525
      Left            =   2640
      TabIndex        =   31
      Top             =   30
      Width           =   7200
      Begin VB.Frame Frame2 
         Caption         =   " Módulos  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1530
         TabIndex        =   36
         Top             =   1000
         Width           =   5550
         Begin VB.CheckBox chkModulo 
            Caption         =   "Punto de Venta"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   2
            Top             =   270
            Width           =   1965
         End
         Begin VB.CheckBox chkModulo 
            Caption         =   "Administración y Configuración"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   3
            Top             =   525
            Width           =   3315
         End
         Begin VB.CheckBox chkModulo 
            Caption         =   "Consultas y Reportes"
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   4
            Top             =   780
            Width           =   2280
         End
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   855
         TabIndex        =   18
         Top             =   6210
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   " Autorización "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         Left            =   1530
         TabIndex        =   35
         Top             =   2115
         Width           =   5550
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Facturación o anulación de Cuentas Corrientes"
            Height          =   195
            Index           =   13
            Left            =   270
            TabIndex        =   17
            Top             =   3690
            Width           =   4770
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Reimpresiones de Documentos"
            Height          =   195
            Index           =   12
            Left            =   270
            TabIndex        =   16
            Top             =   3407
            Width           =   4095
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Cierre de Turnos"
            Height          =   195
            Index           =   11
            Left            =   270
            TabIndex        =   15
            Top             =   3130
            Width           =   4095
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Descuentos y Recargos en Pedidos, Documentos y Cancelaciones"
            Height          =   195
            Index           =   10
            Left            =   270
            TabIndex        =   13
            Top             =   2576
            Width           =   5115
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Cuentas Corrientes"
            Height          =   195
            Index           =   9
            Left            =   270
            TabIndex        =   12
            Top             =   2299
            Width           =   3990
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Cuentas por Cobrar"
            Height          =   195
            Index           =   8
            Left            =   270
            TabIndex        =   11
            Top             =   2022
            Width           =   1875
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Anulación de recibos"
            Height          =   195
            Index           =   7
            Left            =   270
            TabIndex        =   10
            Top             =   1745
            Width           =   1875
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Anulación de Documentos y Recibos de turnos anteriores"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   6
            Left            =   270
            TabIndex        =   14
            Top             =   2853
            Width           =   4485
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Anulación de documentos del turno"
            Height          =   195
            Index           =   5
            Left            =   270
            TabIndex        =   9
            Top             =   1468
            Width           =   2910
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Emisión de Cortesías"
            Height          =   195
            Index           =   4
            Left            =   270
            TabIndex        =   8
            Top             =   1191
            Width           =   1875
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Modificar precios"
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   7
            Top             =   914
            Width           =   1605
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Eliminación items"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   6
            Top             =   637
            Width           =   1605
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Eliminación del pedido"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   5
            Top             =   360
            Width           =   1965
         End
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
         Left            =   1530
         MaxLength       =   15
         TabIndex        =   1
         Text            =   " "
         Top             =   671
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
         Left            =   1545
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   330
         Width           =   1170
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   33
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   855
         TabIndex        =   32
         Top             =   375
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9870
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6585
      Width           =   9930
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4950
         TabIndex        =   24
         Top             =   60
         Width           =   5010
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3510
            Picture         =   "frmUsuarioGrupoDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3990
            Picture         =   "frmUsuarioGrupoDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4470
            Picture         =   "frmUsuarioGrupoDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmUsuarioGrupoDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmUsuarioGrupoDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmUsuarioGrupoDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   34
            Top             =   120
            Width           =   1980
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
         Left            =   6330
         Picture         =   "frmUsuarioGrupoDetalle.frx":23CE
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   5160
         Picture         =   "frmUsuarioGrupoDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   7500
         Picture         =   "frmUsuarioGrupoDetalle.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Left            =   8670
         Picture         =   "frmUsuarioGrupoDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   6540
      Left            =   15
      Picture         =   "frmUsuarioGrupoDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2595
   End
End
Attribute VB_Name = "frmUsuarioGrupoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Asignar()
    With frmUsuarioGrupo.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!tGrupoUsuario), "", !tGrupoUsuario)
        txtDetallado = IIf(IsNull(!tDetallado), "", !tDetallado)
        
        'Check Box
        chkModulo(1) = IIf(!lModulo01 = True, 1, 0) 'Punto de Venta
        chkModulo(2) = IIf(!lModulo02 = True, 1, 0) 'Administracion
        chkModulo(3) = IIf(!lModulo03 = True, 1, 0) 'Reportes
        
        chkOpcion(1) = IIf(!lOpcion01 = True, 1, 0) 'Eliminacion Pedido
        chkOpcion(2) = IIf(!lOpcion02 = True, 1, 0) 'Eliminacion Item
        chkOpcion(3) = IIf(!lOpcion03 = True, 1, 0) 'Modifica Precio
        chkOpcion(4) = IIf(!lOpcion04 = True, 1, 0) 'Emision de cortesias
        chkOpcion(5) = IIf(!lOpcion05 = True, 1, 0) 'Anulacion documentos
        chkOpcion(6) = IIf(!lOpcion06 = True, 1, 0) 'Anulacion doc. otro turno
        chkOpcion(7) = IIf(!lOpcion07 = True, 1, 0) 'Anulacion Recibos
        chkOpcion(8) = IIf(!lOpcion08 = True, 1, 0) 'Cuenta por Cobrar
        chkOpcion(9) = IIf(!lOpcion09 = True, 1, 0) 'Cuenta Corriente
        chkOpcion(10) = IIf(!lOpcion10 = True, 1, 0) 'Desc/Recar Documentos
        chkOpcion(11) = IIf(!lOpcion11 = True, 1, 0) 'Obliga Cierre
        chkOpcion(12) = IIf(!lOpcion12 = True, 1, 0) 'Reimpresiones
        chkOpcion(13) = IIf(!lOpcion13 = True, 1, 0) 'Anulación Cta.Cte.
        
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
End Sub


Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmUsuarioGrupo.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmUsuarioGrupo.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmUsuarioGrupo.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmUsuarioGrupo.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmUsuarioGrupo.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmUsuarioGrupo.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmUsuarioGrupo.RsCabecera.RecordCount = 0, 0, frmUsuarioGrupo.RsCabecera.AbsolutePosition) & " de " & frmUsuarioGrupo.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               Iniciar
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
                    
               If Sw Then
                  'Asignar El Campo de Codificación
                   nCorrela = Calcular("select max(tGrupoUsuario) as Codigo from TGrupoUsuario", Cn)
                   If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                   Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                   End If
                   Sw = False
                   
                   'Cambiar el SQL
                   Isql = "insert into TGrupoUsuario( " & _
                          "tGrupoUsuario, tDetallado, lModulo01, lModulo02, lModulo03, lOpcion01, lOpcion02, lOpcion03, lOpcion04, lOpcion05, lOpcion06, lOpcion07, lOpcion08, lOpcion09, lOpcion10, lOpcion11, lOpcion12, lOpcion13, lActivo) " & _
                          "values ('" & txtCodigo.Text & "', " & _
                                 " '" & txtDetallado.Text & "', " & _
                                        chkModulo(1).Value & ", " & _
                                        chkModulo(2).Value & ", " & _
                                        chkModulo(3).Value & ", " & _
                                        chkOpcion(1).Value & ", " & _
                                        chkOpcion(2).Value & ", " & _
                                        chkOpcion(3).Value & ", " & _
                                        chkOpcion(4).Value & ", " & _
                                        chkOpcion(5).Value & ", " & _
                                        chkOpcion(6).Value & ", " & _
                                        chkOpcion(7).Value & ", " & _
                                        chkOpcion(8).Value & ", " & _
                                        chkOpcion(9).Value & ", " & _
                                        chkOpcion(10).Value & ", " & _
                                        chkOpcion(11).Value & ", " & _
                                        chkOpcion(12).Value & ", " & _
                                        chkOpcion(13).Value & ", " & _
                                        chkActivo.Value & ")"
            
                       Cn.Execute Isql
                       frmUsuarioGrupo.RsCabecera.Sort = "tGrupoUsuario ASC"
                       frmUsuarioGrupo.RsCabecera.Requery
                       frmUsuarioGrupo.RsCabecera.MoveLast
                       MsgBox "Registro Guardado", vbInformation, sMensaje
                       ActivarBotones (True)
                       cmdTexto.Caption = "Registro " & IIf(frmUsuarioGrupo.RsCabecera.RecordCount = 0, 0, frmUsuarioGrupo.RsCabecera.AbsolutePosition) & " de " & frmUsuarioGrupo.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Isql = "update TGrupoUsuario set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "lModulo01 =" & chkModulo(1).Value & ", " & _
                         "lModulo02 =" & chkModulo(2).Value & ", " & _
                         "lModulo03 =" & chkModulo(3).Value & ", " & _
                         "lOpcion01 =" & chkOpcion(1).Value & ", " & _
                         "lOpcion02 =" & chkOpcion(2).Value & ", " & _
                         "lOpcion03 =" & chkOpcion(3).Value & ", " & _
                         "lOpcion04 =" & chkOpcion(4).Value & ", " & _
                         "lOpcion05 =" & chkOpcion(5).Value & ", " & _
                         "lOpcion06 =" & chkOpcion(6).Value & ", " & _
                         "lOpcion07 =" & chkOpcion(7).Value & ", " & _
                         "lOpcion08 =" & chkOpcion(8).Value & ", " & _
                         "lOpcion09 =" & chkOpcion(9).Value & ", " & _
                         "lOpcion10 =" & chkOpcion(10).Value & ", " & _
                         "lOpcion11 =" & chkOpcion(11).Value & ", " & _
                         "lOpcion12 =" & chkOpcion(12).Value & ", " & _
                         "lOpcion13 =" & chkOpcion(13).Value & ", " & _
                         "lActivo =" & chkActivo.Value & " " & _
                         " where tGrupoUsuario = '" & txtCodigo & "'"
                   Cn.Execute Isql
                   nPos = frmUsuarioGrupo.RsCabecera.Bookmark
                   frmUsuarioGrupo.RsCabecera.Requery
                   If frmUsuarioGrupo.RsCabecera.RecordCount = 0 Then
                      frmUsuarioGrupo.RsCabecera.Filter = adFilterNone
                   End If
                   frmUsuarioGrupo.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
               If frmUsuarioGrupo.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el GrupoUsuario " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TGRUPOUSUARIO where tGrupoUsuario = '" & txtCodigo & "'"
               frmUsuarioGrupo.RsCabecera.Requery
               If frmUsuarioGrupo.RsCabecera.RecordCount <> 0 Then
                  frmUsuarioGrupo.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmUsuarioGrupo.RsCabecera.RecordCount = 0, 0, frmUsuarioGrupo.RsCabecera.AbsolutePosition) & " de " & frmUsuarioGrupo.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Grupo de Usuarios "
    fraDetalle.Caption = Me.Caption
              
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       Iniciar
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    cmdTexto.Caption = "Registro " & IIf(frmUsuarioGrupo.RsCabecera.RecordCount = 0, 0, frmUsuarioGrupo.RsCabecera.AbsolutePosition) & " de " & frmUsuarioGrupo.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmUsuarioGrupoDetalle = Nothing
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

Public Sub Iniciar()
   chkModulo(1).Value = 0
   chkModulo(2).Value = 0
   chkModulo(3).Value = 0
   chkOpcion(1).Value = 0
   chkOpcion(2).Value = 0
   chkOpcion(3).Value = 0
   chkOpcion(4).Value = 0
   chkOpcion(5).Value = 0
   chkOpcion(6).Value = 0
   chkOpcion(7).Value = 0
   chkOpcion(8).Value = 0
   chkOpcion(9).Value = 0
   chkOpcion(10).Value = 0
   chkOpcion(11).Value = 0
   chkOpcion(12).Value = 0
   chkOpcion(13).Value = 0
   chkActivo.Value = 1
End Sub
