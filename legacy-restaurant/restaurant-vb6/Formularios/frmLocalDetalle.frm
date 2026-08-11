VERSION 5.00
Begin VB.Form frmLocalDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2865
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9510
   Icon            =   "frmLocalDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   9510
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
      Height          =   2115
      Left            =   1800
      TabIndex        =   17
      Top             =   -30
      Width           =   7665
      Begin VB.TextBox txtEnlaceContable 
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
         Left            =   2100
         MaxLength       =   20
         TabIndex        =   3
         Text            =   " "
         Top             =   1350
         Width           =   5475
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
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   645
         Width           =   5490
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
         Left            =   2100
         MaxLength       =   20
         TabIndex        =   2
         Text            =   " "
         Top             =   990
         Width           =   5475
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
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   300
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1275
         TabIndex        =   4
         Top             =   1815
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enlace Contable :"
         Height          =   195
         Left            =   585
         TabIndex        =   22
         Top             =   1425
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Left            =   165
         TabIndex        =   21
         Top             =   1080
         Width           =   1680
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Left            =   195
         TabIndex        =   19
         Top             =   690
         Width           =   1650
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   1260
         TabIndex        =   18
         Top             =   345
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9450
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2115
      Width           =   9510
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
         Picture         =   "frmLocalDetalle.frx":0442
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
         Left            =   7050
         Picture         =   "frmLocalDetalle.frx":0534
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
         Picture         =   "frmLocalDetalle.frx":0636
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
         Picture         =   "frmLocalDetalle.frx":0B68
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
         TabIndex        =   16
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmLocalDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmLocalDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmLocalDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmLocalDetalle.frx":2060
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
            Picture         =   "frmLocalDetalle.frx":25A2
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
            Picture         =   "frmLocalDetalle.frx":2AE4
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
            Left            =   1500
            TabIndex        =   20
            Top             =   180
            Width           =   1665
         End
      End
   End
   Begin VB.Image Image 
      Height          =   2040
      Left            =   15
      Picture         =   "frmLocalDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1755
   End
End
Attribute VB_Name = "frmLocalDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSector As Recordset


Sub Asignar()
    With frmLocal.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
    
        txtEnlaceContable = IIf(IsNull(!tEnlaceContable), "", !tEnlaceContable)
    
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
        
    End With
End Sub
 
Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmLocal.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmLocal.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmLocal.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmLocal.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmLocal.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmLocal.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmLocal.RsCabecera.RecordCount = 0, 0, frmLocal.RsCabecera.AbsolutePosition) & " de " & frmLocal.RsCabecera.RecordCount
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
                  nCorrela = Calcular("select max(tCodigoLOCAL) as Codigo from TLOCAL ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 3)
                  End If
                  Sw = False
                   
          'Inserta Movimiento auditoria
                              
                    sPasa = txtCodigo.Text
                                       
                   lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TLOCAL", "ESTABLECIMIENTO", "01", sUsuario, sPasa, "", _
                       "TCODIGOLOCAL", "CODIGO LOCAL", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, "tEnlaceContable", "Enlace Contable", Me.txtEnlaceContable.Text, _
                   "lActivo", "Flag Activo", IIf(chkActivo.value, "Verdadero", "Falso"))
                                          
                   
                  'Cambiar el SQL
                  Isql = "INSERT INTO TLOCAL (tcodigolocal,tdetallado,tresumido, tenlaceContable,lactivo) " & _
                           "values ( '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "','" & txtEnlaceContable.Text & "', " & _
                                         chkActivo.value & ") "
                                       

                  Cn.Execute Isql
               '   Cn.Execute Isql
                  frmLocal.RsCabecera.Sort = "Codigo ASC"
                  frmLocal.RsCabecera.Requery
                  frmLocal.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmLocal.RsCabecera.RecordCount = 0, 0, frmLocal.RsCabecera.AbsolutePosition) & " de " & frmLocal.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                 'Inserta Movimiento auditoria
                              
                    sPasa = txtCodigo.Text
                                       
                   lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TLOCAL", "ESTABLECIMIENTO", "02", sUsuario, sPasa, "", _
                       "TCODIGOLOCAL", "CODIGO LOCAL", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, "tEnlaceContable", "Enlace Contable", txtEnlaceContable.Text, _
                   "lActivo", "Flag Activo", IIf(chkActivo.value, "Verdadero", "Falso"))
                                          
                        Isql = "update TLOCAL set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tEnlaceContable='" & txtEnlaceContable.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "lActivo =" & chkActivo.value & ", lReplica=1 " & _
                         " where tCodigoLOCAL = '" & txtCodigo & "'"
                   '  Debug.Print Isql
                  Cn.Execute Isql
   
                  nPos = frmLocal.RsCabecera.Bookmark
                  frmLocal.RsCabecera.Requery
                  If frmLocal.RsCabecera.RecordCount = 0 Then
                     frmLocal.RsCabecera.Filter = adFilterNone
                  End If
                  frmLocal.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          Case Is = 2 ' Eliminar
               If frmLocal.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
                If Calcular("select count(*) as codigo from vsalon where tlocal='" & txtCodigo.Text & "'", Cn) > 0 Then
                    MsgBox "No es posible eliminar Establecimiento por tener salones Asociados"
                    Exit Sub
                End If
                         'Inserta Movimiento auditoria
                              
                    sPasa = txtCodigo.Text
                                       
                   lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TLOCAL", "ESTABLECIMIENTO", "03", sUsuario, sPasa, "", _
                       "TCODIGOLOCAL", "CODIGO LOCAL", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, _
                   "lActivo", "Flag Activo", IIf(chkActivo.value, "Verdadero", "Falso"))
                                          
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Local " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from tlocal where tcodigolocal='" & txtCodigo & "'"
               frmLocal.RsCabecera.Requery
               If frmLocal.RsCabecera.RecordCount <> 0 Then
                  frmLocal.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmLocal.RsCabecera.RecordCount = 0, 0, frmLocal.RsCabecera.AbsolutePosition) & " de " & frmLocal.RsCabecera.RecordCount
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
  
    Me.Caption = " Establecimientos (Locales) "
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
    
    cmdTexto.Caption = "Registro " & IIf(frmLocal.RsCabecera.RecordCount = 0, 0, frmLocal.RsCabecera.AbsolutePosition) & " de " & frmLocal.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmLocalDetalle = Nothing
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

Private Sub txtBd_LostFocus()
 Call ValidaStr(txtResumido)
End Sub

 
Private Sub txtEnlaceContable_LostFocus()
  Call ValidaStr(txtEnlaceContable)
End Sub
