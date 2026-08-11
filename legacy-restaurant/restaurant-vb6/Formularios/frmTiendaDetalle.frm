VERSION 5.00
Begin VB.Form frmTiendaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4155
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9390
   Icon            =   "frmTiendaDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
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
      Height          =   3405
      Left            =   2040
      TabIndex        =   21
      Top             =   -30
      Width           =   7290
      Begin VB.TextBox txtContacto 
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
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   7
         Text            =   " "
         Top             =   2718
         Width           =   5400
      End
      Begin VB.TextBox txtTelefono 
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
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   5
         Text            =   " "
         Top             =   2034
         Width           =   2595
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   6
         Text            =   " "
         Top             =   2376
         Width           =   5400
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1350
         Width           =   5400
      End
      Begin VB.TextBox txtDireccion 
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
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   4
         Text            =   " "
         Top             =   1692
         Width           =   5400
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
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   930
         TabIndex        =   8
         Top             =   3060
         Width           =   840
      End
      Begin VB.Label txtSocial 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1590
         TabIndex        =   1
         Top             =   630
         Width           =   5400
      End
      Begin VB.Label txtComercial 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1590
         TabIndex        =   0
         Top             =   300
         Width           =   5400
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Contacto :"
         Height          =   195
         Index           =   7
         Left            =   750
         TabIndex        =   30
         Top             =   2763
         Width           =   735
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "E-mail :"
         Height          =   195
         Index           =   6
         Left            =   945
         TabIndex        =   29
         Top             =   2421
         Width           =   510
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono :"
         Height          =   195
         Index           =   5
         Left            =   735
         TabIndex        =   28
         Top             =   2079
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social :"
         Height          =   195
         Index           =   4
         Left            =   420
         TabIndex        =   27
         Top             =   675
         Width           =   1035
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Razón Comercial :"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   26
         Top             =   345
         Width           =   1290
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Index           =   2
         Left            =   810
         TabIndex        =   24
         Top             =   1395
         Width           =   645
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         Height          =   195
         Index           =   3
         Left            =   690
         TabIndex        =   23
         Top             =   1737
         Width           =   765
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   1
         Left            =   870
         TabIndex        =   22
         Top             =   1005
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9330
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3405
      Width           =   9390
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4530
         TabIndex        =   14
         Top             =   60
         Width           =   4590
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3090
            Picture         =   "frmTiendaDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3570
            Picture         =   "frmTiendaDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4050
            Picture         =   "frmTiendaDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTiendaDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTiendaDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTiendaDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   15
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
            TabIndex        =   25
            Top             =   180
            Width           =   1545
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
         Left            =   5820
         Picture         =   "frmTiendaDetalle.frx":23CE
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
         Left            =   4650
         Picture         =   "frmTiendaDetalle.frx":2900
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
         Left            =   6990
         Picture         =   "frmTiendaDetalle.frx":2E32
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
         Left            =   8160
         Picture         =   "frmTiendaDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   3345
      Left            =   15
      Picture         =   "frmTiendaDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   15
      Width           =   1980
   End
End
Attribute VB_Name = "frmTiendaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Asignar()
    With frmTienda.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!tCodigoTienda), "", !tCodigoTienda)
        txtNombre = IIf(IsNull(!tNombre), "", !tNombre)
        txtDireccion = IIf(IsNull(!tDireccion), "", !tDireccion)
        txtTelefono = IIf(IsNull(!tTelefono), "", !tTelefono)
        txtEmail = IIf(IsNull(!temail), "", !temail)
        txtContacto = IIf(IsNull(!tContacto), "", !tContacto)
    
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTienda.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTienda.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTienda.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTienda.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTienda.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTienda.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmTienda.RsCabecera.RecordCount = 0, 0, frmTienda.RsCabecera.AbsolutePosition) & " de " & frmTienda.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               
               'Cambia el Nombre del Primer Text
               txtNombre.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               'Chequea Datos
               If txtNombre.Text = "" Then MsgBox "Ingrese el Nombre de la Tienda", vbExclamation, sMensaje: txtNombre.SetFocus: Exit Sub
                       
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigoTienda) as Codigo from TTIENDA where tCodigoDelivery ='" & sCodigo & "'", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 3)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTIENDA( " & _
                         "tCodigoDelivery, tCodigoTienda, tNombre, tDireccion, tTelefono, tEmail, tContacto, lActivo) " & _
                         "values ('" & sCodigo & "', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtNombre.Text & "', " & _
                                " '" & txtDireccion.Text & "', " & _
                                " '" & txtTelefono.Text & "', " & _
                                " '" & txtEmail.Text & "', " & _
                                " '" & txtContacto.Text & "', " & _
                                       chkActivo.value & ") "
           
                  Cn.Execute Isql
                  frmTienda.RsCabecera.Requery
                  frmTienda.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmTienda.RsCabecera.RecordCount = 0, 0, frmTienda.RsCabecera.AbsolutePosition) & " de " & frmTienda.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Isql = "update TTIENDA set " & _
                         "tNombre ='" & txtNombre.Text & "', " & _
                         "tDireccion ='" & txtDireccion.Text & "', " & _
                         "tTelefono ='" & txtTelefono.Text & "', " & _
                         "tEmail ='" & txtEmail.Text & "', " & _
                         "tContacto ='" & txtContacto.Text & "', " & _
                         "lActivo =" & chkActivo.value & _
                         " where tCodigoDelivery = '" & sCodigo & "' and tCodigoTienda = '" & txtCodigo.Text & "'"
                       
                   Cn.Execute Isql
                   nPos = frmTienda.RsCabecera.Bookmark
                   frmTienda.RsCabecera.Requery
                   If frmTienda.RsCabecera.RecordCount = 0 Then
                      frmTienda.RsCabecera.Filter = adFilterNone
                   End If
                   frmTienda.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
                If frmTienda.RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                'Cambia el MsgBox
                If MsgBox("Seguro de Eliminar la Tienda?" & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                   Exit Sub
                End If
                
                'Cambia el Delete
                Cn.Execute "delete from TTIENDA where tCodigoDelivery = '" & sCodigo & "' and tCodigoTienda = '" & txtCodigo.Text & "'"
                frmTienda.RsCabecera.Requery
                If frmTienda.RsCabecera.RecordCount <> 0 Then
                   frmTienda.RsCabecera.MoveLast
                   Asignar
                   cmdTexto.Caption = "Registro " & IIf(frmTienda.RsCabecera.RecordCount = 0, 0, frmTienda.RsCabecera.AbsolutePosition) & " de " & frmTienda.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Tiendas "
    fraDetalle.Caption = Me.Caption
    txtSocial.Caption = sDescrip
    txtComercial.Caption = sTemp
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
       
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       'frmTienda.RsCabecera.Find ("Codigo = '" & frmTienda.RsCabecera!Codigo & "'")
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmTienda.RsCabecera.RecordCount = 0, 0, frmTienda.RsCabecera.AbsolutePosition) & " de " & frmTienda.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmTiendaDetalle = Nothing
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
Private Sub txtContacto_LostFocus()
  Call ValidaStr(txtContacto)
End Sub
Private Sub txtDireccion_LostFocus()
  Call ValidaStr(txtDireccion)
End Sub
Private Sub txtEmail_LostFocus()
  Call ValidaStr(txtEmail)
End Sub
Private Sub txtNombre_Change()
  Call ValidaStr(txtNombre)
End Sub

Private Sub txtTelefono_LostFocus()
 Call ValidaStr(txtTelefono)
End Sub
