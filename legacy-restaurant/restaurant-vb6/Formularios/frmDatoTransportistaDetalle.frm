VERSION 5.00
Begin VB.Form frmDatoTransportistaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3015
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9480
   Icon            =   "frmDatoTransportistaDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
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
      Height          =   2250
      Left            =   1740
      TabIndex        =   17
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtLicencia 
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
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   3
         Text            =   " "
         Top             =   1479
         Width           =   5550
      End
      Begin VB.TextBox txtMarca 
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
         Left            =   1890
         MaxLength       =   80
         TabIndex        =   1
         Text            =   " "
         Top             =   753
         Width           =   5550
      End
      Begin VB.TextBox txtPlaca 
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
         Left            =   1890
         MaxLength       =   15
         TabIndex        =   2
         Text            =   " "
         Top             =   1116
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
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   390
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1230
         TabIndex        =   4
         Top             =   1845
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Licencia de Conducir :"
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   1485
         Width           =   1590
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Marca del Transporte :"
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   780
         Width           =   1605
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Placa :"
         Height          =   195
         Left            =   1320
         TabIndex        =   19
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   1230
         TabIndex        =   18
         Top             =   435
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
      Top             =   2265
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
         Picture         =   "frmDatoTransportistaDetalle.frx":0442
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
         Picture         =   "frmDatoTransportistaDetalle.frx":0534
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
         Picture         =   "frmDatoTransportistaDetalle.frx":0636
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
         Picture         =   "frmDatoTransportistaDetalle.frx":0B68
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
            Picture         =   "frmDatoTransportistaDetalle.frx":109A
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
            Picture         =   "frmDatoTransportistaDetalle.frx":15DC
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
            Picture         =   "frmDatoTransportistaDetalle.frx":1B1E
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
            Picture         =   "frmDatoTransportistaDetalle.frx":2060
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
            Picture         =   "frmDatoTransportistaDetalle.frx":25A2
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
            Picture         =   "frmDatoTransportistaDetalle.frx":2AE4
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
      Height          =   2205
      Left            =   0
      Picture         =   "frmDatoTransportistaDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1710
   End
End
Attribute VB_Name = "frmDatoTransportistaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub Asignar()
    With frmDatoTransportista.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtMarca = IIf(IsNull(!Marca), "", !Marca)
        txtPlaca = IIf(IsNull(!Placa), "", !Placa)
        txtLicencia = IIf(IsNull(!Licencia), "", !Licencia)
    
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
    
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmDatoTransportista.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmDatoTransportista.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmDatoTransportista.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmDatoTransportista.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmDatoTransportista.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmDatoTransportista.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmDatoTransportista.RsCabecera.AbsolutePosition & " de " & frmDatoTransportista.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.value = 1
               'Cambia el Nombre del Primer Text
               txtMarca.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               Dim nPos As Variant
                                        
               'Chequea Datos
               If txtMarca.Text = "" Then MsgBox "Ingrese la Marca del Vehículo", vbExclamation, sMensaje: txtMarca.SetFocus: Exit Sub
               If txtPlaca.Text = "" Then MsgBox "Ingrese la Placa del Vehículo", vbExclamation, sMensaje: txtPlaca.SetFocus: Exit Sub
                                        
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='VEHICULO' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 3)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, tValor, lActivo) " & _
                         "values ('VEHICULO', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtMarca.Text & "', " & _
                                " '" & txtPlaca.Text & "', " & _
                                " '" & txtLicencia.Text & "', " & _
                                       chkActivo.value & ") "
                  Cn.Execute Isql
                  frmDatoTransportista.RsCabecera.Sort = "Codigo ASC"
                  frmDatoTransportista.RsCabecera.Requery
                  frmDatoTransportista.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmDatoTransportista.RsCabecera.RecordCount = 0, 0, frmDatoTransportista.RsCabecera.AbsolutePosition) & " de " & frmDatoTransportista.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtMarca.Text & "', " & _
                         "tResumido ='" & txtPlaca.Text & "', " & _
                         "tValor ='" & txtLicencia.Text & "', " & _
                         "lActivo =" & chkActivo.value & ", lReplica=1 " & _
                         " where tTAbla = 'VEHICULO' and tCodigo = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                   nPos = frmDatoTransportista.RsCabecera.Bookmark
                   frmDatoTransportista.RsCabecera.Requery
                   If frmDatoTransportista.RsCabecera.RecordCount = 0 Then
                      frmDatoTransportista.RsCabecera.Filter = adFilterNone
                   End If
                   frmDatoTransportista.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
               If frmDatoTransportista.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Vehiculo" & txtMarca.Text & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                                                             
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'VEHICULO' and tCodigo = '" & txtCodigo & "'"
               frmDatoTransportista.RsCabecera.Requery
               If frmDatoTransportista.RsCabecera.RecordCount <> 0 Then
                  frmDatoTransportista.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmDatoTransportista.RsCabecera.RecordCount = 0, 0, frmDatoTransportista.RsCabecera.AbsolutePosition) & " de " & frmMotivoEliminacion.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Unidad de Transporte y Conductor "
    fraDetalle.Caption = Me.Caption
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       'frmClienteFactura.RsCabecera.Find ("Codigo = '" & frmClienteFactura.RsCabecera!Codigo & "'")
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & frmDatoTransportista.RsCabecera.AbsolutePosition & " de " & frmDatoTransportista.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmDatoTransportistaDetalle = Nothing
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

Private Sub txtMarca_LostFocus()
   Call ValidaStr(txtMarca)
End Sub

