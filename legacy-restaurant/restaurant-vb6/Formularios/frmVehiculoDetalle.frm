VERSION 5.00
Begin VB.Form frmVehiculoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3165
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9480
   Icon            =   "frmVehiculoDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
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
      Height          =   2385
      Left            =   1755
      TabIndex        =   17
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtSerie 
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
         TabIndex        =   3
         Text            =   " "
         Top             =   1581
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
         Left            =   1515
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   787
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
         Top             =   1184
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
         Top             =   390
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   855
         TabIndex        =   4
         Top             =   1980
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número Serie :"
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   1620
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Vehículo - Placa :"
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   825
         Width           =   1275
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Marca :"
         Height          =   195
         Left            =   870
         TabIndex        =   19
         Top             =   1229
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   825
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
      Top             =   2415
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
         Picture         =   "frmVehiculoDetalle.frx":0442
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
         Picture         =   "frmVehiculoDetalle.frx":0534
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
         Picture         =   "frmVehiculoDetalle.frx":0636
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
         Picture         =   "frmVehiculoDetalle.frx":0B68
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
            Picture         =   "frmVehiculoDetalle.frx":109A
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
            Picture         =   "frmVehiculoDetalle.frx":15DC
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
            Picture         =   "frmVehiculoDetalle.frx":1B1E
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
            Picture         =   "frmVehiculoDetalle.frx":2060
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
            Picture         =   "frmVehiculoDetalle.frx":25A2
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
            Picture         =   "frmVehiculoDetalle.frx":2AE4
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
      Height          =   2340
      Left            =   0
      Picture         =   "frmVehiculoDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1710
   End
End
Attribute VB_Name = "frmVehiculoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Asignar()
    With frmVehiculo.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!Codigo), "", !Codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!Marca), "", !Marca)
        txtSerie = IIf(IsNull(!Serie), "", !Serie)
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
    
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmVehiculo.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmVehiculo.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmVehiculo.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmVehiculo.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmVehiculo.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmVehiculo.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmVehiculo.RsCabecera.AbsolutePosition & " de " & frmVehiculo.RsCabecera.RecordCount
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
               Dim nPos As Integer
                                        
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese el Vehiculo - Placa", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
                                   
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla = 'VEHICULO' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 3)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, tValor, lActivo) " & _
                         "values ('Vehiculo', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                " '" & txtSerie.Text & "', " & _
                                       chkActivo.Value & ") "
            
                  Cn.Execute Isql
                  frmVehiculo.RsCabecera.Sort = "Codigo ASC"
                  frmVehiculo.RsCabecera.Requery
                  frmVehiculo.RsCabecera.MoveLast
                  ActivarBotones (True)
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  cmdTexto.Caption = "Registro " & IIf(frmVehiculo.RsCabecera.RecordCount = 0, 0, frmVehiculo.RsCabecera.AbsolutePosition) & " de " & frmVehiculo.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "tValor ='" & txtSerie.Text & "', " & _
                         "lActivo =" & chkActivo.Value & _
                         ",lreplica=1  where tTabla = 'Vehiculo' and tCodigo = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                   nPos = frmVehiculo.RsCabecera.AbsolutePosition
                   frmVehiculo.RsCabecera.Requery
                   frmVehiculo.RsCabecera.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
               If frmVehiculo.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Vehçiculo " & txtDetallado & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'VEHICULO' and tCodigo = '" & txtCodigo & "'"
               frmVehiculo.RsCabecera.Requery
               If frmVehiculo.RsCabecera.RecordCount <> 0 Then
                  frmVehiculo.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmVehiculo.RsCabecera.RecordCount = 0, 0, frmVehiculo.RsCabecera.AbsolutePosition) & " de " & frmVehiculo.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Vehículos "
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
    
    cmdTexto.Caption = "Registro " & frmVehiculo.RsCabecera.AbsolutePosition & " de " & frmVehiculo.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmVehiculoDetalle = Nothing
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
