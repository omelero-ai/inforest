VERSION 5.00
Begin VB.Form frmMotivoReservaDetalle 
   ClientHeight    =   2685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9420
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1935
      Width           =   9480
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   13
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "frmMotivoReservaDetalle.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmMotivoReservaDetalle.frx":0542
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmMotivoReservaDetalle.frx":0A84
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
            Picture         =   "frmMotivoReservaDetalle.frx":0FC6
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
            Picture         =   "frmMotivoReservaDetalle.frx":1508
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
            Picture         =   "frmMotivoReservaDetalle.frx":1A4A
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
         Picture         =   "frmMotivoReservaDetalle.frx":1F8C
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
         Left            =   4710
         Picture         =   "frmMotivoReservaDetalle.frx":24BE
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
         Left            =   7050
         Picture         =   "frmMotivoReservaDetalle.frx":29F0
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
         Left            =   8220
         Picture         =   "frmMotivoReservaDetalle.frx":2AF2
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
      Height          =   1920
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   7155
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1305
         TabIndex        =   4
         Top             =   1530
         Width           =   840
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
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   3
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
         Left            =   1950
         MaxLength       =   15
         TabIndex        =   2
         Text            =   " "
         Top             =   1094
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
         TabIndex        =   1
         Text            =   " "
         Top             =   720
         Width           =   5070
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   1275
         TabIndex        =   7
         Top             =   375
         Width           =   585
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   1139
         Width           =   1680
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   757
         Width           =   1650
      End
   End
   Begin VB.Image Image 
      Height          =   1905
      Left            =   0
      Picture         =   "frmMotivoReservaDetalle.frx":2BE4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2205
   End
End
Attribute VB_Name = "frmMotivoReservaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Asignar()
    With frmMotivoReserva.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Detallado), "", !Detallado)
        txtResumido = IIf(IsNull(!Resumido), "", !Resumido)
        chkActivo = IIf(!Activo = True, 1, 0)
    End With
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmMotivoReserva.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmMotivoReserva.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmMotivoReserva.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmMotivoReserva.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmMotivoReserva.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmMotivoReserva.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmMotivoReserva.RsCabecera.AbsolutePosition & " de " & frmMotivoReserva.RsCabecera.RecordCount
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
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='MOTIVORESERVA' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, lActivo) " & _
                         "values ('MOTIVORESERVA', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       chkActivo.value & ") "
           
                      Cn.Execute Isql
                      frmMotivoReserva.RsCabecera.Sort = "Codigo ASC"
                      frmMotivoReserva.RsCabecera.Requery
                      frmMotivoReserva.RsCabecera.MoveLast
                      MsgBox "Registro Guardado", vbInformation, sMensaje
                      ActivarBotones (True)
                      cmdTexto.Caption = "Registro " & IIf(frmMotivoReserva.RsCabecera.RecordCount = 0, 0, frmMotivoReserva.RsCabecera.AbsolutePosition) & " de " & frmMotivoReserva.RsCabecera.RecordCount
              Else
                 'Cambiar el SQL
                 Isql = "update TTABLA set " & _
                        "tDetallado ='" & txtDetallado.Text & "', " & _
                        "tResumido ='" & txtResumido.Text & "', " & _
                        "lActivo =" & chkActivo.value & _
                        " where tTAbla = 'MOTIVORESERVA' and tCodigo = '" & txtCodigo & "'"
                      
                  Cn.Execute Isql
                  nPos = frmMotivoReserva.RsCabecera.Bookmark
                  frmMotivoReserva.RsCabecera.Requery
                  If frmMotivoReserva.RsCabecera.RecordCount = 0 Then
                     frmMotivoReserva.RsCabecera.Filter = adFilterNone
                  End If
                  frmMotivoReserva.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
              End If
         
          Case Is = 2 ' Eliminar
               If frmMotivoReserva.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Motivo de Reserva " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'MOTIVORESERVA' and tCodigo = '" & txtCodigo & "'"
               frmMotivoReserva.RsCabecera.Requery
               If frmMotivoReserva.RsCabecera.RecordCount <> 0 Then
                  frmMotivoReserva.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmMotivoReserva.RsCabecera.RecordCount = 0, 0, frmMotivoReserva.RsCabecera.AbsolutePosition) & " de " & frmMotivoReserva.RsCabecera.RecordCount
                  
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
    Me.Caption = " Mantenimiento de Motivos de Reservas "
    fraDetalle.Caption = Me.Caption
       
    'Ingresar la Tabla
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & frmMotivoReserva.RsCabecera.AbsolutePosition & " de " & frmMotivoReserva.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMotivoReservaDetalle = Nothing
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

