VERSION 5.00
Begin VB.Form frmCajonDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3720
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9510
   Icon            =   "frmCajonDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
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
      Height          =   3000
      Left            =   1800
      TabIndex        =   16
      Top             =   -30
      Width           =   7665
      Begin VB.CheckBox Check1 
         Caption         =   "Sin Puerto Serial"
         Height          =   240
         Left            =   5985
         TabIndex        =   30
         Top             =   2602
         Width           =   1500
      End
      Begin VB.Frame Frame1 
         Caption         =   " Puerto Serial "
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
         Left            =   1980
         TabIndex        =   21
         Top             =   1434
         Width           =   5505
         Begin VB.OptionButton Option8 
            Caption         =   "COM 8"
            Height          =   195
            Left            =   4140
            TabIndex        =   29
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option7 
            Caption         =   "COM 7"
            Height          =   195
            Left            =   2850
            TabIndex        =   28
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "COM 6"
            Height          =   195
            Left            =   1560
            TabIndex        =   27
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "COM 5"
            Height          =   195
            Left            =   270
            TabIndex        =   26
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM 4"
            Height          =   195
            Left            =   4140
            TabIndex        =   25
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "COM 3"
            Height          =   195
            Left            =   2850
            TabIndex        =   24
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "COM 2"
            Height          =   195
            Left            =   1560
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "COM 1"
            Height          =   195
            Left            =   270
            TabIndex        =   22
            Top             =   360
            Width           =   1095
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
         Left            =   1980
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   678
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
         Left            =   1980
         MaxLength       =   15
         TabIndex        =   2
         Text            =   " "
         Top             =   1056
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
         Left            =   1980
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
         Left            =   1320
         TabIndex        =   3
         Top             =   2625
         Width           =   840
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   723
         Width           =   1650
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   18
         Top             =   1101
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   1305
         TabIndex        =   17
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2970
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
         Left            =   8265
         Picture         =   "frmCajonDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   7095
         Picture         =   "frmCajonDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   4755
         Picture         =   "frmCajonDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   5925
         Picture         =   "frmCajonDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4635
         TabIndex        =   9
         Top             =   60
         Width           =   4695
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmCajonDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmCajonDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmCajonDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4155
            Picture         =   "frmCajonDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3675
            Picture         =   "frmCajonDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3195
            Picture         =   "frmCajonDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   10
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
      Height          =   2880
      Left            =   15
      Picture         =   "frmCajonDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1755
   End
End
Attribute VB_Name = "frmCajonDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCajon As Recordset
Dim nPuerto As Integer

Sub Asignar()
    With frmCajon.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!Codigo), "", !Codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        LimpiaOpt
        nPuerto = IIf(IsNull(!nValor), 0, !nValor)
        Select Case nPuerto
               Case Is = 1
                    Option1.Value = True
               Case Is = 2
                    Option2.Value = True
               Case Is = 3
                    Option3.Value = True
               Case Is = 4
                    Option4.Value = True
               Case Is = 5
                    Option5.Value = True
               Case Is = 6
                    Option6.Value = True
               Case Is = 7
                    Option7.Value = True
               Case Is = 8
                    Option8.Value = True
        End Select
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
End Sub

Private Sub Check1_Click()
   If Check1.Value Then
      Frame1.Enabled = False
      LimpiaOpt
   Else
      Frame1.Enabled = True
   End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmCajon.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmCajon.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmCajon.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmCajon.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmCajon.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmCajon.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmCajon.RsCabecera.RecordCount = 0, 0, frmCajon.RsCabecera.AbsolutePosition) & " de " & frmCajon.RsCabecera.RecordCount
       
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               LimpiaOpt
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
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='CAJON' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, tValor, nValor, lActivo) " & _
                         "values ('CAJON', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', '" & sCaja & "'," & nPuerto & ", " & _
                                       chkActivo.Value & ") "
                  
                  Cn.Execute Isql
                  frmCajon.RsCabecera.Sort = "Codigo ASC"
                  frmCajon.RsCabecera.Requery
                  frmCajon.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmCajon.RsCabecera.RecordCount = 0, 0, frmCajon.RsCabecera.AbsolutePosition) & " de " & frmCajon.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "nValor =" & nPuerto & ", " & _
                         "lActivo =" & chkActivo.Value & ", lReplica=1 " & _
                         " where tTAbla = 'Cajon' and tCodigo = '" & txtCodigo & "' "
                  Cn.Execute Isql
                  nPos = frmCajon.RsCabecera.Bookmark
                  frmCajon.RsCabecera.Requery
                  If frmCajon.RsCabecera.RecordCount = 0 Then
                     frmCajon.RsCabecera.Filter = adFilterNone
                  End If
                  frmCajon.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          
          Case Is = 2 ' Eliminar
               If frmCajon.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar este Cajon de Dinero " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'CAJON' and tCodigo = '" & txtCodigo & "'"
               frmCajon.RsCabecera.Requery
               If frmCajon.RsCabecera.RecordCount <> 0 Then
                  frmCajon.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmCajon.RsCabecera.RecordCount = 0, 0, frmCajon.RsCabecera.AbsolutePosition) & " de " & frmCajon.RsCabecera.RecordCount
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
    nPuerto = 0
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Cajón de Dineron "
    fraDetalle.Caption = Me.Caption
    
    'Ingresar la Tabla
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       LimpiaOpt
       chkActivo.Value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmCajon.RsCabecera.RecordCount = 0, 0, frmCajon.RsCabecera.AbsolutePosition) & " de " & frmCajon.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set RsCajon = Nothing
    Set frmCajonDetalle = Nothing
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

Private Sub Option1_Click()
    nPuerto = 1
End Sub

Private Sub Option2_Click()
    nPuerto = 2
End Sub

Private Sub Option3_Click()
    nPuerto = 3
End Sub

Private Sub Option4_Click()
    nPuerto = 4
End Sub

Private Sub Option5_Click()
    nPuerto = 5
End Sub

Private Sub Option6_Click()
    nPuerto = 6
End Sub

Private Sub Option7_Click()
    nPuerto = 7
End Sub

Private Sub Option8_Click()
    nPuerto = 8
End Sub

Public Sub LimpiaOpt()
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
    Option4.Value = False
    Option5.Value = False
    Option6.Value = False
    Option7.Value = False
    Option8.Value = False
    nPuerto = 0
End Sub
