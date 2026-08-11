VERSION 5.00
Begin VB.Form frmMaitreDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4095
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11025
   Icon            =   "frmMaitreDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   11025
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10965
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3345
      Width           =   11025
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
         Left            =   9750
         Picture         =   "frmMaitreDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   8580
         Picture         =   "frmMaitreDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   6240
         Picture         =   "frmMaitreDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   7410
         Picture         =   "frmMaitreDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6060
         TabIndex        =   46
         Top             =   60
         Width           =   6120
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmMaitreDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmMaitreDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmMaitreDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5580
            Picture         =   "frmMaitreDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5100
            Picture         =   "frmMaitreDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4620
            Picture         =   "frmMaitreDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   9
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
            TabIndex        =   47
            Top             =   150
            Width           =   3105
         End
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
      Height          =   3300
      Left            =   1410
      TabIndex        =   39
      Top             =   0
      Width           =   6660
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   900
         TabIndex        =   5
         Top             =   2595
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   450
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   2
         Text            =   " "
         Top             =   1274
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   862
         Width           =   4950
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   24
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   " "
         Top             =   1686
         Width           =   2595
      End
      Begin VB.TextBox txtBanda 
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
         IMEMode         =   3  'DISABLE
         Left            =   1575
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   " "
         Top             =   2098
         Width           =   2595
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   3
         Left            =   900
         TabIndex        =   44
         Top             =   495
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Resumida :"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   43
         Top             =   1319
         Width           =   1260
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Detallada :"
         Height          =   195
         Index           =   2
         Left            =   255
         TabIndex        =   42
         Top             =   907
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Password :"
         Height          =   195
         Index           =   0
         Left            =   705
         TabIndex        =   41
         Top             =   1731
         Width           =   780
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Banda Magnética :"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   40
         Top             =   2143
         Width           =   1350
      End
   End
   Begin VB.Frame fraBoton 
      Caption         =   " Botonera "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   8100
      TabIndex        =   36
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdBoton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   2
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   3
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   4
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   5
         Left            =   2310
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   900
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   7
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   900
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   8
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   900
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   9
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   900
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   10
         Left            =   2310
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   900
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   11
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1530
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   12
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1530
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   13
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1530
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   14
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1530
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   15
         Left            =   2310
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1530
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   16
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2160
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   17
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2160
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   18
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2160
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   19
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2160
         Width           =   510
      End
      Begin VB.CommandButton cmdNA 
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2310
         TabIndex        =   35
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boton :"
         Height          =   195
         Left            =   1095
         TabIndex        =   38
         Top             =   2865
         Width           =   510
      End
      Begin VB.Label txtBoton 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1830
         TabIndex        =   37
         Top             =   2820
         Width           =   990
      End
   End
   Begin VB.Image Image 
      Height          =   3255
      Left            =   15
      Picture         =   "frmMaitreDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   15
      Width           =   1380
   End
End
Attribute VB_Name = "frmMaitreDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsBoton As Recordset
Dim sBanda As String

Sub Asignar()
    With frmMaitre.RsCabecera
         'Cuadro de Texto
         txtCodigo = IIf(IsNull(!codigo), "", !codigo)
         txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
         txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
         txtBoton = IIf(IsNull(!nBoton), "", !nBoton)
         txtPassword = Desencapsula(IIf(IsNull(!tValor), "", !tValor))
         txtBanda = Desencapsula(IIf(IsNull(!tBandaMagnetica), "", !tBandaMagnetica))
         'Check Box
'         chkVerPedido = IIf(IsNull(!nValor), 0, !nValor)
         chkActivo = IIf(!lActivo = True, 1, 0)
'         chkCambia = IIf(IsNull(!nTamano), 0, !nTamano)
         Botonera
    End With
End Sub

Private Sub cmdBoton_Click(Index As Integer)
   If Val(txtBoton) <> 0 Then
      cmdBoton(Val(txtBoton)).BackColor = vbButtonFace
      cmdBoton(Val(txtBoton)).Enabled = True
   End If
   cmdBoton(Index).BackColor = vbRed
   cmdBoton(Index).Enabled = False
   txtBoton.Caption = Index
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmMaitre.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmMaitre.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmMaitre.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmMaitre.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmMaitre.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmMaitre.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmMaitre.RsCabecera.RecordCount = 0, 0, frmMaitre.RsCabecera.AbsolutePosition) & " de " & frmMaitre.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               Botonera
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                       
               If Sw Then
                  If Calcular("select count(tValor) as codigo from tTabla where tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MAITRE'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                   If Calcular("select count(tValor) as codigo from tTabla where tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MOZO'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  If Calcular("select count(tPassword) as codigo from tusuario where tpassword='" & Encapsula(txtPassword.Text) & "'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  
                  If txtBanda.Text <> "" And Calcular("select count(tIcono) as codigo from tTabla where tIcono='" & Encapsula(txtBanda.Text) & "' and tTabla='MAITRE'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                  If txtBanda.Text <> "" And Calcular("select count(tIcono) as codigo from tTabla where tIcono='" & Encapsula(txtBanda.Text) & "' and tTabla='MOZO'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                  If txtBanda.Text <> "" And Calcular("select count(TBANDAMAGNETICA) as codigo from TUSUARIO where tbandamagnetica='" & Encapsula(txtBanda.Text) & "'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                             
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='Maitre' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "0001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 4)
                  End If
                  Sw = False
                  sBanda = Extrae(txtBanda.Text)
                  
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, nboton, tValor, nValor, tIcono, nTamano, lActivo) " & _
                         "values ('MAITRE', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       Val(txtBoton.Caption) & ", " & _
                                " '" & Encapsula(txtPassword.Text) & "',0, " & _
                                " '" & Encapsula(sBanda) & "', 0,  " & _
                                       chkActivo.value & ") "
           
                  Cn.Execute Isql
                  RsBoton.Requery
                  frmMaitre.RsCabecera.Sort = "CODIGO ASC"
                  frmMaitre.RsCabecera.Requery
                  frmMaitre.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmMaitre.RsCabecera.RecordCount = 0, 0, frmMaitre.RsCabecera.AbsolutePosition) & " de " & frmMaitre.RsCabecera.RecordCount
               Else
                  If txtPassword.Text <> "" And Calcular("select count(tValor) as codigo from tTabla where tCodigo<>'" & txtCodigo & "' and tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MAITRE'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  If Calcular("select count(tPassword) as codigo from tusuario where tpassword='" & Encapsula(txtPassword.Text) & "'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  If Calcular("select count(tValor) as codigo from tTabla where tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MOZO'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  
                  If txtBanda.Text <> "" And Calcular("select count(tIcono) as Codigo from tTabla where tCodigo<>'" & txtCodigo & "' and tIcono='" & Encapsula(txtBanda.Text) & "' and tTabla='MAITRE'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                  If txtBanda.Text <> "" And Calcular("select count(tIcono) as codigo from tTabla where tIcono='" & Encapsula(txtBanda.Text) & "' and tTabla='MOZO'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                  If txtBanda.Text <> "" And Calcular("select count(TBANDAMAGNETICA) as codigo from TUSUARIO where tbandamagnetica='" & Encapsula(txtBanda.Text) & "'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                  
                  sBanda = Extrae(txtBanda.Text)
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "nBoton =" & Val(txtBoton.Caption) & ", " & _
                         "tValor = '" & Encapsula(txtPassword.Text) & "', " & _
                         "tIcono = '" & Encapsula(sBanda) & "', " & _
                         "nValor =0, " & _
                         "nTamano =0, " & _
                         "lActivo =" & chkActivo.value & ", lReplica=1 " & _
                         " where tTAbla = 'MAITRE' and tCodigo = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                   nPos = frmMaitre.RsCabecera.Bookmark
                   frmMaitre.RsCabecera.Requery
                   If frmMaitre.RsCabecera.RecordCount = 0 Then
                      frmMaitre.RsCabecera.Filter = adFilterNone
                   End If
                   frmMaitre.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
                If frmMaitre.RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                'Cambia el MsgBox
                If MsgBox("Seguro de Eliminar el Maitre" & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                   Exit Sub
                End If
                
               If Calcular("select count(tMaitre) as Codigo From MPEDIDO where tMaitre='" & txtCodigo.Text & "'", Cn) > 0 Then
                  MsgBox "Se tiene movimientos con este Maitre, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               End If
                
                'Cambia el Delete
                Cn.Execute "delete from TTABLA where tTabla = 'MAITRE' and tCodigo = '" & txtCodigo & "'"
                frmMaitre.RsCabecera.Requery
                If frmMaitre.RsCabecera.RecordCount <> 0 Then
                   frmMaitre.RsCabecera.MoveLast
                   Asignar
                   cmdTexto.Caption = "Registro " & IIf(frmMaitre.RsCabecera.RecordCount = 0, 0, frmMaitre.RsCabecera.AbsolutePosition) & " de " & frmMaitre.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Maitres "
    fraDetalle.Caption = Me.Caption
    If lInfhotel Then
       cmdOpcion(0).Enabled = False
       cmdOpcion(2).Enabled = False
    End If
        
    'Botones
    Isql = "select tCodigo, nBoton from TTABLA where nBoton > 0 and TTABLA = 'Maitre' and lActivo=1 order by tCodigo"
    Set RsBoton = Lib.OpenRecordset(Isql, Cn)
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
       Botonera
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmMaitre.RsCabecera.RecordCount = 0, 0, frmMaitre.RsCabecera.AbsolutePosition) & " de " & frmMaitre.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmMaitreDetalle = Nothing
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

Private Sub cmdNA_Click()
   If Val(txtBoton) <> 0 Then
      cmdBoton(Val(txtBoton)).BackColor = vbButtonFace
      cmdBoton(Val(txtBoton)).Enabled = True
      txtBoton.Caption = "NA"
   End If
End Sub

Private Sub Botonera()
    Dim i As Integer
    txtBoton.Caption = "NA"
    If RsBoton.RecordCount <> 0 Then
        For i = 1 To 19
            RsBoton.MoveFirst
            RsBoton.Find ("nBoton=" & i)
            If RsBoton.EOF Then
               cmdBoton(i).BackColor = vbButtonFace
               cmdBoton(i).Enabled = True
            Else
               cmdBoton(i).Enabled = False
               If RsBoton!tCodigo = txtCodigo.Text Then
                  txtBoton.Caption = str(i)
                  cmdBoton(i).BackColor = vbRed
               Else
                  cmdBoton(i).BackColor = vbBlue
               End If
            End If
        Next i
    Else
       For i = 1 To 19
           cmdBoton(i).BackColor = vbButtonFace
           cmdBoton(i).Enabled = True
       Next i
    End If
End Sub

Private Sub txtDetallado_LostFocus()
    Call ValidaStr(txtDetallado)
End Sub

Private Sub txtResumido_LostFocus()
   Call ValidaStr(txtResumido)
End Sub

