VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPropertyDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4455
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11430
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmPropertyDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   11430
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
      Height          =   3720
      Left            =   8415
      TabIndex        =   20
      Top             =   0
      Width           =   2985
      Begin VB.CommandButton cmdBoton 
         Caption         =   "24"
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
         Index           =   24
         Left            =   1818
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3065
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "23"
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
         Index           =   23
         Left            =   1252
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3065
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "22"
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
         Index           =   22
         Left            =   686
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3065
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "21"
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
         Index           =   21
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3065
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "20"
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
         Index           =   20
         Left            =   1818
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2506
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
         Left            =   1252
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2506
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
         Left            =   686
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2506
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
         Left            =   2385
         TabIndex        =   38
         Top             =   3065
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2506
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
         Left            =   1818
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1947
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
         Left            =   1252
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1947
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
         Left            =   686
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1947
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1947
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
         Left            =   1818
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1388
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
         Left            =   1252
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1388
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
         Left            =   686
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1388
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1388
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
         Left            =   1818
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   829
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
         Left            =   1252
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   829
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
         Left            =   686
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   829
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   829
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
         Left            =   1818
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   1252
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   686
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
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
         TabIndex        =   21
         Top             =   270
         Width           =   510
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
      Height          =   3705
      Left            =   1845
      TabIndex        =   16
      Top             =   0
      Width           =   6570
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
         Left            =   1485
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1029
         Width           =   4950
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
         Left            =   1485
         MaxLength       =   24
         TabIndex        =   2
         Text            =   " "
         Top             =   1518
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
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   540
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   855
         TabIndex        =   3
         Top             =   2526
         Width           =   840
      End
      Begin MSDataListLib.DataCombo cboOperador 
         Height          =   315
         Left            =   1485
         TabIndex        =   49
         Top             =   2007
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Operador :"
         Height          =   195
         Index           =   0
         Left            =   645
         TabIndex        =   50
         Top             =   2067
         Width           =   750
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
         Left            =   1485
         TabIndex        =   41
         Top             =   2925
         Width           =   990
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boton :"
         Height          =   195
         Left            =   885
         TabIndex        =   40
         Top             =   2970
         Width           =   510
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Detallada :"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   19
         Top             =   1074
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Resumida :"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   18
         Top             =   1563
         Width           =   1260
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   3
         Left            =   810
         TabIndex        =   17
         Top             =   585
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   690
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   11370
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3765
      Width           =   11430
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6510
         TabIndex        =   9
         Top             =   15
         Width           =   6570
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   5070
            Picture         =   "frmPropertyDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5550
            Picture         =   "frmPropertyDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   6030
            Picture         =   "frmPropertyDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmPropertyDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmPropertyDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmPropertyDetalle.frx":1E8C
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
            Left            =   1440
            TabIndex        =   39
            Top             =   150
            Width           =   3555
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
         Left            =   7860
         Picture         =   "frmPropertyDetalle.frx":23CE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   15
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
         Left            =   6690
         Picture         =   "frmPropertyDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   15
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
         Left            =   9030
         Picture         =   "frmPropertyDetalle.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   15
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
         Left            =   10200
         Picture         =   "frmPropertyDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   15
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   3705
      Left            =   15
      Picture         =   "frmPropertyDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1785
   End
End
Attribute VB_Name = "frmPropertyDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsBoton As Recordset
Dim RsOperador As Recordset

Sub Asignar()
    With frmProperty.RsCabecera
         'Cuadro de Texto
         txtCodigo = IIf(IsNull(!Codigo), "", !Codigo)
         txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
         txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
         txtBoton = IIf(IsNull(!nBoton), "", !nBoton)
         cboOperador.BoundText = IIf(IsNull(!tOperador), "", !tOperador)
         'Check Box
         chkActivo = IIf(!lActivo = True, 1, 0)
         Botonera
    End With
End Sub

Private Sub cboOperador_Change()
   Botonera
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
                MoverPuntero Primero, frmProperty.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmProperty.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmProperty.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmProperty.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmProperty.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmProperty.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmProperty.RsCabecera.RecordCount = 0, 0, frmProperty.RsCabecera.AbsolutePosition) & " de " & frmProperty.RsCabecera.RecordCount
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
               Dim nPos As Integer
               
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
               If cboOperador.Text = "" Then MsgBox "Ingrese el Operador", vbExclamation, sMensaje: cboOperador.SetFocus: Exit Sub
                       
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='PROPERTY' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "0001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 4)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, tValor, nboton, lActivo) " & _
                         "values ('PROPERTY', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                " '" & cboOperador.BoundText & "', " & _
                                       Val(txtBoton.Caption) & ", " & _
                                       chkActivo.Value & ") "
           
                  Cn.Execute Isql
                  RsBoton.Requery
                  frmProperty.RsCabecera.Sort = "CODIGO ASC"
                  frmProperty.RsCabecera.Requery
                  frmProperty.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmProperty.RsCabecera.RecordCount = 0, 0, frmProperty.RsCabecera.AbsolutePosition) & " de " & frmProperty.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "tValor ='" & cboOperador.BoundText & "', " & _
                         "nBoton =" & Val(txtBoton.Caption) & ", " & _
                         "lActivo =" & chkActivo.Value & _
                         " where tTAbla = 'PROPERTY' and tCodigo = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                   RsBoton.Requery
                   nPos = frmProperty.RsCabecera.AbsolutePosition
                   frmProperty.RsCabecera.Requery
                   frmProperty.RsCabecera.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
                If frmProperty.RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                'Cambia el MsgBox
                If MsgBox("Seguro de Eliminar el property" & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                   Exit Sub
                End If
                
               If Calcular("select count(tproperty) as Codigo From MPEDIDO where tproperty='" & txtCodigo.Text & "'", Cn) > 0 Then
                  MsgBox "Se tiene movimientos con este property, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               End If
                
                'Cambia el Delete
                Cn.Execute "delete from TTABLA where tTabla = 'property' and tCodigo = '" & txtCodigo & "'"
                frmProperty.RsCabecera.Requery
                If frmProperty.RsCabecera.RecordCount <> 0 Then
                   frmProperty.RsCabecera.MoveLast
                   Asignar
                   cmdTexto.Caption = "Registro " & IIf(frmProperty.RsCabecera.RecordCount = 0, 0, frmProperty.RsCabecera.AbsolutePosition) & " de " & frmProperty.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Propiedades "
    fraDetalle.Caption = Me.Caption
        
    'Botones
    Isql = "select tCodigo, nBoton, tValor as tOperador from TTABLA where nBoton > 0 and TTABLA = 'PROPERTY' order by tCodigo"
    Set RsBoton = Lib.OpenRecordset(Isql, Cn)
    
    With cboOperador
         Isql = "Select * from vOperador where lActivo = 1 order by Codigo"
         Set RsOperador = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOperador
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
       Botonera
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmProperty.RsCabecera.RecordCount = 0, 0, frmProperty.RsCabecera.AbsolutePosition) & " de " & frmProperty.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmPropertyDetalle = Nothing
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
    RsBoton.Filter = "tOperador='" & cboOperador.BoundText & "'"
    txtBoton.Caption = "NA"
    If RsBoton.RecordCount <> 0 Then
        For i = 1 To 24
            RsBoton.MoveFirst
            RsBoton.Find ("nBoton=" & i)
            If RsBoton.EOF Then
               cmdBoton(i).BackColor = vbButtonFace
               cmdBoton(i).Enabled = True
            Else
               cmdBoton(i).Enabled = False
               If RsBoton!tCodigo = txtCodigo.Text Then
                  txtBoton.Caption = Str(i)
                  cmdBoton(i).BackColor = vbRed
               Else
                  cmdBoton(i).BackColor = vbBlue
               End If
            End If
        Next i
    Else
       For i = 1 To 24
           cmdBoton(i).BackColor = vbButtonFace
           cmdBoton(i).Enabled = True
       Next i
    End If
End Sub
