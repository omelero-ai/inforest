VERSION 5.00
Begin VB.Form frmOperadorDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5910
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   10995
   FillColor       =   &H00808080&
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmOperadorDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   10995
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
      Height          =   5190
      Left            =   9465
      TabIndex        =   26
      Top             =   0
      Width           =   1545
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
         Left            =   855
         TabIndex        =   40
         Top             =   3870
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
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3275
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
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2683
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
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2091
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
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1499
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
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   907
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
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   315
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
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3870
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
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3275
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
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2683
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
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2091
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
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1499
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
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   907
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
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   315
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
      Height          =   5175
      Left            =   2385
      TabIndex        =   22
      Top             =   0
      Width           =   7080
      Begin VB.CheckBox chkObligaProp 
         Alignment       =   1  'Right Justify
         Caption         =   "Obligatoriedad de Propiedad :"
         Height          =   195
         Left            =   60
         TabIndex        =   53
         Top             =   3330
         Width           =   2415
      End
      Begin VB.CheckBox chkImprime 
         Alignment       =   1  'Right Justify
         Caption         =   "Visible en Documentos :"
         Height          =   195
         Left            =   435
         TabIndex        =   7
         Top             =   2940
         Width           =   2055
      End
      Begin VB.CheckBox chkStockMenos 
         Alignment       =   1  'Right Justify
         Caption         =   "Afecta Inventario ( - ) : "
         Height          =   195
         Left            =   615
         TabIndex        =   5
         Top             =   2220
         Width           =   1875
      End
      Begin VB.TextBox txtControl 
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
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   6
         Text            =   " "
         Top             =   2520
         Width           =   990
      End
      Begin VB.CheckBox chkStockMas 
         Alignment       =   1  'Right Justify
         Caption         =   "Afecta Inventario (+) : "
         Height          =   195
         Left            =   660
         TabIndex        =   4
         Top             =   1905
         Width           =   1830
      End
      Begin VB.CheckBox chkVenta 
         Alignment       =   1  'Right Justify
         Caption         =   "Afecta Valor Venta (+) : "
         Height          =   195
         Left            =   525
         TabIndex        =   3
         Top             =   1557
         Width           =   1965
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   1
         Top             =   759
         Width           =   4590
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
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1158
         Width           =   2835
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1650
         TabIndex        =   8
         Top             =   4035
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "(Las propiedades dentro de este operador se imprimiran en las precuentas y documentos)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2595
         TabIndex        =   52
         Top             =   2880
         Width           =   4530
      End
      Begin VB.Label Label3 
         Caption         =   "(permite incrementar el valor venta del producto"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2595
         TabIndex        =   51
         Top             =   1575
         Width           =   3570
      End
      Begin VB.Label Label2 
         Caption         =   "(excluye insumos de la receta de venta original)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2595
         TabIndex        =   50
         Top             =   2220
         Width           =   3570
      End
      Begin VB.Label Label1 
         Caption         =   "(agrega una receta de propiedad a la receta de venta original)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2595
         TabIndex        =   49
         Top             =   1905
         Width           =   4425
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "( Máximo de propiedades concurrentes, 0 desactivado )"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   3360
         TabIndex        =   48
         Top             =   2580
         Width           =   3420
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Max. prop. concurrentes :"
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   47
         Top             =   2565
         Width           =   1815
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Modificación :"
         Height          =   195
         Index           =   31
         Left            =   705
         TabIndex        =   46
         Top             =   4350
         Width           =   1485
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Usuario Modificación :"
         Height          =   195
         Index           =   32
         Left            =   615
         TabIndex        =   45
         Top             =   4740
         Width           =   1575
      End
      Begin VB.Label lblFecha 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   2280
         TabIndex        =   44
         Top             =   4290
         Width           =   2835
      End
      Begin VB.Label lblUsuario 
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
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   43
         Top             =   4710
         Width           =   2835
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boton :"
         Height          =   195
         Left            =   1680
         TabIndex        =   42
         Top             =   3690
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
         Left            =   2280
         TabIndex        =   13
         Top             =   3645
         Width           =   990
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Index           =   2
         Left            =   540
         TabIndex        =   25
         Top             =   810
         Width           =   1650
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   24
         Top             =   1200
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   3
         Left            =   1605
         TabIndex        =   23
         Top             =   405
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   690
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   10935
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5220
      Width           =   10995
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6060
         TabIndex        =   15
         Top             =   15
         Width           =   6120
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4620
            Picture         =   "frmOperadorDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5100
            Picture         =   "frmOperadorDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5580
            Picture         =   "frmOperadorDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmOperadorDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmOperadorDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmOperadorDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   16
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
            TabIndex        =   41
            Top             =   150
            Width           =   3105
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
         Left            =   7365
         Picture         =   "frmOperadorDetalle.frx":23CE
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   6195
         Picture         =   "frmOperadorDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   8535
         Picture         =   "frmOperadorDetalle.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   9705
         Picture         =   "frmOperadorDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   15
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   5130
      Left            =   15
      Picture         =   "frmOperadorDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2325
   End
End
Attribute VB_Name = "frmOperadorDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsBoton As Recordset

Sub Asignar()
    With frmOperador.RsCabecera
         'Cuadro de Texto
         txtCodigo = IIf(IsNull(!codigo), "", !codigo)
         txtDetallado = IIf(IsNull(!descripcion), "", !descripcion)
         txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
         txtControl = IIf(IsNull(!nControl), 0, !nControl)
         
         'Check Box
         chkVenta = IIf(!lValor = True, 1, 0)
         chkStockMas = IIf(!lStockMas = True, 1, 0)
         chkStockMenos = IIf(!lStockMenos = True, 1, 0)
         chkObligaProp = IIf(!lObligaPropiedad = True, 1, 0)
         chkImprime = IIf(!lImprime = True, 1, 0)
         chkActivo = IIf(!lActivo = True, 1, 0)
         
         lblFecha = IIf(IsNull(!fRegistro), "", !fRegistro)
         lblUsuario = IIf(IsNull(!tUsuario), "", !tUsuario)
         Botonera
    End With
End Sub

Private Sub chkStockMas_Click()
   If chkStockMas.Value Then
      chkStockMenos.Value = False
   End If
End Sub

Private Sub chkStockMenos_Click()

   If chkStockMenos.Value Then
      chkStockMas.Value = False
      txtControl.Text = "0"
      chkObligaProp.Value = False
      chkObligaProp.Enabled = False
   Else
      chkObligaProp.Enabled = True
   End If
End Sub

Private Sub cmdBoton_Click(Index As Integer)
   If val(txtBoton) <> 0 Then
      cmdBoton(val(txtBoton)).backColor = vbButtonFace
      cmdBoton(val(txtBoton)).Enabled = True
   End If
   cmdBoton(Index).backColor = vbRed
   cmdBoton(Index).Enabled = False
   txtBoton.Caption = Index
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmOperador.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmOperador.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmOperador.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmOperador.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmOperador.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmOperador.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmOperador.RsCabecera.RecordCount = 0, 0, frmOperador.RsCabecera.AbsolutePosition) & " de " & frmOperador.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               txtBoton.backColor = &H808080
               txtBoton.Caption = "NA"
               chkVenta = 0
               chkStockMas = 0
               chkStockMenos = 0
               txtControl.Text = "0"
               chkActivo.Value = 1
               Botonera
               
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               Dim nPos As Integer
               
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                                                             
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tOperador) as Codigo from TOPERADOR", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                                      
                  'Cambiar el SQL
                  Isql = "insert into TOPERADOR( " & _
                         "tOperador, tDetallado, tResumido, lValor, lStockMas, lStockMenos, lObligaPropiedad, nControl, lImprime, nBoton, lActivo, tUsuario, fRegistro) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       chkVenta.Value & ", " & _
                                       chkStockMas.Value & ", " & _
                                       chkStockMenos.Value & ", " & _
                                       chkObligaProp.Value & ", " & _
                                       val(txtControl.Text) & ", " & _
                                       chkImprime.Value & ", " & _
                                       val(txtBoton.Caption) & ", " & _
                                       chkActivo.Value & ", " & _
                                " '" & sUsuario & "', getdate()) "
                  Cn.Execute Isql
                  RsBoton.Requery
                  frmOperador.RsCabecera.Sort = "CODIGO ASC"
                  frmOperador.RsCabecera.Requery
                  frmOperador.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmOperador.RsCabecera.RecordCount = 0, 0, frmOperador.RsCabecera.AbsolutePosition) & " de " & frmOperador.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Isql = "update TOPERADOR set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "nBoton =" & val(txtBoton.Caption) & ", " & _
                         "nControl =" & val(txtControl.Text) & ", " & _
                         "lValor =" & chkVenta.Value & ", " & _
                         "lStockMas =" & chkStockMas.Value & ", " & _
                         "lStockMenos =" & chkStockMenos.Value & ", " & _
                         "lObligaPropiedad=" & chkObligaProp.Value & ", " & _
                         "lImprime =" & chkImprime.Value & ", " & _
                         "lActivo =" & chkActivo.Value & ", " & _
                         "tUsuario='" & sUsuario & "', fRegistro=getdate(), lReplica=1 " & _
                         "where tOperador = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                   RsBoton.Requery
                   nPos = frmOperador.RsCabecera.AbsolutePosition
                   frmOperador.RsCabecera.Requery
                   frmOperador.RsCabecera.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               
               Cn.Execute "update TOPERADOR set lNuevoModificado=1 where tOperador = '" & txtCodigo & "'"
               
               Asignar
          
          Case Is = 2 ' Eliminar
                If frmOperador.RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                                
                'Cambia el MsgBox
                If MsgBox("Seguro de Eliminar el Operador" & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                   Exit Sub
                End If
                
                'Cambia el Delete
                Cn.Execute "delete from TOPERADOR where tOperador = '" & txtCodigo & "'"
                frmOperador.RsCabecera.Requery
                If frmOperador.RsCabecera.RecordCount <> 0 Then
                   frmOperador.RsCabecera.MoveLast
                   Asignar
                   cmdTexto.Caption = "Registro " & IIf(frmOperador.RsCabecera.RecordCount = 0, 0, frmOperador.RsCabecera.AbsolutePosition) & " de " & frmOperador.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Operadores "
    fraDetalle.Caption = Me.Caption
        
    'Botones
    Isql = "select tOperador, nBoton from TOPERADOR where nBoton > 0 and lActivo=1 order by tOperador"
    Set RsBoton = Lib.OpenRecordset(Isql, Cn)
    
    If Not lAlmacen Then
       chkStockMas.Enabled = False
       chkStockMenos.Enabled = False
    End If
    
    'Botonera
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       txtBoton.backColor = &H808080
       txtBoton.Caption = "NA"
       chkVenta = 0
       chkStockMas = 0
       chkStockMenos = 0
       chkImprime = 0
       txtControl.Text = "0"
       chkActivo.Value = 1
       Botonera
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmOperador.RsCabecera.RecordCount = 0, 0, frmOperador.RsCabecera.AbsolutePosition) & " de " & frmOperador.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmOperadorDetalle = Nothing
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
   If val(txtBoton) <> 0 Then
      cmdBoton(val(txtBoton)).backColor = vbButtonFace
      cmdBoton(val(txtBoton)).Enabled = True
      txtBoton.Caption = "NA"
   End If
End Sub

Private Sub Botonera()
    RsBoton.Requery
    Dim i As Integer
    txtBoton.Caption = "NA"
    If RsBoton.RecordCount <> 0 Then
        For i = 1 To 13
            RsBoton.MoveFirst
            RsBoton.Find ("nBoton=" & i)
            If RsBoton.EOF Then
               cmdBoton(i).backColor = vbButtonFace
               cmdBoton(i).Enabled = True
            Else
               cmdBoton(i).Enabled = False
               If RsBoton!tOperador = txtCodigo.Text Then
                  txtBoton.Caption = str(i)
                  cmdBoton(i).backColor = vbRed
               Else
                  cmdBoton(i).backColor = vbBlue
               End If
            End If
        Next i
    Else
       For i = 1 To 13
           cmdBoton(i).backColor = vbButtonFace
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
