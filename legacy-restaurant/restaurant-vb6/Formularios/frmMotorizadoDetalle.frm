VERSION 5.00
Begin VB.Form frmMotorizadoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4200
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11055
   Icon            =   "frmMotorizadoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   11055
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
      Height          =   3495
      Left            =   8130
      TabIndex        =   46
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   840
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
         TabIndex        =   16
         Top             =   840
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
         TabIndex        =   17
         Top             =   840
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
         TabIndex        =   18
         Top             =   840
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
         TabIndex        =   19
         Top             =   840
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
         TabIndex        =   20
         Top             =   1410
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
         TabIndex        =   21
         Top             =   1410
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
         TabIndex        =   22
         Top             =   1410
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
         TabIndex        =   23
         Top             =   1410
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
         TabIndex        =   25
         Top             =   1410
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
         TabIndex        =   26
         Top             =   1980
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
         TabIndex        =   27
         Top             =   1980
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
         TabIndex        =   28
         Top             =   1980
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
         TabIndex        =   29
         Top             =   1980
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
         TabIndex        =   30
         Top             =   1980
         Width           =   510
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boton :"
         Height          =   195
         Left            =   1095
         TabIndex        =   47
         Top             =   2640
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
         TabIndex        =   39
         Top             =   2595
         Width           =   990
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
      Height          =   3495
      Left            =   1440
      TabIndex        =   42
      Top             =   0
      Width           =   6720
      Begin VB.TextBox txtTarifaES 
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
         Left            =   1485
         MaxLength       =   8
         TabIndex        =   38
         Text            =   " "
         Top             =   3000
         Width           =   2595
      End
      Begin VB.TextBox txtTarifaSD 
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
         Left            =   1485
         MaxLength       =   8
         TabIndex        =   37
         Text            =   " "
         Top             =   2640
         Width           =   2595
      End
      Begin VB.TextBox txtTarifaLV 
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
         Left            =   1485
         MaxLength       =   8
         TabIndex        =   36
         Text            =   " "
         Top             =   2280
         Width           =   2595
      End
      Begin VB.TextBox txtDocumento 
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
         MaxLength       =   12
         TabIndex        =   34
         Text            =   " "
         Top             =   1560
         Width           =   2595
      End
      Begin VB.TextBox txtComision 
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
         Left            =   1485
         MaxLength       =   24
         TabIndex        =   35
         Text            =   "0"
         Top             =   1920
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
         Left            =   1485
         MaxLength       =   50
         TabIndex        =   32
         Text            =   " "
         Top             =   765
         Width           =   5070
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
         TabIndex        =   33
         Text            =   " "
         Top             =   1170
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   360
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   5280
         TabIndex        =   31
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa Especial :"
         Height          =   195
         Left            =   270
         TabIndex        =   54
         Top             =   3000
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa Dom :"
         Height          =   195
         Left            =   540
         TabIndex        =   53
         Top             =   2640
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa Lun - Sab :"
         Height          =   195
         Left            =   180
         TabIndex        =   52
         Top             =   2280
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Identidad :"
         Height          =   195
         Left            =   270
         TabIndex        =   51
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   4080
         TabIndex        =   50
         Top             =   1920
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Porc.Comisión :"
         Height          =   195
         Left            =   315
         TabIndex        =   49
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Detallada :"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   795
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Resumida :"
         Height          =   195
         Left            =   150
         TabIndex        =   44
         Top             =   1170
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   825
         TabIndex        =   43
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10995
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3450
      Width           =   11055
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6120
         TabIndex        =   41
         Top             =   60
         Width           =   6180
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4680
            Picture         =   "frmMotorizadoDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5160
            Picture         =   "frmMotorizadoDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5640
            Picture         =   "frmMotorizadoDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmMotorizadoDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmMotorizadoDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmMotorizadoDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   1
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
            Left            =   1530
            TabIndex        =   48
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
         Left            =   7470
         Picture         =   "frmMotorizadoDetalle.frx":23CE
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
         Left            =   6300
         Picture         =   "frmMotorizadoDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   8640
         Picture         =   "frmMotorizadoDetalle.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   9810
         Picture         =   "frmMotorizadoDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   3465
      Left            =   15
      Picture         =   "frmMotorizadoDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   15
      Width           =   1380
   End
End
Attribute VB_Name = "frmMotorizadoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim frmMotorizado.rscabecera As Recordset
Dim RsBoton As Recordset
Dim rsDatosMotorizado As New ADODB.Recordset

Sub Asignar()
    With frmMotorizado.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        txtBoton = IIf(IsNull(!nBoton), "", !nBoton)
        txtComision.Text = Format(Round(IIf(IsNull(!nValor), 0, !nValor), 2), "###,##0.00")
        'Check Box
     '   chkExterno = IIf(IsNull(!nValor), 0, !nValor)
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
        
    Datos
    Botonera
End Sub
Sub Datos()
    Set rsDatosMotorizado = Lib.OpenRecordset("SELECT * FROM TMOTORIZADODATOS WHERE TCODIGO='" & txtCodigo & "'", Cn)
    If Not (rsDatosMotorizado.EOF Or rsDatosMotorizado.BOF) Then
        With rsDatosMotorizado
             txtDocumento = IIf(IsNull(!tdocumentoidentidad), "", !tdocumentoidentidad)
             txtTarifaLV.Text = Format(Round(IIf(IsNull(!ntarifalv), 0, !ntarifalv), 2), "###,##0.00")
             txtTarifaSD.Text = Format(Round(IIf(IsNull(!ntarifasd), 0, !ntarifasd), 2), "###,##0.00")
              txtTarifaES.Text = Format(Round(IIf(IsNull(!ntarifaes), 0, !ntarifaes), 2), "###,##0.00")
       
        End With
    
    
    Else
    
        txtDocumento.Text = ""
        txtTarifaLV.Text = Format(0, "###,##0.00")
        txtTarifaSD.Text = Format(0, "###,##0.00")
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

Private Sub cmdNA_Click()
   If val(txtBoton) <> 0 Then
      cmdBoton(val(txtBoton)).backColor = vbButtonFace
      cmdBoton(val(txtBoton)).Enabled = True
      txtBoton.Caption = "NA"
   End If
End Sub

Public Sub Botonera()
    Dim i As Integer
    txtBoton.Caption = "NA"
    If RsBoton.RecordCount <> 0 Then
        For i = 1 To 19
            RsBoton.MoveFirst
            RsBoton.Find ("nBoton=" & i)
            If RsBoton.EOF Then
               cmdBoton(i).backColor = vbButtonFace
               cmdBoton(i).Enabled = True
            Else
               cmdBoton(i).Enabled = False
               If RsBoton!tCodigo = txtCodigo.Text Then
                  txtBoton.Caption = str(i)
                  cmdBoton(i).backColor = vbRed
               Else
                  cmdBoton(i).backColor = vbBlue
               End If
            End If
        Next i
    Else
       For i = 1 To 19
           cmdBoton(i).backColor = vbButtonFace
           cmdBoton(i).Enabled = True
       Next i
    End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmMotorizado.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmMotorizado.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmMotorizado.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmMotorizado.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmMotorizado.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmMotorizado.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmMotorizado.RsCabecera.RecordCount = 0, 0, frmMotorizado.RsCabecera.AbsolutePosition) & " de " & frmMotorizado.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
        
               chkActivo.Value = 1
               Botonera
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               'Chequea Datos
               If LTrim(txtDetallado.Text) = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If LTrim(txtResumido.Text) = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                    
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='MOTORIZADO' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "0001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 4)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, nboton, nValor, lActivo) " & _
                         "values ('MOTORIZADO', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       val(txtBoton.Caption) & ", " & _
                                " " & val(Me.txtComision.Text) & ", " & _
                                       chkActivo.Value & ") "
           
                   Cn.Execute Isql
                    'Cambiar el SQL
                  Isql = "insert into TMOTORIZADODATOS( " & _
                         " tCodigo, tDocumentoIdentidad, nTarifaLV, nTarifaSD,nTarifaes ) " & _
                         "values ( '" & txtCodigo.Text & "', " & _
                                " '" & txtDocumento.Text & "', " & _
                                       val(txtTarifaLV.Text) & ", " & _
                                " " & val(Me.txtTarifaSD.Text) & "," & val(Me.txtTarifaES.Text) & ") "
           
                  Cn.Execute Isql
                  RsBoton.Requery
                  frmMotorizado.RsCabecera.Sort = "Codigo ASC"
                  frmMotorizado.RsCabecera.Requery
                  frmMotorizado.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmMotorizado.RsCabecera.RecordCount = 0, 0, frmMotorizado.RsCabecera.AbsolutePosition) & " de " & frmMotorizado.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "nBoton =" & val(txtBoton.Caption) & ", " & _
                         "nValor =" & val(txtComision.Text) & ", " & _
                         "lActivo =" & chkActivo.Value & ", lReplica=1 " & _
                         " where tTAbla = 'MOTORIZADO' and tCodigo = '" & txtCodigo & "'"
                       
                       
                       
                       
                   Cn.Execute Isql
                   
                   
                   If Calcular("SELECT COUNT(*) AS CODIGO FROM TMOTORIZADODATOS WHERE TCODIGO='" & txtCodigo.Text & "'", Cn) = 0 Then
                                'Cambiar el SQL
                                Isql = "insert into TMOTORIZADODATOS( " & _
                                       " tCodigo, tDocumentoIdentidad, nTarifaLV, nTarifaSD,nTarifaes ) " & _
                                       "values ( '" & txtCodigo.Text & "', " & _
                                              " '" & txtDocumento.Text & "', " & _
                                                     val(txtTarifaLV.Text) & ", " & _
                                              " " & val(Me.txtTarifaSD.Text) & "," & val(Me.txtTarifaES.Text) & ") "
                   Else
                   
                                Isql = "update TMOTORIZADODATOS set " & _
                                      "tDocumentoIdentidad ='" & txtDocumento.Text & "', " & _
                                      "nTarifaLV =" & val(txtTarifaLV.Text) & ", " & _
                                      "nTarifaSD =" & val(Me.txtTarifaSD.Text) & ",  " & _
                                      "nTarifaes =" & val(Me.txtTarifaES.Text) & "  " & _
                                      " where tCodigo = '" & txtCodigo & "'"
                                
                   End If
                                    
                   Cn.Execute Isql
                   
                   nPos = frmMotorizado.RsCabecera.Bookmark
                   frmMotorizado.RsCabecera.Requery
                   If frmMotorizado.RsCabecera.RecordCount = 0 Then
                      frmMotorizado.RsCabecera.Filter = adFilterNone
                   End If
                   frmMotorizado.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
               If frmMotorizado.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Motorizado" & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               If Calcular("select count(tMotorizado) as Codigo From MPEDIDO where tMotorizado='" & txtCodigo.Text & "'", Cn) > 0 Then
                  MsgBox "Se tiene movimientos con este Motorizado, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'MOTORIZADO' and tCodigo = '" & txtCodigo & "'"
               Cn.Execute "delete from TMOTORIZADODATOS where   tCodigo = '" & txtCodigo & "'"
               
               frmMotorizado.RsCabecera.Requery
               If frmMotorizado.RsCabecera.RecordCount <> 0 Then
                  frmMotorizado.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmMotorizado.RsCabecera.RecordCount = 0, 0, frmMotorizado.RsCabecera.AbsolutePosition) & " de " & frmMotorizado.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Motorizado "
    fraDetalle.Caption = Me.Caption
    
    'Ingresar la Tabla
    'Isql = "select * from TTABLA where tTabla ='MOTORIZADO' and tCodigo <> '0000' order by tCodigo"
    'Set frmMotorizado.rscabecera = Lib.OpenRecordset(Isql, Cn)
    
    'Botones
    Isql = "select tCodigo, nBoton from TTABLA where nBoton > 0 and TTABLA = 'MOTORIZADO' and lActivo=1 order by tCodigo"
    Set RsBoton = Lib.OpenRecordset(Isql, Cn)
        
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
      chkActivo.Value = 1
       Botonera
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       'frmMotorizado.RsCabecera.Find ("Codigo = '" & frmMotorizado.RsCabecera!Codigo & "'")
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmMotorizado.RsCabecera.RecordCount = 0, 0, frmMotorizado.RsCabecera.AbsolutePosition) & " de " & frmMotorizado.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmMotorizadoDetalle = Nothing
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

Private Sub txtComision_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtComision
   
End Sub

Private Sub txtComision_LostFocus()
If txtComision.Text <> "" Then
txtComision.Text = Format(Round(txtComision, 2), "###,##0.00")
End If


End Sub

Private Sub txtDetallado_LostFocus()
   Call ValidaStr(txtDetallado)
End Sub
 

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtDocumento
   
End Sub

Private Sub txtResumido_LostFocus()
    Call ValidaStr(txtResumido)
End Sub

 
Private Sub txtTarifaES_KeyPress(KeyAscii As Integer)
TabNext KeyAscii
   Numerico KeyAscii, txtTarifaES
End Sub

Private Sub txtTarifaLV_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtTarifaLV
End Sub

 

Private Sub txtTarifaSD_KeyPress(KeyAscii As Integer)
 TabNext KeyAscii
   Numerico KeyAscii, txtTarifaSD
End Sub
