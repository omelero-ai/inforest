VERSION 5.00
Begin VB.Form frmNuevaTienda 
   Caption         =   "Tienda"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2355
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   6795
   End
   Begin VB.TextBox txtZonaDetalle 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   660
      Width           =   6795
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "Activo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Width           =   1620
   End
   Begin VB.TextBox txtApellido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1860
      MaxLength       =   50
      TabIndex        =   9
      Text            =   "RAMOS ONAGA"
      Top             =   4065
      Width           =   6750
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1860
      MaxLength       =   50
      TabIndex        =   8
      Text            =   "JOSE LUIS"
      Top             =   1800
      Width           =   6750
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
      Height          =   525
      Left            =   1860
      MaxLength       =   80
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmNuevaTienda.frx":0000
      Top             =   2355
      Width           =   6750
   End
   Begin VB.TextBox txtReferencia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1860
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2910
      Width           =   6750
   End
   Begin VB.TextBox txtObservacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1860
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3465
      Width           =   6750
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Teléfono"
      Height          =   555
      Index           =   4
      Left            =   8790
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2895
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Email"
      Height          =   555
      Index           =   7
      Left            =   8790
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3450
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Dirección"
      Height          =   555
      Index           =   2
      Left            =   8790
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Nombres"
      Height          =   555
      Index           =   1
      Left            =   8790
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1770
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Contacto"
      Height          =   555
      Index           =   0
      Left            =   8790
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4035
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombres :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   21
      Top             =   1260
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombres :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   20
      Top             =   660
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombres :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   19
      Top             =   180
      Width           =   1515
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Apellidos :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   14
      Top             =   4095
      Width           =   1515
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombres :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   13
      Top             =   1830
      Width           =   1515
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Dirección :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   1515
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Referencia :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   11
      Top             =   2955
      Width           =   1515
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Observación :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   10
      Top             =   3510
      Width           =   1755
   End
End
Attribute VB_Name = "frmNuevaTienda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
