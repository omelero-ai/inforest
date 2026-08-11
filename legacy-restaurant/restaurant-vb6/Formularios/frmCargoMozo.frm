VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCargoMozo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "frmCargoMozo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargoMozo.frx":57E2
   ScaleHeight     =   8625
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtBarra 
      Height          =   345
      Left            =   1440
      TabIndex        =   334
      Top             =   8760
      Width           =   1785
   End
   Begin VB.CommandButton cmdBuscarPedido 
      Caption         =   "Codigo Pedido"
      Height          =   585
      Left            =   5880
      TabIndex        =   332
      Top             =   7920
      Width           =   615
   End
   Begin VB.Timer TimerMESA247 
      Interval        =   280
      Left            =   11400
      Top             =   11400
   End
   Begin VB.TextBox txtEnvioAntes 
      Height          =   285
      Left            =   120
      TabIndex        =   306
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraProductoCombo 
      Caption         =   " Productos de Combos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   8610
      Left            =   7680
      TabIndex        =   197
      Top             =   0
      Width           =   4200
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   246
         Top             =   315
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   245
         Top             =   315
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   244
         Top             =   315
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   243
         Top             =   315
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   242
         Top             =   315
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   241
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   240
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   239
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   238
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   237
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   236
         Top             =   1965
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   235
         Top             =   1965
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   234
         Top             =   1965
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   14
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   233
         Top             =   1965
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   15
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   232
         Top             =   1965
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   231
         Top             =   2790
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   230
         Top             =   2790
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   18
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   229
         Top             =   2790
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   19
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   228
         Top             =   2790
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   20
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   227
         Top             =   2790
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   21
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   226
         Top             =   3615
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   22
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   225
         Top             =   3615
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   23
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   224
         Top             =   3615
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   24
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   223
         Top             =   3615
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   25
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   222
         Top             =   3615
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   26
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   221
         Top             =   4440
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   27
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   220
         Top             =   4440
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   28
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   219
         Top             =   4440
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   29
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   218
         Top             =   4440
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   30
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   217
         Top             =   4440
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   31
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   216
         Top             =   5265
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   32
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   5265
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   33
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   5265
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   34
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   5265
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   35
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   5265
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   36
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   6090
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   37
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   6090
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   38
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   209
         Top             =   6090
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   39
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   208
         Top             =   6090
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   40
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   207
         Top             =   6090
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   41
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   206
         Top             =   6915
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   42
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   205
         Top             =   6915
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   43
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   204
         Top             =   6915
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   44
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   6915
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   45
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   6915
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   46
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   201
         Top             =   7740
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   47
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   7740
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   48
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   7740
         Width           =   720
      End
      Begin VB.CommandButton cmdBuscar 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   2529
         Picture         =   "frmCargoMozo.frx":5B24
         Style           =   1  'Graphical
         TabIndex        =   198
         Top             =   7740
         Width           =   1530
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Ofertas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   5
      Left            =   3600
      Picture         =   "frmCargoMozo.frx":5F66
      Style           =   1  'Graphical
      TabIndex        =   249
      Top             =   8040
      Width           =   1200
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Pre Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   4
      Left            =   6480
      Picture         =   "frmCargoMozo.frx":60B0
      Style           =   1  'Graphical
      TabIndex        =   196
      Top             =   7920
      Width           =   1200
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   8
      Left            =   6510
      Picture         =   "frmCargoMozo.frx":65E2
      Style           =   1  'Graphical
      TabIndex        =   195
      Top             =   7140
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   6
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   194
      Top             =   7140
      Width           =   615
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   7
      Left            =   6510
      Picture         =   "frmCargoMozo.frx":66E4
      Style           =   1  'Graphical
      TabIndex        =   193
      Top             =   6555
      Width           =   1185
   End
   Begin VB.Frame fraCombo 
      Caption         =   " Combo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3585
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
      Begin VB.Timer Timer 
         Enabled         =   0   'False
         Left            =   6480
         Top             =   240
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   8
         Left            =   6390
         Style           =   1  'Graphical
         TabIndex        =   190
         Top             =   2955
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   6
         Left            =   6390
         Picture         =   "frmCargoMozo.frx":67E6
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   1845
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   7
         Left            =   6390
         Picture         =   "frmCargoMozo.frx":68E8
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   2400
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Propiedad y Observación"
         Height          =   555
         Index           =   5
         Left            =   6390
         TabIndex        =   105
         Top             =   1290
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Cantidad"
         Height          =   555
         Index           =   4
         Left            =   5152
         TabIndex        =   104
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   17
         Left            =   4530
         Picture         =   "frmCargoMozo.frx":69EA
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   735
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   16
         Left            =   4530
         Picture         =   "frmCargoMozo.frx":72B4
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "&Aumentar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   5152
         Picture         =   "frmCargoMozo.frx":7B7E
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   735
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "&Disminuir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   5152
         Picture         =   "frmCargoMozo.frx":7C80
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   1290
         Width           =   1230
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   12
         Left            =   4530
         Picture         =   "frmCargoMozo.frx":7D82
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   15
         Left            =   4530
         Picture         =   "frmCargoMozo.frx":864C
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   2955
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   13
         Left            =   4530
         Picture         =   "frmCargoMozo.frx":8F16
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   1290
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   14
         Left            =   4530
         Picture         =   "frmCargoMozo.frx":97E0
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1845
         Width           =   615
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Elimina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   5152
         Picture         =   "frmCargoMozo.frx":A0AA
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1845
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
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
         Height          =   555
         Index           =   0
         Left            =   5152
         Picture         =   "frmCargoMozo.frx":A1AC
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2955
         Width           =   1230
      End
      Begin TrueOleDBGrid80.TDBGrid grdCombo 
         Height          =   3285
         Left            =   60
         TabIndex        =   74
         Top             =   180
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   5794
         _LayoutType     =   4
         _RowHeight      =   26
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "True/False"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   0
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   2
         Splits(0).DividerColor=   32768
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
         PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
         _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(9)   =   ":id=4,.fontname=Arial"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
         _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
         _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
         _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
         _StyleDefs(43)  =   "Named:id=34:Heading"
         _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=34,.fontname=Arial"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(53)  =   ":id=36,.fontname=Arial"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(56)  =   "Named:id=38:HighlightRow"
         _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(60)  =   "Named:id=39:EvenRow"
         _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(62)  =   "Named:id=40:OddRow"
         _StyleDefs(63)  =   ":id=40,.parent=33"
         _StyleDefs(64)  =   "Named:id=41:RecordSelector"
         _StyleDefs(65)  =   ":id=41,.parent=34"
         _StyleDefs(66)  =   "Named:id=42:FilterBar"
         _StyleDefs(67)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame fraPedido 
      Caption         =   " Pedido "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3555
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Width           =   2805
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto Max :"
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
         Index           =   10
         Left            =   80
         TabIndex        =   309
         Top             =   3300
         Width           =   795
      End
      Begin VB.Label txtMontoMax 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   308
         Top             =   3270
         Width           =   1800
      End
      Begin VB.Label txtEntregar 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   303
         Top             =   3030
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Entregar A :"
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
         Index           =   9
         Left            =   165
         TabIndex        =   304
         Top             =   3080
         Width           =   720
      End
      Begin VB.Label txtMozo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   102
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mesero :"
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
         Index           =   8
         Left            =   90
         TabIndex        =   101
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label txtDireccion 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   915
         TabIndex        =   100
         Top             =   2300
         Width           =   1800
      End
      Begin VB.Label txtCliente 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   915
         TabIndex        =   99
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label txtMotorizado 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   98
         Top             =   1440
         Width           =   1800
      End
      Begin VB.Label txtTelefono 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   97
         Top             =   2060
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono :"
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
         Left            =   255
         TabIndex        =   96
         Top             =   2080
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
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
         Index           =   5
         Left            =   225
         TabIndex        =   95
         Top             =   2400
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observac. :"
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
         Index           =   6
         Left            =   195
         TabIndex        =   94
         Top             =   2760
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
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
         Index           =   3
         Left            =   390
         TabIndex        =   93
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Motorizado :"
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
         Index           =   7
         Left            =   90
         TabIndex        =   92
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
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
         Index           =   1
         Left            =   435
         TabIndex        =   91
         Top             =   263
         Width           =   450
      End
      Begin VB.Label txtObservacion 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   915
         TabIndex        =   90
         Top             =   2690
         Width           =   1800
      End
      Begin VB.Label txtDescuento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1365
         TabIndex        =   89
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label txtFecha 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   88
         Top             =   225
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descuento :"
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
         Index           =   2
         Left            =   135
         TabIndex        =   87
         Top             =   960
         Width           =   750
      End
      Begin VB.Label txtPuntoVenta 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   86
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Programac. :"
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
         Index           =   25
         Left            =   90
         TabIndex        =   85
         Top             =   720
         Width           =   795
      End
      Begin VB.Label txtFechaProg 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   84
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido :"
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
         Index           =   0
         Left            =   60
         TabIndex        =   83
         Top             =   480
         Width           =   825
      End
      Begin VB.Label txtTipoPedido 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   82
         Top             =   480
         Width           =   1800
      End
      Begin VB.Label lblmoneda 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   960
         TabIndex        =   81
         Top             =   960
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Enviar Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   3
      Left            =   2400
      TabIndex        =   72
      Top             =   8040
      Width           =   1200
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Combos"
      Height          =   585
      Index           =   12
      Left            =   1200
      TabIndex        =   65
      Top             =   8040
      Width           =   1200
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Visualizar Pedido"
      Height          =   585
      Index           =   13
      Left            =   0
      TabIndex        =   64
      Top             =   8040
      Width           =   1200
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Propiedad y Observación"
      Height          =   585
      Index           =   4
      Left            =   6510
      TabIndex        =   25
      Top             =   5970
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Disminuir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   3
      Left            =   6510
      Picture         =   "frmCargoMozo.frx":A29E
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5385
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Aumentar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   2
      Left            =   6510
      Picture         =   "frmCargoMozo.frx":A3A0
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Cantidad"
      Height          =   585
      Index           =   1
      Left            =   6510
      TabIndex        =   22
      Top             =   4220
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Elimina Prod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   0
      Left            =   6510
      Picture         =   "frmCargoMozo.frx":A4A2
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3615
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "&Observación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   5
      Left            =   6510
      Picture         =   "frmCargoMozo.frx":A5A4
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2985
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Transferencia"
      Height          =   585
      Index           =   4
      Left            =   6510
      TabIndex        =   19
      Top             =   2400
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Camb Mesa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   3
      Left            =   6510
      Picture         =   "frmCargoMozo.frx":A6E6
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1815
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Elimina Ped"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   2
      Left            =   6510
      Picture         =   "frmCargoMozo.frx":A7E0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1230
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Modifica"
      Height          =   585
      Index           =   1
      Left            =   6510
      TabIndex        =   16
      Top             =   645
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   0
      Left            =   6510
      Picture         =   "frmCargoMozo.frx":A8E2
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   60
      Width           =   1185
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   11
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":AE14
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6555
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   10
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":B6DE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5970
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   9
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":BFA8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5385
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   8
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":C872
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   7
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":D13C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4221
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   6
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":DA06
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3630
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   5
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":E2D0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2985
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   4
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":EB9A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   3
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":F464
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1815
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   2
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":FD2E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1230
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   1
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":105F8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   645
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   0
      Left            =   5880
      Picture         =   "frmCargoMozo.frx":10EC2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   615
   End
   Begin TrueOleDBGrid80.TDBGrid grdDetalle 
      Height          =   4335
      Left            =   0
      TabIndex        =   73
      Top             =   3645
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   7646
      _LayoutType     =   4
      _RowHeight      =   26
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0).NumberFormat=   "True/False"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).AnchorRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).ScrollBars=   0
      Splits(0).AllowColSelect=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerStyle=   2
      Splits(0).DividerColor=   32768
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
      PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
      _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(9)   =   ":id=4,.fontname=Arial"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
      _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Arial"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
      _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
      _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
      _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
      _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
      _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
      _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
      _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
      _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
      _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
      _StyleDefs(43)  =   "Named:id=34:Heading"
      _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
      _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
      _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(47)  =   ":id=34,.fontname=Arial"
      _StyleDefs(48)  =   "Named:id=35:Footing"
      _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   "Named:id=36:Selected"
      _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
      _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(53)  =   ":id=36,.fontname=Arial"
      _StyleDefs(54)  =   "Named:id=37:Caption"
      _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(56)  =   "Named:id=38:HighlightRow"
      _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
      _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
      _StyleDefs(60)  =   "Named:id=39:EvenRow"
      _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(62)  =   "Named:id=40:OddRow"
      _StyleDefs(63)  =   ":id=40,.parent=33"
      _StyleDefs(64)  =   "Named:id=41:RecordSelector"
      _StyleDefs(65)  =   ":id=41,.parent=34"
      _StyleDefs(66)  =   "Named:id=42:FilterBar"
      _StyleDefs(67)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBGrid grdCabecera 
      Height          =   3465
      Left            =   2880
      TabIndex        =   75
      Top             =   90
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   6112
      _LayoutType     =   4
      _RowHeight      =   26
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0).NumberFormat=   "True/False"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).ScrollBars=   0
      Splits(0).AllowColSelect=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerStyle=   2
      Splits(0).DividerColor=   32768
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
      PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "Cabecera de Pedidos"
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
      _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(9)   =   ":id=4,.fontname=Arial"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
      _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Arial"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
      _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
      _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
      _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
      _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
      _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
      _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
      _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
      _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
      _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
      _StyleDefs(43)  =   "Named:id=34:Heading"
      _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
      _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
      _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(47)  =   ":id=34,.fontname=Arial"
      _StyleDefs(48)  =   "Named:id=35:Footing"
      _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   "Named:id=36:Selected"
      _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
      _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(53)  =   ":id=36,.fontname=Arial"
      _StyleDefs(54)  =   "Named:id=37:Caption"
      _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(56)  =   "Named:id=38:HighlightRow"
      _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
      _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
      _StyleDefs(60)  =   "Named:id=39:EvenRow"
      _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(62)  =   "Named:id=40:OddRow"
      _StyleDefs(63)  =   ":id=40,.parent=33"
      _StyleDefs(64)  =   "Named:id=41:RecordSelector"
      _StyleDefs(65)  =   ":id=41,.parent=34"
      _StyleDefs(66)  =   "Named:id=42:FilterBar"
      _StyleDefs(67)  =   ":id=42,.parent=33"
   End
   Begin VB.Frame fraPropiedad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   8610
      Left            =   7680
      TabIndex        =   146
      Top             =   0
      Width           =   4200
      Begin VB.CommandButton cmdBuscar 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   3345
         Picture         =   "frmCargoMozo.frx":1178C
         Style           =   1  'Graphical
         TabIndex        =   302
         Top             =   4950
         Width           =   720
      End
      Begin VB.TextBox lblObservacion 
         Height          =   555
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   186
         Top             =   7155
         Width           =   3975
      End
      Begin VB.TextBox lblResumen 
         Height          =   870
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   188
         Top             =   5940
         Width           =   3975
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   183
         Top             =   4950
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   19
         Left            =   1815
         TabIndex        =   182
         Top             =   4920
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   20
         Left            =   2580
         TabIndex        =   181
         Top             =   4950
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   18
         Left            =   3345
         TabIndex        =   180
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   1815
         TabIndex        =   177
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   2580
         TabIndex        =   176
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   15
         Left            =   3345
         TabIndex        =   175
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   1815
         TabIndex        =   172
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   14
         Left            =   2580
         TabIndex        =   171
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   3345
         TabIndex        =   170
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   1815
         TabIndex        =   167
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   2580
         TabIndex        =   166
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   3345
         TabIndex        =   165
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   1815
         TabIndex        =   162
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   2580
         TabIndex        =   161
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   3345
         TabIndex        =   160
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   1815
         TabIndex        =   157
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   2580
         TabIndex        =   156
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   3345
         TabIndex        =   155
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   1815
         TabIndex        =   152
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   2580
         TabIndex        =   151
         Top             =   1100
         Width           =   720
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
         Height          =   720
         Index           =   8
         Left            =   2970
         Picture         =   "frmCargoMozo.frx":11BCE
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   7785
         Width           =   1110
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   1440
         Picture         =   "frmCargoMozo.frx":11CC0
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   7785
         Width           =   1470
      End
      Begin VB.CommandButton cmdBuscar 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   870
         Picture         =   "frmCargoMozo.frx":11E02
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   4950
         Width           =   720
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   5415
         Left            =   1672
         TabIndex        =   147
         Top             =   225
         Width           =   60
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Observacion :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   192
         Top             =   6885
         Width           =   1200
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Propiedad :"
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
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   191
         Top             =   5715
         Width           =   990
      End
      Begin VB.Label lblPropiedad 
         AutoSize        =   -1  'True
         Caption         =   "  Operador   "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   185
         Top             =   45
         Width           =   1095
      End
      Begin VB.Label lblPropiedad 
         AutoSize        =   -1  'True
         Caption         =   "  Propiedad "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   3
         Left            =   1980
         TabIndex        =   184
         Top             =   45
         Width           =   1050
      End
   End
   Begin VB.Frame fraEliminacion 
      Caption         =   " Motivo de Eliminación "
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
      Height          =   6650
      Left            =   7680
      TabIndex        =   106
      Top             =   120
      Width           =   4155
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   2505
         Picture         =   "frmCargoMozo.frx":12244
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   5625
         Width           =   1515
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   2505
         TabIndex        =   144
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   1725
         TabIndex        =   143
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   930
         TabIndex        =   142
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   150
         TabIndex        =   141
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   3300
         TabIndex        =   140
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   2505
         TabIndex        =   139
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   1725
         TabIndex        =   138
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   930
         TabIndex        =   137
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   150
         TabIndex        =   136
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   3300
         TabIndex        =   135
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   14
         Left            =   2505
         TabIndex        =   134
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   1725
         TabIndex        =   133
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   930
         TabIndex        =   132
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   150
         TabIndex        =   131
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   15
         Left            =   3300
         TabIndex        =   130
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   19
         Left            =   2505
         TabIndex        =   129
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   18
         Left            =   1725
         TabIndex        =   128
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   930
         TabIndex        =   127
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   150
         TabIndex        =   126
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   20
         Left            =   3300
         TabIndex        =   125
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   24
         Left            =   2505
         TabIndex        =   124
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   23
         Left            =   1725
         TabIndex        =   123
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   22
         Left            =   930
         TabIndex        =   122
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   21
         Left            =   150
         TabIndex        =   121
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   25
         Left            =   3300
         TabIndex        =   120
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   29
         Left            =   2505
         TabIndex        =   119
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   28
         Left            =   1725
         TabIndex        =   118
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   27
         Left            =   930
         TabIndex        =   117
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   26
         Left            =   150
         TabIndex        =   116
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   30
         Left            =   3300
         TabIndex        =   115
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   34
         Left            =   2505
         TabIndex        =   114
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   33
         Left            =   1710
         TabIndex        =   113
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   32
         Left            =   930
         TabIndex        =   112
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   31
         Left            =   150
         TabIndex        =   111
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   35
         Left            =   3300
         TabIndex        =   110
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   36
         Left            =   150
         TabIndex        =   109
         Top             =   5625
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   37
         Left            =   930
         TabIndex        =   108
         Top             =   5625
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   38
         Left            =   1710
         TabIndex        =   107
         Top             =   5625
         Width           =   720
      End
   End
   Begin VB.Frame fraPlato 
      Caption         =   " Platos de Venta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6585
      Left            =   7680
      TabIndex        =   250
      Top             =   0
      Width           =   4200
      Begin VB.Frame fraProducto 
         Caption         =   " Productos "
         ForeColor       =   &H000000C0&
         Height          =   5925
         Left            =   1740
         TabIndex        =   268
         Top             =   210
         Width           =   2325
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   1530
            Picture         =   "frmCargoMozo.frx":127CE
            Style           =   1  'Graphical
            TabIndex        =   289
            Top             =   5070
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   20
            Left            =   810
            TabIndex        =   288
            Top             =   5070
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   19
            Left            =   90
            TabIndex        =   287
            Top             =   5070
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   18
            Left            =   1530
            TabIndex        =   286
            Top             =   4270
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   17
            Left            =   810
            TabIndex        =   285
            Top             =   4270
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   16
            Left            =   90
            TabIndex        =   284
            Top             =   4270
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   15
            Left            =   1530
            TabIndex        =   283
            Top             =   3470
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   14
            Left            =   810
            TabIndex        =   282
            Top             =   3470
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   13
            Left            =   90
            TabIndex        =   281
            Top             =   3470
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   12
            Left            =   1530
            TabIndex        =   280
            Top             =   2670
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   11
            Left            =   810
            TabIndex        =   279
            Top             =   2670
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   10
            Left            =   90
            TabIndex        =   278
            Top             =   2670
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   1530
            TabIndex        =   277
            Top             =   1870
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   810
            TabIndex        =   276
            Top             =   1870
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   90
            TabIndex        =   275
            Top             =   1870
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   1530
            TabIndex        =   274
            Top             =   1070
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   810
            TabIndex        =   273
            Top             =   1070
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   90
            TabIndex        =   272
            Top             =   1070
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   1530
            TabIndex        =   271
            Top             =   270
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   810
            TabIndex        =   270
            Top             =   270
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   90
            TabIndex        =   269
            Top             =   270
            Width           =   720
         End
      End
      Begin VB.Frame fraGrupo 
         Caption         =   " Grupos "
         ForeColor       =   &H000000C0&
         Height          =   2595
         Left            =   60
         TabIndex        =   261
         Top             =   210
         Width           =   1635
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   0
            Left            =   840
            Picture         =   "frmCargoMozo.frx":12C10
            Style           =   1  'Graphical
            TabIndex        =   267
            Top             =   1710
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   266
            Top             =   1710
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   265
            Top             =   990
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   264
            Top             =   990
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   263
            Top             =   270
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   262
            Top             =   270
            Width           =   720
         End
      End
      Begin VB.Frame fraSubGrupo 
         Caption         =   " SubGrupos "
         ForeColor       =   &H000000C0&
         Height          =   3315
         Left            =   60
         TabIndex        =   252
         Top             =   2820
         Width           =   1635
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   840
            Picture         =   "frmCargoMozo.frx":13052
            Style           =   1  'Graphical
            TabIndex        =   260
            Top             =   2460
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   259
            Top             =   2460
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   258
            Top             =   1740
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   257
            Top             =   1740
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   256
            Top             =   1020
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   255
            Top             =   1020
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   254
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   253
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Código Directo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   120
         TabIndex        =   251
         Top             =   6180
         Width           =   3945
      End
   End
   Begin VB.Frame fraOpcion 
      Caption         =   " Opcion "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   7740
      TabIndex        =   66
      Top             =   6570
      Width           =   4155
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Buscar Reservas"
         Height          =   585
         Index           =   9
         Left            =   150
         TabIndex        =   335
         Top             =   1030
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdMesa247 
         BackColor       =   &H0000FF00&
         Caption         =   "Pedidos de MESA 24/7"
         Height          =   615
         Left            =   1440
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   310
         Top             =   360
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Despachar Pedidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   7
         Left            =   2790
         Picture         =   "frmCargoMozo.frx":13494
         Style           =   1  'Graphical
         TabIndex        =   301
         ToolTipText     =   "Confirmar Entrega de Pedidos"
         Top             =   360
         Width           =   1230
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
         Height          =   585
         Index           =   0
         Left            =   2790
         Picture         =   "frmCargoMozo.frx":13596
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   1020
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Busca Pedido"
         Height          =   585
         Index           =   10
         Left            =   150
         TabIndex        =   68
         Top             =   360
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Refrescar"
         Height          =   585
         Index           =   11
         Left            =   1470
         TabIndex        =   67
         Top             =   360
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Aceptar"
         Height          =   585
         Index           =   1
         Left            =   150
         TabIndex        =   70
         Top             =   1020
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Cancelar"
         Height          =   585
         Index           =   2
         Left            =   1470
         TabIndex        =   71
         Top             =   1020
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.Frame fraCabecera 
      Caption         =   " Opciones de Pedido "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   6540
      Left            =   7695
      TabIndex        =   26
      Top             =   0
      Width           =   4185
      Begin VB.Frame fraOrigenVentas 
         Caption         =   "Origen de Ventas"
         ForeColor       =   &H00800000&
         Height          =   3405
         Left            =   90
         TabIndex        =   311
         Top             =   1320
         Width           =   3975
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   3150
            Picture         =   "frmCargoMozo.frx":13688
            Style           =   1  'Graphical
            TabIndex        =   333
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   19
            Left            =   2385
            TabIndex        =   330
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   18
            Left            =   1635
            TabIndex        =   329
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   17
            Left            =   870
            TabIndex        =   328
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   16
            Left            =   120
            TabIndex        =   327
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   15
            Left            =   3150
            TabIndex        =   326
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   14
            Left            =   2385
            TabIndex        =   325
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   13
            Left            =   1635
            TabIndex        =   324
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   12
            Left            =   870
            TabIndex        =   323
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   11
            Left            =   120
            TabIndex        =   322
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   10
            Left            =   3150
            TabIndex        =   321
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   2385
            TabIndex        =   320
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   1635
            TabIndex        =   319
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   870
            TabIndex        =   318
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   120
            TabIndex        =   317
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   3150
            TabIndex        =   316
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   2385
            TabIndex        =   315
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   1635
            TabIndex        =   314
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   870
            TabIndex        =   313
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            TabIndex        =   312
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "Monto Max/P : 0"
         Height          =   515
         Index           =   21
         Left            =   2820
         TabIndex        =   307
         Top             =   5965
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "&Entregar A"
         Height          =   515
         Index           =   19
         Left            =   1515
         TabIndex        =   305
         Top             =   5965
         Width           =   1155
      End
      Begin VB.Frame fraPuntoVenta 
         Caption         =   " Punto de Venta "
         ForeColor       =   &H00800000&
         Height          =   1830
         Left            =   90
         TabIndex        =   290
         Top             =   1305
         Width           =   3975
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            TabIndex        =   300
            Top             =   315
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   870
            TabIndex        =   299
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   1635
            TabIndex        =   298
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   2385
            TabIndex        =   297
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   3150
            TabIndex        =   296
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   120
            TabIndex        =   295
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   870
            TabIndex        =   294
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   1635
            TabIndex        =   293
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   2385
            TabIndex        =   292
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   3150
            Picture         =   "frmCargoMozo.frx":13ACA
            Style           =   1  'Graphical
            TabIndex        =   291
            Top             =   1050
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "Alta Prioridad"
         Height          =   515
         Index           =   9
         Left            =   210
         TabIndex        =   56
         Top             =   4910
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "Programación"
         Height          =   515
         Index           =   10
         Left            =   1515
         TabIndex        =   55
         Top             =   4910
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "&Observación"
         Height          =   515
         Index           =   11
         Left            =   2820
         TabIndex        =   54
         Top             =   4910
         Width           =   1155
      End
      Begin VB.Frame fraMotorizado 
         Caption         =   " Motorizado "
         ForeColor       =   &H00800000&
         Height          =   3350
         Left            =   90
         TabIndex        =   31
         Top             =   1305
         Width           =   3975
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   870
            TabIndex        =   50
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   1620
            TabIndex        =   49
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   2370
            TabIndex        =   48
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   3120
            TabIndex        =   47
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   870
            TabIndex        =   45
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   1620
            TabIndex        =   44
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   2370
            TabIndex        =   43
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   10
            Left            =   3120
            TabIndex        =   42
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   11
            Left            =   120
            TabIndex        =   41
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   12
            Left            =   870
            TabIndex        =   40
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   13
            Left            =   1620
            TabIndex        =   39
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   14
            Left            =   2370
            TabIndex        =   38
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   15
            Left            =   3120
            TabIndex        =   37
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   16
            Left            =   120
            TabIndex        =   36
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   17
            Left            =   870
            TabIndex        =   35
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   18
            Left            =   1620
            TabIndex        =   34
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   19
            Left            =   2370
            TabIndex        =   33
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   3120
            Picture         =   "frmCargoMozo.frx":13F0C
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2550
            Width           =   720
         End
      End
      Begin VB.Frame fraTipoPedido 
         Caption         =   " Tipo de Pedido "
         ForeColor       =   &H00800000&
         Height          =   1110
         Left            =   90
         TabIndex        =   27
         Top             =   300
         Width           =   3975
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "&Delivery"
            Height          =   720
            Index           =   16
            Left            =   3120
            TabIndex        =   248
            Top             =   240
            Width           =   720
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "&Delivery"
            Height          =   720
            Index           =   14
            Left            =   2355
            TabIndex        =   247
            Top             =   240
            Width           =   720
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "&Para Llevar"
            Height          =   720
            Index           =   6
            Left            =   60
            TabIndex        =   30
            Top             =   240
            Width           =   720
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "En el &Local"
            Height          =   720
            Index           =   7
            Left            =   825
            TabIndex        =   29
            Top             =   240
            Width           =   720
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "&Delivery"
            Height          =   720
            Index           =   8
            Left            =   1590
            TabIndex        =   28
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "&Mesa"
         Height          =   515
         Index           =   13
         Left            =   1515
         TabIndex        =   52
         Top             =   5430
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "&Cliente"
         Height          =   515
         Index           =   12
         Left            =   210
         TabIndex        =   53
         Top             =   5430
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "&Descuento"
         Height          =   515
         Index           =   15
         Left            =   2820
         TabIndex        =   103
         Top             =   5430
         Width           =   1155
      End
   End
   Begin VB.TextBox txtbuscarpedidobarra 
      Height          =   375
      Left            =   3360
      TabIndex        =   331
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Image imagepIE 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imageCab 
      Height          =   135
      Left            =   0
      Top             =   7920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblFecha 
      Alignment       =   1  'Right Justify
      Caption         =   "10/01/1970"
      Height          =   255
      Left            =   7710
      TabIndex        =   14
      Top             =   8340
      Width           =   4155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Pedido"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   24
      Left            =   4853
      TabIndex        =   1
      Top             =   8040
      Width           =   1005
   End
   Begin VB.Label txtSuma 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4860
      TabIndex        =   0
      Top             =   8280
      Width           =   990
   End
End
Attribute VB_Name = "frmCargoMozo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xOperador As String
Dim nombreMozo As String
Dim wAgregarPropiedad As Boolean
Public RsCabecera As Recordset
Public RsDetalle As Recordset
Dim RsCombo As Recordset
Dim RsImpresion As Recordset
Dim RsGrupo As Recordset
Dim RsSubgrupo As Recordset
Dim RsProducto As Recordset
Dim RsMotorizado As Recordset
Dim RsArea As Recordset
Dim RsPropiedad As Recordset
Dim RsT3 As Recordset
Dim RsOferta As Recordset
Dim RsMotivoEliminacion As Recordset
Dim RsProductoPropiedad As Recordset
Dim RsComboPropiedad As Recordset
Dim Recordset As Recordset
Dim RsOperador As Recordset
Dim RsProductoCombo As Recordset
Dim RsCanalesVenta As Recordset
Dim rsPuntoVenta As Recordset
Dim RsExisteOferta As Recordset
Dim RsPedido As Recordset '-- MESA247
Dim contadormesa As Integer '-- refresar
'Variables Generales
Dim wDetalle As Boolean
Dim wAgrega As Boolean
'entregarA
Dim lObligaEntregarA As Boolean
Dim lObligaOrigenVenta As Boolean

'Variables para oferta
Dim ValorOferta As Recordset
Dim respuestaOferta As String

'Variables Combo
Dim wCombo As Boolean
Dim wAgregaCombo As Boolean
Dim nCombo As Integer
Dim nCCombo As Double
'-- VARIABLE COLOR MESA247
'Dim mesacolor As Integer


'origen de ventas
Dim RscanalOrigenVentas As Recordset
Dim vOrigenVentas As String
Dim RsOrigenVentas As Recordset
Dim lOrigenVentas As Boolean


Dim ValProductoAsociado As Recordset


'Variables Cabecera
Dim sTipoPedido As String
Dim sTipoAtencion As String
Dim sMotorizado As String
Dim sCortesia As String
Dim sFechaProg As String
Dim nCorrela As Long
Dim xMontoMaximo As Double '---- monto maximo erick
'Dim ntTiempo As Integer
Dim sDetalleConsumo As String
Public lIncluido As Boolean
Dim sUsuarioAutoriza As String
Dim wCabecera As Boolean

Dim sPuntoVenta As String

'Variables Detalles
Dim nPVenta As Double
Dim nPBase As Double
Dim nImpuesto1 As Double
Dim nImpuesto2 As Double
Dim nImpuesto3 As Double
Dim nRecargo As Double
Dim nDescuento As Double
Dim nOficial As Double
Dim nCantidad As Double
Dim sItem As String
Dim xItem As String
Dim sProducto As String
Dim sProductoCombo As String
Dim sCombo As String
Dim sGrupo As String
Dim xGrupo As String
Dim sSubGrupo As String
Dim xSubGrupo As String
Dim sTD As String
Dim xSuma As Double
Dim xLinea As Double
Dim xConsumo As Double
Dim lPrecuenta As Boolean
Dim sMozoD As String
Dim lPropiedad As Boolean
Dim lComboActivo As Boolean
Dim lCombo As Boolean
Dim wSalir As Boolean

Dim xDescuento As Double
Dim sDescuento As Double
Dim Acumulado As Double
Dim sCodigoDescuento As String
Dim sDescripcionDescuento As String
Dim wCalcula As Boolean
Dim tAutorizaDescuento As String
Dim nOrden As Integer
Dim nOperadorPropiedad As Integer


'===================== tope
Dim lRatio As Boolean
Dim ltope As Boolean
Dim nTope As Double
Dim procedeDescuento As Boolean
Dim codigoanteriordescuento As String
Dim montoanteriorDescuento As Double

'insumo critico23
Dim muestra As String

'Canales de Venta
Dim lActivaMozo As Boolean
Dim lActivaMotorizado As Boolean
Dim lObligaMesa As Boolean
Dim lObligaMotorizado As Boolean
Dim lObligaMozo As Boolean
Dim lObligaFechaEntrega As Boolean
Dim lObligaClienteFrecuente As Boolean
Dim lCanalDelivery As Boolean
Dim lCanalCentralPedidos As Boolean

'---MODIFICAR DESCUENTO
Dim sAutoriza As String
 
'auditoria
Dim tNombreMozo As String

'invitado2013
Dim RsTemporal As New Recordset
'invitado2013

Public RsTimer As Recordset
'Verificacion de Combo
Dim iProductoVerificaComboMaximo As Integer
Dim iProductoVerificaComboMinimo As Integer
Dim iProductoVerificaComboUsado As Integer
Dim sProductoVerificaComboUsado As String

Dim lEstadoBuscarPedido As Boolean
Private Sub cmdBuscar_Click(Index As Integer)
    Dim i As Integer
    frmMozoUsuario.ReseteaTimer
    Select Case Index
           Case Is = 0 'Grupo
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                sTemp = ""
                Isql = "select * from vGrupo where lActivo = 1 Order by Descripcion"
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Grupo", 2, "Descripcion", 5800, 0, 0, "", _
                                                                "Boton", 2, "nBoton", 1200, 1, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                If wEnter = True Then
                   For i = 1 To 5
                       cmdGrupo(i).backColor = vbButtonFace
                       cmdSubGrupo(i).backColor = vbButtonFace
                   Next i
                   cmdSubGrupo(6).backColor = vbButtonFace
                   cmdSubGrupo(7).backColor = vbButtonFace
                   
                   sGrupo = sCodigo
                   xGrupo = sCodigo
                   RsGrupo.MoveFirst
                   RsGrupo.Find ("Codigo ='" & sCodigo & "'")
                   If Not RsGrupo.EOF And RsGrupo!nBoton > 0 Then
                      cmdGrupo(RsGrupo!nBoton).backColor = vbBlue
                   End If
                   RsSubgrupo.Filter = ("tGrupo = '" & xGrupo & "'")
                   Call AsignaBoton(7, RsSubgrupo, cmdSubGrupo())
                                      
                   If RsSubgrupo.RecordCount <> 0 Then
                      RsSubgrupo.MoveFirst
                      RsSubgrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vSubGrupo where nBoton>0 and tGrupo = '" & sGrupo & "'", Cn) & "'"
                      sSubGrupo = RsSubgrupo!codigo
                      xSubGrupo = RsSubgrupo!codigo
                      If Not RsSubgrupo.EOF And RsSubgrupo!nBoton > 0 Then
                         cmdSubGrupo(RsSubgrupo!nBoton).backColor = vbRed
                      End If
                   Else
                      sSubGrupo = ""
                      xSubGrupo = ""
                   End If
                   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
                   Call AsignaBotonProducto(20, RsProducto, cmdProducto(), sTipoPedido, sUnidadNegocio)
                End If
           
           Case Is = 1 'SubGrupo
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                sTemp = ""
                Isql = "select * from vSubGrupo where lActivo = 1 and tGrupo ='" & xGrupo & "' Order by Descripcion"
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "SubGrupo", 2, "Descripcion", 5800, 0, 0, "", _
                                                                "Boton", 2, "nBoton", 1200, 1, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
   
                If wEnter Then
                   For i = 1 To 7
                       cmdSubGrupo(i).backColor = vbButtonFace
                   Next i
                
                   sSubGrupo = sCodigo
                   xSubGrupo = sCodigo
                   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
                   
                   RsSubgrupo.MoveFirst
                   RsSubgrupo.Find ("Codigo ='" & sCodigo & "'")
                   If Not RsSubgrupo.EOF And RsSubgrupo!nBoton > 0 Then
                         cmdSubGrupo(RsSubgrupo!nBoton).backColor = vbRed
                   End If
                   Call AsignaBotonProducto(20, RsProducto, cmdProducto(), sTipoPedido, sUnidadNegocio)
                End If
           
           Case Is = 2 'Producto
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                sTemp = ""
                'Cn
                Isql = "select * from vProducto where lActivo = 1 and tunidadnegocio='" & sUnidadNegocio & "'  and tSubGrupo ='" & xSubGrupo & "' and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lcanal4=1", "lcanal5=1")))) & " Order by Descripcion"
                Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                                "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                                "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                                "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                                "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                
                If wEnter Then
                
                    'INSUMOCRITICO23
                    If validadIngresoProducto(sCodigo) = False Then
                        Exit Sub
                    End If
                     'INSUMOCRITICO23
 
                
                
                
                   sProducto = sCodigo
                   'Correccion
                   ListarOperadoresConFiltro (sProducto)
                   RsProducto.MoveFirst
                   RsProducto.Find "Codigo = '" & sProducto & "'"
                             
                   If vComanda Then
                      sTipo = "Comanda"
                      sDescrip = sComanda
                      frmNumPad.Caption = "Comanda"
                      frmNumPad.Show vbModal
                    
                      sComanda = IIf(sDescrip = "" Or sDescrip = "0", "", Mid("0000000000", 1, 10 - Len(Trim(sDescrip))) & Trim(sDescrip))
                      
                     If wComanda And sComanda = "" Then
                        MsgBox "No ha ingresado el número de la comanda", vbExclamation, sMensaje
                        Exit Sub
                     End If
                   
                     If wEnter Then
                     
                     Else
                        Exit Sub
                     End If
                   End If
                    nCantidad = 1
                   InsertaProducto 1
                   If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                       lPropiedad = True
                   End If
                  
                End If
                               
           Case Is = 4 'Motorizado
                sTemp = ""
                Isql = "select * from vMotorizado where lActivo = 1 Order by Descripcion"
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                If wEnter = True Then
                   sMotorizado = sCodigo
                   txtMotorizado.Caption = sDescrip
                End If
                
           Case Is = 5 'Punto de Venta
                sTipo = "Infhotel"
                sTemp = ""
                Isql = "Select distinct vPuntoVenta.tPuntoVenta as Codigo, vPuntoVenta.tDescripcion as Descripcion, nUltimoComanda, vPuntoVenta.tmoneda " & _
                       "FROM dbo.vPuntoVenta INNER JOIN dbo.vTipoProductoPuntoVenta ON dbo.vPuntoVenta.tPuntoVenta = dbo.vTipoProductoPuntoVenta.tPuntoVenta " & _
                       "where tTipoCargo='02' and vPuntoVenta.tHotel='" & sHotel & "' AND ACTIVO='1' and lInforest=1"
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                If wEnter = True Then
                   txtPuntoVenta.Caption = sDescrip
                End If
                
           Case Is = 6 'Operadores
                sTipo = ""
                sTemp = ""
                ListarOperadoresConFiltro (sProducto)
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                
                If wEnter = True Then
                    Screen.MousePointer = vbHourglass
                    For i = 1 To 13
                        cmdOperador(i).backColor = vbButtonFace
                    Next i
                    RsOperador.MoveFirst
                    RsOperador.Find "Codigo='" & sCodigo & "'"
                    If Not RsOperador.EOF And RsOperador!nBoton > 0 Then
                       cmdOperador(RsOperador!nBoton).backColor = vbRed
                    End If
                    AsignaPropiedad
                    xOperador = sCodigo
                    Screen.MousePointer = vbDefault
                End If
                
           Case Is = 7 'ProductoCombo
                 If RsCabecera.RecordCount = 0 Then
                    Exit Sub
                 End If
                 
                 sTemp = ""
                 If lComboGeneral Then
                    Isql = "select * from vProducto where lActivo = 1 and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " Order by Descripcion"
                 Else
                    Isql = "SELECT * FROM dbo.vProducto INNER JOIN dbo.TCOMBO ON dbo.vProducto.Codigo = dbo.TCOMBO.tCodigoProducto Where (dbo.vProducto.lActivo = 1) And tCombo='" & sProductoCombo & "' And " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " ORDER BY dbo.vProducto.Descripcion "
                 End If
                 Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                                 "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                                 "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                                 "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                                 "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
                 frmBusquedaRapida.nPredeterm = 1
                 frmBusquedaRapida.Show vbModal
        
                 If wEnter Then
                    sProducto = sCodigo
                    Dim xxx As String
                    xxx = RsProducto.Filter
                    RsProducto.Filter = adFilterNone
                    RsProducto.MoveFirst
                    RsProducto.Find "Codigo = '" & sProducto & "'"

                     '----erick de la cruz ------------------------------------------
                        'Obtener Si es LUnico
                        Dim oRsProductoDeCombo As Recordset
                        Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sProducto)
                        If oRsProductoDeCombo.RecordCount > 0 Then
                             If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                                 'Obtener Suma de cantidades
                                 Dim nCantidadEnElCombo As Integer
                                 nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sPedido, sItem, oRsProductoDeCombo!tEtiqueta)
                                 'Suma de cantidades < que nCantidad
                                 If nCantidadEnElCombo < nCantidad Then
                                     InsertaCombo sProducto
                                 Else
                                     MsgBox "Solo es permitido " & nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                                 End If
                             Else
                                 InsertaCombo sProducto
                             End If
                         Else
                             InsertaCombo sProducto
                         End If
                        '----------------------------------------------------------
                        'InsertaCombo sProducto

'                    nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
'                                       "FROM dbo.CPEDIDO WHERE CPEDIDO.tCodigoPedido='" & sPedido & "' and tItem='" & sitem & "'", Cn)
'                    If nCCombo < nCombo * RsDetalle!nCantidad Then
'                       InsertaCombo sProducto
'                    Else
'                       MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
'                    End If
                    RsProducto.Filter = IIf(xxx = "0", "", xxx)
                 End If
      Case 8 'propiedad
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                            
                sTipo = ""
                sTemp = ""
                'Isql = "select * from vOperador where lActivo = 1 Order by Descripcion "
                'ListarOperadoresConFiltro (sProducto)
                'CASO CORP.FERNANDEZ
                If BDLink = "1" Then
                    If wAgregaCombo = False Then
                            Isql = "usp_Inforest_ObtienePropiedadesBusqueda '" & xOperador & "','" & sProducto & "','" & sPedido & "','" & sItem & "','" & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & "','" & sLocal & "'"
                    Else
                            Isql = "usp_Inforest_ObtienePropiedadesComboBusqueda '" & xOperador & "','" & sCombo & "','" & sPedido & "','" & sItem & "','" & xItem & "','" & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & "','" & sLocal & "'"
                    End If
                Else
                    If wAgregaCombo = False Then
                            Isql = "usp_Inforest_ObtienePropiedadesBusqueda '" & xOperador & "','" & sProducto & "','" & sPedido & "','" & sItem & "','" & sAlmacenMDB & "','" & sLocal & "'"
                    Else
                            Isql = "usp_Inforest_ObtienePropiedadesComboBusqueda '" & xOperador & "','" & sCombo & "','" & sPedido & "','" & sItem & "','" & xItem & "','" & sAlmacenMDB & "','" & sLocal & "'"
                    End If
                End If
                
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Estado", 2, "tEstado", 1500, 0, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 5500, 0, 0, "")
                                                                
                frmBusquedaRapida.nPredeterm = 2
                frmBusquedaRapida.tipoBusqueda = "BusquedaVenta"
                frmBusquedaRapida.Show vbModal
                
                If wEnter = True Then
                   ' Screen.MousePointer = vbHourglass
                    If wAgregaCombo = False Then
                            If Calcular("SELECT COUNT(*) AS CODIGO FROM TPRODUCTOPROPIEDAD WHERE tCodigoPedido='" & sPedido & "' AND tItem='" & sItem & "' AND tCodigoPropiedad='" & sCodigo & "' AND TPRODUCTO='" & sProducto & "'  and tEnlace='" & tenlacebusquedaVenta & "'", Cn) = 0 Then
                                    wAgregarPropiedad = True
                            Else
                                    wAgregarPropiedad = False
                            End If
                            AgregarPropiedadBusqueda sCodigo, sDescrip
                    Else
                    
                           If Calcular("SELECT COUNT(*) AS CODIGO FROM tcombopropiedad WHERE tCodigoPedido='" & sPedido & "' AND tItem='" & sItem & "' AND tCodigoPropiedad='" & sCodigo & "' AND TPRODUCTO='" & sCombo & "' and titemcombo='" & xItem & "'", Cn) = 0 Then
                                    wAgregarPropiedad = True
                            Else
                                    wAgregarPropiedad = False
                            End If
                            AgregarPropiedadBusqueda sCodigo, sDescrip
                    End If
                    
                    
                  '  Screen.MousePointer = vbDefault
                End If
              '  txtBarra.SetFocus

       Case Is = 9 'Origen de ventas
                sTemp = ""
                Isql = "select CodOrigenVenta as Codigo,* from vOrigenVenta where Activo = 1 and CodCanalVenta='" & sTipoPedido & "' and CodOrigenVenta not in (select CodigoOrigenVenta from TCAJAORIGEN_BLOQUEO  where tcaja='" & sCaja & "' ) Order by Descripcion"
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.tipoBusqueda = "OrigenVentas"
                frmBusquedaRapida.Show vbModal
                If wEnter = True Then
                   vOrigenVentas = sCodigo
                    Call validarOrigenVentas
                End If
'                txtBarra.SetFocus
     
   End Select
End Sub

Private Sub cmdBuscarPedido_Click()
On Error GoTo fin

    frmKeyBoard.Caption = "Buscar por Codigo Pedido"
    frmKeyBoard.txtResultado.Text = ""
    frmKeyBoard.Show vbModal
    
    If wEnter = True Then

    End If
    
    Dim rsPed As Recordset
    lEstadoBuscarPedido = False
    
    If wEnter Then
          cmdOpcion_Click (2)
          
          Isql = ""
          Isql = "select * from mpedido where tcodigopedido = '" & Trim(Mid(sDescrip, 1, 10)) & "'"
          Set rsPed = Lib.OpenRecordset(Isql, Cn)
    
          If rsPed.RecordCount = 0 Then MsgBox "Pedido no encontrado : " & Trim(sDescrip): sDescrip = "": Exit Sub
    
          Isql = ""
          Isql = "select * from mpedido c inner join dpedido d on c.tcodigopedido = d.tcodigopedido " & _
          "and c.tcodigopedido = '" & Trim(Mid(sDescrip, 1, 10)) & "' " & _
          "and d.limprime = 0"
    
          Set rsPed = Lib.OpenRecordset(Isql, Cn)
          If Not (Not rsPed.BOF And Not rsPed.EOF) Then
             MsgBox "No se puede mostrar Pedido: " & Trim(Mid(sDescrip, 1, 10)) & Chr(10) & Chr(13) & _
             "Pedido enviado a las areas de producción en su totalidad."
             txtbuscarpedidobarra.Text = ""
             Exit Sub
          End If
          
          Cn.Execute "update mpedido set tmozo='" & sMozo & "' where tcodigopedido = '" & Trim(Mid(sDescrip, 1, 10)) & "'"
          Isql = ""
          Isql = "select * from vPedidoCabecera " & _
                 "where tEstadoPedido <> '03' and " & _
                 "codigo = '" & Trim(Mid(sDescrip, 1, 10)) & "'"
    
          Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
          Set grdCabecera.DataSource = RsCabecera
    
          Isql = ""
          Isql = "select * from vPedidoDetalle " & _
                 "where tEstadoItem ='N' and " & _
                 "tcodigopedido = '" & Trim(Mid(sDescrip, 1, 10)) & "'"
       
          Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
          Set grdDetalle.DataSource = RsDetalle
          
          Call ProcesoBuscarPedidoBarra(False)
          lEstadoBuscarPedido = True

    End If

Exit Sub
fin:

End Sub

Private Sub cmdCabecera_Click(Index As Integer)
 
    Dim sEstado As String
    frmMozoUsuario.ReseteaTimer
    wCalcula = False
    Me.Caption = muestra
    Select Case Index
           Case Is = 0 'Nuevo
                'txtSuma.Text = ""
                grdCabecera.Enabled = False
                ActivaCabecera False
                ActivaFrame fraCabecera
                ActivaOpcion False
                fraTipoPedido.Enabled = True
                InicializaVariableCabecera
                cmdCabecera(13).Enabled = True
                cmdCabecera(9).FontBold = False
                
                'Llena el Detalle
                RsDetalle.Filter = "tCodigoPedido = '" & sPedido & "'"
                
                Select Case sTipoPedidoPD
                       Case Is = "01"
                            cmdCabecera_Click (6)
                            vOrigenVentas = sDefaultOrigenVenta
                       Case Is = "02"
                            cmdCabecera_Click (7)
                            vOrigenVentas = sDefaultOrigenVenta
                       Case Is = "03"
                            cmdCabecera_Click (8)
                            vOrigenVentas = sDefaultOrigenVenta
                       Case Is = "04"
                            cmdCabecera_Click (14)
                            vOrigenVentas = sDefaultOrigenVenta
                       Case Is = "05"
                            cmdCabecera_Click (16)
                            vOrigenVentas = sDefaultOrigenVenta
                       Case Else
                End Select
                
                 'If lcover Then
                        cmdCabecera(21).Caption = "Monto Max/P : 0"
                 'End If
                 
                ActivaMesa247 (False)
                               
           Case Is = 1 'Modifica
                If RsCabecera.RecordCount <> 0 Then
                    If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                   ActivaCabecera False
                   ActivaFrame fraCabecera
                   ActivaOpcion False
                   wDetalle = False
                   cmdCabecera(13).Enabled = False
                   cmdOpcion(1).Enabled = True
                   
                   If lmodificatipoPedido = False Then
                        fraTipoPedido.Enabled = False
                   Else
                        fraTipoPedido.Enabled = True
                   End If
                   
                   'Canales de Venta
                   RsCanalesVenta.MoveFirst
                   RsCanalesVenta.Filter = "Codigo = '" & sTipoPedido & "'"
                   'lActivaMozo = IIf(IsNull(RsCanalesVenta!lActivaMozo), False, RsCanalesVenta!lActivaMozo)
                   lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
                   lCanalDelivery = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
                   lCanalCentralPedidos = IIf(IsNull(RsCanalesVenta!lCanalCentralPedidos), False, RsCanalesVenta!lCanalCentralPedidos)
                   lObligaMotorizado = IIf(IsNull(RsCanalesVenta!lObligaMotorizado), False, RsCanalesVenta!lObligaMotorizado)
                   lObligaMesa = IIf(IsNull(RsCanalesVenta!lObligaMesa), False, RsCanalesVenta!lObligaMesa)
                   lObligaPax = IIf(IsNull(RsCanalesVenta!lObligaPax), False, RsCanalesVenta!lObligaPax)
                   lObligaFechaEntrega = IIf(IsNull(RsCanalesVenta!lObligaIngresoFechaEntrega), False, RsCanalesVenta!lObligaIngresoFechaEntrega)
                   lObligaClienteFrecuente = IIf(IsNull(RsCanalesVenta!lObligaClienteFrecuente), False, RsCanalesVenta!lObligaClienteFrecuente)
                   lObligaOrigenVenta = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
                   
                   If lActivaMotorizado Then
                      fraMotorizado.Visible = True
                   Else
                      fraMotorizado.Visible = False
                   End If
                   
                    'If lcover Then
                        cmdCabecera(21).Caption = "Monto Max/P : " & Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
                    'End If
                   
                End If
                ActivaMesa247 (False)
           Case Is = 2 'Elimina
           nombreMozo = sUsuario
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
                Dim i As Integer
                If sPedido <> "" Then
                   If Calcular("select count(tFacturado) as Codigo from DPEDIDO where tCodigoPedido ='" & sPedido & "' and isnull(tFacturado,'0')<>'0' and len(ltrim(tFacturado)) <> 0", Cn) > 0 Then
                      MsgBox "Imposible eliminar pedidos con items Facturados", vbExclamation, sMensaje
                      Exit Sub
                   End If
                   
                   'GCAA 10/07/2023 - REST-012-C-2023
                   If lBloqueoAnulaPedido Then
                        MsgBox "Esta opcion se encuentra Bloqueada, favor de comunicar al Supervisor", vbExclamation, sMensaje
                        Exit Sub
                   End If
                   
                   If MsgBox("Seguro de Eliminar el Pedido?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                      sUsuarioAutoriza = sUsuario
                      If lPasswordC Then
                         If Supervisor("01") = False Then
                            MsgBox "Clave no permitida", vbExclamation, sMensaje
                            Exit Sub
                         End If
                         sUsuarioAutoriza = sVar1
                         sUsuario = UCase(sVar1)
                      End If
                       
                      If lEliminaC Then
                         wCabecera = True
                         ActivaFrame fraEliminacion
                         ActivaCabecera False
                         ActivaOpcion False
                         cmdOpcion(1).Visible = False
                         cmdOpcion(2).Visible = False
                      Else
                         sCodigo = ""
                         sDescrip = ""
                         EliminaCabecera
                      End If
                   End If
                End If
           
           Case Is = 3 'Mesa
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                
                If Not lBuscarPedidoBarraMesero Then
                    If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
                 If lcover Then
                   If Supervisor("26") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                End If
                
                
                Dim xSql As String
                sTipo = "M"
                frmMesa.Show vbModal
                If wEnter Then
                   Cn.Execute "Update MPEDIDO set nAdulto = " & Val(frmVenta.txtAdulto.Text) & ", nNino = " & Val(frmVenta.txtNino.Text) & ", nMesa= " & Val(frmVenta.txtJuntar.Text) & " where tCodigoPedido ='" & sPedido & "'"
                   RsCabecera.Requery
                   RsCabecera.Find "Codigo = '" & sPedido & "'"
                End If
                
                If lBuscarPedidoBarraMesero And lEstadoBuscarPedido Then
                    Cn.Execute "update tmesa set testadomesa ='04' where tcodigomesa='" & sMesa & "'"
                End If
                
                If wMesa Then
                   xSql = "nAdulto = " & Val(frmVenta.txtAdulto.Text) & ", nNino = " & Val(frmVenta.txtNino.Text) & ", nMesa= " & Val(frmVenta.txtJuntar.Text)
                   If wCambioMesa Then
                      If Calcular("select count(tMesa) as Codigo from TPEDIDOMESA where tCodigoPedido='" & sPedido & "'", Cn) > 0 Then
                         MsgBox "No se puede mover la Mesa cuando tiene mesas asociadas", vbCritical, sMensaje
                         Exit Sub
                      End If
                   
                      sEstado = Calcular("select tEstadoMesa as Codigo from TMESA where tCodigoMesa ='" & sCodigo & "'", Cn)
                      If sEstado = "01" Or sEstado = "04" Or sEstado = "0" Then
                         xSql = xSql & ", tMesa = '" & sCodigo & "', tSalon='" & IIf(sCodigo <> "", Calcular("select tsalon as Codigo from tmesa where tCodigoMesa='" & sCodigo & "'", Cn), sSalon) & "'"
                         Cn.Execute "Update TMESA set tEstadoMesa='02' where tCodigoMesa ='" & sCodigo & "'"
                         Cn.Execute "Update TMESA set tEstadoMesa='04' where tCodigoMesa ='" & sTempMesa & "'"
                         sMesa = sCodigo
                         
                         If lPrinter And lCambioMesa Then
                            sDescrip = "Cambio de Mesa/Ubicacion a " & Calcular("SELECT dbo.vSalon.Descripcion + ' - ' + dbo.TMESA.tDetallado AS Codigo FROM dbo.TMESA INNER JOIN dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo where tCodigoMesa='" & sCodigo & "'", Cn)
                            Dim RsTemp As Recordset
                            frmMozoUsuario.timSalida.Interval = 0
                            frmPrecuentaImpresora.cmdOpcion(0).Enabled = False
                            frmPrecuentaImpresora.Show vbModal
                            frmMozoUsuario.timSalida.Interval = nSalir
                            If Not wEnter Then
                               Exit Sub
                            End If
                               
                            Set RsTemp = Lib.OpenRecordset("SELECT dbo.vTipoPedido.Descripcion AS TipoPedido, dbo.vMozo.Descripcion AS Mozo, dbo.vSalon.Descripcion + ' - ' + dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tObservacion AS Observacion, dbo.MPEDIDO.nAdulto AS Pax FROM dbo.vMozo RIGHT OUTER JOIN dbo.MPEDIDO LEFT OUTER JOIN dbo.TMESA INNER JOIN dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo LEFT OUTER JOIN dbo.vTipoPedido ON dbo.MPEDIDO.tTipoPedido = dbo.vTipoPedido.Codigo where tCodigoPedido='" & sPedido & "'", Cn)
                            ImprimeMensaje sCodigo, sDescrip, RsTemp
                            Set RsTemp = Nothing
                         End If
                         
                         Cn.Execute "Update MPEDIDO set " & xSql & "  where tCodigoPedido ='" & sPedido & "'"
                         RsCabecera.Requery
                         RsCabecera.Find "Codigo = '" & sPedido & "'"
                         wMesa = False
                      Else
                         MsgBox "Mesa Ocupada", vbExclamation, sMensaje
                         Exit Sub
                      End If
                   End If
                End If
                
           Case Is = 4 'Transferencia
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
           
                If lPasswordTransferencia Then
                   If Supervisor("15") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                     'inicia lg 25/06
                   sUsuarioAutoriza = sVar1
                   'fin
                End If
           
                wEnter = False
                frmDivision.txtFecha = txtFecha.Caption
                frmDivision.txtTipoPedido.ForeColor = txtTipoPedido.ForeColor
                frmDivision.txtTipoPedido.Caption = txtTipoPedido.Caption
                frmDivision.txtDescuento.Caption = txtDescuento.Caption
                frmDivision.txtMozo.Caption = txtMozo.Caption
                frmDivision.txtMotorizado.Caption = txtMotorizado.Caption
                frmDivision.txtObservacion.Caption = txtObservacion.Caption
                'inicia 25/06
                frmDivision.usuarioautoriza sUsuarioAutoriza
                'fin
                  
                frmDivision.txtMesa.Caption = grdCabecera.Columns(1).Text
                frmDivision.txtOrden.Caption = grdCabecera.Columns(0).Text
                frmDivision.Show vbModal
                If wEnter Then
                   RsCabecera.Requery
                   RsDetalle.Requery
                   RsCombo.Requery
                   RsProductoPropiedad.Requery
                   If RsCabecera.RecordCount <> 0 Then
                      RsCabecera.Find "Codigo = '" & sPedido & "'"
                      AsignaPedido
                   End If
                    'MESA247
                    If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                          Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                    End If
                End If

           Case Is = 5 'Observación
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
           
                frmKeyBoard.Caption = "Nombre / Observación"
                frmKeyBoard.txtResultado.Text = txtObservacion.Caption
                frmKeyBoard.Show vbModal
                
                If wEnter = True Then
                   txtObservacion.Caption = sDescrip
                   Cn.Execute "Update MPEDIDO set tObservacion = '" & sDescrip & "' where tCodigoPedido ='" & sPedido & "'"
                   RsCabecera.Requery
                   RsCabecera.Find "Codigo = '" & sPedido & "'"
                   
                    'MESA247
                    If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                        Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                    End If
                End If
                
           Case Is = 6 'Primer Boton
                txtTipoPedido.Caption = sBoton1
                txtTipoPedido.ForeColor = &H800000
                sTipoPedido = "01"
                InicializaVariableCanales
                cmdCabecera(6).FontBold = True
                cmdCabecera(7).FontBold = False
                cmdCabecera(8).FontBold = False
                cmdCabecera(14).FontBold = False
                cmdCabecera(16).FontBold = False
                vOrigenVentas = ""
                                            
           Case Is = 7 'Segundo Boton
                txtTipoPedido.Caption = sBoton2
                txtTipoPedido.ForeColor = &H800000
                sTipoPedido = "02"
                InicializaVariableCanales
                cmdCabecera(6).FontBold = False
                cmdCabecera(7).FontBold = True
                cmdCabecera(8).FontBold = False
                cmdCabecera(14).FontBold = False
                cmdCabecera(16).FontBold = False
                vOrigenVentas = ""
                               
           Case Is = 8 'Tercer Boton
                txtTipoPedido.Caption = sBoton3
                txtTipoPedido.ForeColor = &H800000
                sTipoPedido = "03"
                InicializaVariableCanales
                cmdCabecera(6).FontBold = False
                cmdCabecera(7).FontBold = False
                cmdCabecera(8).FontBold = True
                cmdCabecera(14).FontBold = False
                cmdCabecera(16).FontBold = False
                vOrigenVentas = ""
           
           Case Is = 14 'Canal4
                txtTipoPedido.Caption = sBoton4
                txtTipoPedido.ForeColor = &H800000
                sTipoPedido = "04"
                InicializaVariableCanales
                cmdCabecera(6).FontBold = False
                cmdCabecera(7).FontBold = False
                cmdCabecera(8).FontBold = False
                cmdCabecera(14).FontBold = True
                cmdCabecera(16).FontBold = False
                vOrigenVentas = ""
                
           Case Is = 16 'Canal5
                txtTipoPedido.Caption = sBoton5
                txtTipoPedido.ForeColor = &H800000
                sTipoPedido = "05"
                InicializaVariableCanales
                cmdCabecera(6).FontBold = False
                cmdCabecera(7).FontBold = False
                cmdCabecera(8).FontBold = False
                cmdCabecera(14).FontBold = False
                cmdCabecera(16).FontBold = True
                vOrigenVentas = ""
                                
           Case Is = 9 'Alta Prioridad
                txtTipoPedido.ForeColor = IIf(txtTipoPedido.ForeColor = vbRed, &H800000, vbRed)
                cmdCabecera(9).FontBold = IIf(cmdCabecera(9).FontBold = True, False, True)
                
           Case Is = 10 'Programa
                frmPrograma.Show vbModal
                If wEnter = True Then
                   txtFechaProg.Caption = sCodigo
                   txtEnvioAntes.Text = lMinutoEnvioAntes
                Else
                   txtFechaProg.Caption = ""
                End If
                lMinutoEnvioAntes = 0
                                
           Case Is = 11 'Observación
                frmKeyBoard.Caption = "Nombre / Observación"
                frmKeyBoard.txtResultado.Text = txtObservacion.Caption
                frmKeyBoard.Show vbModal
                If wEnter = True Then
                   txtObservacion.Caption = sDescrip
                End If
                      
           Case Is = 12 'Cliente
           
          If Not Sw Then
                If Calcular("select isnull(nDescuento,0) as codigo from MPEDIDO where tCodigoPedido='" & sPedido & "' ", Cn) > 0 Then
                    MsgBox "No se puede cambiar Cliente del Pedido, Pedido con Descuento!! Revertir Descuento!!", vbInformation
                    Exit Sub
                End If
           End If
          
           
                If lClub = False Then
                    If Not Sw Then
                      sTemp = txtTelefono.Caption
                    End If
                    
                    wEnter = False
                    sTipo = sTipoPedido
                    frmBusquedaDelivery.txtResultado.Caption = txtTelefono.Caption
                    frmBusquedaDelivery.Show vbModal
                Else
                    If Not Sw Then
                        frmBusquedaSocio.sCodigoPariente = sCodigoParienteSeleccionado
                        frmBusquedaSocio.sCodigoSocio = txtTelefono.Caption
                        frmBusquedaSocio.sCodigoInvitadoB = sCodigoInvitado
                    
                    Else
                        frmBusquedaSocio.sCodigoPariente = ""
                        frmBusquedaSocio.sCodigoSocio = ""
                        frmBusquedaSocio.sCodigoInvitadoB = ""
                    
                    End If

                    frmBusquedaSocio.Show vbModal
                End If
                        

                If wEnter = True Then
                
                   sCliente = sCodigo
                   txtCliente.Caption = ""
                   
                   If sCodigoParienteSeleccionado <> "" Then
                        txtCliente.Caption = Calcular("select pariente + ' ' +case when lconyugue=1 then ' Conyugue de ' else 'Hijo de ' end as codigo from vPariente where tCodigoPariente='" & sCodigoParienteSeleccionado & "' and tCodigoDelivery='" & sCliente & "'", Cn)
                   End If
                   
                   txtCliente.Caption = txtCliente.Caption & " " & sDescrip

                   Set RsTemporal = Lib.OpenRecordset("Select * from TDELIVERY where tcodigoDelivery='" & sCliente & "'", Cn)
                   If Not RsTemporal.EOF Then
                    If lClub Then
                        txtTelefono.Caption = IIf(IsNull(RsTemporal!tCodigoDelivery), "", RsTemporal!tCodigoDelivery)
                    Else
                        txtTelefono.Caption = IIf(IsNull(RsTemporal!tTelefono), "", RsTemporal!tTelefono)
                    End If
                        txtDireccion.Caption = IIf(IsNull(RsTemporal!tDireccion), "", RsTemporal!tDireccion)
                        xDescuento = IIf(IsNull(RsTemporal!nDescuento), 0, RsTemporal!nDescuento)
                        sCodigoDescuento = "000"
                   Else
                        txtTelefono.Caption = ""
                        txtDireccion.Caption = ""
                        xDescuento = 0
                        sCodigoDescuento = ""
                   End If
                   
                   If sTienda <> "" Then
                      Set RsTemporal = Lib.OpenRecordset("Select * from TTIENDA where tCodigoDelivery='" & sCliente & "' and tCodigoTienda='" & sTienda & "'", Cn)
                      If Not RsTemporal.EOF Then
                         txtCliente.Caption = txtCliente.Caption + " - " + IIf(IsNull(RsTemporal!tNombre), "", RsTemporal!tNombre)
                         txtDireccion.Caption = IIf(IsNull(RsTemporal!tDireccion), "", RsTemporal!tDireccion)
                      
                         xDescuento = IIf(IsNull(RsTemporal!nDescuento), 0, RsTemporal!nDescuento)
                         sCodigoDescuento = "000"
                      Else
                         txtCliente.Caption = ""
                         txtDireccion.Caption = ""
                         xDescuento = 0
                         sCodigoDescuento = ""
                      End If
                   End If
                   
                   'invitado2013
                   If sCodigoInvitado <> "" Then
                      Set RsTemporal = Lib.OpenRecordset("Select isnull(tnombre,'')   + ' ' + isnull(tapellido,'') as invitado from tdeliveryinvitado where tCodigoinvitado='" & sCodigoInvitado & "'", Cn)
                      If Not RsTemporal.EOF Then
                        txtObservacion.Caption = "Inv: " & IIf(IsNull(RsTemporal!Invitado), "", RsTemporal!Invitado)
                         
                      Else
                         txtObservacion.Caption = ""
                      End If
                   End If
                   'invitado2013
                   
                   If xDescuento > 0 Then
                      sCodigoDescuento = "000"
                      sDescripcionDescuento = "DESCUENTO POR CLIENTE DELIVERY FRECUENTE"
                   Else
                      sCodigoDescuento = ""
                      sDescripcionDescuento = ""
                   End If
                   wCalcula = True
                   
                End If
                
           Case Is = 13 'Mesa
                sTipo = "A"
                frmMesa.cmdOpcion(6).Enabled = False
                frmMesa.Show vbModal
                If wMesa Then
                   sEstado = Calcular("select tEstadoMesa as Codigo from TMESA where tCodigoMesa ='" & sCodigo & "'", Cn)
                   If sEstado = "01" Or sEstado = "04" Then
                      sMesa = sCodigo
                   Else
                      MsgBox "Mesa Ocupada", vbExclamation, sMensaje
                      Exit Sub
                   End If
                End If
                
           Case Is = 14
           
           Case Is = 15 'Descuento
                If Supervisor("10") = False Then
                   MsgBox "Clave no permitida", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                sUsuarioAutoriza = sVar1
                tAutorizaDescuento = sUsuarioAutoriza
                sTemp = ""
                Isql = "SELECT Codigo, LTRIM(RTRIM(Descripcion)) as Descripcion, case lRatio when 1 then nRatio else 0 END as nRatio, case lRatio when 0 then nRatio else 0 END as nMonto FROM vMotivoDescuento WHERE lActivo = '1' ORDER BY Descripcion"
                Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1500, 2, 0, "", _
                                                                "Descripción", 2, "Descripcion", 4300, 0, 0, "", _
                                                                "Porcentaje", 2, "nRatio", 1200, 1, 0, "###,##0.00", _
                                                                "Monto", 2, "nMonto", 1200, 1, 0, "###,##0.00")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                wCalcula = True
                
                If wEnter Then
                   sCodigoDescuento = sCodigo
                   Dim RsDesc As ADODB.Recordset
                   Set RsDesc = Lib.OpenRecordset("select * from vMotivoDescuento where Codigo = '" & sCodigo & "'", Cn)
                   If RsDesc.EOF Then
                      Exit Sub
                   End If
                   RsDesc.MoveFirst
                   xDescuento = RsDesc!nRatio
                   lRatio = RsDesc!lRatio
                   
                   Dim SumTotalPedido As Double
                   SumTotalPedido = Calcular("select sum(d.nventa) as codigo from DPEDIDO d inner join TPRODUCTO p on d.tCodigoProducto = p.tCodigoProducto where d.tCodigoPedido='" & sPedido & "' and p.lDescuento = 1", Cn)
                   
                   If Not RsDesc!lRatio And (RsDesc!nRatio > CDbl(txtSuma.Caption)) Then
                       sCodigoDescuento = ""
                       xDescuento = 0
                       MsgBox "Descuento mayor al Pedido", vbCritical, sMensaje
                       Exit Sub
                   End If
                   
                   If sCodigo = "000" Then
                      sTipo = "Prepintado"
                      frmNumPad.Show vbModal
                      xDescuento = Val(sDescrip)

                      If (RsDesc!lRatio) Then
                            If xDescuento > 99 Then
                                    MsgBox "Descuento Incorrecto"
                                    Exit Sub
                            End If
                      Else
                            If xDescuento > SumTotalPedido Then
                                    MsgBox "Descuento mayor al Pedido", vbCritical, sMensaje
                                    Exit Sub
                            End If
                      End If
                      
                      If xDescuento > 99 Then
                        xDescuento = 0
                        MsgBox "Porcentaje no válido", vbCritical, sMensaje
                        Exit Sub
                      End If
                      
                      
                      frmKeyBoard.Caption = "Descripcion del Descuento"
                      frmKeyBoard.Show vbModal
                      sDescripcionDescuento = sDescrip
                      
                   Else
                      If RsDesc!lBloqueo Then
                         sTipo = "Prepintado"
                         sCodigo = xDescuento
                         frmNumPad.Show vbModal
                         
                            'Cambio de monto o % Descuento
                            xDescuento = Val(sDescrip)
                            
                            If (RsDesc!lRatio) Then
                                If xDescuento > 99 Then
                                    MsgBox "Descuento Incorrecto"
                                    Exit Sub
                                End If
                            Else
                                If xDescuento > SumTotalPedido Then
                                    MsgBox "Descuento mayor al Pedido", vbCritical, sMensaje
                                    Exit Sub
                                End If
                            End If
                      Else
                         wEnter = False
                      End If
                   End If
                Else
                   Exit Sub
                End If
                
                
            Case Is = 19 ' entregar A
                frmKeyBoard.Caption = "Entregar A"
                frmKeyBoard.txtResultado.Text = txtEntregar.Caption
                frmKeyBoard.Show vbModal
                If wEnter = True Then
                    txtEntregar.Caption = sDescrip
                End If
           Case Is = 21 ' monto maximo pedido
                If lcover Then
                    If Val(frmVenta.txtAdulto.Text) = 0 Then
                       MsgBox ("#Pax del pedido no ingresado!"), vbExclamation
                       Exit Sub
                    End If
                End If
                sTipo = ""
                frmNumPad.Show vbModal
                xMontoMaximo = Val(sDescrip)
'                Dim nAdultoPed As Integer
'                nAdultoPed = Val(Calcular("select nAdulto as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
'
                If lcover Then
                    If (sMontoMinCover * Val(frmVenta.txtAdulto.Text)) > xMontoMaximo Then
                        MsgBox "El Monto Minimo por Pedido es Mayor al Monto Maximo Ingresado!", vbInformation
                        xMontoMaximo = 0
                    End If
                End If
                'cmdCabecera(20).FontBold = IIf(cmdCabecera(20).FontBold = True, False, True)
                cmdCabecera(21).Caption = "Monto Max/P : " & CStr(IIf(xMontoMaximo = 0, 0, xMontoMaximo))
                                            
    End Select
End Sub

Private Sub cmdCombo_Click(Index As Integer)
   Dim nPos As Integer
   Dim nOrd As Integer
   frmMozoUsuario.ReseteaTimer
   
   Select Case Index
          Case Is = 0 ' Salir
               fraCombo.Visible = False
               fraProductoCombo.Visible = False
               wAgregaCombo = False
               ActivaCabecera True
               ActivaOpcion True
               AsignaProducto
               If fraPropiedad.Visible = True Then
                  cmdOpcion_Click (8)
               End If
                RsDetalle.Requery
                ActivaMesa247 (True)
          Case Is = 1 ' Elimina
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               If obtieneEliminaItemFijoCombo(RsCombo.Fields("tproducto"), RsCombo.Fields("tproductocombo")) = True Then
                    MsgBox "No se puede quitar este producto del combo. Consulte con el Administrador"
                    Exit Sub
               End If
               
                If MsgBox("Seguro de Eliminar el Producto?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                  sUsuarioAutoriza = sUsuario
                  If (lPassword And lPrinter = False) Or (lPassword And lPrinter And RsCombo!lImprime) Then
                     If Supervisor("02") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                     End If
                     sUsuarioAutoriza = sVar1
                  End If
                                                                            
                  If lPrinter = False Or (lPrinter = True And RsCombo!lImprime) Then
                     'Impresion del Pedidos Anulados
                     Isql = "SELECT TPRODUCTO_1.tDetallado AS Producto, dbo.vSalon.tResumido + ' - ' + dbo.TMESA.tResumido AS Mesa, dbo.TPRODUCTOAREA.tArea, dbo.MPEDIDO.tTipoPedido AS TipoPedido, dbo.MPEDIDO.nAdulto, dbo.MPEDIDO.lPrioridad AS Prioridad, dbo.MPEDIDO.tObservacion AS Observacion, dbo.vMozo.Descripcion AS Mozo, dbo.CPEDIDO.nCantidad AS nCombo, dbo.CPEDIDO.tItem, dbo.CPEDIDO.tItemCombo, dbo.CPEDIDO.tObservacion AS tObservacionCombo, TPRODUCTO_2.tDetallado AS Combo, dbo.vDelivery.Cliente " & _
                            "FROM dbo.TPRODUCTO TPRODUCTO_2 LEFT OUTER JOIN dbo.TPRODUCTOAREA ON TPRODUCTO_2.tCodigoProducto = dbo.TPRODUCTOAREA.tCodigoProducto RIGHT OUTER JOIN dbo.TMESA LEFT OUTER JOIN dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo RIGHT OUTER JOIN dbo.vMozo RIGHT OUTER JOIN dbo.vDelivery RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vDelivery.Codigo = dbo.MPEDIDO.tClienteDelivery LEFT OUTER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo ON " & _
                            "dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.tMesa ON TPRODUCTO_2.tCodigoProducto = dbo.CPEDIDO.tProductoCombo LEFT OUTER JOIN dbo.TPRODUCTO TPRODUCTO_1 ON dbo.CPEDIDO.tProducto = TPRODUCTO_1.tCodigoProducto " & _
                            "Where dbo.CPEDIDO.lImprime = 1 And dbo.CPEDIDO.lImprimeArea = 1 and dbo.CPEDIDO.tCodigoPedido = '" & sPedido & "' and dbo.CPEDIDO.tItem ='" & RsCombo!tItem & "' and dbo.CPEDIDO.tItemCombo='" & RsCombo!tItemCombo & "' ORDER BY dbo.CPEDIDO.tItem"
              
                     Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                     Dim i As Integer
                     If RsImpresion.RecordCount = 0 Then
                        LimpiaRs
                     Else
                        RsArea.MoveFirst
                        For i = 1 To RsArea.RecordCount
                            RsImpresion.Filter = "tArea = '" & RsArea!tArea & "'"
                            If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                               If RsImpresion.RecordCount <> 0 Then
                                  RsImpresion.MoveFirst
                                  ImprimeAnulaCombo RsImpresion, RsArea!timpresora, RsArea!Area, RsComboPropiedad
                               End If
                            End If
                            RsArea.MoveNext
                        Next i
                     End If
                  End If
                  'Oscar Ortega----------------------------------------------
                  Dim RstCombo As Recordset
                  Isql = "Select c.nCantidad, t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & xItem & "'"
                  Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                  If RstCombo.RecordCount > 0 Then
                    If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                        txtSuma.Caption = CambiaPrecio(nPVenta - ((RstCombo!nAumento / nCantidad) * RstCombo!nCantidad), sPedido)
                    End If
                  End If
                  'Fin Oscar Ortega------------------------------------------
                  'KDS
                  If lKDS Then
                    Call KDS_EliminarProductoDeCombo(RsCabecera, sItem, xItem)
                  End If
                  
                          
                 'insumoCOMBO
                 'INSUMOCRITICO23
                    Dim rstItems As New ADODB.Recordset
                    Set rstItems = New ADODB.Recordset
                    Set rstItems = Lib.OpenRecordset("SELECT     dbo.TPRODUCTO.tCodigoInsumo, dbo.DPEDIDO.nCantidad * dbo.CPEDIDO.nCantidad AS nCantidad FROM         dbo.DPEDIDO INNER JOIN                       dbo.CPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.CPEDIDO.tItem INNER JOIN                       dbo.TPRODUCTO ON dbo.CPEDIDO.tProductoCombo = dbo.TPRODUCTO.tCodigoProducto WHERE    cpedido.tcodigopedido='" & sPedido & "'  and cpedido.tItem ='" & sItem & "' and tItemCombo='" & xItem & "' and  (dbo.TPRODUCTO.lControlInsumoCritico = 1) AND (ISNULL(dbo.TPRODUCTO.tCodigoInsumo, '') <> '') AND (ISNULL(dbo.CPEDIDO.lImprime, 0) = 1) ", Cn)
                
                    If Not (rstItems.EOF Or rstItems.BOF) Then
                        modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
                    End If
    
                   Dim cMax As String
                   cMax = Calcular("select max(tItem) as Codigo from APEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
                   cMax = Lib.Correlativo(cMax, 3)
                   Isql = "insert into APEDIDO (tCodigoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                          "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
                          "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, " & _
                          "tComanda, lImprime, tUsuario, fRegistro, tUsuarioAnulado, fRegistroAnulado, " & _
                          "tObservacion, tObservacionAnulado, tEstadoItem, lImprimeArea, tArea, tMotivoEliminacion, tTurnoAnulado,fDiaContable) " & _
                          "select '" & sPedido & "' as tCodigoPedido, '" & cMax & "' as tItem, cpedido.tProductocombo, cpedido.tCodigoGRupo, cpedido.tCodigoSubGrupo, " & _
                          "cpedido.nPrecioNeto, cpedido.nImpuesto1/cpedido.ncantidad, cpedido.nImpuesto2/cpedido.ncantidad, cpedido.nImpuesto3/cpedido.ncantidad, cpedido.nVenta/cpedido.ncantidad, " & _
                          "0, cpedido.nPrecioNeto, cpedido.nCantidad, cpedido.nImpuesto1, cpedido.nImpuesto2, cpedido.nImpuesto3, cpedido.nVenta, '', cpedido.lImprime, " & _
                          "'" & RsCabecera!tUsuario & "' as tUsuario, dpedido.fregistro as fRegistro, " & _
                          "'" & sUsuarioAutoriza & "' as tUsuarioAnulado, getDate() as fRegistroAnulado, " & _
                          "'Anulado de Combo' as tObservacion, 'Anul. de Combo:' + t.tResumido as tObservacionAnulado, 'N', cpedido.lImprimeArea, '', '000', '" & sTurno & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
                          "from  dpedido inner join cpedido on dpedido.tCodigoPedido=cpedido.tcodigopedido and  dpedido.tItem = CPEDIDO.tItem inner join tproducto t on t.tcodigoproducto = dpedido.tCodigoProducto " & _
                          "where cpedido.tCodigoPedido = '" & sPedido & "' and cpedido.tItem ='" & sItem & "' and cpedido.tItemCombo='" & xItem & "'"
                   Cn.Execute Isql
                   
                  
                  Cn.Execute "delete from CPEDIDO where tCodigoPedido ='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & xItem & "'"
                  Cn.Execute "delete from TCOMBOPROPIEDAD where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & xItem & "'"
                  RsComboPropiedad.Requery
                  RsCombo.Requery
                  If RsCombo.RecordCount > 0 Then
                     RsCombo.MoveLast
                  End If

               End If
   
          Case Is = 2 'Aumentar
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                                 "FROM dbo.CPEDIDO WHERE CPEDIDO.tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'", Cn)
               If nCCombo + 1 > nCombo * RsDetalle!nCantidad Then
                  MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
                  Exit Sub
               End If
               'Oscar Ortega----------------------------------------------
                Dim oRsProductoDeCombo As Recordset
                Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sCombo)
                If oRsProductoDeCombo.RecordCount > 0 Then
                    If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                         Dim nCantidadEnElCombo As Integer
                         nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sPedido, sItem, oRsProductoDeCombo!tEtiqueta)
                         If nCantidadEnElCombo >= RsDetalle!nCantidad Then
                             MsgBox "Solo es permitido " & nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                             Exit Sub
                         End If
                    End If
                End If
                '----------------------------------------------------------
               'If lcover Then
                    Isql = "Select t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
                    Dim RstCombo2 As Recordset
                    Set RstCombo2 = Lib.OpenRecordset(Isql, Cn)
                    If RstCombo2.RecordCount > 0 Then
                            
                             Dim maxped As Double
                             maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
                                 If maxped > 0 Then
                                        If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + (RstCombo2!nAumento)) Then
                                          MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                                           RsCombo.Requery
                                           'RsCombo.AbsolutePosition = nPos
                                          Exit Sub
                
                                        End If
                                 End If
                            
                    End If
                
                ' End If
                
                '--------------------------
                nPos = RsCombo.AbsolutePosition
               Cn.Execute "update CPEDIDO set nCantidad = " & RsCombo!nCantidad + 1 & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               'Oscar Ortega------------------------------------------------------------
               Isql = "Select t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
               Set RstCombo = Lib.OpenRecordset(Isql, Cn)
               If RstCombo.RecordCount > 0 Then
                    If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                        txtSuma.Caption = CambiaPrecio(nPVenta + RstCombo!nAumento / nCantidad, sPedido)
                    End If
               End If
               'Fin Oscar Ortega--------------------------------------------------------
               RsCombo.Requery
               RsCombo.AbsolutePosition = nPos
               
          Case Is = 3 'Disminuir
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
          
               If grdCombo.Columns(2).Text > 1 Then
                  nPos = RsCombo.AbsolutePosition
                  Cn.Execute "update CPEDIDO set nCantidad = " & RsCombo!nCantidad - 1 & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                  'Oscar Ortega------------------------------------------------------------
                  Isql = "Select t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
                  Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                  If RstCombo.RecordCount > 0 Then
                    If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                      txtSuma.Caption = CambiaPrecio(nPVenta - RstCombo!nAumento / nCantidad, sPedido)
                    End If
                  End If
                  'Fin Oscar Ortega--------------------------------------------------------
                  RsCombo.Requery
                  RsCombo.AbsolutePosition = nPos
               End If
               
          Case Is = 4 'Cantidad
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter And Val(sDescrip) > 0 Then
                  nCantidad = Val(sDescrip)
                  
                  nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                                    "FROM dbo.CPEDIDO WHERE CPEDIDO.tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'", Cn)
                  If nCCombo + nCantidad - RsCombo!nCantidad > nCombo * RsDetalle!nCantidad Then
                     MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
                     nCantidad = 1
                     Exit Sub
                  End If
                  'Oscar Ortega----------------------------------------------
                  Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sCombo)
                  If oRsProductoDeCombo.RecordCount > 0 Then
                     If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                         nCantidadEnElCombo = ObtenerSumaCantidadesEnElComboSinEsteProducto(sPedido, sItem, xItem, oRsProductoDeCombo!tEtiqueta)
                         If nCantidad > RsDetalle!nCantidad - nCantidadEnElCombo Then
                             MsgBox "Solo es permitido " & RsDetalle!nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                             nCantidad = 1
                             Exit Sub
                         End If
                     End If
                  End If
                  '----------------------------------------------------------
                    nCantidad = RsDetalle!nCantidad '-----no tocar
                    
                ' If lcover Then
                      Isql = "Select c.nCantidad, t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & xItem & "'"
                    'Dim RstCombo As Recordset
                    Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                    If RstCombo.RecordCount > 0 And RstCombo!nCantidad < Val(sDescrip) Then
                            
                             Dim maxped2 As Double
                             maxped2 = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
                                 If maxped2 > 0 Then
                                        If maxped2 < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + (RstCombo!nAumento * (CDbl(sDescrip) - RstCombo!nCantidad))) Then
                                          MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                                           RsCombo.Requery
                                           'RsCombo.AbsolutePosition = nPos
                                          Exit Sub
                
                                        End If
                                 End If
                             
                    End If
                  
                 ' End If
                  
                  
                  
                  'Oscar Ortega----------------------------------------------Cambia Precio
                
                  Isql = "Select c.nCantidad, t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & xItem & "'"
                  Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                  If RstCombo.RecordCount > 0 Then
                    If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                        txtSuma.Caption = CambiaPrecio(nPVenta - ((RstCombo!nAumento / nCantidad) * RstCombo!nCantidad), sPedido)
                        txtSuma.Caption = CambiaPrecio(nPVenta + ((RstCombo!nAumento / nCantidad) * Val(sDescrip)), sPedido)
                    End If
                  End If
                  'Fin Oscar Ortega------------------------------------------
                  nPos = RsDetalle.AbsolutePosition
                  Cn.Execute "update CPEDIDO set nCantidad = " & Val(sDescrip) & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                  RsCombo.Requery
                  RsCombo.AbsolutePosition = nPos
               End If
               
          Case Is = 5 'Propiedad Combos
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
          
               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               ActivaFrame fraPropiedad
   
           Case Is = 6  'Orden +
                If RsCombo.RecordCount = 0 Then
                   Exit Sub
                End If
           
                nPos = RsCombo.AbsolutePosition
                nOrd = IIf(IsNull(RsCombo!nOrden), 0, RsCombo!nOrden)
                Cn.Execute "update CPEDIDO set nOrden = " & nOrd + 1 & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                RsCombo.Requery
                RsCombo.AbsolutePosition = nPos
                
           Case Is = 7  'Orden -
                If RsCombo.RecordCount = 0 Then
                   Exit Sub
                End If
           
                nPos = RsCombo.AbsolutePosition
                nOrd = IIf(IsNull(RsCombo!nOrden), 0, RsCombo!nOrden)
                If nOrd > 1 Then
                   Cn.Execute "update CPEDIDO set nOrden = " & nOrd - 1 & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                   RsCombo.Requery
                   RsCombo.AbsolutePosition = nPos
                End If
          
          Case Is = 8  'Linea Corte
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
           
               nPos = RsCombo.AbsolutePosition
               If IIf(IsNull(RsCombo!lCorte), False, RsCombo!lCorte) Then
                  Cn.Execute "update CPEDIDO set lCorte = 0 where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               Else
                  Cn.Execute "update CPEDIDO set lCorte = 1 where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               End If
               RsCombo.Requery
               RsCombo.AbsolutePosition = nPos
   End Select

End Sub

Private Sub cmdDetalle_Click(Index As Integer)
   Dim nPos As Integer
   
   Dim maxped As Double '---erick
   
   frmMozoUsuario.ReseteaTimer
   If RsDetalle.RecordCount = 0 Then
      Exit Sub
   End If
   If RsDetalle!tFacturado = "F" Or RsDetalle!tFacturado = "P" Then
      MsgBox "Producto ya Facturado ", vbExclamation, sMensaje
      Exit Sub
   End If
   
    If wAgregaCombo = False Then
    'VALIDA ENVIO DE PRODUCTO MOBILE
    Dim xImprimeProducto As Boolean
    xImprimeProducto = Calcular("Select lImprime as codigo From DPEDIDO Where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'", Cn)
    
    If Index <> 0 Then
        If xImprimeProducto = True Then
           MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
           Exit Sub
        End If
    End If
    End If
   '-------------------------------

   Select Case Index
          Case Is = 0 ' Elimina

                nombreMozo = sUsuario
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
                If lActivaAnticipo Then
                  If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
                      MsgBox "No es posible modificar/elimnar el producto, anticipo aplicado al pedido!!!", vbExclamation, sMensaje
                      Exit Sub
                  End If
                End If
               'GCAA 10/07/2023 - REST-012-C-2023
               If lBloqueoAnulaItemsPedido And Calcular("Select isnull(lImprime,0) as codigo From DPEDIDO Where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'", Cn) = True Then
                   MsgBox "Esta opcion se encuentra Bloqueada, favor de comunicar al Supervisor", vbExclamation, sMensaje
                   Exit Sub
               End If
                
               If MsgBox("Seguro de Eliminar el Producto?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                  sUsuarioAutoriza = sUsuario
                  
                  xImprimeProducto = Calcular("Select lImprime as codigo From DPEDIDO Where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'", Cn)

                  'If (lPassword And lPrinter = False) Or (lPassword And lPrinter And RsDetalle!lImprime) Then
                  If (lPassword And lPrinter = False) Or (lPassword And lPrinter And xImprimeProducto) Then
                     If Supervisor("02") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                     End If
                     sUsuarioAutoriza = sVar1
                     sUsuario = UCase(sVar1)
                  End If
                                                                            
                  'If (lElimina And lPrinter = False) Or (lElimina And lPrinter = True And RsDetalle!lImprime) Then
                  If (lElimina And lPrinter = False) Or (lElimina And lPrinter = True And xImprimeProducto) Then
                     wCabecera = False
                     ActivaFrame fraEliminacion
                     ActivaCabecera False
                     ActivaOpcion False
                     cmdOpcion(1).Enabled = False
                     cmdOpcion(2).Enabled = False
                  Else
                     sCodigo = ""
                     sDescrip = ""
                     'oo
                     'EliminaItem
                     'KDS
                     If lKDS Then
                         Call KDS_EliminarProducto(RsCabecera, sItem, 0)
                     End If
                     EliminaItem
                  End If
                  
                  'MESA247
                  If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                        Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                  End If
                       
                  
              End If
                Exit Sub
          Case Is = 1 ' Cantidad
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
                If lActivaAnticipo Then
                  If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
                      MsgBox "No es posible modificar/elimnar el producto, anticipo aplicado al pedido!!!", vbExclamation, sMensaje
                      Exit Sub
                  End If
                End If
                If Calcular("select count(*) as codigo FROM cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where cpedido.tCodigoPedido='" & sPedido & "' and titem='" & sItem & "' and isnull(tcombo.lfijo,0)=0 and isnull(tcombo.naumento,0)>0", Cn) > 0 Then
                    MsgBox ("No es posible modificar la cantidad!!! " & vbNewLine & "Hay al menos 1 item combo con aumenta precio no fijo que restringue el proceso!!!"), vbInformation
                    Exit Sub
                End If
               
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter And Val(sDescrip) > 0 Then
                    'Oscar Ortega------------
                    Dim oRsDetalleProducto As Recordset
                    Set oRsDetalleProducto = ObtenerDetalleProducto(sPedido, sItem)
                    
                    
                      'If lcover Then
                        maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
                            If maxped > 0 Then
                                If (nCantidad < CDbl(sDescrip)) Then ' Sumar
                                       If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + CDbl((CDbl(sDescrip) - nCantidad) * nPVenta)) Then
                                         MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                                         
        '                                 Screen.MousePointer = vbDefault
        '                                 RsDetalle.MoveLast
        '                                 RsDetalle.Requery
                                         'RsProducto.MoveFirst
                                         Exit Sub
                    
                                       End If
                                Else ' restar
'                                        If maxped < (Val(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + Val((nCantidad - Val(sDescrip)) * nPVenta)) Then
'                                         MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
'
'        '                                 Screen.MousePointer = vbDefault
'        '                                 RsDetalle.MoveLast
'        '                                 RsDetalle.Requery
'                                         'RsProducto.MoveFirst
'                                         Exit Sub
'
'                                       End If
                                End If
                            End If
                     'End If
                    
                    
                    If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                            If verificaCantidadDeItemsCombos(sPedido, sItem, oRsDetalleProducto!nCombinacion, Val(sDescrip)) = False Then
                                MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                                Exit Sub
                            End If
                            Dim Xse As Integer
                            If (nCantidad < Val(sDescrip)) Then ' Sumar
                                Xse = Val(sDescrip) - nCantidad
                                Dim i As Integer
                                For i = 1 To Xse
                                    nPos = RsDetalle.AbsolutePosition
                                    nCantidad = nCantidad + 1
                                    Set oRsDetalleProducto = ObtenerDetalleProducto(sPedido, sItem)
                                    If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                                        'txtSuma.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad - 1), sPedido)
                                        Dim AAcombo As String
                                        AAcombo = CambiaPrecioCombo(((oRsDetalleProducto!nVenta - oRsDetalleProducto!nDescuento) + oRsDetalleProducto!nPrecioOficial) / nCantidad, sPedido)
                                    Else
                                        Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
                                    End If
                                Next i
                            Else 'Restar
                                Dim nCantidadMax As Double
                                nCantidadMax = Obtener_CantidadMaximaDeUnicoEtiqueta(sPedido, sItem, nCantidad)
                                
                                If nCantidad > nCantidadMax Then

                                Xse = nCantidad - Val(sDescrip)
                                For i = 1 To Xse
                                    Set oRsDetalleProducto = ObtenerDetalleProducto(sPedido, sItem)
                                    If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                                        nCantidad = nCantidad - 1
                                        If nCantidad = 0 Then
                                            MsgBox "NO es posible Hacer este proceso para Combos!!!!"
                                            Exit Sub
                                        End If
                                       'txtSuma.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad + 1), sPedido)
                                       'Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sitem & "'"
                                        Dim DDcombo As String
                                        DDcombo = CambiaPrecioCombo((oRsDetalleProducto!nVenta - (oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento)) / nCantidad, sPedido)
                                    Else
                                      Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
                                    End If
                                Next i
                               Else
                                MsgBox ("No puedes reducir la cantidad de combos con elementos únicos" & Chr(13) & "Disminuya primero la cantidad de productos dentro del combo"), vbExclamation
                                Exit Sub
                                End If
                            End If
                            'Cambiar Combo
                            nCantidad = Val(sDescrip)
                            nPos = RsDetalle.AbsolutePosition
                            oRsDetalleProducto.Requery
                            RsDetalle.Requery
 
                    Else
                        nCantidad = Val(sDescrip)
                        nPos = RsDetalle.AbsolutePosition
                        Cn.Execute "update DPEDIDO set nCantidad = " & Val(sDescrip) & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1 * " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2 * " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3 * " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
                    End If
                    '------------------------
                    RsDetalle.Requery
                    
                    RsDetalle.AbsolutePosition = nPos
                    fxCombo "M", nCantidad, sProducto
                    txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
                    verificatitulo
      
               End If
               
          Case Is = 2 ' Aumentar
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
                If lActivaAnticipo Then
                  If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
                      MsgBox "No es posible modificar/elimnar el producto, anticipo aplicado al pedido!!!", vbExclamation, sMensaje
                      Exit Sub
                  End If
                End If
               'GCAA
                 If Calcular("select count(*) as codigo FROM cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where cpedido.tCodigoPedido='" & sPedido & "' and titem='" & sItem & "' and isnull(tcombo.lfijo,0)=0 and isnull(tcombo.naumento,0)>0 ", Cn) > 0 Then
                    MsgBox ("No es posible modificar la cantidad!!! " & vbNewLine & "Hay al menos 1 item combo con aumenta precio no fijo que restringue el proceso!!!"), vbInformation
                    Exit Sub
                End If
                'If lcover Then
                    
                    maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
                        If maxped > 0 Then
                               If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + CDbl(nPVenta)) Then
                                 MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                                 
'                                 Screen.MousePointer = vbDefault
'                                 RsDetalle.MoveLast
'                                 RsDetalle.Requery
                                 'RsProducto.MoveFirst
                                 Exit Sub
            
                               End If
                        End If
                'End If
               
               
               
               nPos = RsDetalle.AbsolutePosition
               nCantidad = nCantidad + 1
               'Oscar Ortega Aumenta Combo---------------
               Set oRsDetalleProducto = ObtenerDetalleProducto(sPedido, sItem)
               If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                  If verificaCantidadDeItemsCombos(sPedido, sItem, oRsDetalleProducto!nCombinacion, nCantidad) = False Then
                     MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                     Exit Sub
                  End If
                    'txtSuma.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad - 1), sPedido)
                    'Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sitem & "'"
                    Dim Acombo As String
                    Acombo = CambiaPrecioCombo(((oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento) + oRsDetalleProducto!nVenta) / nCantidad, sPedido)
                  Else
                  Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
               End If
               '-----------------------------------------

               RsDetalle.Requery
               RsDetalle.AbsolutePosition = nPos
               fxCombo "M", nCantidad, sProducto
               txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
               verificatitulo
               

          Case Is = 3 ' Disminuir
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
                If lActivaAnticipo Then
                  If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
                      MsgBox "No es posible modificar/elimnar el producto, anticipo aplicado al pedido!!!", vbExclamation, sMensaje
                      Exit Sub
                  End If
                End If
               'GCAA
                If Calcular("select count(*) as codigo FROM cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where cpedido.tCodigoPedido='" & sPedido & "' and titem='" & sItem & "' and isnull(tcombo.lfijo,0)=0 and isnull(tcombo.naumento,0)>0", Cn) > 0 Then
                    MsgBox ("No es posible modificar la cantidad!!! " & vbNewLine & "Hay al menos 1 item combo con aumenta precio no fijo que restringue el proceso!!!"), vbInformation
                    Exit Sub
                End If
               If grdDetalle.Columns(4).Text > 1 Then
                  nCantidad = nCantidad - 1
                  nPos = RsDetalle.AbsolutePosition
                  'Oscar Ortega Disminuir Combo---------------
                  Set oRsDetalleProducto = ObtenerDetalleProducto(sPedido, sItem)
                  If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                        If verificaCantidadDeItemsCombos(sPedido, sItem, oRsDetalleProducto!nCombinacion, nCantidad) = False Then
                            MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                            nCantidad = nCantidad + 1 'oo03
                              Exit Sub
                       End If
                    'Dim nCantidadMax As Double
                    nCantidadMax = Obtener_CantidadMaximaDeUnicoEtiqueta(sPedido, sItem, nCantidad + 1)
                    If nCantidad + 1 > nCantidadMax Then
                       'Disminuir Combo
                       'txtSuma.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad + 1), sPedido)
                       'Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sitem & "'"
                        Dim Dcombo As String
                        Dcombo = CambiaPrecioCombo((oRsDetalleProducto!nVenta - (oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento)) / nCantidad, sPedido)
                    Else
                        MsgBox ("No puedes reducir la cantidad de combos con elementos únicos" & Chr(13) & "Disminuya primero la cantidad de productos dentro del combo"), vbExclamation
                        nCantidad = nCantidad + 1
                        Exit Sub
                    End If
                  Else
                    Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
                  End If
                  RsDetalle.Requery
                  RsDetalle.AbsolutePosition = nPos
                  fxCombo "M", nCantidad, sProducto
                  txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
                  verificatitulo
               End If
               
                    
          Case Is = 4 ' Propiedad
          
               If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
               End If
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
                If lActivaAnticipo Then
                  If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
                      MsgBox "No es posible modificar/elimnar el producto, anticipo aplicado al pedido!!!", vbExclamation, sMensaje
                      Exit Sub
                  End If
                End If
               'Oscar Ortega------------------------------
               If Obtener_PropiedadesPendientesPorItems(sProducto, sItem) = False Then
                    grdDetalle.Enabled = False
                    grdCabecera.Enabled = False
                Else
                    grdDetalle.Enabled = True
                    grdCabecera.Enabled = True
               End If
               'Fin Oscar Ortega--------------------------
               ActivaFrame fraPropiedad
               ActivaCabecera False
               ActivaOpcion False
               ActivaMesa247 (False)
              ' ActivaMesa247 (True)
          Case Is = 5 ' Observacion
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
               If wAgregaCombo Then
                  If RsCombo!lImprime = True Then
                     MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                     Exit Sub
                  End If
               Else
                  If RsDetalle!lImprime = True Then
                     MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                     Exit Sub
                  End If
               End If
          
               frmKeyBoard.Caption = "Observación del Producto"
               frmKeyBoard.txtResultado.Text = lblObservacion.Text
               frmKeyBoard.Show vbModal
               If wEnter = True Then
                  If wAgregaCombo Then
                     nPos = RsCombo.AbsolutePosition
                     Cn.Execute "Update CPEDIDO set tObservacion = '" & sDescrip & "' where tCodigoPedido ='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & xItem & "'"
                  Else
                        'VALIDA ENVIO DE PRODUCTO MOBILE
                        Dim xImprimeProducto1 As Boolean
                        xImprimeProducto1 = Calcular("Select lImprime as codigo From DPEDIDO Where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'", Cn)
                        
                        If xImprimeProducto1 = True Then
                           MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                           cmdOpcion_Click (8)
                           Exit Sub
                        End If
  
                     nPos = RsDetalle.AbsolutePosition
                     Cn.Execute "Update DPEDIDO set tObservacion = '" & sDescrip & "' where tCodigoPedido ='" & sPedido & "' and tItem ='" & sItem & "'"
                  End If
                  lblObservacion.Text = sDescrip
               End If
                         
          Case Is = 6 ' Linea
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
               
               nPos = RsDetalle.AbsolutePosition
               If IIf(IsNull(RsDetalle!lCorte), False, RsDetalle!lCorte) Then
                  Cn.Execute "update DPEDIDO set lCorte = 0 where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
               Else
                  Cn.Execute "update DPEDIDO set lCorte = 1 where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
               End If
               RsDetalle.Requery
               RsDetalle.AbsolutePosition = nPos

          Case Is = 7  'Orden +
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               nPos = RsDetalle.AbsolutePosition
               nOrden = nOrden + 1
               Cn.Execute "update DPEDIDO set nOrden = " & nOrden & " where tCodigoPedido='" & sPedido & "' and tItem ='" & RsDetalle!tItem & "'"
               
               RsDetalle.Requery
               RsDetalle.AbsolutePosition = nPos
              
          Case Is = 8  'Orden -
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               nPos = RsDetalle.AbsolutePosition
               
               If nOrden > 1 Then
                  If lOrden Then
                     If Calcular("select max(nOrden) as Codigo from DPEDIDO where tCodigoPedido='" & sPedido & "' and lImprime=1", Cn) >= nOrden - 1 Then
                        If MsgBox("Deseas Juntar el Producto a la Orden " & nOrden - 1 & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                           Exit Sub
                        End If
                     End If
                  End If
                  nOrden = nOrden - 1
                  Cn.Execute "update DPEDIDO set nOrden = " & nOrden & " where tCodigoPedido='" & sPedido & "' and tItem ='" & RsDetalle!tItem & "'"
                  RsDetalle.Requery
                  RsDetalle.AbsolutePosition = nPos
               End If
               
         
            End Select
          ActivaMesa247 (False)
End Sub

Private Sub cmdEliminacion_Click(Index As Integer)

    RsMotivoEliminacion.MoveFirst
    RsMotivoEliminacion.Find ("Descripcion = '" & cmdEliminacion(Index).Caption & "'")
    
    If RsMotivoEliminacion.EOF Then
       RsMotivoEliminacion.MoveFirst
    End If
    
    If RsMotivoEliminacion!codigo = "000" Then
       frmKeyBoard.txtResultado = ""
       frmKeyBoard.Show vbModal
       If Not wEnter Then
          Exit Sub
       End If
       sCodigo = "000"
    Else
       sCodigo = RsMotivoEliminacion!codigo
       sDescrip = ""
    End If
    
    If wCabecera Then
       EliminaCabecera
    Else
        'KDS
        If lKDS Then
            Call KDS_EliminarProducto(RsCabecera, sItem, 0)
        End If
       EliminaItem
    End If
        
    fraEliminacion.Visible = False
    wDetalle = True
    wAgrega = False
    wAgregaCombo = False
End Sub

Private Sub cmdGrupo_Click(Index As Integer)
   Dim i As Integer
   frmMozoUsuario.ReseteaTimer
   Screen.MousePointer = vbHourglass
   For i = 1 To 5
       cmdGrupo(i).backColor = vbButtonFace
       cmdSubGrupo(i).backColor = vbButtonFace
   Next i
   cmdSubGrupo(6).backColor = vbButtonFace
   cmdSubGrupo(7).backColor = vbButtonFace
   
   RsGrupo.MoveFirst
   RsGrupo.Find "nboton = " & Trim(str(Index))
   sGrupo = RsGrupo!codigo
   xGrupo = RsGrupo!codigo
   cmdGrupo(RsGrupo!nBoton).backColor = vbBlue
   RsSubgrupo.Filter = "tGrupo = '" & sGrupo & "'"
   AsignaBoton 7, RsSubgrupo, cmdSubGrupo()
   
   If RsSubgrupo.RecordCount = 0 Then
      sSubGrupo = ""
   Else
      RsSubgrupo.MoveFirst
      RsSubgrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vSubGrupo where nBoton>0 and tGrupo = '" & sGrupo & "'", Cn) & "'"
      sSubGrupo = RsSubgrupo!codigo
      xSubGrupo = RsSubgrupo!codigo
      If RsSubgrupo!nBoton > 0 Then
         cmdSubGrupo(RsSubgrupo!nBoton).backColor = vbRed
      End If
 
      RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
      AsignaBotonProducto 20, RsProducto, cmdProducto(), sTipoPedido, sUnidadNegocio
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdMotorizado_Click(Index As Integer)
   RsMotorizado.MoveFirst
   RsMotorizado.Find "nboton = " & Trim(str(Index))
   txtMotorizado.Caption = RsMotorizado!descripcion
   sMotorizado = RsMotorizado!codigo
End Sub


Private Sub cmdNavegar_Click(Index As Integer)
    frmMozoUsuario.ReseteaTimer
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, grdCabecera
           Case Is = 1 'PgUp
                MoverPuntero pgup, grdCabecera
           Case Is = 2 'Previo
                MoverPuntero previo, grdCabecera
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, grdCabecera
           Case Is = 4 'PgDn
                MoverPuntero pgdn, grdCabecera
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, grdCabecera
           Case Is = 6 'Primero
                MoverPuntero Primero, grdDetalle
           Case Is = 7 'PgUp
              grdDetalle.MoveRelative -7
               If grdDetalle.BOF Then
                  grdDetalle.MoveFirst
               End If
           Case Is = 8 'Previo
                MoverPuntero previo, grdDetalle
           Case Is = 9 'Siguiente
                MoverPuntero siguiente, grdDetalle
           Case Is = 10 'PgDn
             grdDetalle.MoveRelative 7
               If grdDetalle.EOF Then
                  grdDetalle.MoveLast
               End If
            Case Is = 11 'Ultimo
                MoverPuntero Ultimo, grdDetalle
           Case Is = 12 'Primero
                MoverPuntero Primero, grdCombo
           Case Is = 13 'Previo
                MoverPuntero previo, grdCombo
           Case Is = 14 'Siguiente
                MoverPuntero siguiente, grdCombo
           Case Is = 15 'Ultimo
                MoverPuntero Ultimo, grdCombo
           Case Is = 16 'PgDn
                MoverPuntero pgdn, grdCombo
           Case Is = 17 'PgUp
                MoverPuntero pgup, grdCombo
                
    End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   frmMozoUsuario.ReseteaTimer
   cmdOpcion(1).Enabled = True
   Dim i As Integer
   Dim oComando As clsComando
   
   Select Case Index
                   
       Case Is = 0  'Exit

            If lEstadoBuscarPedido Then
                Cn.Execute "Update mpedido set tmesa='' where tcodigopedido='" & sPedido & "'"
                Cn.Execute "Update tmesa set testadomesa='04' where tcodigomesa='" & sMesa & "'"
                Call TerminaprocesoBarra
            Else
                registroAccesoAuditoria "S", tNombreMozo
                'auditoria
                Unload Me
            End If
       Case Is = 1  'Grabar
            If wDetalle Then
               ActivaOpcion True
               wDetalle = False
               ActivaFrame fraPlato
               ActivaCabecera True
               GrabaProducto
            Else
               'Obligatoriedad de Mozo
               If lObligaMozo Then
                  If lMCPV Then
                      sMozo = ObtenerCodigoMozo(sVar1)
                      If sMozo = "" Then
                         MsgBox "Configure un Cajero-Mesero", vbExclamation, sMensaje
                         Exit Sub
                      End If
                  Else
                      If sMozo = "" Or sMozo = "0000" Then
                         MsgBox "Asigne al " & Mesero, vbExclamation, sMensaje
                         Exit Sub
                      End If
                  End If
               End If
            
                'GCAA 12072021
                If lObligaOrigenVenta = True Then
                        'If vOrigenVentas = "" And sTipoPedido = "02" Then
                        ' Validamos que si no se ha seleccionado el origen de venta se agrege el defautl por caja, siempre y cuando sea el canal de venta por default
                        If sTipoPedidoPD = sTipoPedido And (vOrigenVentas = "" Or vOrigenVentas = Null) Then
                            vOrigenVentas = sDefaultOrigenVenta
                        End If
                        If vOrigenVentas = "" Or vOrigenVentas = Null Then
                          MsgBox "Debe seleccionar un origen de venta", vbExclamation, sMensaje
                          Call ActivarOrigenVentas
        '                  cmdCabecera_Click (9)
                          Exit Sub
                        End If
                   End If
            
            
               'Obligatoriedad de Mesa
'               If lObligaMesa And sMesa = "" And txtObservacion.Caption = "" Then
'                  MsgBox "Asigne una Mesa", vbExclamation, sMensaje
'                  cmdCabecera_Click (13)
'                  Exit Sub
'               End If

               'Obligatoriedad de Mesa
               If lObligaMesa Then 'And sMesa = "" And txtObservacion.Caption = "" Then
                    If sMesa = "" Then
                      MsgBox "Asigne una Mesa", vbExclamation, sMensaje
                      cmdCabecera_Click (13)
                      Exit Sub
                    End If
               End If
               
               
               'Obligatoriedad de Cliente Frecuente
               If sCliente = "" And lObligaClienteFrecuente Then
                  MsgBox "Asigne el Cliente Delivery", vbExclamation, sMensaje
                  cmdCabecera_Click (12)
                  Exit Sub
               End If
            
               'Obligatoriedad de Fecha de Entrega
               If Me.txtFechaProg.Caption = "" And lObligaFechaEntrega Then
                  MsgBox "Asigne la Fecha de Entrega", vbExclamation, sMensaje
                  cmdCabecera_Click (10)
                  Exit Sub
               End If
            
               
               'entregar A
               If lObligaEntregarA = True And Me.txtEntregar.Caption = "" Then
                  MsgBox "Asigne información en Entregar A", vbExclamation, sMensaje
                  cmdCabecera_Click (9)
                  Exit Sub
               End If
                           
            
               ActivaOpcion True
               ActivaFrame fraPlato
               ActivaCabecera True
               
               If wAgrega Then
                  'Actualiza el estado de la mesa
                  If wMesa Then
                     Isql = "Update TMESA set tEstadoMesa='02' where tCodigoMesa ='" & sMesa & "'"
                     Cn.Execute Isql
                  Else
                     sMesa = ""
                  End If
                  wMesa = False
                  wAgrega = False
                  
                  'Valida si hay datos a modificar, para guardar la modificación en el LOG_INFOREST
                  ValidaModifica
                  
                  Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("spIns_MPEDIDO", Cn) Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  
                  oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 7, sCliente
                  oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, sTipoPedido
                  oComando.CreateParameter "@lPrioridad", adBoolean, adParamInput, 1, IIf(txtTipoPedido.ForeColor = &H800000, 0, 1)
                  oComando.CreateParameter "@tTipoAtencion", adVarChar, adParamInput, 2, sTipoAtencion
                  oComando.CreateParameter "@tMesa", adVarChar, adParamInput, 3, sMesa
                  oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, sMozo
                  oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 4, sMotorizado
                  oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
                  oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 2, sSalon
                  oComando.CreateParameter "@tTurno", adVarChar, adParamInput, 10, sTurno
                  oComando.CreateParameter "@tObservacion", adVarChar, adParamInput, 250, txtObservacion.Caption
                  oComando.CreateParameter "@nTiempo", adInteger, adParamInput, 10, 0
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, Right(sUsuario, 15)
                  oComando.CreateParameter "@nAdulto", adInteger, adParamInput, 10, Val(frmVenta.txtAdulto.Text)
                  oComando.CreateParameter "@nNino", adInteger, adParamInput, 10, Val(frmVenta.txtNino.Text)
                  oComando.CreateParameter "@nMesa", adInteger, adParamInput, 10, Val(frmVenta.txtJuntar.Text)
                  oComando.CreateParameter "@tPuntoVenta", adVarChar, adParamInput, 2, sPuntoVenta
                  oComando.CreateParameter "@tHabitacion", adVarChar, adParamInput, 6, ""
                  oComando.CreateParameter "@tReserva", adVarChar, adParamInput, 6, ""
                  oComando.CreateParameter "@tPasajero", adVarChar, adParamInput, 50, ""
                  oComando.CreateParameter "@tCompania", adVarChar, adParamInput, 5, ""
                  oComando.CreateParameter "@tContacto", adVarChar, adParamInput, 4, ""
                  oComando.CreateParameter "@nDescuento", adDouble, adParamInput, 10, xDescuento
                  oComando.CreateParameter "@tDescuento", adVarChar, adParamInput, 3, sCodigoDescuento
                  oComando.CreateParameter "@tObservacionDescuento", adVarChar, adParamInput, 250, IIf(sCodigoDescuento = "000", sDescripcionDescuento, "")
                  oComando.CreateParameter "@tAutorizaDescuento", adVarChar, adParamInput, 15, Right(tAutorizaDescuento, 15)
                  oComando.CreateParameter "@nTiempoDelivery", adInteger, adParamInput, 10, nTiempoDelivery
                  oComando.CreateParameter "@tTienda", adVarChar, adParamInput, 3, sTienda
                  oComando.CreateParameter "@fDiaContable", adDate, adParamInput, 10, obtieneDiaContable
                  oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 10, IIf(txtFechaProg.Caption = "", Null, Format(txtFechaProg.Caption, "dd/MM/yyyy HH:nn"))
                  'invitado2013
                  oComando.CreateParameter "@tCodigoInvitado", adVarChar, adParamInput, 10, sCodigoInvitado
                  'invitado2013
                    'pariente2013
                  oComando.CreateParameter "@tCodigopariente", adVarChar, adParamInput, 7, sCodigoParienteSeleccionado
                  'pariente2013
                    'entregara
                  oComando.CreateParameter "@tEntregarA", adVarChar, adParamInput, 20, IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20))
                  
                  oComando.CreateParameter "@nTiempoAntesEnvio", adInteger, adParamInput, 10, Val(txtEnvioAntes.Text)
                  
                  oComando.CreateParameter "@nMontoMaximo", adInteger, adParamInput, 250, xMontoMaximo
                  
                  oComando.CreateParameter "@tPedido", adVarChar, adParamOutput, 10, sPedido
                  
                 'origen de ventas poner validacion
                  If vOrigenVentas = Null Or vOrigenVentas = "" Then
                    vOrigenVentas = "00"
                  End If
                  
                  oComando.CreateParameter "@codigoOrigenVentas", adVarChar, adParamInput, 2, vOrigenVentas
                  oComando.CreateParameter "@tTotemMesa", adVarChar, adParamInput, 10, ""
                  
                  If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  Else
                     sPedido = oComando.GetParameterValue("@tPedido")
                  End If
                                                      
                  RsCabecera.Requery
                  RsCabecera.MoveFirst
                  RsCabecera.Find "Codigo = '" & sPedido & "'"
                  
               Else
                    
                    If wMesa Then
                     Isql = "Update TMESA set tEstadoMesa='02' where tCodigoMesa ='" & sMesa & "'"
                     Cn.Execute Isql
                    End If
                    wMesa = False
                    wAgrega = False
                    
                    Set oComando = New clsComando
                    If Not oComando.CreateCmdSp("spUpd_MPEDIDO", Cn) Then
                       Set oComando = Nothing
                       Exit Sub
                    End If
                    
                    'Valida si hay datos a modificar, para guardar la modificación en el LOG_INFOREST
                    ValidaModifica
                    
                    If sCodigoDescuento <> "" Then
                       CalculaDescuento
                       RsDetalle.Requery
                       wCalcula = False
                    End If
                    
                    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 7, sCliente
                    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, sTipoPedido
                    oComando.CreateParameter "@lPrioridad", adBoolean, adParamInput, 1, IIf(txtTipoPedido.ForeColor = &H800000, 0, 1)
                    oComando.CreateParameter "@tTipoAtencion", adVarChar, adParamInput, 2, sTipoAtencion
                    oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, sMozo
                    oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 4, sMotorizado
                    oComando.CreateParameter "@tObservacion", adVarChar, adParamInput, 250, txtObservacion.Caption
                    oComando.CreateParameter "@nTiempo", adInteger, adParamInput, 10, 0
                    oComando.CreateParameter "@tPuntoVenta", adVarChar, adParamInput, 2, sPuntoVenta
                    oComando.CreateParameter "@tHabitacion", adVarChar, adParamInput, 6, ""
                    oComando.CreateParameter "@tReserva", adVarChar, adParamInput, 6, ""
                    oComando.CreateParameter "@tPasajero", adVarChar, adParamInput, 50, ""
                    oComando.CreateParameter "@tCompania", adVarChar, adParamInput, 5, ""
                    oComando.CreateParameter "@tContacto", adVarChar, adParamInput, 4, ""
                    oComando.CreateParameter "@nDescuento", adDouble, adParamInput, 10, xDescuento
                    oComando.CreateParameter "@tDescuento", adVarChar, adParamInput, 3, sCodigoDescuento
                    oComando.CreateParameter "@tObservacionDescuento", adVarChar, adParamInput, 250, IIf(sCodigoDescuento = "000", sDescripcionDescuento, "")
                    oComando.CreateParameter "@tAutorizaDescuento", adVarChar, adParamInput, 15, Mid(tAutorizaDescuento, 1, 15)
                    oComando.CreateParameter "@tTienda", adVarChar, adParamInput, 3, sTienda
                    'oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 10, Format(txtFechaProg.Caption, "dd/MM/yyyy HH:nn")
                    oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 20, IIf(txtFechaProg.Caption = "", Null, Format(txtFechaProg.Caption, "dd/MM/yyyy HH:mm"))
                  
                    'invitado2013
                    oComando.CreateParameter "@tCodigoInvitado", adVarChar, adParamInput, 10, sCodigoInvitado
                    'invitado2013
                    
                    'pariente2014
                    oComando.CreateParameter "@tCodigopariente", adVarChar, adParamInput, 7, sCodigoParienteSeleccionado
                    'pariente2013
                    
                    'entregara
                    oComando.CreateParameter "@tEntregarA", adVarChar, adParamInput, 20, IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20))
                  
                    oComando.CreateParameter "@nTiempoAntesEnvio", adInteger, adParamInput, 10, Val(txtEnvioAntes.Text)
                    
                    oComando.CreateParameter "@nMontoMaximo", adInteger, adParamInput, 250, xMontoMaximo
                    
                    oComando.CreateParameter "@tPedido", adVarChar, adParamInput, 10, sPedido
                    
                    
                    'origen de ventas poner validacion
                    If vOrigenVentas = Null Or vOrigenVentas = "" Then
                        vOrigenVentas = "00"
                    End If
                    
                    oComando.CreateParameter "@codigoOrigenVentas", adVarChar, adParamInput, 2, vOrigenVentas
                    oComando.CreateParameter "@tMesa", adVarChar, adParamInput, 3, sMesa
                    oComando.CreateParameter "@tTotemMesa", adVarChar, adParamInput, 10, ""
                    
                    If Not oComando.GetParamOK Then
                       Set oComando = Nothing
                       Exit Sub
                    End If
                    If Not oComando.ExecSP Then
                       Set oComando = Nothing
                       Exit Sub
                    End If
                    
                    RsCabecera.Requery
                    RsCabecera.Find "Codigo = '" & sPedido & "'"
                    
                    'MESA247
                    If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                          Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                    End If
               End If
            End If
            
            ActivaOpcion True
            grdCabecera.Enabled = True
            ActivaMesa247 (True)
       Case Is = 2  'Cancelar
            ActivaOpcion True
            wMesa = False
            
            'Juntar Mesa
            Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where ISNULL(tCodigoPedido,'')='')"
            Cn.Execute "delete from TPEDIDOMESA where tCodigoPedido=''"
            
            If wDetalle Then
               wDetalle = False
               ActivaFrame fraPlato
               ActivaCabecera True
            Else
               If RsCabecera.RecordCount = 0 Then
                  fraCabecera.Visible = False
                  ActivaCabecera False
                  cmdCabecera(0).Visible = True
               Else
                  ActivaCabecera True
                  ActivaFrame fraPlato
               End If
               AsignaPedido
               wAgrega = False
            End If
            grdCabecera.Enabled = True
            
            ActivaMesa247 (True)

       Case Is = 3  'Imp.Pedido
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
            ' Miguel Mendoza - 10062019
            If Not VerificaComboMaximo() Then Exit Sub
            ' *************************
            If RsArea.RecordCount = 0 Then
               MsgBox "No existe area definida", vbInformation, sMensaje
               Exit Sub
            End If
            
            If lEnvioAutomatico Then
                Dim nMinutosEnvioAntes As Integer
                nMinutosEnvioAntes = Calcular("Select ISNULL(nMinutosAntesEnvio,0) as codigo from MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
                If nMinutosEnvioAntes <> 0 Then
                   MsgBox "El Pedido tiene una programacion de impresión automática", vbInformation, sMensaje
                   Exit Sub
                End If
            End If
            
            If lObservacion And Trim(txtObservacion.Caption) = "" Then
               MsgBox "Debes ingresar la Observación", vbInformation, sMensaje
               cmdCabecera_Click (5)
               If Trim(txtObservacion.Caption) = "" Then
                  Exit Sub
               End If
            End If
            
            If lBuscarPedidoBarraMesero And lEstadoBuscarPedido Then
                If Calcular("select isnull(tmesa,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) = "" Then
                    MsgBox "Debe asignar una mesa al pedido!!!", vbInformation, sMensaje
                    Exit Sub
                End If
                
            End If

            'ENVIO DE PEDIDO AUTORIZADO
            Dim xClaveAutoriza As String
            Dim xUsuarioAutorizaenvio As String
            Dim xPerfilUsusario As String
            Dim cantProductoAutoriza As Integer
            
            xPerfilUsusario = ""
            xUsuarioAutorizaenvio = ""
            xClaveAutoriza = ""
            
            If lEnvioProduccionUsuario Then
            
            cantProductoAutoriza = Calcular("select count(*) as codigo from DPEDIDO d inner join TPRODUCTO t ON d.tCodigoProducto = t.tCodigoProducto where d.tCodigoPedido='" & sPedido & "' and d.lImprime <> 1 and t.lSolicitaAutorizacionEnvio = 1", Cn)
            
                 If cantProductoAutoriza > 0 And lEnvioProduccionCaja Then
                     frmPassword.Show vbModal
                     If wEnter Then
                         xClaveAutoriza = sDescrip
                         
                         If xClaveAutoriza <> "" Then
                            
                             xUsuarioAutorizaenvio = VerificarAutorizacionEnvio(xClaveAutoriza)
                             
                             If xUsuarioAutorizaenvio = "" Then
                                 Exit Sub
                             Else
                                xPerfilUsusario = Mid(xUsuarioAutorizaenvio, 1, 1)
                                 xUsuarioAutorizaenvio = Mid(xUsuarioAutorizaenvio, 2, Len(xUsuarioAutorizaenvio))
                                 
                             End If
                             
                         Else
                             Exit Sub
                         End If
                     Else
                        Exit Sub
                     End If
                     
                End If
    
           End If
           '-------------------------
            
            'Oscar Ortega----------------------------------------------------
            If ExistenPropiedadesPendientesEnPedido(sPedido) Then
                If ExistenPropiedadesPendientesEnCombos(sPedido) Then
                    '----------------------------------------------------------------
                    Screen.MousePointer = vbHourglass
                    'InsumosCriticos ' 23
                    If lPrinter = True Then
                                Dim cmdInsumo          As ADODB.Command
                                Dim X As Integer
                                Dim rstItems As New Recordset
                                Dim cadenaInsumos As String
                                Dim cadenaAEnviar As String
                                Dim resultado As String
                                'cadenaInsumos = "select sum(ncantidad) ncantidad, TPRODUCTO.TCODIGOINSUMO from dpedido inner join tproducto on dpedido.tcodigoproducto=tproducto.tcodigoproducto INNER JOIN  dbo.TINSUMO ON dbo.TPRODUCTO.tcodigoInsumo = dbo.TINSUMO.tcodigo  where tcodigopedido='" & sPedido & "' and lcontrolinsumocritico=1 and isnull(limprime,0)=0 AND ISNULL(TCODIGOINSUMO,''  )<>'' and (tinsumo.lactivo=1) group by tcodigoinsumo"
                                cadenaInsumos = "usp_Inforest_ObtenerInsumosCriticos '" & sPedido & "'"

                                Set rstItems = Lib.OpenRecordset(cadenaInsumos, Cn)
                                        If Not (rstItems.EOF Or rstItems.BOF) Then
                                             rstItems.MoveFirst
                                             For X = 0 To rstItems.RecordCount - 1
                                                cadenaAEnviar = cadenaAEnviar + rstItems.Fields(1) + "|" + str(rstItems.Fields(0)) + "$"
                                                rstItems.MoveNext
                                             Next X
                                           
                                            Set cmdInsumo = New ADODB.Command
 
                                            With cmdInsumo
                                                 .ActiveConnection = Cn
                                                 .CommandType = adCmdStoredProc
                                                 .CommandText = "USP_actualizaStockInsumo"
                                                 .Parameters.Refresh
                                                 .Parameters("@vi_detalles") = cadenaAEnviar
                                                 .Parameters("@vi_numdet") = rstItems.RecordCount
                                                 .Parameters("@vch_Salida") = ""
                                            End With
                                            cmdInsumo.Execute
                                            resultado = cmdInsumo.Parameters("@vch_Salida").value
                                            If resultado <> "1" Then:  MsgBox "No hay cantidad disponible de : " & resultado:   Screen.MousePointer = vbDefault: Exit Sub
                                        End If
                    End If
                    'InsumosCriticos
                    
                    
                    
      
     
                    
                    'CESAR IMPRESION ALERTA COMBO
                    Dim RsCombox As Recordset
                    Dim CountComb As Integer
                        If lOrden Then
                           Isql = "select vpedido.*, isnull(dpedido.lcombinacion,0) lcombinacion FROM dbo.vPedido INNER JOIN dbo.DPEDIDO ON dbo.vPedido.Codigo = dbo.DPEDIDO.tCodigoPedido AND dbo.vPedido.tItem = dbo.DPEDIDO.tItem " & _
                                  "Where vpedido.Codigo = '" & sPedido & "' and vpedido.nOrden in (select nOrden from DPEDIDO where tCodigoPedido='" & sPedido & "' And tArea IS NULL and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))) " & _
                                  "ORDER BY vpedido.nOrden, vpedido.tItem, vpedido.nOrdenCombo,vpedido.tetiqueta,vpedido.combo " ' tItemCombo"
                        Else
                           Isql = "select vpedido.*, isnull(dpedido.lcombinacion,0) lcombinacion FROM dbo.vPedido INNER JOIN dbo.DPEDIDO ON dbo.vPedido.Codigo = dbo.DPEDIDO.tCodigoPedido AND dbo.vPedido.tItem = dbo.DPEDIDO.tItem " & _
                                  "Where vpedido.Codigo = '" & sPedido & "' And vpedido.lImprimeArea = 1 And vpedido.tArea IS NULL and (vpedido.lImprime = 0 or (isnull(vpedido.lImprimeAreaCombo,0) = 1  and isnull(vpedido.lImprimeCombo,0) = 0 ))" & _
                                  "ORDER BY vpedido.nOrden, vpedido.tItem, vpedido.nOrdenCombo,vpedido.tetiqueta,vpedido.combo" 'tItemCombo"
                        End If
                    Set RsCombox = Lib.OpenRecordset(Isql, Cn)
                    
                    If RsCombox.RecordCount > 0 Then
                       If IIf(IsNull(RsCombox!Combo), "", RsCombox!Combo) = "" And RsCombox!lCombinacion = 0 Then
                            MsgBox "El Producto " & RsCombox!Producto & " no tiene un area de impresion configurada", vbExclamation, sMensaje
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        
                                
                       CountComb = Calcular("Select count(*) As Codigo From CPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
                       If CountComb > 0 Then
                          MsgBox "El Producto Combo " & RsCombox!Producto & " contiene Items sin un area de impresion configurada", vbExclamation, sMensaje
                       Else
                          MsgBox "El Producto Combo " & RsCombox!Producto & " no contiene ningun Item como detalle", vbExclamation, sMensaje
                       End If
                       Screen.MousePointer = vbDefault
                       Exit Sub
                       
                    End If
                    '--------------------



                    If lOrden Then
                       Isql = "select * from vPedido " & _
                              "Where Codigo = '" & sPedido & "' and nOrden in (select nOrden from DPEDIDO where tCodigoPedido='" & sPedido & "' and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))) " & _
                              "ORDER BY nOrden, tItem, nOrdenCombo, tetiqueta,combo"
                              '"ORDER BY nOrden, tItem, nOrdenCombo, tItemCombo"
                    Else
                       Isql = "select * from vPedido " & _
                              "Where Codigo = '" & sPedido & "' and lImprimeArea = 1 and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))" & _
                              "ORDER BY nOrden, tItem, nOrdenCombo, tetiqueta,combo"
                              '"ORDER BY nOrden, tItem, nOrdenCombo, tItemCombo"
                    End If


                    Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                    frmMozoUsuario.timSalida.Enabled = False
                    If Not RsImpresion.EOF Then
                       RsArea.MoveFirst
                       For i = 1 To RsArea.RecordCount
                           RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
                           If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                              If RsImpresion.RecordCount <> 0 Then
                                 RsImpresion.MoveFirst
                                 ImprimePedido RsImpresion, "N", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Inforest"
                                    'If (lCompatibilidadTVS) Then 'TVS
                                    '    Call TVS_EnviarTicket(RsImpresion, RsArea!Area, "Nuevo Ticket", sPedido)
                                    'End If
                              End If
                           End If
                           RsArea.MoveNext
                       Next i
                    End If
                    'KDS
                    If lKDS And lKDSInforest = False Then
                       Call KDS_AnadirNuevaOrden(RsCabecera, 0)
                    End If
                    
                    
                   '---------------------------------
                   '--- GCAA 23112021 KDS INFOREST
                   '---------------------------------
                   Dim activo_kds As Boolean
                   Dim x_orden_area As Boolean
                   activo_kds = True
                   x_orden_area = True
                   
                   If lKDSInforest Then
                        'ACTUALIZA
                        Dim oComando_ As clsComando
                        Set oComando_ = New clsComando
                        
                        If Not oComando_.CreateCmdSp("KDS_Usp_spUpdate_DPEDIDO_Orden", Cn) Then
                           Set oComando_ = Nothing
                           Exit Sub
                        End If
                   
                        If lKDSxArea Then
                            oComando_.CreateParameter "@tipo", adVarChar, adParamInput, 10, 2
                        Else
                            oComando_.CreateParameter "@tipo", adVarChar, adParamInput, 10, 1
                        End If
                        
                        oComando_.CreateParameter "@pedido", adVarChar, adParamInput, 10, sPedido
                        oComando_.CreateParameter "@tturnoKDS", adVarChar, adParamInput, 10, sPedido
                        
                        If Not oComando_.GetParamOK Then
                           Set oComando_ = Nothing
                           Exit Sub
                        End If
                    
                        If Not oComando_.ExecSP Then
                           Set oComando_ = Nothing
                           Exit Sub
                         End If
                   End If
                    
                    
                    'CESAR----CHEF CONTROL
                    Dim ChefEnvio As Boolean
                    ChefEnvio = Calcular("select lEnvioChef as Codigo FROM TPARAMETRO", Cn)
                    
                    Cn.Execute "Update DPEDIDO Set lNoCantado=0 where tCodigoPedido = '" & sPedido & "' and lImprime<>1"
                    
                    If ChefEnvio Then
                    Cn.Execute "Update DPEDIDO Set lCantadoc=1,fCantadoC=GetDate(), lTipoEnvio=0 where tCodigoPedido = '" & sPedido & "' and lImprime<>1"
                    End If
                    '---------------------------------
                    

                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    'MESA247
                    If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                          Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                    End If
                                    
                    frmMozoUsuario.timSalida.Enabled = True
                    Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1  where tCodigoPedido='" & sPedido & "'"
                    Cn.Execute "Update DPEDIDO Set lImprime = 1, tUsuarioAutorizaEnvio = '" & xUsuarioAutorizaenvio & "', tPerfilUsuarioAutoriza = '" & xPerfilUsusario & "', fEnvio = getdate(), nEnvio = 1 where tCodigoPedido = '" & sPedido & "' and lImprime<>1"
                    Cn.Execute "Update CPEDIDO Set lImprime = 1 where tCodigoPedido = '" & sPedido & "'"
                End If
            End If
            RsDetalle.Requery
            RsCombo.Requery
            Me.Caption = muestra
            LimpiaRs
            
            If lBuscarPedidoBarraMesero And lEstadoBuscarPedido Then
                Call TerminaprocesoBarra
                Cn.Execute "update tmesa set testadomesa ='04' where tcodigomesa='" & sMesa & "'"
            End If
                                           
       Case Is = 4
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
            
            If lObservacion And Trim(txtObservacion.Caption) = "" Then
               MsgBox "Debes ingresar la Observación", vbInformation, sMensaje
               cmdCabecera_Click (5)
               If Trim(txtObservacion.Caption) = "" Then
                  Exit Sub
               End If
            End If
                                   
            Dim nLP As Integer
            nLP = Calcular("select nPrecuenta as Codigo FROM MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
            If nLimitePrecuenta <> 0 And nLimitePrecuenta <= nLP Then
               If Supervisor("14") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
                                               
            If lPrecuentaImpresora Then
               frmMozoUsuario.timSalida.Interval = 0
               frmPrecuentaImpresora.Show vbModal
               ' GCAA 06012021
'               frmMozoUsuario.timSalida.Interval = True
               If Not wEnter Then
                  Exit Sub
               End If
            Else
               sCodigo = sPreCuenta
            End If
                                                
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from DPEDIDO where lImprime=0 and tCodigoPedido ='" & sPedido & "'", Cn)
               If i > 0 Then
                  MsgBox "Falta Enviar platos a Producción", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
                                    
            'Calcula Descuento
            If CDbl(txtDescuento) > 0 Then
               Dim nTope As Double
               Dim nTotalMes As Double
               
               Dim nConsumo As Double
               Dim aplicaTope As Boolean
               
               nTope = Calcular("select nTope as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
               If nTope > 0 Then
                  If Calcular("select lTopePedido as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn) Then
                     If CDbl(txtDescuento) > nTope Then
                        If MsgBox("El Descuento a aplicar Supera El Tope Registrado por Pedido" & Chr(13) & "¿Desea aplicar el Tope de " & sMonN & " " & nTope & "?", vbQuestion + vbYesNo) = vbYes Then
                           CalculaAplicaTope (nTope)
                        Else
                           Exit Sub
                        End If
                     End If
                  Else
                     nTotalMes = Calcular("select sum(DPEDIDO.nDescuento*nCantidad) as Codigo FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                                          "WHERE month(MPEDIDO.fFecha) = month(getdate()) and year(MPEDIDO.fFecha)=year(getdate()) and mPedido.tDescuento='" & sCodigoDescuento & "' and tEstadoPedido<>'01' and tEstadoPedido<>'03'", Cn)
                     
                     If CDbl(txtDescuento) + nTotalMes > nTope Then
                        If CDbl(txtDescuento) < nTope Then
                           If MsgBox("El Descuento a aplicar Supera El Tope Registrado dentro de un mes" & Chr(13) & "¿Desea aplicar el saldo " & sMonN & " " & nTope - nTotalMes & "?", vbQuestion + vbYesNo) = vbYes Then
                              CalculaAplicaTope (nTope - nTotalMes)
                           Else
                              Exit Sub
                           End If
                        Else
                            MsgBox "El Descuento a aplicar Supera El Tope Registrado dentro de un mes", vbExclamation
                            Exit Sub
                        End If
                     End If
                  End If
               End If
               txtDescuento.Caption = Format(Calcular("select sum(nDescuento*nCantidad) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
            End If
                                    
                                    
            wEnter = False
            Screen.MousePointer = vbHourglass
            Cn.Execute "update MPEDIDO set nPrecuenta = isnull(nPrecuenta,0) + 1  where tCodigoPedido='" & sPedido & "'"
            
            
            'CESAR VALIDACION DELIVERY
            lCanalDelivery = Calcular("select v.lCanalDelivery As Codigo from vTipoPedido v INNER JOIN MPEDIDO m ON Codigo = m.tTipoPedido Where tCodigoPedido='" & sPedido & "'", Cn)
            '------------------------------------
            
            If lCanalDelivery Then
               Isql = "select * from vPreCuentaDelivery WHERE Codigo='" & sPedido & "' ORDER BY tItem, tItemCombo"
            Else
               If lPrecuenta Then
                  Isql = "select * from vPrecuentaDetallada WHERE Codigo='" & sPedido & "'"
               Else
                  If lPrecuentaAgrupada Then
                     Isql = "select * from vPrecuentaAgrupada WHERE Codigo='" & sPedido & "' order by tItem"
                  Else
                     Isql = "select * from vPrecuenta WHERE Codigo='" & sPedido & "' order by tItem"
                  End If
               End If
            End If

 
            Dim rstFuente As New ADODB.Recordset
            Set rstFuente = New ADODB.Recordset
            imageCab.Picture = Nothing
            imagepIE.Picture = Nothing
            Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
            If lImprimeImagCabPrecuenta Then
            
                    imageCab.DataField = "foto"
                    Set imageCab.DataSource = rstFuente
            End If
            If lImprimeImagPiePrecuenta Then
                    imagepIE.DataField = "fotoPie"
                    Set imagepIE.DataSource = rstFuente
            End If
          

            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
            If RsImpresion.RecordCount = 0 Then
               LimpiaRs
               MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
            Else
               If lCanalDelivery Then
                  ImprimeDelivery RsImpresion, RsProductoPropiedad, RsComboPropiedad, imageCab, imagepIE
               Else
                  If lPrecuenta Then
                     ImprimePreCuentaDetallada RsImpresion, sPreCuenta, imageCab, imagepIE
                  Else
                     If lInfhotel Then
                        ImprimeInfhotel RsImpresion, sPreCuenta, imageCab, imagepIE
                     Else
                            If lPrecuentaNoValorizada Then
                                ImprimePrecuentaNoValorizada RsImpresion, sPreCuenta, imageCab, imagepIE
                            Else
                                ImprimePreCuenta RsImpresion, sPreCuenta, imageCab, imagepIE
                            End If
                     End If
                  End If
               End If
               'TVS
                'If (lCompatibilidadTVS) Then 'TVS
                '    Call TVS_EnviarTicket(RsImpresion, sCodigo, "Imp Precuenta", sPedido)
                'End If
               Cn.Execute "update MPEDIDO set fRegCuenta = getDate() where tCodigoPedido='" & sPedido & "'"
            End If
            LimpiaRs
            
       Case Is = 5  'OFERTAS
        If Supervisor("10") = False Then
                   MsgBox "Clave no permitida", vbExclamation, sMensaje
                   Exit Sub
                End If
                                
                Dim sCriterio As String
                Dim nOferta As Double
                
                If sTipoPedido = "01" Then
                   sCriterio = " and lLocal=1"
                ElseIf sTipoPedido = "02" Then
                   sCriterio = " and lDelivery=1"
                ElseIf sTipoPedido = "03" Then
                   sCriterio = " and lLlevar=1"
                ElseIf sTipoPedido = "04" Then
                   sCriterio = " and lCanal4=1"
                Else
                   sCriterio = " and lCanal5=1"
                End If
                
                Isql = "SELECT tOferta as Codigo, tNombre as Descripcion, " & _
                       "case when nRatio>0 then 'Descuento del ' + str(nRatio,2) + '%' else " & _
                       "case when nMonto>0 then 'Descuento de ' + str(nMonto,2) + ' " & sMonedaN & "' else 'Producto al Precio de ' + ' " & sMonN & " ' + str(nPrecio,2) end end as Oferta " & _
                       "From dbo.TOFERTA WHERE lAutomatica=0 and tCodigoProducto = '" & sProducto & "' and lActivo=1 " & _
                       " and (substring(tFrecuencia," & Weekday(FechaServidor(), vbMonday) & "+1,1) = '1' or (substring(tFrecuencia,1,1)='1') and MONTH(fFecha) = " & Month(FechaServidor()) & " AND DAY(fFecha)= " & Day(FechaServidor()) & ") and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "'" & _
                       " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "')) " & sCriterio

                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                "Oferta", 2, "Descripcion", 3000, 0, 0, "", _
                                                                "Descripcion de la Oferta", 2, "Oferta", 4200, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                wCalcula = True
                                
                                
                If wEnter Then
                    
                     Dim stockOferta, stockFinal As Integer
                                        
                   'Valida stock de Oferta
                    Isql = "exec CalcularStockOfertas '" & sProducto & "', '" & sCodigo & "', " & nCantidad & ", '" & sTipoPedido & "'"
                    
                    Set ValorOferta = Lib.OpenRecordset(Isql, Cn)
                    respuestaOferta = ValorOferta!respuesta
                    
                    If UCase(respuestaOferta) = "EXITO" Then
                    
                        stockOferta = ValorOferta!stockOferta
                        
                        If stockOferta = 0 Then
                            'MsgBox "Actualmente, dispones de cantidad ilimitada para este producto", vbExclamation, sMensaje
                        Else
                            stockFinal = stockOferta - nCantidad
                        
                            MsgBox "Actualmente, dispone de una cantidad de " + CStr(stockFinal) + " unidades disponibles en oferta.", vbExclamation, sMensaje
                        End If
                        
                    Else
                        MsgBox respuestaOferta, vbExclamation, sMensaje
                        Exit Sub
                    End If
                    
                    'Fin Valida stock de Oferta
                
                   Isql = "select * from TOFERTA where tOferta='" & sCodigo & "' and tCodigoProducto='" & sProducto & "'"
                   Set RsOferta = Lib.OpenRecordset(Isql, Cn)
                   If RsOferta.RecordCount > 0 Then
                      RsOferta.MoveFirst
                      If RsOferta!nPrecio > 0 Then
                         nOferta = nOficial - IIf(IsNull(RsOferta!nPrecio), 0, RsOferta!nPrecio)
                      ElseIf RsOferta!nMonto > 0 Then
                         nOferta = RsOferta!nMonto
                      Else
                         nOferta = nOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
                      End If
                   End If
                   
                   nPVenta = nOficial - nOferta
                   nDescuento = nOficial - nPVenta

                   Dim Acumulado As Double
                   'extranjero bolivia
                   Select Case pais 'ok
                        Case "001" 'Bolivia
                                    Acumulado = 0
                                    Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                                    Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                                    Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                                    Acumulado = (Acumulado / 100)
                                    nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta * nPorcentaje1 / 100, 0)
                                    nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta * nPorcentaje2 / 100, 0)
                                    nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta * nPorcentaje3 / 100, 0)
                                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                                    
                        Case Else 'Peru, Ecuador
                                    Acumulado = 0
                                    Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                                    Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                                    Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                                    Acumulado = 1 + (Acumulado / 100)
                                    nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                                    nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                                    nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                     End Select
                     
                     Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
                           "nDescuento = " & nDescuento & ", " & _
                           "nRecargo = " & nRecargo & ", " & _
                           "nPrecioOficial = " & nOficial & ", " & _
                           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
                           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
                           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
                           "nPrecioVenta = " & nPVenta & ", " & _
                           "nventa = " & nPVenta * nCantidad & ", " & _
                           "nCantidad = " & nCantidad & ", " & _
                           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
                           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
                           "nImpuesto3 = " & nImpuesto3 * nCantidad & ", tOferta='" & sCodigo & "', tAutorizaOferta='" & sVar1 & "' " & _
                           "where tItem = '" & sItem & "' and tCodigoPedido ='" & sPedido & "'"
                   Cn.Execute Isql
                   
                   nPos = RsDetalle.AbsolutePosition
                   RsDetalle.Requery
                   RsDetalle.AbsolutePosition = nPos
                   txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
                
                    'MESA247
                    If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                          Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                    End If
               
                End If
                
       Case Is = 6  'Código Directo
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
            
            If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
               MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
               RsCabecera.Requery
               RsDetalle.Requery
               Exit Sub
            End If
            
            sTemp = ""
            
            If Calcular("Select count(*) as Codigo from tclienteproducto where tcodigodelivery='" & sCliente & "' ", Cn) = 0 Then
                Isql = "select * from vProducto where lActivo = 1 and (tUnidadNegocio='" & sUnidadNegocio & "')and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lcanal4=1", "lcanal5=1")))) & " Order by Descripcion"
                frmBusquedaRapida.cmdOpcion(4).Visible = False
            Else
                Isql = "select vProducto.Grupo, vProducto.Descripcion , tclienteproducto.nprecio As nPrecioVenta , vProducto.nBoton, vProducto.SubGrupo, vProducto.Codigo from vProducto inner join tclienteproducto on vProducto.codigo=tclienteproducto.tcodigoproducto where vProducto.lActivo = 1 and (vProducto.tUnidadNegocio='" & sUnidadNegocio & "' ) and tclienteproducto.tcodigodelivery='" & sCliente & "'  Order by vProducto.Descripcion"
                frmBusquedaRapida.cmdOpcion(4).Visible = True
                frmBusquedaRapida.cmdOpcion(4).FontBold = True
            End If

            'Isql = "select * from vProducto where lActivo = 1 and (tUnidadNegocio='" & sUnidadNegocio & "')and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lcanal4=1", "lcanal5=1")))) & " Order by Descripcion"

            Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                            "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                            "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                            "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                            "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
            frmBusquedaRapida.nPredeterm = 1
            frmBusquedaRapida.cmdOpcion(0).Visible = True
            frmBusquedaRapida.Show vbModal
            
            If wEnter Then
                    'INSUMOCRITICO23
                If validadIngresoProducto(sCodigo) = False Then
                    Exit Sub
                End If
                    'INSUMOCRITICO23
             
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                   MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                   RsCabecera.Requery
                   RsDetalle.Requery
                   Exit Sub
                End If
                
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                   MsgBox "El Pedido ya fue facturado", vbExclamation, sMensaje
                   RsCabecera.Requery
                   RsDetalle.Requery
                   Exit Sub
                End If
                
               sProducto = sCodigo
               'Correccion
               ListarOperadoresConFiltro (sProducto)
               Dim xxx As String
               xxx = RsProducto.Filter
               RsProducto.Filter = adFilterNone
               RsProducto.MoveFirst
               RsProducto.Find ("Codigo='" & sProducto & "'")
               
               If vComanda Then
                  sTipo = "Comanda"
                  sDescrip = sComanda
                  frmNumPad.Caption = "Comanda"
                  frmNumPad.Show vbModal
                
                  sComanda = IIf(sDescrip = "" Or sDescrip = "0", "", Mid("0000000000", 1, 10 - Len(Trim(sDescrip))) & Trim(sDescrip))
                  
                 If wComanda And sComanda = "" Then
                    MsgBox "No ha ingresado el número de la comanda", vbExclamation, sMensaje
                    Exit Sub
                 End If
               
                 If Not wEnter Then
                    Exit Sub
                 End If
               End If
              nCantidad = 1
               InsertaProducto 1
               If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                  lPropiedad = True
               End If
               RsProducto.Filter = IIf(xxx = "0", "", xxx)
               
               wEnter = False
            End If
       Case Is = 7
                If Calcular("select COUNT(*) as codigo from vtipopedido where lcanalcentralpedidos=1", Cn) > 0 Then
                    frmBusquedaCentralPedido.Show vbModal
                Else
                    MsgBox "No tiene configurado Canal de Venta con Opción de Central de Pedidos"
            
               End If
       Case Is = 8  'Salir Propiedades
            ActivaMesa247 (True)
            If wAgregaCombo Then
                'Oscar Ortega---------------------------------
                If ObligaPropiedad(sCombo) = False Then
                    Exit Sub
                End If
                'Fin Oscar Ortega-----------------------------
               cmdOpcion(1).Enabled = False
               RsCombo.Requery
               RsCombo.MoveFirst
               RsCombo.Find "titemCombo = '" & xItem & "'"
               ActivaFrame fraProductoCombo
            Else
                'Oscar Ortega---------------------------------
               If ObligaPropiedad(sProducto) = False Then
                    Exit Sub
                Else
                    grdDetalle.Enabled = True
                    grdCabecera.Enabled = True
               End If
               'Fin Oscar Ortega-----------------------------
               RsDetalle.Requery
               RsDetalle.MoveFirst
               RsDetalle.Find "titem = '" & sItem & "'"
               ActivaOpcion True
               ActivaCabecera True
               ActivaFrame fraPlato
            End If
            
        Case Is = 9
               
                frmReserva.Show vbModal
            
        Case Is = 10  'IR
            If Not RsCabecera.EOF Then
               wEnter = False
               Sw = False
               sDescrip = ""
               
               'busquedasocio
               'busquedaxSocio
               sTemp = ""
               lCargaDesdePedido = False
               Call ConfGrilla(6, frmIr.grdGrillaPedido, "Pedido", 2, "Pedido", 1200, 0, 0, "", _
                                        "Tipo Pedido", 2, "TipoPedido", 1300, 0, 0, "", _
                                        "Salon", 2, "Salon", 1700, 0, 0, "", _
                                        "Mesa", 2, "Mesa", 800, 0, 0, "", _
                                        "Cliente", 2, "Cliente", 3800, 0, 0, "", _
                                        "Total", 2, "Venta", 1000, 1, 0, "")
                If lBuscarPedidoFiltrarMesa = True Then
                    frmIr.nPredeterm = 3
                Else
                    frmIr.nPredeterm = 4
                End If
               
               
               
               frmIr.Show vbModal
               
               If wEnter Then
                  RsCabecera.MoveFirst
                  
                  If lCargaDesdePedido = True Then
                     RsCabecera.Find ("Codigo='" & sCodigo & "'")
                  Else
                     RsCabecera.Find ("tMesa='" & sCodigo & "'")
                  End If

                  
                  
                  If RsCabecera.EOF Then
                     RsCabecera.MoveFirst
                     MsgBox "El pedido pertenece a otra caja ", vbExclamation, sMensaje
                  End If
                  grdCabecera.Bookmark = RsCabecera.Bookmark
               End If
               
               If Sw Then
                  RsCabecera.MoveFirst
                  If Len(Trim(sDescrip)) > 0 Then
                     RsCabecera.Find ("tObservacion like '*" & Trim(sDescrip) & "*'")
                     If RsCabecera.EOF = True Then
                        RsCabecera.MoveFirst
                        MsgBox "No se encontro " & sDescrip, vbExclamation, sMensaje
                        Exit Sub
                     End If
                     grdCabecera.Bookmark = RsCabecera.Bookmark
                  End If
                  Sw = False
               End If
            End If
            
       Case Is = 11 ' Refrescar
            RsCabecera.Requery
            RsDetalle.Requery
            RsCombo.Requery
            If RsCabecera.EOF Then
               LimpiaCabecera
            End If
            If lBuscarPedidoBarraMesero Then
                Me.txtbuscarpedidobarra.SetFocus
            Else
                On Error Resume Next
                Me.txtBarra.SetFocus
                
            End If
       Case Is = 12 ' Combo
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
       
            If wCombo Then
               fraCombo.Visible = True
               fraProductoCombo.Visible = True
               wAgregaCombo = True
               ActivaCabecera False
               ActivaOpcion False
               If Not RsCombo.EOF Then
                  RsCombo.MoveFirst
               End If
               
               AsignaProductoCombo
               cmdOpcion(1).Enabled = False
               cmdOpcion(2).Enabled = False
            End If
            ActivaMesa247 (False)
       Case Is = 13 ' Visualizar Pedido
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
            
            If lPrecuenta Then
               sTipo = "ConMozo"
            Else
               sTipo = "SinMozo"
            End If
            sTipo = ""
            frmPedido.Show vbModal
            If wEnter Then
               RsDetalle.Requery
            End If
                                    
       Case Is = 17 ' Cancelar del Motivo de Eliminacion
            ActivaFrame fraPlato
            ActivaCabecera True
            ActivaOpcion True
                                    
   End Select
End Sub
Public Sub ValidaModifica()
    On Error GoTo fin
    If sPedido <> "" Then

        Set RsPedido = Lib.OpenRecordset("SELECT * FROM MPEDIDO WHERE  tCodigoPedido= '" & sPedido & "'", Cn)
        Dim valorActual As String
        'Modifica Cliente Delivery
        valorActual = IIf(IsNull(RsPedido!tClienteDelivery), "", RsPedido!tClienteDelivery)
        If valorActual <> sCliente Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA CLIENTE DELIVERY", sPedido, Error, "", "", "Modificación del Cliente Delivery. Valor modificado:" & valorActual & " - Valor Actual: " & sCliente, "Modificación del cliente delivery del pedido " & sPedido, sUsuario)
        End If
    
        'Modifica Tipo Pedido
        valorActual = IIf(IsNull(RsPedido!tTipoPedido), "", RsPedido!tTipoPedido)
        If valorActual <> sTipoPedido Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA TIPO PEDIDO", sPedido, Error, "", "", "Modificación del tipo de pedido. Valor modificado:" & valorActual & " - Valor actual: " & sTipoPedido, "Modificación del tipo pedido (" & sPedido & ") ", sUsuario)
        End If
    
        'Modifica lPrioridad
        valorActual = IIf(IsNull(RsPedido!lPrioridad), "", RsPedido!lPrioridad)
        If valorActual <> IIf(txtTipoPedido.ForeColor = &H800000, False, True) Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA PRIORIDAD", sPedido, Error, "", "", "Modificación de la perioridad del pedido. Valor modificado:" & valorActual & " - Valor actual: " & IIf(txtTipoPedido.ForeColor = &H800000, 0, 1), "Modificación de la prioridad del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        'Modifica Tipo Atencion
        valorActual = IIf(IsNull(RsPedido!tTipoAtencion), "", RsPedido!tTipoAtencion)
        If valorActual <> sTipoAtencion Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA TIPO ATENCIÓN", sPedido, Error, "", "", "Modificación del tipo de atención. Valor modificado:" & valorActual & " - Valor actual: " & sTipoAtencion, "Modificación del tipo atención del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        'Modifica Mozo
        valorActual = IIf(IsNull(RsPedido!tMozo), "", RsPedido!tMozo)
        If valorActual <> sMozo Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA MOZO ASIGNADO", sPedido, Error, "", "", "Modificación del Mozo. Valor modificado:" & valorActual & " - Valor actual: " & sMozo, "Modificación del mozo del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tMotorizado), "", RsPedido!tMotorizado)
        If valorActual <> sMotorizado Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA MOTORIZADO", sPedido, Error, "", "", "Modificación del Motorizado. Valor modificado:" & valorActual & " - Valor actual: " & sMotorizado, "Modificación del motorizado del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tObservacion), "", RsPedido!tObservacion)
        If valorActual <> txtObservacion.Caption Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA OBSERVACIÓN", sPedido, Error, "", "", "Modificación de la observación. Valor modificado:" & valorActual & " - Valor actual: " & txtObservacion.Caption, "Modificación de la observación del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tcodigoinvitado), "", RsPedido!tcodigoinvitado)
        If valorActual <> sCodigoInvitado Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA INVITADO", sPedido, Error, "", "", "Modifica invitado. Valor modificado:" & valorActual & " - Valor actual: " & sCodigoInvitado, "Modificación invitado del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tPuntoVenta), "", RsPedido!tPuntoVenta)
        If valorActual <> sPuntoVenta Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA PUNTO VENTA", sPedido, Error, "", "", "Modificación del punto de venta. Valor modificado:" & valorActual & " - Valor actual: " & sPuntoVenta, "Modificación del punto de venta del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tHabitacion), "", RsPedido!tHabitacion)
        If valorActual <> sHabitacion Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA HABITACIÓN", sPedido, Error, "", "", "Modificación de Nro Habitación. Valor modificado:" & valorActual & " - Valor actual: " & sHabitacion, "Modificación del número de habitación del pedido (" & sPedido & ") ", sUsuario)
        End If
            
        valorActual = IIf(IsNull(RsPedido!TRESERVA), "", RsPedido!TRESERVA)
        If valorActual <> sReserva Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA RESERVA", sPedido, Error, "", "", "Modificación de Nro Reserva. Valor modificado:" & valorActual & " - Valor actual: " & sReserva, "Modificación del número de reserva del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tPasajero), "", RsPedido!tPasajero)
        If valorActual <> sPasajero Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA PASAJERO", sPedido, Error, "", "", "Modificación de Pasajero. Valor modificado:" & valorActual & " - Valor actual: " & sPasajero, "Modificación del pasajero del pedido (" & sPedido & ") ", sUsuario)
        End If

        valorActual = IIf(IsNull(RsPedido!nDescuento), "", RsPedido!nDescuento)
        If valorActual <> xDescuento Or RsPedido!tDescuento <> sCodigoDescuento Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA DESCUENTO", sPedido, Error, "", "", "Modificación de Descuento. Valor modificado:" & valorActual & "(Cod. desct" & RsPedido!tDescuento & ") - Valor actual: " & xDescuento & "(Cod. Descuento" & sCodigoDescuento & ")", "Modificación de descuento del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tUsuarioDescuento), "", RsPedido!tUsuarioDescuento)
        If valorActual <> Mid(tAutorizaDescuento, 1, 15) Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA AUTORIZA DESCUENTO", sPedido, Error, "", "", "Modifica autoriza descuento. Valor modificado:" & valorActual & " - Valor actual: " & Mid(tAutorizaDescuento, 1, 15), "Modificación de autoriza descuenta ( Ped. " & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tObservacionDescuento), "", RsPedido!tObservacionDescuento)
        If valorActual <> IIf(sCodigoDescuento = "000", sDescripcionDescuento, "") Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA OBS DESCUENTO", sPedido, Error, "", "", "Modificación de obs. de descuento. Valor modificado:" & valorActual & " - Valor actual: " & IIf(sCodigoDescuento = "000", sDescripcionDescuento, ""), "Modificación de obs. de descuento (Ped. " & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tTienda), "", RsPedido!tTienda)
        If valorActual <> sTienda Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA TIENDA", sPedido, Error, "", "", "Modificación de tienda. Valor modificado:" & valorActual & " - Valor actual: " & sTienda, "Modificación de tienda (Ped. " & sPedido & ") ", sUsuario)
        End If
    
        valorActual = IIf(IsNull(RsPedido!fProgramacion), "", RsPedido!fProgramacion)
        If Format(valorActual, "dd/MM/yyyy HH:mm") <> IIf(txtFechaProg.Caption = "", Null, Format(txtFechaProg.Caption, "dd/MM/yyyy HH:mm")) Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA FECHA PROGRAMACIÓN", sPedido, Error, "", "", "Modificación fecha programación. Valor modificado:" & valorActual & " - Valor actual: " & IIf(txtFechaProg.Caption = "", Null, Format(txtFechaProg.Caption, "dd/MM/yyyy HH:mm")), "Modifica fecha programación (Ped. " & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tcodigopariente), "", RsPedido!tcodigopariente)
        If valorActual <> sCodigoParienteSeleccionado Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA PARIENTE", sPedido, Error, "", "", "Modificación pariente. Valor modificado:" & valorActual & " - Valor actual: " & sCodigoParienteSeleccionado, "Modificación pariente (Ped. " & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tentregara), "", RsPedido!tentregara)
        If valorActual <> IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20)) Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA ENTREGAR A ", sPedido, Error, "", "", "Modificación entregar a. Valor modificado:" & valorActual & " - Valor actual: " & IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20)), "Modificación 'entregar a' (Ped. " & sPedido & ") ", sUsuario)
        End If
    
        valorActual = IIf(IsNull(RsPedido!nMinutosAntesEnvio), 0, RsPedido!nMinutosAntesEnvio)
        If valorActual <> Val(txtEnvioAntes.Text) Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA MINUTOS ANTES ENVIO", sPedido, Error, "", "", "Modificación de tienda. Valor modificado:" & valorActual & " - Valor actual: " & Val(txtEnvioAntes.Text), "Modificación de minutos antes del envío (Ped. " & sPedido & ") ", sUsuario)
        End If
    
    
        valorActual = IIf(IsNull(RsPedido!CodigoOrigenVenta), "", RsPedido!CodigoOrigenVenta)
        If valorActual <> vOrigenVentas Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA ORIGEN VENTA", sPedido, Error, "", "", "Modificación de Origen Venta. Valor modificado:" & valorActual & " - Valor actual: " & vOrigenVentas, "Modificación de origen de venta (Ped. " & sPedido & ") ", sUsuario)
        End If
    
        valorActual = IIf(IsNull(RsPedido!tMesa), "", RsPedido!tMesa)
        If valorActual <> sMesa Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA MESA", sPedido, Error, "", "", "Modificación de mesa. Valor modificado:" & valorActual & " - Valor actual: " & sMesa, "Modificación de mesa (Ped. " & sPedido & ") ", sUsuario)
        End If
        
        If wAgrega Then
        
            Dim nSalon As String
            nSalon = Calcular("select tSalon as Codigo from TMESA where tCodigoMesa='" & sMesa & "' ", Cn)
            valorActual = IIf(IsNull(RsPedido!tSalon), "", RsPedido!tSalon)
            If valorActual <> sSalon Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA SALON", sPedido, Error, "", "", "Modificación de salon. Valor modificado:" & valorActual & " - Valor actual: " & sSalon, "Modificación de Salón (Ped. " & sPedido & ") ", sUsuario)
            End If
        
            valorActual = IIf(IsNull(RsPedido!tCaja), "", RsPedido!tCaja)
            If valorActual <> sCaja Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA NRO DE CAJA", sPedido, Error, "", "", "Modificación de NRO DE CAJA. Valor modificado:" & valorActual & " - Valor actual: " & sCaja, "Modificación de nro de caja (Ped. " & sPedido & ") ", sUsuario)
            End If
            
            valorActual = IIf(IsNull(RsPedido!tTurno), "", RsPedido!tTurno)
            If valorActual <> sTurno Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA TURNO", sPedido, Error, "", "", "Modificación de turno. Valor modificado:" & valorActual & " - Valor actual: " & sTurno, "Modificación de Turno (Ped. " & sPedido & ") ", sUsuario)
            End If
            
            valorActual = IIf(IsNull(RsPedido!tUsuario), "", RsPedido!tUsuario)
            If valorActual <> sUsuario Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA USUARIO", sPedido, Error, "", "", "Modificación de usuario. Valor modificado:" & valorActual & " - Valor actual: " & sUsuario, "Modificación de usuario (Ped. " & sPedido & ") ", sUsuario)
            End If
   
            
            valorActual = IIf(IsNull(RsPedido!nTiempoDelivery), "", RsPedido!nTiempoDelivery)
            If valorActual <> nTiempoDelivery Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA TIEMPO DELIVERY", sPedido, Error, "", "", "Modificación tiempo delivery. Valor modificado:" & valorActual & " - Valor actual: " & nTiempoDelivery, "Modificación tiempo delivery (Ped. " & sPedido & ") ", sUsuario)
            End If
            
            valorActual = IIf(IsNull(RsPedido!fdiacontable), "", RsPedido!fdiacontable)
            If valorActual <> obtieneDiaContable Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA DIA CONTABLE", sPedido, Error, "", "", "Modificación del día contable. Valor modificado:" & valorActual & " - Valor actual: " & obtieneDiaContable, "Modificación del día contable (Ped. " & sPedido & ") ", sUsuario)
            End If
        
        End If
    End If
fin:
    
End Sub

Public Function VerificarAutorizacionEnvio(ByVal tClave As String) As String
    Dim RsUsuario As New ADODB.Recordset
    Dim RsMozo As New ADODB.Recordset
    Dim tCodigoPersona As String
    Dim tDescPersona As String
    Dim ttipoPersona As String 'USUARIO 'MESERO
    
    
    tDescPersona = ""
    ttipoPersona = ""
    
    Set RsUsuario = Lib.OpenRecordset("SELECT u.tCodigoUsuario, u.tResumido FROM TUSUARIO u INNER JOIN TGRUPOUSUARIO g ON u.tGrupoUsuario = g.tGrupoUsuario  WHERE u.tPassword = '" & Encapsula(tClave) & "' and u.lActivo = 1 and g.lopcion24 = 1", Cn)
    
    If RsUsuario.RecordCount > 0 Then
        If Not (RsUsuario.EOF Or RsUsuario.BOF) Then
            RsUsuario.MoveFirst
            tCodigoPersona = RsUsuario.Fields("tCodigoUsuario")
            tDescPersona = "U" + RsUsuario.Fields("tResumido")
            ttipoPersona = "USUARIO"
        End If
    End If
   
    Set RsMozo = Lib.OpenRecordset("SELECT * FROM vmozo WHERE tvalor='" & Encapsula(tClave) & "' and lactivo=1 and lAutorizaEnviosMozo=1", Cn)
    If RsMozo.RecordCount > 0 Then
        If Not (RsMozo.EOF Or RsMozo.BOF) Then
            RsMozo.MoveFirst
            tCodigoPersona = RsMozo.Fields("Codigo")
            tDescPersona = "M" + RsMozo.Fields("tResumido")
            ttipoPersona = Mesero
        End If
    End If
    
    VerificarAutorizacionEnvio = tDescPersona
  
End Function

Private Sub cmdOperador_Click(Index As Integer)
   Dim i As Integer
   Screen.MousePointer = vbHourglass
   For i = 1 To 13
       cmdOperador(i).backColor = vbButtonFace
   Next i
   RsOperador.MoveFirst
   RsOperador.Find "nboton = " & Trim(str(Index))
   nOperadorPropiedad = RsOperador!nControl
        xOperador = RsOperador!codigo
   cmdOperador(Index).backColor = vbRed
   If wAgregaCombo Then
      AsignaComboPropiedad
   Else
      AsignaPropiedad
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOrigen_Click(Index As Integer)
' origen de ventas
  ' HabilitaTimerColor (False)
    
   RsOrigenVentas.MoveFirst
   RsOrigenVentas.Find "boton = " & Trim(str(Index))
   vOrigenVentas = RsOrigenVentas!CodOrigenVenta
   'sMotorizado = RsMotorizado!codigo
   'HabilitaTimerColor (True)
   validarOrigenVentas
End Sub
Private Function validarOrigenVentas()
' origen de ventas
    lActivaMozo = IIf(IsNull(RsCanalesVenta!lActivaMozo), False, RsCanalesVenta!lActivaMozo)
    lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
    lCanalDelivery = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
    lCanalCentralPedidos = IIf(IsNull(RsCanalesVenta!lCanalCentralPedidos), False, RsCanalesVenta!lCanalCentralPedidos)
    'entregarA
    lObligaEntregarA = IIf(IsNull(RsCanalesVenta!lObligaEntregarA), False, RsCanalesVenta!lObligaEntregarA)
    
    'origen de ventas
     lOrigenVentas = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
    
    If lMCPV Then
        lObligaMozo = False
        lActivaMozo = False
    Else
        lObligaMozo = IIf(IsNull(RsCanalesVenta!lObligaMozo), False, RsCanalesVenta!lObligaMozo)
    End If
    lObligaMotorizado = IIf(IsNull(RsCanalesVenta!lObligaMotorizado), False, RsCanalesVenta!lObligaMotorizado)
    lObligaMesa = IIf(IsNull(RsCanalesVenta!lObligaMesa), False, RsCanalesVenta!lObligaMesa)
    lObligaPax = IIf(IsNull(RsCanalesVenta!lObligaPax), False, RsCanalesVenta!lObligaPax)
    lObligaFechaEntrega = IIf(IsNull(RsCanalesVenta!lObligaIngresoFechaEntrega), False, RsCanalesVenta!lObligaIngresoFechaEntrega)
    lObligaClienteFrecuente = IIf(IsNull(RsCanalesVenta!lObligaClienteFrecuente), False, RsCanalesVenta!lObligaClienteFrecuente)
    
    
    'If lActivaMozo Then
        'fraMozo.Visible = True
    'Else
       ' fraMozo.Visible = False
   ' End If
                    
    If lActivaMotorizado Then
        fraMotorizado.Visible = True
    Else
        fraMotorizado.Visible = False
    End If
    
    Me.fraOrigenVentas.Visible = False
    
    
    
End Function



Private Sub cmdProducto_Click(Index As Integer)
    frmMozoUsuario.ReseteaTimer
    
    If RsCabecera.RecordCount = 0 Then
       Exit Sub
    End If
        
    If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
       MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
       RsCabecera.Requery
       RsDetalle.Requery
       Exit Sub
    End If
    
  '  RsProducto.Requery
    RsProducto.MoveFirst
    RsProducto.Find "nboton = " & Trim(str(Index))
    sProducto = RsProducto!codigo
        
    If validadIngresoProducto(sProducto) = False Then
       Exit Sub
    End If
    
    'Correccion
    ListarOperadoresConFiltro (sProducto)
    If vComanda Then
       sTipo = "Comanda"
       sDescrip = sComanda
       frmNumPad.Caption = "Comanda"
       frmNumPad.Show vbModal
       
       sComanda = IIf(sDescrip = "" Or sDescrip = "0", "", Mid("0000000000", 1, 10 - Len(Trim(sDescrip))) & Trim(sDescrip))
       
      If wComanda And sComanda = "" Then
         MsgBox "No ha ingresado el número de la comanda", vbExclamation, sMensaje
         Exit Sub
      End If
      
      If Not wEnter Then
         Exit Sub
       End If
    End If
    
    If lPrecuenta Then
       sTemp = ""
       Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' and lActivo = 1 Order by Descripcion"
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                       "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
       
       frmBusquedaRapida.nPredeterm = 1
       frmBusquedaRapida.Show vbModal
       If wEnter = True Then
          sMozoD = sCodigo
       Else
          sMozoD = sMozo
       End If
    End If
    nCantidad = 1
    InsertaProducto 1
    If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
       lPropiedad = True
    End If
    
End Sub

Private Sub cmdProductoCombo_Click(Index As Integer)
    If RsCabecera.RecordCount = 0 Then
       Exit Sub
    End If
            
    Dim xxx As String
    xxx = RsProducto.Filter
    RsProducto.Filter = adFilterNone
    'RsProducto.Requery
    RsProducto.MoveFirst
    RsProducto.Find "tResumido = '" & cmdProductoCombo(Index).Caption & "'"
    sProducto = RsProducto!codigo
    
    nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                       "FROM dbo.CPEDIDO WHERE CPEDIDO.tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'", Cn)
    If nCCombo < nCombo * RsDetalle!nCantidad Then
       If lPrecuenta Then
          sTemp = ""
          Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' and lActivo = 1 Order by Descripcion"
          
          Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                          "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
          frmBusquedaRapida.nPredeterm = 1
          frmBusquedaRapida.Show vbModal
          If wEnter = True Then
             sMozoD = sCodigo
          Else
             sMozoD = sMozo
          End If
       End If
       'Oscar Ortega----------------------------------------------
       'Obtener Si es LUnico
       Dim oRsProductoDeCombo As Recordset
       Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sProducto)
       If oRsProductoDeCombo.RecordCount > 0 Then
            If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                'Obtener Suma de cantidades
                Dim nCantidadEnElCombo As Integer
                nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sPedido, sItem, oRsProductoDeCombo!tEtiqueta)
                'Suma de cantidades < que nCantidad
                If nCantidadEnElCombo < nCantidad Then
                    InsertaCombo sProducto
                Else
                    MsgBox "Solo es permitido " & nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                End If
            Else
                InsertaCombo sProducto
            End If
        Else
            InsertaCombo sProducto
        End If
       '----------------------------------------------------------
       'InsertaCombo sProducto
    Else
       MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
    End If
    'RsProducto.Requery
    RsProducto.MoveFirst
    RsProducto.Filter = IIf(xxx = "0", "", xxx)
End Sub

Private Sub cmdPropiedad_Click(Index As Integer)
    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMObra As Double
    Dim ncantidadPropiedad As Double
    Dim Cantidad As Double
    
    'VALIDA ENVIO DE PRODUCTO MOBILE
    Dim xImprimeProducto2 As Boolean
    xImprimeProducto2 = Calcular("Select lImprime as codigo From DPEDIDO Where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'", Cn)
    
    If xImprimeProducto2 = True And wAgregaCombo = False Then
       MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
       cmdOpcion_Click (8)
       Exit Sub
    End If
    
    RsPropiedad.MoveFirst
    RsPropiedad.Find ("Descripcion = '" & cmdPropiedad(Index).Caption & "'")
    
    If Not (RsOperador.EOF Or RsOperador.BOF) Then
        nOperadorPropiedad = Calcular("select isnull(ncontrol,0) as codigo from voperador where codigo='" & RsOperador!codigo & "'", Cn)
     End If
    
    If cmdPropiedad(Index).FontBold = True Then
       cmdPropiedad(Index).FontBold = False
       If Not RsPropiedad.EOF Then
          If wAgregaCombo Then
             Cantidad = Calcular("select isnull(ncantidad,1) as codigo from tcombopropiedad where tcodigopedido='" & sPedido & "' and titem='" & sItem & "' and titemcombo='" & xItem & "' and  tproducto='" & sCombo & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' ", Cn)
             If RsPropiedad!nPrecio <> 0 Then
                txtSuma.Caption = CambiaPrecio(nPVenta - (RsPropiedad!nPrecio * Cantidad), sPedido)
             End If
             Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and tProducto='" & sCombo & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "'"
          Else
             Cantidad = Calcular("select isnull(ncantidad,1) as codigo from tproductopropiedad where tcodigopedido='" & sPedido & "' and titem='" & sItem & "' and tproducto='" & sProducto & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' and tenlace='" & RsPropiedad!tEnlace & "'", Cn)
             If RsPropiedad!nPrecio <> 0 Then
                txtSuma.Caption = CambiaPrecio(nPVenta - (RsPropiedad!nPrecio * Cantidad), sPedido)
             End If
             'Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido = '" & sPedido & "' and tItem = '" & sitem & "' and tProducto='" & sProducto & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "'"
             Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tProducto='" & sProducto & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "' and tEnlace='" & RsPropiedad!tEnlace & "'"
          End If
          If Cantidad <> 1 Then
'             lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", ", "")
             lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & cmdPropiedad(Index).Caption & ": (" & Cantidad & "), ", "")
          Else
             lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & cmdPropiedad(Index).Caption & ", ", "")
          End If
       End If
    Else
        ncantidadPropiedad = 1
        If RsPropiedad!lsolicitacantidad = 1 Or RsPropiedad!lsolicitacantidad = True Then
            sTipo = "Prepintado"
            sCodigo = ncantidadPropiedad
            frmNumPad.Show vbModal
            
            If wEnter And Val(sDescrip) > 0 Then
                ncantidadPropiedad = sDescrip
            End If
        End If

        If nOperadorPropiedad > 0 Then
          If wAgregaCombo Then
             Isql = "SELECT COUNT(dbo.TCOMBOPROPIEDAD.tCodigoPedido) AS codigo " & _
                    "FROM dbo.TCOMBOPROPIEDAD INNER JOIN dbo.TPROPIEDAD ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TCOMBOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
                    "where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and dbo.TCOMBOPROPIEDAD.tProducto='" & sCombo & "'  and tOperador='" & RsOperador!codigo & "'"
             If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                Exit Sub
             End If
          Else
             Isql = "SELECT COUNT(dbo.TPRODUCTOPROPIEDAD.tCodigoPedido) AS codigo FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN " & _
                    "dbo.TPROPIEDAD ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
                    "where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tOperador='" & RsOperador!codigo & "'"
             If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                Exit Sub
             End If
          End If
        End If

        Dim maxped As Double
        maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
        If maxped > 0 Then
            If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + ((RsPropiedad!nPrecio * ncantidadPropiedad) * nCantidad)) Then
              MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
              Exit Sub
            End If
        End If

       cmdPropiedad(Index).FontBold = True
       If Not RsPropiedad.EOF Then
          nInsumo = IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo)
          nGasto = IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto)
          nMObra = IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra)
          If wAgregaCombo Then
             Cn.Execute "Insert into TCOMBOPROPIEDAD values ('" & sPedido & "', '" & sItem & "', '" & xItem & "', '" & RsPropiedad!codigo & "', '" & sCombo & "', '" & RsPropiedad!tEnlace & "', " & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nInsumo * -1, ncantidadPropiedad * nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nGasto * -1, ncantidadPropiedad * nGasto) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nMObra * -1, ncantidadPropiedad * nMObra) & ", 1," & ncantidadPropiedad & " , " & IIf(RsPropiedad!codigo = "9999", nInsumo * -1, nInsumo) & ", " & IIf(RsPropiedad!codigo = "9999", nGasto * -1, nGasto) & " , " & IIf(RsPropiedad!codigo = "9999", nMObra * -1, nMObra) & ")"
             If RsPropiedad!nPrecio <> 0 Then
                txtSuma.Caption = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad), sPedido)
             End If
          Else
             Cn.Execute "Insert into TPRODUCTOPROPIEDAD values ('" & sPedido & "', '" & sItem & "', '" & RsPropiedad!codigo & "', '" & sProducto & "', '" & RsPropiedad!tEnlace & "', " & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nInsumo * -1, ncantidadPropiedad * nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nGasto * -1, ncantidadPropiedad * nGasto) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nMObra * -1, ncantidadPropiedad * nMObra) & ", 1, " & ncantidadPropiedad & ",  " & IIf(RsPropiedad!codigo = "9999", nInsumo * -1, nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", nGasto * -1, nGasto) & "," & IIf(RsPropiedad!codigo = "9999", nMObra * -1, nMObra) & ")"
             If RsPropiedad!nPrecio <> 0 Then
                txtSuma.Caption = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad), sPedido)
             End If
          End If
          If ncantidadPropiedad <> 1 Then
                lblResumen.Text = lblResumen.Text & RsOperador!descripcion & " " & cmdPropiedad(Index).Caption & ": (" & ncantidadPropiedad & "), "
          Else
                lblResumen.Text = lblResumen.Text & RsOperador!descripcion & " " & cmdPropiedad(Index).Caption & ", "
          End If
       End If
    End If
                   
    If wAgregaCombo Then
       RsComboPropiedad.Requery
    Else
       RsProductoPropiedad.Requery
    End If
End Sub


Private Sub cmdPunto_Click(Index As Integer)
  
   Dim i As Integer
   For i = 1 To 9
       cmdPunto(i).FontBold = False
   Next i
   cmdPunto(Index).FontBold = True
   rsPuntoVenta.MoveFirst
   rsPuntoVenta.Find "Descripcion = '" & cmdPunto(Index).Caption & "'"
   txtPuntoVenta.Caption = rsPuntoVenta!descripcion
   sPuntoVenta = rsPuntoVenta!codigo
  
End Sub

Private Sub cmdSubGrupo_Click(Index As Integer)
   Dim i As Integer
   frmMozoUsuario.ReseteaTimer
   Screen.MousePointer = vbHourglass
   For i = 1 To 7
       cmdSubGrupo(i).backColor = vbButtonFace
   Next i
   
   RsSubgrupo.MoveFirst
   RsSubgrupo.Find "nboton = " & Trim(str(Index))
   xSubGrupo = RsSubgrupo!codigo
   cmdSubGrupo(RsSubgrupo!nBoton).backColor = vbRed
   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
   AsignaBotonProducto 20, RsProducto, cmdProducto(), sTipoPedido, sUnidadNegocio
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
On Error GoTo fin
    If lBuscarPedidoBarraMesero Then
        Me.txtbuscarpedidobarra.SetFocus
    Else
            On Error Resume Next
            Me.txtBarra.SetFocus
    End If
    Exit Sub
fin:
'MsgBox "Error: " & error
End Sub

Private Sub Form_Load()
   wSalir = False
   txtMozo.Caption = sVar1
   sUsuario = sVar1
   tNombreMozo = sVar1
   frmCargoMozo.Caption = " Caja : " & sCaja
   muestra = Me.Caption
   sComanda = ""
   sMozoD = ""
   sTurno = "MOZO"
   lEstadoBuscarPedido = False
   If MESA247 Then
    ActivaMesa247 (True)
   Else
    ActivaMesa247 (False)
   End If
   'cmdCabecera(10).Enabled = False
   
   cmdCabecera(6).Caption = sBoton1
   cmdCabecera(7).Caption = sBoton2
   cmdCabecera(8).Caption = sBoton3
   cmdCabecera(14).Caption = sBoton4
   cmdCabecera(16).Caption = sBoton5
   
   cmdCabecera(6).Enabled = IIf(sBoton1 <> "", True, False)
   cmdCabecera(7).Enabled = IIf(sBoton2 <> "", True, False)
   cmdCabecera(8).Enabled = IIf(sBoton3 <> "", True, False)
   cmdCabecera(14).Enabled = IIf(sBoton4 <> "", True, False)
   cmdCabecera(16).Enabled = IIf(sBoton5 <> "", True, False)
   
   If lBloqueoCanal1 Then
    cmdCabecera(6).Enabled = False
   End If
    If lBloqueoCanal2 Then
    cmdCabecera(7).Enabled = False
   End If
    If lBloqueoCanal3 Then
    cmdCabecera(8).Enabled = False
   End If
    If lBloqueoCanal4 Then
    cmdCabecera(14).Enabled = False
   End If
   If lBloqueoCanal5 Then
    cmdCabecera(16).Enabled = False
   End If
   
   If lBloqueaPrecuenta Then
    cmdOpcion(4).Enabled = False
   Else
    cmdOpcion(4).Enabled = True
   End If
      
   Label1(8).Caption = Mesero
   
   If lBusquedaReservas Then
        Me.cmdOpcion(9).Visible = True
        
        If lImportacionReservas Then
            Me.cmdOpcion(9).Enabled = True
        Else
            Me.cmdOpcion(9).Enabled = False
        End If
   Else
        Me.cmdOpcion(9).Visible = False
   End If
   
   'Cabecera
   If lSomelier Then
      Isql = "SELECT * from vPedidoCabecera " & _
             "where tEstadoPedido = '01' and tTurno = 'MOZO' order by Codigo"
   Else
      Isql = "SELECT * from vPedidoCabecera " & _
             "where tEstadoPedido = '01' and tTurno = 'MOZO' and tMozo='" & sMozo & "' order by Codigo"
   End If
   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
   
   If lEnvioAutomatico Then
      Isql = "SELECT * from vPedidoCabecera " & _
          "where tEstadoPedido = '01' and tCaja = '" & sCaja & "' and FecProg is not null and isnull(nMinutosAntesEnvio,0) <> 0 order by Codigo"
      Set RsTimer = Lib.OpenRecordset(Isql, Cn)
      
      'Timer.Enabled = True
      'Timer.Interval = 1000
   End If
   
     ' origenes de venta
  Me.fraOrigenVentas.Visible = False
     'Origen de ventas
   Isql = "select * from vOrigenVenta where Activo = 1 and Visible = 1 and CodOrigenVenta not in (select CodigoOrigenVenta from TCAJAORIGEN_BLOQUEO  where tcaja='" & sCaja & "' ) Order by Boton"
   Set RsOrigenVentas = Lib.OpenRecordset(Isql, Cn)
   
   Isql = "select * from vTipoPedido where Codigo = '02'"
   Set RscanalOrigenVentas = Lib.OpenRecordset(Isql, Cn)
    AsignaBotonOrigenVentas 19, RsOrigenVentas, Me.cmdOrigen()
   '-----------------------------------------------------
  '----------------------------------
  
   'Detalle
   If lSomelier Then
      Isql = "select * from vPedidoDetalle " & _
             "where tEstadoItem ='N' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tTurno = 'MOZO') " & _
             "ORDER BY tCodigoPedido, tItem"
   Else
      Isql = "select * from vPedidoDetalle " & _
             "where tEstadoItem ='N' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tTurno = 'MOZO' and tMozo='" & sMozo & "') " & _
             "ORDER BY tCodigoPedido, tItem"
   End If
   Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
   If lAccesoDespachoPedido = False Then
        cmdOpcion(7).Enabled = False
   End If
            
            
   'Grupo
   Isql = "select * from vGrupo where lActivo = 1 order by nBoton"
   Set RsGrupo = Lib.OpenRecordset(Isql, Cn)
   If RsGrupo.RecordCount = 0 Then
      MsgBox "Error: Se necesita al menos un Grupo creado", vbCritical, sMensaje
      End
   End If
   
   'Sub Grupo
   Isql = "select * from vSubGrupo where lActivo = 1 Order by nBoton"
   Set RsSubgrupo = Lib.OpenRecordset(Isql, Cn)
   If RsSubgrupo.RecordCount = 0 Then
      MsgBox "Error: Se necesita al menos un SubGrupo creado", vbCritical, sMensaje
      End
   End If
      
     'Producto
   
   'Isql = "select * from vProducto where lActivo = 1 Order by nBoton"
   Set RsProducto = Lib.OpenRecordset("[usp_Inforest_ObtieneProductos]", Cn)
   'Set RsProducto = Lib.OpenRecordset(Isql, Cn)
   If RsProducto.RecordCount = 0 Then
      MsgBox "Error: Se necesita al menos un producto creado", vbCritical, sMensaje
      End
   End If
                                    
   'Motorizado
   Isql = "select * from vMotorizado where lActivo = 1 Order by nBoton"
   Set RsMotorizado = Lib.OpenRecordset(Isql, Cn)
      
   'Operador Correccion
   ListarOperadoresConFiltro (sProducto)
      
   'Propiedades
   Dim xSql As String
   If lAlmacen Then
      Dim RsOp As Recordset
      Set RsOp = Lib.OpenRecordset("select Codigo, Descripcion from vOperador where lStockMenos=1", Cn)
      If RsOp.RecordCount > 0 Then
        'CASO CORP.FERNANDEZ
        If BDLink = "1" Then
            xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                   "nInsumo, nGasto, nManoObra, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                   "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                   "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                   "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!descripcion & "' as Operador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0,0 " & _
                   "FROM " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.DRECETAVENTA t1x INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t1x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto " & _
                   "Where lNoDescargo = 1 and " & "t1x.tLocal='" & sLocal & "'"
        Else
             xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                   "nInsumo, nGasto, nManoObra, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                   "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                   "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                   "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!descripcion & "' as Operador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0,0 " & _
                   "FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA t1x INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t1x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto " & _
                   "Where lNoDescargo = 1 and " & "t1x.tLocal='" & sLocal & "'"
            End If
      Else
         xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                "nInsumo, nGasto, nManoObra, tOperador.tDetallado AS Operador, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
      End If
   Else
         xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
   End If
   'CASO CORP.FERNANDEZ
   If BDLink = "1" Then
        Set RsPropiedad = Lib.OpenRecordset("usp_Inforest_ObtienePropiedad '" & IIf(lAlmacen = True, 1, 0) & "','" & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & "','" & sLocal & "'", Cn)
    Else
        Set RsPropiedad = Lib.OpenRecordset("usp_Inforest_ObtienePropiedad '" & IIf(lAlmacen = True, 1, 0) & "','" & sAlmacenMDB & "','" & sLocal & "'", Cn)
    End If
      
    If lBuscarPedidoBarraMesero Then
    
       'Producto Propiedad
        Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, t1.Descripcion AS Descripcion, Operador,isnull(tproductopropiedad.ncantidad,1) ncantidad " & _
                    " FROM    dbo.TOPERADOR INNER JOIN dbo.TPRODUCTOPROPIEDAD INNER JOIN (" & xSql & ")T1 ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TPRODUCTOPROPIEDAD.tProducto = T1.tProducto AND dbo.TPRODUCTOPROPIEDAD.tEnlace = T1.tEnlace ON dbo.TOPERADOR.tOperador = T1.tOperador COLLATE Modern_Spanish_CI_AS LEFT OUTER JOIN dbo.TPROPIEDAD ON dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador AND dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
              " where tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido in ('01','02') AND fregistro>getdate()-1  ) order by toperador.nboton "
    
       Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
          
       'Combo Propiedad
       Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, T1.Descripcion, T1.Operador, isnull(tcombopropiedad.ncantidad,1) ncantidad  " & _
              " FROM         dbo.TOPERADOR INNER JOIN dbo.TPROPIEDAD ON dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador RIGHT OUTER JOIN dbo.TCOMBOPROPIEDAD INNER JOIN  (" & xSql & ") T1 ON  dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TCOMBOPROPIEDAD.tProducto = T1.tProducto AND                       dbo.TCOMBOPROPIEDAD.tEnlace = T1.tEnlace ON dbo.TOPERADOR.tOperador = T1.tOperador COLLATE Modern_Spanish_CI_AS AND dbo.TPROPIEDAD.tCodigoPropiedad = dbo.TCOMBOPROPIEDAD.tCodigoPropiedad AND dbo.TPROPIEDAD.tProducto = dbo.TCOMBOPROPIEDAD.tProducto  " & _
              " where tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido in ('01','02') AND fregistro>getdate()-1 ) order by toperador.nboton"
       Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
    
    Else
    
       'Producto Propiedad
        Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, t1.Descripcion AS Descripcion, Operador,isnull(tproductopropiedad.ncantidad,1) ncantidad " & _
                    " FROM    dbo.TOPERADOR INNER JOIN dbo.TPRODUCTOPROPIEDAD INNER JOIN (" & xSql & ")T1 ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TPRODUCTOPROPIEDAD.tProducto = T1.tProducto AND dbo.TPRODUCTOPROPIEDAD.tEnlace = T1.tEnlace ON dbo.TOPERADOR.tOperador = T1.tOperador COLLATE Modern_Spanish_CI_AS LEFT OUTER JOIN dbo.TPROPIEDAD ON dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador AND dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
              " where tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tTurno = 'MOZO' ) order by toperador.nboton "
    
       Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
          
       'Combo Propiedad
       Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, T1.Descripcion, T1.Operador, isnull(tcombopropiedad.ncantidad,1) ncantidad  " & _
              " FROM         dbo.TOPERADOR INNER JOIN dbo.TPROPIEDAD ON dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador RIGHT OUTER JOIN dbo.TCOMBOPROPIEDAD INNER JOIN  (" & xSql & ") T1 ON  dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TCOMBOPROPIEDAD.tProducto = T1.tProducto AND                       dbo.TCOMBOPROPIEDAD.tEnlace = T1.tEnlace ON dbo.TOPERADOR.tOperador = T1.tOperador COLLATE Modern_Spanish_CI_AS AND dbo.TPROPIEDAD.tCodigoPropiedad = dbo.TCOMBOPROPIEDAD.tCodigoPropiedad AND dbo.TPROPIEDAD.tProducto = dbo.TCOMBOPROPIEDAD.tProducto  " & _
              " where tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tTurno = 'MOZO') order by toperador.nboton"
       Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
    
    End If
      



    Set RsProductoCombo = Lib.OpenRecordset("usp_Inforest_ObtieneCombos", Cn)
   'Motivo de Eliminacion
   Isql = "select * from vMotivoEliminacion where lActivo = 1 order by Codigo"
   Set RsMotivoEliminacion = Lib.OpenRecordset(Isql, Cn)
   AsignaComando 38, RsMotivoEliminacion, cmdEliminacion()

   'Areas
   Set RsArea = Lib.OpenRecordset("select * from vAreaImpresora where tCaja ='" & sCaja & "'", Cn)
                                                                                                                                                          
   'Canales de Venta
   Set RsCanalesVenta = Lib.OpenRecordset("select * from vTipoPedido", Cn)
   If Not RsCanalesVenta.RecordCount > 0 Then
      MsgBox "Debe tener configurado al menos un Canal de Venta", vbCritical + vbOKOnly, sMensaje
      Unload Me
   End If
                                                                                                                                                          
   'Combo
   'GCCA 02082021 - REQ REST-N014-2021 - TORI
   If (lNoVisualizarItemCombos) Then
        Isql = "SELECT * from vPedidoCombox " & _
               "WHERE tEstadoPedido='01' " & _
               "ORDER BY tCodigoPedido, tItemCombo"
        Set RsCombo = Lib.OpenRecordset(Isql, Cn)
    Else
         Isql = "SELECT * from vPedidoCombo " & _
               "WHERE tEstadoPedido='01' " & _
               "ORDER BY tCodigoPedido, tItemCombo"
        Set RsCombo = Lib.OpenRecordset(Isql, Cn)
    End If
    
   'Configuración de Grillas
   Call ConfGrilla(3, grdCabecera, "Pax", 2, "nAdulto", 430, 0, 0, "", _
                                   "Mesa", 2, "Mesa", 1000, 0, 0, "", _
                                   "Observacion", 2, "tObservacion", 1200, 0, 0, "")
   
   If (lActivaTresDecimales) Then
   Call ConfGrilla(10, grdDetalle, "Or", 2, "nOrden", 300, 1, 0, "#0", _
                                   "-", 2, "lCorte", 250, 0, 4, "", _
                                   "Producto", 2, "Producto", 1820, 0, 0, "", _
                                   "Precio", 2, "nPrecioVenta", 700, 1, 0, "###,###,##0.00", _
                                   "Cant.", 2, "nCantidad", 590, 1, 0, "##0.000", _
                                   "SubTotal", 2, "nVenta", 790, 1, 0, "###,###,##0.00", _
                                   "F", 2, "tFacturado", 250, 0, 0, "", _
                                   "E", 2, "lImprime", 250, 0, 4, "", _
                                   "P", 2, "lPropiedad", 250, 0, 4, "", _
                                   "O", 2, "lObservacion", 250, 0, 4, "")
   Else
      Call ConfGrilla(10, grdDetalle, "Or", 2, "nOrden", 300, 1, 0, "#0", _
                                   "-", 2, "lCorte", 250, 0, 4, "", _
                                   "Producto", 2, "Producto", 1820, 0, 0, "", _
                                   "Precio", 2, "nPrecioVenta", 700, 1, 0, "###,###,##0.00", _
                                   "Cant.", 2, "nCantidad", 590, 1, 0, "###0.00", _
                                   "SubTotal", 2, "nVenta", 790, 1, 0, "###,###,##0.00", _
                                   "F", 2, "tFacturado", 250, 0, 0, "", _
                                   "E", 2, "lImprime", 250, 0, 4, "", _
                                   "P", 2, "lPropiedad", 250, 0, 4, "", _
                                   "O", 2, "lObservacion", 250, 0, 4, "")
   
   End If
   Call ConfGrilla(7, grdCombo, "-", 2, "lCorte", 250, 0, 4, "", _
                                "Producto", 2, "Producto", 1950, 0, 0, "", _
                                "Cant.", 2, "nCantidad", 650, 1, 0, "#,##0.00", _
                                "E", 2, "lImprime", 250, 0, 4, "", _
                                "P", 2, "lPropiedad", 250, 0, 4, "", _
                                "O", 2, "lObservacion", 250, 0, 4, "", _
                                "Ord", 2, "nOrden", 400, 1, 0, "#0")
                                     
   Set grdCabecera.DataSource = RsCabecera
   Set grdDetalle.DataSource = RsDetalle
   Set grdCombo.DataSource = RsCombo
      
   'Asignar Grupo
   AsignaBoton 5, RsGrupo, cmdGrupo()
   RsGrupo.MoveFirst
   RsGrupo.Find "Codigo='" & sGrupoDefault & "'"
   If RsGrupo.EOF Then
      MsgBox "Error: Se necesita configurar en caja, un Grupo por defecto", vbCritical, sMensaje
      End
   End If
   If Not RsGrupo.EOF And RsGrupo!nBoton > 0 Then
      cmdGrupo(RsGrupo!nBoton).backColor = vbBlue
   Else
      RsGrupo.MoveFirst
      RsGrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vGrupo where nBoton>0 and lActivo=1", Cn) & "'"
      cmdGrupo(RsGrupo!nBoton).backColor = vbBlue
   End If
      
   xGrupo = RsGrupo!codigo
   RsSubgrupo.Filter = "tGrupo = '" & xGrupo & "'"
         
   'Asignar Subgrupo
   AsignaBoton 7, RsSubgrupo, cmdSubGrupo()
   RsSubgrupo.MoveFirst
   RsSubgrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vSubGrupo where nBoton>0 and tGrupo = '" & xGrupo & "' and lActivo=1", Cn) & "'"
   If RsSubgrupo.EOF Or RsSubgrupo!nBoton = 0 Then
      MsgBox "Error: Se necesita configurar un SubGrupo con Boton", vbCritical, sMensaje
      End
   End If
   cmdSubGrupo(RsSubgrupo!nBoton).backColor = vbRed
   xSubGrupo = RsSubgrupo!codigo
   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
            
   'Asignar Producto
   sTipoPedido = sTipoPedidoPD
   AsignaBotonProducto 20, RsProducto, cmdProducto(), sTipoPedido, sUnidadNegocio
   AsignaBoton 19, RsMotorizado, cmdMotorizado()
   
   'Asigna Operador
   AsignaBoton 13, RsOperador, cmdOperador()
   If RsOperador.RecordCount > 0 Then
      RsOperador.MoveFirst
      If Not IsNull(RsOperador!nBoton) And RsOperador!nBoton > 0 Then
         cmdOperador_Click (RsOperador!nBoton)
      End If
   End If
   
   fraCabecera.Visible = False
   fraPlato.Visible = False
   fraPropiedad.Visible = False
   fraEliminacion.Visible = False
   ActivaFrame fraPlato
   wDetalle = True
   wAgrega = False
   wAgregaCombo = False
   lPropiedad = False
   
   If RsCabecera.RecordCount = 0 Then
      sPedido = ""
      LimpiaCabecera
   Else
      RsCabecera.MoveFirst
      sPedido = RsCabecera!codigo
      ActivaFrame fraPlato
   End If
   
   'nuevoo
   
   If Not lInfhotel Then
     fraPuntoVenta.Visible = False
     txtPuntoVenta.Visible = False
     
   Else
   
      'Punto de Venta
      Isql = "Select tPuntoVenta as Codigo, tDescripcion as Descripcion, nUltimoComanda, tmoneda" & _
             " From tPuntoVenta " & _
             " where tHotel='" & sHotel & "' AND lActivo=1 and lInforest=1"
      Set rsPuntoVenta = Lib.OpenRecordset(Isql, CnInfhotel)
   
      If rsPuntoVenta.RecordCount = 0 Then
         MsgBox "Error: Se necesita configurar los puntos de venta en Infhotel", vbCritical, sMensaje
         End
      End If
   
      Dim i As Integer
      For i = 10 To 19
          
          cmdMotorizado(i).Visible = False
      Next i
      
      txtFechaProg.Visible = False
      txtPuntoVenta.Visible = True
      Label1(25).Caption = "Pto. Venta :"
      cmdBuscar(4).Top = 1050
      fraMotorizado.Top = 3105
      fraMotorizado.Height = 1865
      AsignaBoton 9, RsMotorizado, cmdMotorizado()
      AsignaComando 9, rsPuntoVenta, cmdPunto()
      
   End If
      '-- cover pedido----
  ' If lcover Then
        cmdCabecera(21).Visible = True
    'Else
     '   cmdCabecera(21).Visible = False
   'End If
   '------------------------
    ActivaMesa247 (True)
  
   lblmoneda.Caption = sMonN
   
   lblFecha.Caption = Format(FechaServidor(), "dddd, dd MMMM yyyy")
   
   If lBuscarPedidoBarraMesero Then
     cmdBuscarPedido.Visible = True
   Else
     cmdBuscarPedido.Visible = False
   End If
   
   Screen.MousePointer = vbDefault
   
End Sub

Public Sub AsignaPedido()
   Screen.MousePointer = vbHourglass
   
   'Llena los valores
   If RsCabecera.RecordCount = 0 Then
      sTipoPedido = "01"
      sPedido = ""
      LimpiaCabecera
   Else
         
      txtEnvioAntes.Text = IIf(IsNull(RsCabecera!nMinutosAntesEnvio), 0, RsCabecera!nMinutosAntesEnvio)
      
      sPedido = IIf(IsNull(RsCabecera!codigo), "", RsCabecera!codigo)
      sMesa = IIf(IsNull(RsCabecera!tMesa), "", RsCabecera!tMesa)
      'grdDetalle.Caption = "Pedidos de " & IIf(sMesa = "000" Or Len(Trim(sMesa)) = 0, Mid(grdCabecera.Columns(2).Text, 1, 50), grdCabecera.Columns(1).Text)
       If lcover Then
         grdDetalle.Caption = "Pedidos de " & IIf(sMesa = "000" Or Len(Trim(sMesa)) = 0, Mid(grdCabecera.Columns(2).Text, 1, 50), grdCabecera.Columns(1).Text) & " / Cover : " & Val(RsCabecera!nAdulto) * sMontoMinCover
      Else
         grdDetalle.Caption = "Pedidos de " & IIf(sMesa = "000" Or Len(Trim(sMesa)) = 0, Mid(grdCabecera.Columns(2).Text, 1, 50), grdCabecera.Columns(1).Text)
      End If
            'invitado2013
      txtObservacion.Caption = ""
      'invitado2013
      'pariente2013
      txtCliente.Caption = ""
      'pariente2013
      sMotorizado = IIf(IsNull(RsCabecera!tMotorizado), "0000", RsCabecera!tMotorizado)
      sCliente = IIf(IsNull(RsCabecera!tClienteDelivery), "", RsCabecera!tClienteDelivery)
      sTipoAtencion = IIf(IsNull(RsCabecera!tTipoAtencion), "01", RsCabecera!tTipoAtencion)
      sTipoPedido = IIf(IsNull(RsCabecera!tTipoPedido), "01", RsCabecera!tTipoPedido)
      nCorrela = IIf(IsNull(RsCabecera!nCorrelativo), "01", RsCabecera!nCorrelativo)
      
      frmVenta.txtAdulto.Text = IIf(IsNull(RsCabecera!nAdulto), 0, RsCabecera!nAdulto)
      frmVenta.txtNino.Text = IIf(IsNull(RsCabecera!nNino), 0, RsCabecera!nNino)
      frmVenta.txtJuntar.Text = IIf(IsNull(RsCabecera!nMesa), 0, RsCabecera!nMesa)
   
      fraPedido.Caption = " Pedido " & IIf(RsCabecera.RecordCount = 0, "", IIf(IsNull(RsCabecera!codigo), "", RsCabecera!codigo)) & " "
      txtFecha.Caption = IIf(IsNull(RsCabecera!fFecha), "", RsCabecera!fFecha)
      txtTipoPedido.ForeColor = IIf(RsCabecera!lPrioridad, vbRed, &H800000)
'      txtTipoPedido.Caption = IIf(IsNull(RsCabecera!TipoPedido), "", RsCabecera!TipoPedido)
      Dim DescripcionOrigen As String
      DescripcionOrigen = Calcular("select isnull(Descripcion,'') as codigo from vOrigenVenta where CodOrigenVenta='" & vOrigenVentas & "'", Cn)
      txtTipoPedido.Caption = IIf(IsNull(RsCabecera!TipoPedido), "", RsCabecera!TipoPedido) & IIf(DescripcionOrigen = "0" Or DescripcionOrigen = "", "", " / " & DescripcionOrigen)
      
      
      txtFechaProg.Caption = IIf(IsNull(RsCabecera!fProgramacion), "", RsCabecera!fProgramacion)
      txtMozo.Caption = IIf(IsNull(RsCabecera!Mozo), "", RsCabecera!Mozo)
      txtMotorizado.Caption = IIf(IsNull(RsCabecera!Motorizado), "", RsCabecera!Motorizado)
      txtMontoMax.Caption = IIf(IsNull(RsCabecera!nMontoMaximo), 0, RsCabecera!nMontoMaximo)
      'pariente'
      sCodigoParienteSeleccionado = IIf(IsNull(RsCabecera!tcodigopariente), "", RsCabecera!tcodigopariente)
      'entregarA
      txtEntregar.Caption = IIf(IsNull(RsCabecera!tentregara), "", RsCabecera!tentregara)
            
      If sCodigoParienteSeleccionado <> "" Then

        txtCliente.Caption = Calcular("select pariente + ' ' +case when lconyugue=1 then ' Conyugue de ' else 'Hijo de ' end as codigo from vPariente where tCodigoPariente='" & sCodigoParienteSeleccionado & "' and tCodigoDelivery='" & sCliente & "'", Cn)

      End If
            txtCliente.Caption = txtCliente.Caption & " " & IIf(IsNull(RsCabecera!Cliente), "", RsCabecera!Cliente) & IIf(IsNull(RsCabecera!Tienda), "", " - " & RsCabecera!Tienda)

      'pariente2013
      txtDireccion.Caption = IIf(IsNull(RsCabecera!direccion), "", RsCabecera!direccion)
      If lClub Then
      txtTelefono.Caption = sCliente
      Else
      txtTelefono.Caption = IIf(IsNull(RsCabecera!Telefono), "", RsCabecera!Telefono)
      End If
      txtObservacion.Caption = IIf(IsNull(RsCabecera!tObservacion), "", RsCabecera!tObservacion)
      xDescuento = IIf(IsNull(RsCabecera!nDescuento), 0, RsCabecera!nDescuento)
      sDescripcionDescuento = IIf(IsNull(RsCabecera!tObservacionDescuento), 0, RsCabecera!tObservacionDescuento)
      sCodigoDescuento = IIf(IsNull(RsCabecera!tDescuento), 0, RsCabecera!tDescuento)
      txtDescuento.Caption = Format(Calcular("select sum(nDescuento*nCantidad) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
      'invitado2013
      sCodigoInvitado = IIf(IsNull(RsCabecera!tcodigoinvitado), "", RsCabecera!tcodigoinvitado)
      'invitado2013
      'MODIFICAR DESCUENTO
      sAutoriza = IIf(IsNull(RsCabecera!tUsuarioDescuento), 0, RsCabecera!tUsuarioDescuento)
      tAutorizaDescuento = IIf(IsNull(RsCabecera!tUsuarioDescuento), 0, RsCabecera!tUsuarioDescuento)
      
      '--- MontoMinimo pedido-------
      xMontoMaximo = IIf(IsNull(RsCabecera!nMontoMaximo), 0, RsCabecera!nMontoMaximo)
      '-----------------------------
       vOrigenVentas = Calcular("select isnull(codigoorigenventa,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
       
      If lInfhotel Then
         sPuntoVenta = IIf(IsNull(RsCabecera!tPuntoVenta), "", RsCabecera!tPuntoVenta)
         
         rsPuntoVenta.MoveFirst
         rsPuntoVenta.Find "Codigo='" & sPuntoVenta & "'"
         
         If Not rsPuntoVenta.EOF Then
            txtPuntoVenta.Caption = rsPuntoVenta!descripcion
         Else
            txtPuntoVenta.Caption = ""
         End If
      End If
                        
      cmdCabecera(9).FontBold = IIf(txtTipoPedido.ForeColor = True, True, False)
      cmdCabecera(6).FontBold = IIf(sTipoPedido = "01", True, False)
      cmdCabecera(7).FontBold = IIf(sTipoPedido = "02", True, False)
      cmdCabecera(8).FontBold = IIf(sTipoPedido = "03", True, False)
      cmdCabecera(14).FontBold = IIf(sTipoPedido = "04", True, False) 'OO
      cmdCabecera(16).FontBold = IIf(sTipoPedido = "05", True, False) 'OO
      
      If sMesa <> "" Then
         If Calcular("select nValor as Codigo from vSalon where Codigo in(select tSalon from tmesa where tCodigoMesa='" & sMesa & "')", Cn) > 0 Then
            lPrecuenta = True
            sMozoD = ""
         Else
            lPrecuenta = False
            sMozoD = sMozo
         End If
      Else
         lPrecuenta = False
         sMozoD = sMozo
      End If
   End If
      
   RsDetalle.Filter = "tCodigoPedido = '" & sPedido & "'"
   
   'GCCA 03082021 - REQ REST-N014-2021 - TORI
   If (lNoVisualizarItemCombos) Then
        RsCombo.Filter = "tCodigoPedido = '" & sPedido & "' and tItem='" & sItem & "' and automatico='" & 0 & "'"
   Else
        RsCombo.Filter = "tCodigoPedido = '" & sPedido & "' and tItem='" & sItem & "'"
   End If
   
   If RsDetalle.RecordCount <> 0 Then
      RsDetalle.MoveFirst
      sProducto = IIf(IsNull(RsDetalle!tCodigoProducto), "", RsDetalle!tCodigoProducto)
      sSubGrupo = IIf(IsNull(RsDetalle!tCodigoSubGrupo), "", RsDetalle!tCodigoSubGrupo)
      sItem = RsDetalle!tItem
   Else
      sProducto = ""
      wCombo = False
      nCombo = 0
      sItem = ""
      
      ActivaCabecera True
      ActivaOpcion True
      ActivaFrame fraPlato
   End If
   
   'Imprime el Total
   txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
   AsignaBotonProducto 20, RsProducto, cmdProducto(), sTipoPedido, sUnidadNegocio
   Screen.MousePointer = vbDefault
      
    On Error GoTo fin
     If lBuscarPedidoBarraMesero Then
        txtbuscarpedidobarra.SetFocus
     Else
        On Error Resume Next
        Me.txtBarra.SetFocus
     End If
     
   Exit Sub
fin:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   frmMozoUsuario.ReseteaTimer
   
  On Error GoTo fin
     If lBuscarPedidoBarraMesero Then
        txtbuscarpedidobarra.SetFocus
     Else
       On Error Resume Next
       Me.txtBarra.SetFocus
     End If
fin:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsT3 = Nothing
    Set RsCabecera = Nothing
    Set RsDetalle = Nothing
    Set RsCombo = Nothing
    Set RsImpresion = Nothing
    Set RsGrupo = Nothing
    Set RsSubgrupo = Nothing
    Set RsProducto = Nothing
    Set RsMotorizado = Nothing
    Set RsArea = Nothing
    Set RsPropiedad = Nothing
    Set frmCargoMozo = Nothing
    Set RsProductoCombo = Nothing
              registroAccesoAuditoria "S", tNombreMozo

End Sub


Private Sub grdCabecera_DblClick()
On Error GoTo fin

    If MESA247 Then
        If RsCabecera.RecordCount > 0 Then
           sPedido = RsCabecera!codigo
         Else
           Exit Sub
         End If
        If Calcular("select count(*) as codigo from mpedido where isnull(idpedidomesa,'') <> '' and tcodigopedido='" & sPedido & "'", Cn) > 0 Then 'Calcular("SELECT CASE WHEN COUNT(*)>0 THEN 1 ELSE CASE WHEN (SELECT COUNT(*)  FROM MPEDIDO M WHERE M.TCODIGOPEDIDO='" & sPedido & "' AND isnull(M.idPedidoMesa,'')='' )>0 THEN 1 ELSE 0 END END  AS CODIGO FROM MPEDIDO MP INNER JOIN DPEDIDO DP ON MP.tCodigoPedido=DP.tCodigoPedido WHERE  MP.tcodigopedido='" & sPedido & "'", Cn) = 0 Then
           If MsgBox("Retornar Checking MESA24/7", vbYesNo) = vbYes Then
            If lPasswordC Then
               If Supervisor("01") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
            Isql = " exec usp_WS_MESA_Inforest '" & sPedido & "','',4"
'            Isql = "update mpedido set tTurno = '' , tCaja = '' ,tMozo=''  where tCodigoPedido ='" & sPedido & "'"
'            Cn.Execute Isql
'            Isql = "UPDATE TMESA SET TESTADOMESA='01' WHERE TCODIGOMESA= (SELECT TMESA FROM MPEDIDO WHERE TCODIGOPEDIDO ='" & sPedido & "')"
            Cn.Execute Isql
            RsCabecera.Requery
           End If
           
        End If
        
    End If
fin:

End Sub

Private Sub grdCabecera_HeadClick(ByVal ColIndex As Integer)
    Screen.MousePointer = vbHourglass
    frmMozoUsuario.ReseteaTimer
    RsCabecera.Sort = grdCabecera.Columns(ColIndex).DataField & " ASC"
    RsCabecera.Requery
    AsignaPedido
    Screen.MousePointer = vbDefault
End Sub

Public Sub ActivaCabecera(Activa As Boolean)
   cmdCabecera(0).Visible = Activa
   cmdCabecera(1).Visible = Activa
   cmdCabecera(2).Visible = Activa
   cmdCabecera(3).Visible = Activa
   cmdCabecera(4).Visible = Activa
   cmdCabecera(5).Visible = Activa
   
   cmdDetalle(0).Visible = Activa
   cmdDetalle(1).Visible = Activa
   cmdDetalle(2).Visible = Activa
   cmdDetalle(3).Visible = Activa
   cmdDetalle(4).Visible = Activa
   'cmdDetalle(5).Visible = Activa
   cmdDetalle(6).Visible = Activa
   cmdDetalle(7).Visible = Activa
   cmdDetalle(8).Visible = Activa
   
   cmdNavegar(0).Visible = Activa
   cmdNavegar(1).Visible = Activa
   cmdNavegar(2).Visible = Activa
   cmdNavegar(3).Visible = Activa
   cmdNavegar(4).Visible = Activa
   cmdNavegar(5).Visible = Activa
   cmdNavegar(6).Visible = Activa
   cmdNavegar(7).Visible = Activa
   cmdNavegar(8).Visible = Activa
   cmdNavegar(9).Visible = Activa
   cmdNavegar(10).Visible = Activa
   cmdNavegar(11).Visible = Activa

End Sub

Public Sub ActivaFrame(Frame As Frame)
    fraCabecera.Visible = False
    fraPlato.Visible = False
    fraPropiedad.Visible = False
    fraEliminacion.Visible = False
    fraProductoCombo.Visible = False
    Frame.Visible = True
End Sub

Public Sub ActivaOpcion(Activa As Boolean)
    cmdOpcion(0).Visible = Activa
    cmdOpcion(7).Visible = Activa

    cmdOpcion(1).Enabled = Not Activa
    cmdOpcion(2).Enabled = Not Activa
    
    cmdOpcion(1).Visible = Not Activa
    cmdOpcion(2).Visible = Not Activa
    cmdOpcion(3).Visible = IIf(lPrinter, Activa, False)
    cmdOpcion(4).Visible = Activa
    cmdOpcion(5).Visible = Activa
    cmdOpcion(6).Visible = Activa
    cmdOpcion(9).Visible = Activa
    cmdOpcion(10).Visible = Activa
    cmdOpcion(11).Visible = Activa
    cmdOpcion(12).Visible = Activa
    cmdOpcion(13).Visible = Activa
End Sub

Public Sub InsertaProducto(xCantidad As Double)
    Dim precioventa As Double
    Screen.MousePointer = vbHourglass
    Dim nValor As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    Dim RsOrd As Recordset
    Dim nOrden As Integer
    
    ' variables MULTIAREAPRODUCCION
        Dim lProductoMultiArea As Boolean
        Dim tsubalmacen As String
        Dim tAreaProduccion As String
        'CPvalicacion central d pedido LG
   ' Dim codigoClienteF As String
    Dim lClienteExcluyeProducto As Boolean
    Dim lProductoPermiteDescuento As Boolean
    Dim lClienteControlaProducto As Boolean
  '  codigoClienteF = Calcular("select tclientedelivery as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
    If lActivaAnticipo Then
      If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
          MsgBox "No es posible modificar/elimnar el producto, anticipo aplicado al pedido!!!", vbExclamation, sMensaje
           Screen.MousePointer = vbDefault
          Exit Sub
      End If
    End If
    If sCliente <> "" Then ' verificamos si tiene cliente frecuente
        If Calcular("select count(tcodigodelivery) as codigo from TCLIENTEPRODUCTO where tcodigoDelivery='" & sCliente & "'", Cn) > 0 Then
            lClienteExcluyeProducto = Calcular("select isnull(lexcluyeproductos,0) as codigo from tdelivery where tcodigodelivery='" & sCliente & "' ", Cn)
            If lClienteExcluyeProducto Then
                If Calcular("select count(tcodigodelivery) as codigo from tclienteproducto where tcodigodelivery='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'  ", Cn) = 0 Then ' si el cliente tiene productos asociados
                            MsgBox "Según Configuración, este Producto no puede seleccionarse para el Cliente Frecuente Indicado"
                            Screen.MousePointer = vbDefault
                            Exit Sub
                End If
            End If
            
            lClienteControlaProducto = False
            If Calcular("select count(tcodigoproducto) as codigo from tclienteproducto where tcodigodelivery='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'", Cn) > 0 Then
                lClienteControlaProducto = True
            End If
            
        End If
    End If
    
        'INSUMOCRITICO23
        Dim rsInsumo As New ADODB.Recordset
        If Calcular("select isnull(lControlInsumoCritico,0) as codigo from tproducto where tcodigoproducto='" & sProducto & "'", Cn) = True Then
            Set rsInsumo = Lib.OpenRecordset("select isnull(tcodigoinsumo,'') tcodigoinsumo , isnull(tinsumo.descripcion,'') ,isnull(nstock,0) from tproducto inner join tinsumo on tproducto.tcodigoinsumo =tinsumo.tcodigo where tproducto.tcodigoproducto='" & sProducto & "' and tinsumo.lactivo=1", Cn)
            If Not (rsInsumo.EOF Or rsInsumo.BOF) Then
                Me.Caption = muestra & "   Insumo Crítico --->   " & rsInsumo.Fields(1) & " =  Stock: " & str(rsInsumo.Fields(2)) & " / Solicitado: " + str(xCantidad)
            End If
        Else
            Me.Caption = muestra
        End If
        'INSUMOCRITICO
    sItem = Lib.Correlativo(Calcular("select max(tItem) as codigo from DPEDIDO where tCodigoPedido = '" & sPedido & "'", Cn), 3)
    If RsDetalle.RecordCount = 0 Then
       'sitem = "001"
       nOrden = 1
    Else
       'sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from DPEDIDO where tCodigoPedido = '" & sPedido & "'", Cn), 3)
       If lOrden Then
          Set RsOrd = Lib.OpenRecordset("select nOrden, lImprime from DPEDIDO where tCodigoPedido='" & sPedido & "' Order by nOrden DESC", Cn)
          If RsOrd.RecordCount > 0 Then
             If IIf(IsNull(RsOrd!lImprime), False, RsOrd!lImprime) Then
                nOrden = RsOrd!nOrden + 1
             Else
                nOrden = RsOrd!nOrden
             End If
          Else
             nOrden = 1
          End If
       Else
          nOrden = RsProducto!nOrden
       End If
    End If
    
    'Precios con Recargos / Descargos por Tipo de Pedido
    nRecargo = 0
    nDescuento = 0
    nValor = 0
    nValor = nValor + IIf(RsProducto!lImpuesto1, nPorcentaje1, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto2, nPorcentaje2, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto3, nPorcentaje3, 0)
    
    lImp1 = RsProducto!lImpuesto1
    lImp2 = RsProducto!lImpuesto2
    lImp3 = RsProducto!lImpuesto3
    
    If sTipoPedido = "02" Then
       If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioDelivery), 0, RsProducto!nPrecioDelivery)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto4, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto5, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto6, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto4
          lImp2 = RsProducto!lImpuesto5
          lImp3 = RsProducto!lImpuesto6
       End If
    ElseIf sTipoPedido = "03" Then
       If IsNull(RsProducto!nPreciollevar) Or RsProducto!nPreciollevar = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPreciollevar), 0, RsProducto!nPreciollevar)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto7, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto8, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto9, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto7
          lImp2 = RsProducto!lImpuesto8
          lImp3 = RsProducto!lImpuesto9
       End If
    ElseIf sTipoPedido = "04" Then
       If IsNull(RsProducto!nPrecioCanal4) Or RsProducto!nPrecioCanal4 = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioCanal4), 0, RsProducto!nPrecioCanal4)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto10
          lImp2 = RsProducto!lImpuesto11
          lImp3 = RsProducto!lImpuesto12
       End If
    ElseIf sTipoPedido = "05" Then
       If IsNull(RsProducto!nPrecioCanal5) Or RsProducto!nPrecioCanal5 = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioCanal5), 0, RsProducto!nPrecioCanal5)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto13, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto14, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto15, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto13
          lImp2 = RsProducto!lImpuesto14
          lImp3 = RsProducto!lImpuesto15
       End If
    
    Else
       nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
    End If
    
    nOficial = IIf(RsProducto!tMONEDA = "02", nOficial * nTC, nOficial)
    
    'CPvalicacion LG
    If lClienteControlaProducto Then
        'GCAA 28/02/2020
        Dim vajaja As Integer
          Set ValProductoAsociado = Lib.OpenRecordset("  SP_INFOREST_VALPRODASOCIADOS '" & sProducto & "', '" & sTipoPedido & "', '" & sCliente & "','" & sPedido & "'", Cn)
             If Not (ValProductoAsociado.EOF Or ValProductoAsociado.BOF) Then
                vajaja = ValProductoAsociado.Fields(0)
            End If
        If vajaja = 1 Then
            nOficial = Calcular("SELECT ISNULL(NPRECIO,0) AS CODIGO FROM TCLIENTEPRODUCTO WHERE TCODIGODELIVERY='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'", Cn)
            nOficial = IIf(RsProducto!tMONEDA = "02", nOficial * nTC, nOficial)
            lProductoPermiteDescuento = Calcular("select isnull(lPermiteDescuentos,0) as codigo from tclienteproducto where TCODIGODELIVERY='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'", Cn)
        End If
    End If
    'CPvalicacion LG
    
    'Busca Oferta
    Dim sCriterio As String
    Dim nOferta As Double
    Dim tOferta As String
    Dim lAcumulable As Boolean
    sCriterio = "tCodigoProducto ='" & sProducto & "' and lActivo=1 "
    sCriterio = sCriterio & " and (substring(tFrecuencia," & Weekday(FechaServidor(), vbMonday) & "+1,1) = '1' or (substring(tFrecuencia,1,1)='1') and MONTH(fFecha) = " & Month(FechaServidor()) & " AND DAY(fFecha)= " & Day(FechaServidor()) & ") and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "'"
    sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "')) "
    If sTipoPedido = "01" Then
       sCriterio = sCriterio & " and lLocal=1"
    ElseIf sTipoPedido = "02" Then
       sCriterio = sCriterio & " and lDelivery=1"
    ElseIf sTipoPedido = "03" Then
       sCriterio = sCriterio & " and lLlevar=1"
    ElseIf sTipoPedido = "04" Then
       sCriterio = sCriterio & " and lCanal4=1"
    Else
       sCriterio = sCriterio & " and lCanal5=1"
    End If
    tOferta = ""
    
    Isql = "select  *  from TOFERTA where tCodigoProducto ='" & sProducto & "' and lActivo=1  "
    Set RsExisteOferta = Lib.OpenRecordset(Isql, Cn)
    
    
    
    
    Isql = "select * from TOFERTA where " & sCriterio
'     Debug.Print Isql
    Set RsOferta = Lib.OpenRecordset(Isql, Cn)
                
    'inserta descto
    nOferta = 0
    lAcumulable = False
    If RsOferta.RecordCount > 0 And RsProducto!lDescuento Then
        If (RsOferta!lAutomatica) Then
            RsOferta.MoveFirst
            tOferta = RsOferta!tOferta
            If RsOferta!nPrecio > 0 Then
               nOferta = nOficial - IIf(IsNull(RsOferta!nPrecio), 0, RsOferta!nPrecio)
            ElseIf RsOferta!nMonto > 0 Then
               nOferta = RsOferta!nMonto
            Else
               nOferta = nOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
            End If
        End If
    Else
       'GCAA 05012021
       If RsExisteOferta.RecordCount > 0 Then
            If Calcular("select lExcluyente as Codigo from TOFERTA where tCodigoProducto ='" & sProducto & "' and lActivo=1", Cn) Then
              Screen.MousePointer = vbDefault
              MsgBox "Este producto no puede ser cargado en esta franja horaria" & Chr(13) & "Consulte con el Manager", vbCritical, sMensaje
              Exit Sub
'            Else
'               MsgBox "Fuera de rango" & Chr(13) & "Consulte con el Manager", vbCritical, sMensaje
'              Exit Sub
            End If
       End If
    End If
     If lClienteControlaProducto Then
        nPVenta = nOficial
        If lProductoPermiteDescuento Then
            If xDescuento <> 0 And RsProducto!lDescuento Then
               If RsOferta.RecordCount > 0 Then
                  If RsOferta!lAcumulable Then
                     nPVenta = (nPVenta - nOferta) - ((nPVenta - nOferta) * xDescuento / 100)
                     nDescuento = nOficial - nPVenta
                  Else
                     nPVenta = nPVenta - nOferta
                     nDescuento = nOficial - nPVenta
                  End If
               Else
                  nPVenta = nPVenta - (nPVenta * xDescuento / 100)
                  nDescuento = nOficial - nPVenta
               End If
               
            Else
               nPVenta = nPVenta - nOferta
               nDescuento = nOficial - nPVenta
            End If
        Else
            nOferta = 0
            
        End If
    Else
        If xDescuento <> 0 And RsProducto!lDescuento Then
           If RsOferta.RecordCount > 0 Then
              If RsOferta!lAcumulable Then
                 nPVenta = (nOficial - nOferta) - ((nOficial - nOferta) * xDescuento / 100)
                 nDescuento = nOficial - nPVenta
              Else
                 nPVenta = nOficial - nOferta
                 nDescuento = nOficial - nPVenta
              End If
           Else
              nPVenta = nOficial - (nOficial * xDescuento / 100)
              nDescuento = nOficial - nPVenta
           End If
        Else
           nPVenta = nOficial - nOferta
           nDescuento = nOficial - nPVenta
        End If
    
    End If
    
'    If xDescuento <> 0 And RsProducto!lDescuento Then
'       If RsOferta.RecordCount > 0 Then
'          If RsOferta!lAcumulable Then
'             nPVenta = (nOficial - nOferta) - ((nOficial - nOferta) * xDescuento / 100)
'             nDescuento = nOficial - nPVenta
'          Else
'             nPVenta = nOficial - nOferta
'             nDescuento = nOficial - nPVenta
'          End If
'       Else
'          nPVenta = nOficial - (nOficial * xDescuento / 100)
'          nDescuento = nOficial - nPVenta
'       End If
'    Else
'       nPVenta = nOficial - nOferta
'       nDescuento = nOficial - nPVenta
'    End If
                                    
   'extranjero bolivia
    Select Case pais ' ok
        Case "001" 'Bolivia
                    nValor = (nValor / 100)
                    nImpuesto1 = IIf(lImp1, nPVenta * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(lImp2, nPVenta * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(lImp3, nPVenta * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                    
        Case Else 'Peru, Ecuador
                    nValor = 1 + (nValor / 100)
                    nImpuesto1 = IIf(lImp1, nPVenta / nValor * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(lImp2, nPVenta / nValor * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(lImp3, nPVenta / nValor * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
    End Select
    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMObra As Double
    
    If sTipoPedido = "01" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo)
       nGasto = IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto)
       nMObra = IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra)
    ElseIf sTipoPedido = "02" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo2), 0, RsProducto!nInsumo2)
       nGasto = IIf(IsNull(RsProducto!nGasto2), 0, RsProducto!nGasto2)
       nMObra = IIf(IsNull(RsProducto!nManoObra2), 0, RsProducto!nManoObra2)
    ElseIf sTipoPedido = "03" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo3), 0, RsProducto!nInsumo3)
       nGasto = IIf(IsNull(RsProducto!nGasto3), 0, RsProducto!nGasto3)
       nMObra = IIf(IsNull(RsProducto!nManoObra3), 0, RsProducto!nManoObra3)
    ElseIf sTipoPedido = "04" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo4), 0, RsProducto!nInsumo4)
       nGasto = IIf(IsNull(RsProducto!nGasto4), 0, RsProducto!nGasto4)
       nMObra = IIf(IsNull(RsProducto!nManoObra4), 0, RsProducto!nManoObra4)
    Else
       nInsumo = IIf(IsNull(RsProducto!nInsumo5), 0, RsProducto!nInsumo5)
       nGasto = IIf(IsNull(RsProducto!nGasto5), 0, RsProducto!nGasto5)
       nMObra = IIf(IsNull(RsProducto!nManoObra5), 0, RsProducto!nManoObra5)
    End If
                                
  'Dim tAreaProduccion As String
    'multiarea produccion
    'lProductoMultiArea = Calcular("select isnull(lmultiarea,0) as codigo from tproducto where tcodigoproducto='" & RsProducto.Fields("codigo") & "'", Cn)
    If Not RsProducto!lMultiArea Then
    'If lProductoMultiArea = False Then
        tsubalmacen = ""
    Else
    
        Dim rstAlmacenArea As New ADODB.Recordset
        Set rstAlmacenArea = New ADODB.Recordset
        Dim xPara As String
        xPara = IIf(lMultiAreaSubGrupo, "1", "0") + "," + IIf(lMultiAreaCaja, "1", "0") + ",'" + sCaja + "','" + IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) + "'"
        Set rstAlmacenArea = Lib.OpenRecordset("usp_Inforest_ObtieneSubAlmacen " + xPara, Cn)
        If Not (rstAlmacenArea.EOF Or rstAlmacenArea.BOF) Then
                tsubalmacen = rstAlmacenArea.Fields(0)
        End If
'        tsubalmacen = ""
'        If lMultiAreaSubGrupo = True Then
'            tAreaProduccion = Calcular("select isnull(tarea,'') codigo from TAREASUBGRUPO where tcaja='" & sCaja & "' and tSubGrupo='" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "'", Cn)
'            tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
'        End If
'        If lMultiAreaCaja = True Then
'            tAreaProduccion = Calcular("select isnull(tsubalmacen,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
'            tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
'        End If
'
    
        
        If tsubalmacen = "0" Then
            tsubalmacen = ""
        End If
    
    End If
    
    If lBloqInafecto And Calcular("select count(*) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn) > 0 And Round(nPVenta, 2) > 0 Then
        Dim IgvPED As Double
        IgvPED = CDbl(Calcular("select top 1 round(nprecioimpuesto1,2) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn))
        If IgvPED > 0 Then
            If Round(nImpuesto1, 2) <= 0 Then
                MsgBox ("No es posible ingresar el producto, no se puede ingresar Items inafecto y afecto en un mismo pedido!! "), vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Else
            If Round(nImpuesto1, 2) > 0 Then
                MsgBox ("No es posible ingresar el producto, no se puede ingresar Items inafecto y afecto en un mismo pedido!! "), vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
    
     'If lcover Then
        Dim maxped As Double
        maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
            If maxped > 0 Then
                   If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + (nPVenta * xCantidad)) Then
                     MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                     
                     Screen.MousePointer = vbDefault
                     'RsDetalle.MoveLast
                     RsDetalle.Requery
                     'RsProducto.MoveFirst
                     Exit Sub

                   End If
            End If
    'End If
    
      
    fxCombo "A", 1, sProducto

    
    
                  Dim oComandoD
                  Set oComandoD = New clsComando
                  If Not oComandoD.CreateCmdSp("spIns_DPEDIDO", Cn) Then
                     Set oComandoD = Nothing
                     Exit Sub
                  End If
                  
                  oComandoD.CreateParameter "@tPedido", adVarChar, adParamInput, 10, sPedido
                  oComandoD.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, sTipoPedido
                  oComandoD.CreateParameter "@tItem", adVarChar, adParamInput, 3, sItem
                  oComandoD.CreateParameter "@tProducto", adVarChar, adParamInput, 7, sProducto
                  oComandoD.CreateParameter "@tGrupo", adVarChar, adParamInput, 2, IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo)
                  oComandoD.CreateParameter "@tSubGrupo", adVarChar, adParamInput, 4, IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo)
                  oComandoD.CreateParameter "@nPBase", adDouble, adParamInput, 10, nPBase
                  oComandoD.CreateParameter "@nRecargo", adDouble, adParamInput, 10, nRecargo
                  oComandoD.CreateParameter "@nDescuento", adDouble, adParamInput, 10, nDescuento
                  oComandoD.CreateParameter "@nOficial", adDouble, adParamInput, 10, nOficial
                  oComandoD.CreateParameter "@nImpuesto1", adDouble, adParamInput, 10, nImpuesto1
                  oComandoD.CreateParameter "@nImpuesto2", adDouble, adParamInput, 10, nImpuesto2
                  oComandoD.CreateParameter "@nImpuesto3", adDouble, adParamInput, 10, nImpuesto3
                  
                  oComandoD.CreateParameter "@nPVenta", adDouble, adParamInput, 10, nPVenta
                  oComandoD.CreateParameter "@nCantidad", adDouble, adParamInput, 10, xCantidad
                  oComandoD.CreateParameter "@tArea", adVarChar, adParamInput, 3, RsProducto!tArea
                  oComandoD.CreateParameter "@lImprimeArea", adBoolean, adParamInput, 1, IIf(RsProducto!lImprimeArea, True, False)
                  oComandoD.CreateParameter "@lCombinacion", adBoolean, adParamInput, 1, IIf(RsProducto!lCombinacion, True, False)
                  oComandoD.CreateParameter "@nCombinacion", adDouble, adParamInput, 10, IIf(IsNull(RsProducto!nCombinacion), 0, RsProducto!nCombinacion)
                  oComandoD.CreateParameter "@tComanda", adVarChar, adParamInput, 10, sComanda
                  oComandoD.CreateParameter "@tMozo", adVarChar, adParamInput, 4, sMozoD
                  oComandoD.CreateParameter "@tUsuario", adVarChar, adParamInput, 30, sUsuario
                  oComandoD.CreateParameter "@nInsumo", adDouble, adParamInput, 10, nInsumo
                  oComandoD.CreateParameter "@nGasto", adDouble, adParamInput, 10, nGasto
                  oComandoD.CreateParameter "@nMObra", adDouble, adParamInput, 10, nMObra
                  oComandoD.CreateParameter "@nOrden", adDouble, adParamInput, 10, nOrden
                  oComandoD.CreateParameter "@tUnidadNegocio", adVarChar, adParamInput, 2, sUnidadNegocio
                  oComandoD.CreateParameter "@tOferta", adVarChar, adParamInput, 5, tOferta
                  oComandoD.CreateParameter "@tSubAlmacen", adVarChar, adParamInput, 3, tsubalmacen
                  oComandoD.CreateParameter "@tFechaContable", adDBDate, adParamInput, 10, obtieneDiaContable
                  oComandoD.CreateParameter "@sCaja", adVarChar, adParamInput, 3, sCaja
                                     
                  If Not oComandoD.GetParamOK Then
                     Set oComandoD = Nothing
                     Exit Sub
                  End If
                  If Not oComandoD.ExecSP Then
                     Set oComandoD = Nothing
                     Exit Sub
                  End If
    
    'CalculaDescuento

    RsDetalle.Requery
    nCombo = IIf(IsNull(RsDetalle!nCombinacion), 1, RsDetalle!nCombinacion)
    txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
    
    RsDetalle.MoveLast
    Screen.MousePointer = vbDefault
    
    
    If (RsProducto!lCombinacion) And (RsProducto!lComboVisible) Then
    lComboActivo = True
''       MsgBox ("es un combo !!!"), vbInformation
'        If RsCabecera.RecordCount = 0 Then
'           Exit Sub
'        End If
'
'        wAgregaCombo = True
'        ActivaCabecera False
'        ActivaOpcion False
'        If Not RsCombo.EOF Then
'           RsCombo.MoveFirst
'        End If
'        AsignaProductoCombo
'
'        CmdOpcion(1).Visible = False
'        CmdOpcion(2).Visible = False
'        txtBarra.SetFocus
''                 fraCombo.Visible = True
''                fraProductoCombo.Visible = True
'        ActivaFrame fraCombo
'        ActivaFrame fraProductoCombo
    End If
End Sub

Public Sub AsignaProducto()
   If Not RsDetalle.EOF Then
      sProducto = IIf(IsNull(RsDetalle!tCodigoProducto), "", RsDetalle!tCodigoProducto)
      'Correccion
      ListarOperadoresConFiltro (sProducto)
      sGrupo = IIf(IsNull(RsDetalle!tCodigoGrupo), "", RsDetalle!tCodigoGrupo)
      sSubGrupo = IIf(IsNull(RsDetalle!tCodigoSubGrupo), "", RsDetalle!tCodigoSubGrupo)
      sItem = IIf(IsNull(RsDetalle!tItem), "001", RsDetalle!tItem)
      sCortesia = IIf(IsNull(RsDetalle!tCortesia), "", RsDetalle!tCortesia)
    
      nOficial = IIf(IsNull(RsDetalle!nPrecioOficial), 0, RsDetalle!nPrecioOficial)
      nOrden = IIf(IsNull(RsDetalle!nOrden), 0, RsDetalle!nOrden)
      
      nDescuento = IIf(IsNull(RsDetalle!nDescuento), 0, RsDetalle!nDescuento)
      nRecargo = IIf(IsNull(RsDetalle!nRecargo), 0, RsDetalle!nRecargo)
      nPBase = IIf(IsNull(RsDetalle!nPrecioNeto), 0, RsDetalle!nPrecioNeto)
      nImpuesto1 = IIf(IsNull(RsDetalle!nprecioImpuesto1), 0, RsDetalle!nprecioImpuesto1)
      nImpuesto2 = IIf(IsNull(RsDetalle!nprecioImpuesto2), 0, RsDetalle!nprecioImpuesto2)
      nImpuesto3 = IIf(IsNull(RsDetalle!nprecioImpuesto3), 0, RsDetalle!nprecioImpuesto3)
      nPVenta = IIf(IsNull(RsDetalle!nprecioVenta), 0, RsDetalle!nprecioVenta)
      nCantidad = IIf(IsNull(RsDetalle!nCantidad), 0, RsDetalle!nCantidad)
      lblObservacion.Text = IIf(IsNull(RsDetalle!tObservacion), "", RsDetalle!tObservacion)
      txtDescuento.Caption = Format(Calcular("select sum(nDescuento*nCantidad) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
                
                    
        If IIf(IsNull(RsDetalle!lImprime), False, RsDetalle!lImprime) = False Then
            'luchoinsumos
             verificatitulo
             'luchoinsumos
        Else
        Me.Caption = muestra
        End If
                                        
                    
      'Llena el Combo
      fraCombo.Caption = IIf(IsNull(RsDetalle!Producto), "", " " & RsDetalle!Producto & " ")
      wCombo = IIf(IsNull(RsDetalle!lCombinacion), False, RsDetalle!lCombinacion)
      nCombo = IIf(IsNull(RsDetalle!nCombinacion), 1, RsDetalle!nCombinacion)
      
      'GCCA 03082021 - REQ REST-N014-2021 - TORI
      If (lNoVisualizarItemCombos) Then
        RsCombo.Filter = "[tCodigoPedido]='" & sPedido & "' and [tItem]='" & sItem & "' and [automatico]='" & 0 & "'    "
      Else
        RsCombo.Filter = "[tCodigoPedido]='" & sPedido & "' and [tItem]='" & sItem & "'"
      End If
      
      fraCombo.Visible = False
      
      If wCombo = True Then
        sProductoCombo = sProducto
      End If
      
      
      wAgregaCombo = False
      ActivaCabecera True
      ActivaOpcion True
      ActivaFrame fraPlato
      AsignaPropiedad
   End If
   
End Sub

Public Sub GrabaProducto()
   Screen.MousePointer = vbHourglass
   Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
           "nDescuento = " & nDescuento & ", " & _
           "nRecargo = " & nRecargo & ", " & _
           "nPrecioOficial = " & nOficial & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & " " & _
           "where tItem = '" & sItem & "' and tCodigoPedido ='" & sPedido & "'"
           Cn.Execute Isql
           RsDetalle.Requery
           RsDetalle.MoveFirst
           RsDetalle.Find "tItem = '" & sItem & "'"
           'AsignaProducto
   txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
   Screen.MousePointer = vbDefault

End Sub

Public Sub AsignaPropiedad()
    Dim i As Integer
    If RsOperador.RecordCount > 0 Then
       RsPropiedad.Filter = "tOperador = '" & RsOperador!codigo & "' and tProducto='" & sProducto & "'"
       nOperadorPropiedad = IIf(IsNull(RsOperador!nControl), 0, RsOperador!nControl)
    Else
       RsPropiedad.Filter = "tOperador = '  ' and tProducto='" & sProducto & "'"
       nOperadorPropiedad = 0
    End If
    
    AsignaComando 20, RsPropiedad, cmdPropiedad()
    
'       Validaciòn de Visibilidad de Propiedades -BR
        With RsPropiedad
            If .RecordCount > 0 Then
                .MoveFirst
                For i = 1 To IIf(.RecordCount > 20, 20, .RecordCount)
                    If IsNull(!visibilidad) Or !visibilidad = False Then
                        cmdPropiedad(i).Visible = True
                    Else
                        If !visibilidad Then
                            Select Case RsCabecera!tTipoPedido
                                Case "01"
                                    cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal1) And RsPropiedad!canal1
                                Case "02"
                                    cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal2) And RsPropiedad!canal2
                                Case "03"
                                    cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal3) And RsPropiedad!canal3
                                Case "04"
                                    cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal4) And RsPropiedad!canal4
                                Case "05"
                                    cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal5) And RsPropiedad!canal5
                                Case Else
                                    cmdPropiedad(i).Visible = False
                            End Select
                        Else
                            cmdPropiedad(i).Visible = False
                        End If
                    End If
                    cmdPropiedad(i).Caption = !descripcion
                    .MoveNext
                Next i
            Else
                For i = 1 To 20
                    cmdPropiedad(i).Visible = False
                Next i
            End If
        End With
        
    For i = 1 To 20
        cmdPropiedad(i).FontBold = False
    Next i
    
    lblResumen.Text = ""
    RsProductoPropiedad.Filter = "tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'"
    If Not RsProductoPropiedad.EOF Then
       RsProductoPropiedad.MoveFirst
       Do While Not RsProductoPropiedad.EOF
          For i = 1 To 20
              If cmdPropiedad(i).Caption = RsProductoPropiedad!descripcion And RsOperador!descripcion = RsProductoPropiedad!Operador Then
                 cmdPropiedad(i).FontBold = True
                 Exit For
              End If
          Next i
'          lblResumen.Text = lblResumen.Text & LTrim(RsProductoPropiedad!Operador) & " " & LTrim(RsProductoPropiedad!Descripcion) & ", "

          If RsProductoPropiedad!nCantidad = 1 Then
            lblResumen.Text = lblResumen.Text & LTrim(RsProductoPropiedad!Operador) & " " & LTrim(RsProductoPropiedad!descripcion) & ", "
          Else
            lblResumen.Text = lblResumen.Text & LTrim(RsProductoPropiedad!Operador) & " " & LTrim(RsProductoPropiedad!descripcion) & ": (" & RsProductoPropiedad!nCantidad & "), "
          End If
          
          RsProductoPropiedad.MoveNext
       Loop
    End If
End Sub
Private Sub grdCabecera_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If RsCabecera.RecordCount > 0 And RsCabecera.AbsolutePosition > 0 Then
      AsignaPedido
   End If
End Sub

Private Sub grdCombo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   frmMozoUsuario.ReseteaTimer
   AsignaCombo
   If fraPropiedad.Visible = True Then
      cmdOpcion_Click (8)
   End If
End Sub



Private Sub grdDetalle_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   frmMozoUsuario.ReseteaTimer
   If RsCabecera.RecordCount <> 0 Then
      If fraPropiedad.Visible = True Then
         nPos = RsDetalle.AbsolutePosition
         RsDetalle.Requery
         RsDetalle.AbsolutePosition = nPos
       End If
   
      AsignaProducto
      
      If lPropiedad Then
         lPropiedad = False
         cmdDetalle_Click (4)
      End If
      
    If lComboActivo Then
         lComboActivo = False
         cmdOpcion_Click (12)
      End If
      
      
   End If
   
    On Error GoTo fin
     If lBuscarPedidoBarraMesero Then
        txtbuscarpedidobarra.SetFocus
     Else
       On Error Resume Next
        Me.txtBarra.SetFocus
     End If

   Exit Sub
fin:
End Sub

Public Sub InsertaCombo(wProducto As String)
    Screen.MousePointer = vbHourglass
    Dim xItem As String
    Dim nValor As Double
    Dim nCNeto As Double
    Dim nCImp1 As Double
    Dim nCImp2 As Double
    Dim nCImp3 As Double
    Dim nCVenta As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMano As Double
    
    If RsCombo.RecordCount = 0 Then
       xItem = "001"
    Else
       xItem = Lib.Correlativo(Calcular("select max(tItemCombo) as codigo from CPEDIDO where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'", Cn), 3)
    End If
                          
    nValor = 0
    nValor = nValor + IIf(RsProducto!lImpuesto1, nPorcentaje1, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto2, nPorcentaje2, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto3, nPorcentaje3, 0)
      
    lImp1 = RsProducto!lImpuesto1
    lImp2 = RsProducto!lImpuesto2
    lImp3 = RsProducto!lImpuesto3
          
    If sTipoPedido = "02" Then
       If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPrecioDelivery), 0, RsProducto!nPrecioDelivery)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto4, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto5, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto6, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto4
          lImp2 = RsProducto!lImpuesto5
          lImp3 = RsProducto!lImpuesto6
       End If
    ElseIf sTipoPedido = "03" Then
       If IsNull(RsProducto!nPreciollevar) Or RsProducto!nPreciollevar = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPreciollevar), 0, RsProducto!nPreciollevar)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto7, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto8, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto9, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto7
          lImp2 = RsProducto!lImpuesto8
          lImp3 = RsProducto!lImpuesto9
       End If
    ElseIf sTipoPedido = "04" Then
       If IsNull(RsProducto!nPrecioCanal4) Or RsProducto!nPrecioCanal4 = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPrecioCanal4), 0, RsProducto!nPrecioCanal4)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto10
          lImp2 = RsProducto!lImpuesto11
          lImp3 = RsProducto!lImpuesto12
       End If
    ElseIf sTipoPedido = "05" Then
       If IsNull(RsProducto!nPrecioCanal5) Or RsProducto!nPrecioCanal5 = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPrecioCanal5), 0, RsProducto!nPrecioCanal5)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto13, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto14, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto15, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto13
          lImp2 = RsProducto!lImpuesto14
          lImp3 = RsProducto!lImpuesto15
       End If
    
    Else
       nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
    End If
    If sTipoPedido = "01" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo)
       nGasto = IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto)
       nMano = IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra)
    ElseIf sTipoPedido = "02" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo2), 0, RsProducto!nInsumo2)
       nGasto = IIf(IsNull(RsProducto!nGasto2), 0, RsProducto!nGasto2)
       nMano = IIf(IsNull(RsProducto!nManoObra2), 0, RsProducto!nManoObra2)
    ElseIf sTipoPedido = "03" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo3), 0, RsProducto!nInsumo3)
       nGasto = IIf(IsNull(RsProducto!nGasto3), 0, RsProducto!nGasto3)
       nMano = IIf(IsNull(RsProducto!nManoObra3), 0, RsProducto!nManoObra3)
    ElseIf sTipoPedido = "04" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo4), 0, RsProducto!nInsumo4)
       nGasto = IIf(IsNull(RsProducto!nGasto4), 0, RsProducto!nGasto4)
       nMano = IIf(IsNull(RsProducto!nManoObra4), 0, RsProducto!nManoObra4)
    Else
       nInsumo = IIf(IsNull(RsProducto!nInsumo5), 0, RsProducto!nInsumo5)
       nGasto = IIf(IsNull(RsProducto!nGasto5), 0, RsProducto!nGasto5)
       nMano = IIf(IsNull(RsProducto!nManoObra5), 0, RsProducto!nManoObra5)
    End If
      
    nCVenta = IIf(RsProducto!tMONEDA = "02", nCVenta * nTC, nCVenta)
    
    'extranjero bolivia
    Select Case pais ' ok
        Case "001" 'Bolivia
                nValor = (nValor / 100)
                nCImp1 = IIf(lImp1, nCVenta * nPorcentaje1 / 100, 0)
                nCImp2 = IIf(lImp2, nCVenta * nPorcentaje2 / 100, 0)
                nCImp3 = IIf(lImp3, nCVenta * nPorcentaje3 / 100, 0)
                nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
                
        Case Else 'Peru, Ecuador
                nValor = 1 + (nValor / 100)
                nCImp1 = IIf(lImp1, nCVenta / nValor * nPorcentaje1 / 100, 0)
                nCImp2 = IIf(lImp2, nCVenta / nValor * nPorcentaje2 / 100, 0)
                nCImp3 = IIf(lImp3, nCVenta / nValor * nPorcentaje3 / 100, 0)
                nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
    
    End Select
    
    
       'If lcover Then
        Isql = "select * from TCombo Where tCombo = '" & RsDetalle!tCodigoProducto & "' And tCodigoProducto = '" & wProducto & "'"
        Dim RstCombo2 As Recordset
        Set RstCombo2 = Lib.OpenRecordset(Isql, Cn)
        If RstCombo2.RecordCount > 0 Then
                
                 Dim maxped As Double
                 maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
                     If maxped > 0 Then
                            If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + (RstCombo2!nAumento)) Then
                              MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                               RsCombo.Requery
                              Exit Sub
    
                            End If
                     End If
                 
        End If
    'End If
    
    Dim xCantCombo As Double
    xCantCombo = Calcular("Select ISNULL(nCantidad,0) as Codigo From TCOMBO where tCombo = '" & RsDetalle!tCodigoProducto & "' and tCodigoProducto='" & wProducto & "'", Cn)

    Isql = "insert into CPEDIDO " & _
           "(tCodigoPedido, tProducto, tItem, tItemCombo, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden ) " & _
           "Values(   '" & sPedido & "', " _
                   & "'" & RsDetalle!tCodigoProducto & "', " _
                   & "'" & sItem & "', " _
                   & "'" & xItem & "', " _
                   & "'" & wProducto & "', " & xCantCombo & ", " _
                   & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                   & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                   & nCNeto & ", " _
                   & nCImp1 & ", " _
                   & nCImp2 & ", " _
                   & nCImp3 & ", " _
                   & nCVenta & ", " _
                   & nInsumo & ", " & nGasto & ", " & nMano & ", " _
                   & IIf(RsProducto!lImprimeArea, -1, 0) & ", 0, " _
                   & RsProducto!nOrden & ") "
    Cn.Execute Isql
    'Oscar Ortega------------------------------------------------------------
    Isql = "select * from TCombo Where tCombo = '" & RsDetalle!tCodigoProducto & "' And tCodigoProducto = '" & wProducto & "'"
    Dim RstCombo As Recordset
    Set RstCombo = Lib.OpenRecordset(Isql, Cn)
    If RstCombo.RecordCount > 0 Then
        If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
            txtSuma.Caption = CambiaPrecio(nPVenta + (RstCombo!nAumento * RstCombo!nCantidad) / nCantidad, sPedido)
        End If
    End If
    'Fin Oscar Ortega--------------------------------------------------------
          
    RsCombo.Requery
    RsCombo.MoveLast
    AsignaCombo
    Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaCombo()
   If Not RsCombo.EOF Then
      sCombo = IIf(IsNull(RsCombo!tProductoCombo), "", RsCombo!tProductoCombo)
      sProducto = sCombo
      xItem = IIf(IsNull(RsCombo!tItemCombo), "001", RsCombo!tItemCombo)
      lblObservacion.Text = IIf(IsNull(RsCombo!tObservacion), "", RsCombo!tObservacion)
      'Correccion
      ListarOperadoresConFiltro (sCombo)
      AsignaComboPropiedad
   End If
End Sub

Public Sub AsignaProductoCombo()
    Dim i As Integer
    RsProductoCombo.Filter = "tCombo = '" & sProducto & "'"
    AsignaComandoColor 48, RsProductoCombo, cmdProductoCombo()
End Sub

Public Sub AsignaComboPropiedad()
    Dim i As Integer
    If RsOperador.RecordCount > 0 Then
       RsPropiedad.Filter = "tOperador = '" & RsOperador!codigo & "' and tProducto='" & sCombo & "'"
    Else
       RsPropiedad.Filter = "tOperador = '  ' and tProducto='" & sCombo & "'"
    End If
    AsignaComando 20, RsPropiedad, cmdPropiedad()
'Validaciòn de Visibilidad de Propiedades -BR
    With RsPropiedad
        If .RecordCount > 0 Then
            .MoveFirst
            For i = 1 To IIf(.RecordCount > 20, 20, .RecordCount)
                If IsNull(!visibilidad) Or !visibilidad = False Then
                    cmdPropiedad(i).Visible = True
                Else
                    If (!visibilidad) Then
                         Select Case RsCabecera!tTipoPedido
                            Case "01"

                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal1) And RsPropiedad!canal1
                            Case "02"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal2) And RsPropiedad!canal2
                            Case "03"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal3) And RsPropiedad!canal3
                            Case "04"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal4) And RsPropiedad!canal4
                            Case "05"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal5) And RsPropiedad!canal5
                            Case Else
                                cmdPropiedad(i).Visible = False
                        End Select
                    Else
                         cmdPropiedad(i).Visible = False
                    End If
                End If
                cmdPropiedad(i).Caption = !descripcion
                .MoveNext
            Next i
        Else
               For i = 1 To 20
                   cmdPropiedad(i).Visible = False
               Next i
        End If
    End With
    For i = 1 To 20
        cmdPropiedad(i).FontBold = False
    Next i
    
    lblResumen.Text = ""
    RsComboPropiedad.Filter = "tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "' and tItemCombo='" & xItem & "'"
    If Not RsComboPropiedad.EOF Then
       RsComboPropiedad.MoveFirst
       Do While Not RsComboPropiedad.EOF
          For i = 1 To 20
              If cmdPropiedad(i).Caption = RsComboPropiedad!descripcion And RsOperador!descripcion = RsComboPropiedad!Operador Then
                 cmdPropiedad(i).FontBold = True
                 Exit For
              End If
          Next i
          'lblResumen.Text = lblResumen.Text & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!Descripcion) & ", "
          
          If RsComboPropiedad!nCantidad = 1 Then
              lblResumen.Text = lblResumen.Text & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!descripcion) & ", "
          Else
              lblResumen.Text = lblResumen.Text & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!descripcion) & ": (" & RsComboPropiedad!nCantidad & "), "
          End If
          
          RsComboPropiedad.MoveNext
       Loop
    End If
End Sub


Public Sub LimpiaCabecera()
   If RsCabecera.RecordCount = 0 Then
      sPedido = ""
      txtFecha.Caption = ""
      txtTipoPedido.Caption = ""
      txtMozo.Caption = ""
      txtDescuento.Caption = "0.00"
      txtMotorizado.Caption = ""
      txtCliente.Caption = ""
      txtTelefono.Caption = ""
      txtDireccion.Caption = ""
      txtObservacion.Caption = ""
      txtMontoMax.Caption = "0"
      fraPlato.Visible = True
       'entregarA
      txtEntregar.Caption = ""
   End If
End Sub

Public Sub fxCombo(Funcion As String, Cantidad As Double, Combo As String)
   Dim i As Integer
   Dim xItem As String
   Dim RsTemp As Recordset
   Dim nValor As Double
   Dim nCNeto As Double
   Dim nCImp1 As Double
   Dim nCImp2 As Double
   Dim nCImp3 As Double
   Dim nCVenta As Double
   Dim lImp1 As Boolean
   Dim lImp2 As Boolean
   Dim lImp3 As Boolean
   Dim nInsumo As Double
   Dim nGasto As Double
   Dim nMano As Double
   

   If sTipoPedido = "01" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo As nInsumo, dbo.TPRODUCTO.nGasto As nGasto, dbo.TPRODUCTO.nManoObra As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   ElseIf sTipoPedido = "02" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo2 As nInsumo, dbo.TPRODUCTO.nGasto2 As nGasto, dbo.TPRODUCTO.nManoObra2 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
   
   ElseIf sTipoPedido = "03" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo3 As nInsumo, dbo.TPRODUCTO.nGasto3 As nGasto, dbo.TPRODUCTO.nManoObra3 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   ElseIf sTipoPedido = "04" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo4 As nInsumo, dbo.TPRODUCTO.nGasto4 As nGasto, dbo.TPRODUCTO.nManoObra4 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   ElseIf sTipoPedido = "05" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo5 As nInsumo, dbo.TPRODUCTO.nGasto5 As nGasto, dbo.TPRODUCTO.nManoObra5 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   End If
   
   Set RsTemp = Lib.OpenRecordset(Isql, Cn)
   
   If RsTemp.RecordCount = 0 Then
      Exit Sub
   End If
   
   'GCCA 03082021 - REQ REST-N014-2021 - TORI
   If (lNoVisualizarItemCombos) Then
    RsCombo.Filter = "tCodigoPedido = '" & sPedido & "' and tItem='" & sItem & "' and automatico='" & 0 & "'"
   Else
    RsCombo.Filter = "tCodigoPedido = '" & sPedido & "' and tItem='" & sItem & "'"
   End If
   
   Select Case Funcion
          Case Is = "A"
               Do While Not RsTemp.EOF
                  If RsCombo.RecordCount = 0 Then
                     xItem = "001"
                  Else
                     xItem = Lib.Correlativo(Calcular("select max(tItemCombo) as codigo from CPEDIDO where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'", Cn), 3)
                  End If
                  
                  nValor = 0
                  nValor = nValor + IIf(RsTemp!lImpuesto1, nPorcentaje1, 0)
                  nValor = nValor + IIf(RsTemp!lImpuesto2, nPorcentaje2, 0)
                  nValor = nValor + IIf(RsTemp!lImpuesto3, nPorcentaje3, 0)
                    
                  lImp1 = RsTemp!lImpuesto1
                  lImp2 = RsTemp!lImpuesto2
                  lImp3 = RsTemp!lImpuesto3
                        
                  If sTipoPedido = "02" Then
                     If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
                        nCVenta = IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) / 100)
                     Else
                        nCVenta = IIf(IsNull(RsTemp!nPrecioDelivery), 0, RsTemp!nPrecioDelivery)
                        nValor = 0
                        nValor = nValor + IIf(RsTemp!lImpuesto4, nPorcentaje1, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto5, nPorcentaje2, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto6, nPorcentaje3, 0)
                        lImp1 = RsTemp!lImpuesto4
                        lImp2 = RsTemp!lImpuesto5
                        lImp3 = RsTemp!lImpuesto6
                     End If
                  ElseIf sTipoPedido = "03" Then
                     If IsNull(RsTemp!nPreciollevar) Or RsTemp!nPreciollevar = 0 Then
                        nCVenta = IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) + (nLlevar * IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) / 100)
                     Else
                        nCVenta = IIf(IsNull(RsTemp!nPreciollevar), 0, RsTemp!nPreciollevar)
                        nValor = 0
                        nValor = nValor + IIf(RsTemp!lImpuesto7, nPorcentaje1, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto8, nPorcentaje2, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto9, nPorcentaje3, 0)
                        lImp1 = RsTemp!lImpuesto7
                        lImp2 = RsTemp!lImpuesto8
                        lImp3 = RsTemp!lImpuesto9
                     End If
                  Else
                     nCVenta = IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta)
                  End If
                  nInsumo = IIf(IsNull(RsTemp!nInsumo), 0, RsTemp!nInsumo)
                  nGasto = IIf(IsNull(RsTemp!nGasto), 0, RsTemp!nGasto)
                  nMano = IIf(IsNull(RsTemp!nManoObra), 0, RsTemp!nManoObra)
                    
                  nCVenta = IIf(RsTemp!tMONEDA = "02", nCVenta * nTC, nCVenta)
                    'extranjero bolivia
                  Select Case pais ' ok
                        Case "001" 'Bolivia
                                    nValor = (nValor / 100)
                                    nCImp1 = IIf(lImp1, nCVenta * nPorcentaje1 / 100, 0)
                                    nCImp2 = IIf(lImp2, nCVenta * nPorcentaje2 / 100, 0)
                                    nCImp3 = IIf(lImp3, nCVenta * nPorcentaje3 / 100, 0)
                                    nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
                        
                        Case Else 'Peru, Ecuador
                                    nValor = 1 + (nValor / 100)
                                    nCImp1 = IIf(lImp1, nCVenta / nValor * nPorcentaje1 / 100, 0)
                                    nCImp2 = IIf(lImp2, nCVenta / nValor * nPorcentaje2 / 100, 0)
                                    nCImp3 = IIf(lImp3, nCVenta / nValor * nPorcentaje3 / 100, 0)
                                    nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
                  
                  End Select
                
                  Isql = "insert into CPEDIDO " & _
                         "(tCodigoPedido, tProducto, tItem, tItemCombo, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden ) " & _
                         "Values(   '" & sPedido & "', " _
                                 & "'" & Combo & "', " _
                                 & "'" & sItem & "', " _
                                 & "'" & xItem & "', " _
                                 & "'" & RsTemp!tCodigoProducto & "', " & RsTemp!nCantidad & ", " _
                                 & "'" & IIf(IsNull(RsTemp!tgrupo), "", RsTemp!tgrupo) & "', " _
                                 & "'" & IIf(IsNull(RsTemp!tSubGrupo), "", RsTemp!tSubGrupo) & "', " _
                                 & nCNeto & ", " _
                                 & nCImp1 & ", " _
                                 & nCImp2 & ", " _
                                 & nCImp3 & ", " _
                                 & nCVenta & ", " _
                                 & nInsumo & ", " & nGasto & ", " & nMano & ", " _
                                 & IIf(RsTemp!lImprimeArea, -1, 0) & ", 0, " _
                                 & RsTemp!nOrden & ") "
                  Cn.Execute Isql
                  RsCombo.Requery
                  RsTemp.MoveNext
               Loop
          Case Is = "M"
               Do While Not RsTemp.EOF
                  Dim X As Double
                  X = Calcular("select nCantidad as Codigo FROM TCOMBO where tCombo='" & RsTemp!tCombo & "' and tCodigoproducto ='" & RsTemp!tCodigoProducto & "'", Cn)
                  Isql = "update CPEDIDO set nCantidad = " & X * Cantidad & " where tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "' and tProductocombo='" & RsTemp!tCodigoProducto & "'"
                  Cn.Execute Isql
                  RsCombo.Requery
                  RsTemp.MoveNext
               Loop
          
          Case Is = "D"
               For i = 1 To Cantidad
                   Isql = "DELETE from CPEDIDO " & _
                          "where tCodigoPedido ='" & sPedido & "' and tProducto='" & Combo & "' and tItem='" & sItem & "'"
                   Cn.Execute Isql
                   RsCombo.Requery
               Next i
   End Select
   Set RsTemp = Nothing
End Sub

Public Function SupervisorCM() As Boolean
    Dim RsPW As Recordset
    
    Set RsPW = Lib.OpenRecordset("select * from TUSUARIO where lActivo = 1 and tGrupoUsuario ='01'", Cn)
    If RsPW.RecordCount = 0 Then
       Set RsPW = Nothing
       SupervisorCM = False
       Exit Function
    End If
    frmPassword.Show vbModal
    If Not wEnter Then
       Set RsPW = Nothing
       SupervisorCM = False
       Exit Function
    End If
    
    RsPW.MoveFirst
    Do While Not RsPW.EOF
       If sDescrip = Desencapsula(RsPW!tpassword) Then
          SupervisorCM = True
          Set RsPW = Nothing
          Exit Function
       End If
       RsPW.MoveNext
    Loop
    SupervisorCM = False
    Set RsPW = Nothing
End Function

Private Function CalculaDescuento() As Boolean
    Dim sCriterio As String
    Dim lAcumulable As Boolean
    Dim nOferta As Double
    Dim nSuma As Double
    
    nSuma = Calcular("SELECT sum(nPrecioOficial*nCantidad) as Codigo FROM dbo.DPEDIDO LEFT OUTER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where tCodigoPedido='" & sPedido & "' and lDescuento=1", Cn)

If RsDetalle.RecordCount <> 0 Then
   RsDetalle.MoveFirst
   
   Do While Not RsDetalle.EOF
      'Busca Oferta
      nPVenta = 0
      sCriterio = "tCodigoProducto ='" & RsDetalle!tCodigoProducto & "' and lActivo=1"
      sCriterio = sCriterio & " and (tFrecuencia='00' or tFrecuencia='0" & Weekday(FechaServidor(), vbMonday) & "' or (tFrecuencia='99' and fFecha='" & Format(FechaServidor(), "yyyy/MM/dd 00:00") & "') and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "')"
      sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "'))"
        
      Isql = "select * from TOFERTA where " & sCriterio
      Set RsOferta = Lib.OpenRecordset(Isql, Cn)
      
      lAcumulable = True
      nOferta = 0
      Acumulado = 0
      
      If RsOferta.RecordCount > 0 Then
         RsOferta.MoveFirst
         lAcumulable = RsOferta!lAcumulable
         nOferta = RsDetalle!nPrecioOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
      End If
      
      If RsDetalle!lDescuento And lAcumulable = True Then
         If Calcular("select lRatio as Codigo FROM vMotivoDescuento where Codigo='" & sCodigoDescuento & "'", Cn) Then
            nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((RsDetalle!nPrecioOficial - nOferta) * xDescuento / 100)
         Else
            Dim xPorc As Double
            xPorc = (RsDetalle!nPrecioOficial - nOferta) * RsDetalle!nCantidad * 100 / nSuma
            nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((xPorc * xDescuento / 100) / RsDetalle!nCantidad)
         End If
        
        ' extranjero bolivia
            Select Case pais ' ok
                Case "001" 'Bolivia
                     Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                     Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                     Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                     Acumulado = (Acumulado / 100)
                    
                     nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta * nPorcentaje1 / 100, 0)
                     nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta * nPorcentaje2 / 100, 0)
                     nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta * nPorcentaje3 / 100, 0)
                     nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                
                Case Else 'Peru, Ecuador
                     Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                     Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                     Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                     Acumulado = 1 + (Acumulado / 100)
                    
                     nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                     nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                     nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                     nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                
            End Select
            Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
                "nDescuento = " & RsDetalle!nPrecioOficial - nPVenta & ", " & _
                "nRecargo = " & nRecargo & ", " & _
                "nPrecioOficial = " & RsDetalle!nPrecioOficial & ", " & _
                "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
                "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
                "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
                "nPrecioVenta = " & nPVenta & ", " & _
                "nventa = " & nPVenta * RsDetalle!nCantidad & ", " & _
                "nCantidad = " & RsDetalle!nCantidad & ", " & _
                "nImpuesto1 = " & nImpuesto1 * RsDetalle!nCantidad & ", " & _
                "nImpuesto2 = " & nImpuesto2 * RsDetalle!nCantidad & ", " & _
                "nImpuesto3 = " & nImpuesto3 * RsDetalle!nCantidad & ", " & _
                "tCortesia = '" & sCortesia & "' " & _
                "where tItem = '" & RsDetalle!tItem & "' and tCodigoPedido ='" & sPedido & "'"
                Cn.Execute Isql
      End If
   RsDetalle.MoveNext
   Loop
End If

End Function
Public Sub EliminaCabecera()
    Call KDS_EliminarOrden(RsCabecera) 'KDS
   Dim i As Integer
    
   Dim sMotivo As String
    'cambio mango
    Screen.MousePointer = vbHourglass
    'Impresion del Pedidos Anulados
    sMotivo = Calcular("select Descripcion as Codigo from vMotivoEliminacion where Codigo='" & sCodigo & "'", Cn)
    If sCodigo = "000" Then
    sMotivo = sMotivo & " (" & sDescrip & ")"
    End If
    
   If lPrinter Then
       Isql = "select *,'" & sMotivo & "' as  MotivoEliminacion FROM dbo.vPedido LEFT OUTER JOIN dbo.vMotivoEliminacion ON dbo.vPedido.tMotivoEliminacion = dbo.vMotivoEliminacion.Codigo " & _
             "WHERE vPedido.Codigo='" & sPedido & "' AND lImprime=1 AND lImprimeArea=1 " & _
             "ORDER BY tItem,tetiqueta,combo"
       Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
              
       If Not RsImpresion.EOF Then
          RsArea.MoveFirst
          For i = 1 To RsArea.RecordCount
              RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
              If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                 If RsImpresion.RecordCount <> 0 Then
                    RsImpresion.MoveFirst
                    ImprimePedido RsImpresion, "A", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Inforest"
                    'TVS
                    'If (lCompatibilidadTVS) Then 'TVS
                    '    Call TVS_EnviarTicket(RsImpresion, RsArea!Area, "Anulacion", sPedido)
                    'End If
                 End If
              End If
              RsArea.MoveNext
          Next i
       End If
   End If
      
   If sMesa <> "" Or sMesa <> "000" Then
      Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa ='" & sMesa & "'"
   End If
   
   If lInfhotel Then
      CnInfhotel.Execute "update MCOMANDA set TESTADO ='04', TOBSERVACIONANULA = 'Anulado por Inforest - " & sUsuarioAutoriza & " " & RsCabecera!codigo & " - " & Trim(sDescrip) & "' where tComanda ='" & RsCabecera!tComanda & "'"
   End If
   
   'Juntar Mesa
   Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
   Cn.Execute "delete from TPEDIDOMESA where tCodigoPedido='" & sPedido & "'"
   
    'INSUMOCRITICO23
    Dim rstItems As New ADODB.Recordset
    Dim j As Integer
    Set rstItems = New ADODB.Recordset
    'Set rstItems = Lib.OpenRecordset("select tcodigoinsumo,ncantidad from dpedido inner join tproducto on dpedido.tcodigoproducto=tproducto.tcodigoproducto where tcodigopedido='" & sPedido & "'  and tproducto.lControlInsumoCritico=1 and isnull(tproducto.tcodigoinsumo,'')<>''  and isnull(dpedido.limprime,0)=1", Cn)
    Set rstItems = Lib.OpenRecordset("usp_Inforest_RevertirInsumosCriticosCabecera '" & sPedido & "'", Cn)

    If Not (rstItems.EOF Or rstItems.BOF) Then
        rstItems.MoveFirst
        For j = 0 To rstItems.RecordCount - 1
            modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
            rstItems.MoveNext
        Next j
    End If
   'INSUMOCRITICO
   
   Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='" & sCodigo & "', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='" & sDescrip & "' where tCodigoPedido ='" & sPedido & "'"
   Cn.Execute "Update DPEDIDO Set tEstadoItem = 'A' where tCodigoPedido = '" & sPedido & "'"
   Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido='" & sPedido & "'"
   Cn.Execute "delete CPEDIDO where tCodigoPedido='" & sPedido & "'"
   Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido='" & sPedido & "'"
   
'   Dim CodReserva As String
'
'   CodReserva = Calcular("Select ISNULL(tReservaInf,'') as codigo from Mpedido where tCodigoPedido = '" & sPedido & "'", Cn)
'
'   If CodReserva <> "" Then
'        If Calcular("Select count(*) as codigo from TRESERVA where tReserva='" & CodReserva & "'", Cn) > 0 Then
'            Cn.Execute "Update TRESERVA Set tEstadoReserva='01' where tReserva = '" & CodReserva & "'"
'        End If
'   End If
   
   RsCabecera.Requery
   RsDetalle.Requery
   RsProductoPropiedad.Requery
   RsCombo.Requery
   RsComboPropiedad.Requery
   AsignaPedido
   ActivaOpcion True
   
   If RsCabecera.RecordCount = 0 Then
      sPedido = ""
   Else
      RsCabecera.MoveFirst
      sPedido = RsCabecera!codigo
      ActivaFrame fraPlato
   End If
  sUsuario = nombreMozo
   Screen.MousePointer = vbDefault
End Sub

Public Sub EliminaItem()
    Dim xMax As String
    Dim sMotivo As String
    
    If lPrinter Then
       'Impresion del Pedidos Anulados
       sMotivo = Calcular("select Descripcion as Codigo from vMotivoEliminacion where Codigo='" & sCodigo & "'", Cn)
       If sCodigo = "000" Then
        sMotivo = sMotivo & " (" & sDescrip & ")"
       End If
       
       Isql = "select *, '" & sMotivo & "' as MotivoEliminacion FROM dbo.vPedido " & _
              "WHERE Codigo = '" & sPedido & "' and tItem = '" & sItem & "' and lImprime = 1 And lImprimeArea = 1 " & _
              "ORDER BY tItem,tetiqueta,combo"
                    
       Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
       Dim i As Integer
       If RsImpresion.RecordCount = 0 Then
          LimpiaRs
       Else
          RsArea.MoveFirst
          For i = 1 To RsArea.RecordCount
              RsImpresion.Filter = "tArea = '" & RsArea!tArea & "'"
              If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                 If RsImpresion.RecordCount <> 0 Then
                    RsImpresion.MoveFirst
                    ImprimePedido RsImpresion, "A", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Inforest"
                    'If (lCompatibilidadTVS) Then 'TVS
                    '    Call TVS_EnviarTicket(RsImpresion, RsArea!Area, "Anulacion", sPedido)
                    'End If
                 End If
              End If
              RsArea.MoveNext
          Next i
          RsDetalle.Requery
       End If
    End If
    
    xMax = Calcular("select max(tItem) as Codigo from APEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
    xMax = Lib.Correlativo(xMax, 3)
           
           'diacontable
     Isql = "insert into APEDIDO (tCodigoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, " & _
           "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, " & _
           "tComanda, lImprime, tUsuario, fRegistro, tUsuarioAnulado, fRegistroAnulado, " & _
           "tObservacion, tObservacionAnulado, tEstadoItem, lImprimeArea, tArea, tMotivoEliminacion, tTurnoAnulado,fdiaContable) " & _
           "select '" & sPedido & "' as tCodigoPedido, '" & xMax & "' as tItem, tCodigoProducto, tCodigoGRupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, " & _
           "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tComanda, lImprime, " & _
           "'" & RsCabecera!tUsuario & "' as tUsuario, dpedido.fregistro as fRegistro, " & _
           "'" & sUsuarioAutoriza & "' as tUsuarioAnulado, getDate() as fRegistroAnulado, " & _
           "tObservacion, '" & sDescrip & "' as tObservacion, tEstadoItem, lImprimeArea, tArea, '" & sCodigo & "', '" & sTurno & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
           "from DPEDIDO " & _
           "where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'"
           'diacontable
    Cn.Execute Isql
    
    'INSUMOCRITICO23
    Dim rstItems As New ADODB.Recordset
    Set rstItems = New ADODB.Recordset
    'Set rstItems = Lib.OpenRecordset("select tcodigoinsumo,ncantidad from dpedido inner join tproducto on dpedido.tcodigoproducto=tproducto.tcodigoproducto where tcodigopedido='" & sPedido & "' and titem='" & sitem & "' and tproducto.lControlInsumoCritico=1 and isnull(tproducto.tcodigoinsumo,'')<>''  and isnull(dpedido.limprime,0)=1 ", Cn)
    Set rstItems = Lib.OpenRecordset(" usp_Inforest_RevertirInsumosCriticos '" & sPedido & "','" & sItem & "' ", Cn)
    If Not (rstItems.EOF Or rstItems.BOF) Then
        modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
    End If
    
    'INSUMOCRITICO
    
    Cn.Execute "delete from TPRODUCTOPROPIEDAD where tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'"
    Cn.Execute "delete from DPEDIDO where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'"
    Cn.Execute "delete from CPEDIDO where tCodigoPedido ='" & sPedido & "' and tItem ='" & sItem & "'"
    Cn.Execute "delete from TCOMBOPROPIEDAD where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
    
    RsProductoPropiedad.Requery
    RsCombo.Requery
    RsComboPropiedad.Requery
    RsDetalle.Requery
    
    If RsDetalle.RecordCount = 0 Then
       txtSuma.Caption = "0.00"
       sProducto = ""
       wCombo = False
       nCombo = 0
       sItem = ""
              
       ActivaFrame fraPlato
       wAgregaCombo = False
       ActivaCabecera True
       ActivaOpcion True
    Else
       txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
       RsDetalle.MoveLast
       sItem = RsDetalle!tItem
    End If
                 sUsuario = nombreMozo
End Sub

Public Function CambiaPrecio(Valor As Double, Pedido As String)
    nPVenta = Val(Valor)
    nOficial = nPVenta
    Dim Acumulado As Double
           Select Case pais 'ok
            Case "001" 'Bolivia
                Acumulado = 0
                Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = (Acumulado / 100)
                
                nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta * nPorcentaje3 / 100, 0)
                
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                
            Case Else 'Peru, Ecuador
                Acumulado = 0
                Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = 1 + (Acumulado / 100)
                
                nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                
        End Select
        

    Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
           "nPrecioOficial = " & nOficial & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & " " & _
           "where tItem = '" & sItem & "' and tCodigoPedido ='" & Pedido & "'"
           Cn.Execute Isql
    CambiaPrecio = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & Pedido & "'", Cn), "#,###,##0.00")
End Function




Public Function CambiaPrecioCombo(Valor As Double, Pedido As String)
    nPVenta = Val(Valor)
    'nOficial = nPVenta
    Dim Acumulado As Double
         Select Case pais 'ok
            Case "001" 'Bolivia
                Acumulado = 0
                Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = (Acumulado / 100)
                
                nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta * nPorcentaje3 / 100, 0)
                
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                
            Case Else 'Peru, Ecuador
                Acumulado = 0
                Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = 1 + (Acumulado / 100)
                
                nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                
        End Select
        
        '"nPrecioOficial = " & nOficial & ", " &
    Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & " " & _
           "where tItem = '" & sItem & "' and tCodigoPedido ='" & Pedido & "'"
           Cn.Execute Isql
           
    CambiaPrecioCombo = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & Pedido & "'", Cn), "#,###,##0.00")
End Function
'Oscar Ortega---------------------------------
Private Function ObligaPropiedad(ByVal tProducto As String) As Boolean
    Dim j As Integer
    Dim i As Integer
    Dim RstProductoPropiedad As Recordset
    Dim flag As Boolean
    Dim oPos As Integer
    oPos = RsOperador.AbsolutePosition
    Dim mensajeOperador As String
    flag = True 'Si permite salir
    ObligaPropiedad = True
    If RsOperador.RecordCount > 0 Then
        If RsOperador.EOF Then
            RsOperador.MoveFirst
            For i = 1 To 13
                cmdOperador(i).backColor = vbButtonFace
            Next i
        End If
        RsOperador.MoveFirst
        While RsOperador.EOF = False
            If RsOperador!lObligaPropiedad = True Then
                If wAgregaCombo Then
                    Isql = "Select * From tComboPropiedad Where tCodigoPropiedad In (Select tCodigoPropiedad from TPropiedad Where tOperador = '" & RsOperador!codigo & "' And tProducto = '" & tProducto & "' ) And tCodigoPedido = '" & sPedido & "' And tItem = '" & sItem & "' And tItemCombo = '" & xItem & "'"
                Else
                    Isql = "Select * From tProductoPropiedad Where tCodigoPropiedad In (Select tCodigoPropiedad from TPropiedad Where tOperador = '" & RsOperador!codigo & "' And tProducto = '" & tProducto & "' ) And tCodigoPedido = '" & sPedido & "' And tItem = '" & sItem & "' "
                End If
                Set RstProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
                If RstProductoPropiedad.RecordCount = 0 Then
                    flag = False 'Esta Obligado y no ha elegido Propiedad
                    mensajeOperador = mensajeOperador + "(" + RsOperador!descripcion + ")"
                End If
            End If
            RsOperador.MoveNext
        Wend
        If flag = False Then
            MsgBox "Propiedades obligadas " & mensajeOperador, vbExclamation, sMensaje
        End If
        ObligaPropiedad = flag
        RsOperador.AbsolutePosition = oPos
        RsOperador.Find "nboton = " & Trim(str(RsOperador!nBoton))
        nOperadorPropiedad = RsOperador!nControl
        For i = 1 To 13
            cmdOperador(i).backColor = vbButtonFace
        Next i
        If RsOperador!nBoton Then
            cmdOperador(RsOperador!nBoton).backColor = vbRed
        End If
        'AsignaPropiedad
    End If
End Function

Private Sub ListarOperadoresConFiltro(ByVal tProducto As String)
'carlos la rosa 2014
  '  Isql = "select * from vOperador where lActivo = 1 " & _
           "AND ((select Count(tCodigoPropiedad) " & _
           "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
           "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 " & _
           "And TPROPIEDAD.tOperador = vOperador.Codigo and tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) " & _
           "order by Codigo"
'carlos la rosa 2014

'lg 2016
    Isql = "usp_Inforest_ObtieneOperadoresXProducto '" & tProducto & "'"

    Set RsOperador = Lib.OpenRecordset("usp_Inforest_ObtieneOperadoresXProducto '" & tProducto & "'", Cn)
    'If RsOperador.RecordCount > 0 Then
    AsignaBoton 13, RsOperador, cmdOperador()
    If RsOperador.RecordCount > 0 Then
        RsOperador.MoveFirst
        Dim i As Integer
        For i = 1 To RsOperador.RecordCount
            If RsOperador!nBoton > 0 Then
                cmdOperador(RsOperador!nBoton).backColor = vbButtonFace
            End If
            RsOperador.MoveNext
        Next i
        
        RsOperador.MoveFirst
            xOperador = RsOperador!codigo
        If RsOperador!nBoton > 0 Then
            cmdOperador(RsOperador!nBoton).backColor = vbRed
        End If
    End If
End Sub
Public Sub AgregarPropiedadBusqueda(ByVal CodigoPropiedad As String, ByVal DescripcionPropiedad As String)
    If lMCPV Then
        If Not ValidaExistenciaProducto() Then
               MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
               RsCabecera.Requery
                RsDetalle.Requery
               Exit Sub
            End If
    End If

  '   HabilitaTimerColor (False)
   Dim ncantidadPropiedad As Double
   Dim Cantidad As Double

    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMObra As Double
    Dim nBotonPropiedad As Double
    nBotonPropiedad = 999
    RsPropiedad.MoveFirst
    RsPropiedad.Find ("Descripcion = '" & DescripcionPropiedad & "'")
     RsPropiedad.Find ("Codigo = '" & CodigoPropiedad & "'")
   
     If Not (RsOperador.EOF Or RsOperador.BOF) Then
        nOperadorPropiedad = Calcular("select isnull(ncontrol,0) as codigo from voperador where codigo='" & RsOperador!codigo & "'", Cn)
     End If
    Dim k As Integer
    For k = 1 To 20
        If cmdPropiedad(k).Caption = DescripcionPropiedad Then
        nBotonPropiedad = k
        Exit For
        End If
    Next k
    
    If nBotonPropiedad <> 999 Then
        If cmdPropiedad(nBotonPropiedad).FontBold = True Then
            cmdPropiedad(nBotonPropiedad).FontBold = False
        Else
            cmdPropiedad(nBotonPropiedad).FontBold = True

        End If
    
    End If
    
    
     
                 
        If wAgregarPropiedad = False Then
          '  cmdPropiedad(nBotonPropiedad).FontBold = False
            If Not RsPropiedad.EOF Then
                If wAgregaCombo Then
                    Cantidad = Calcular("select isnull(ncantidad,1) as codigo from tcombopropiedad where tcodigopedido='" & sPedido & "' and titem='" & sItem & "' and titemcombo='" & xItem & "' and  tproducto='" & sCombo & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' ", Cn)
                    Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and tProducto='" & sCombo & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "'"
                Else
                    Cantidad = Calcular("select isnull(ncantidad,1) as codigo from tproductopropiedad where tcodigopedido='" & sPedido & "' and titem='" & sItem & "' and tproducto='" & sProducto & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' and tenlace='" & RsPropiedad!tEnlace & "'", Cn)
                    If RsPropiedad!nPrecio <> 0 Then
                        txtSuma.Caption = CambiaPrecio(nPVenta - RsPropiedad!nPrecio, sPedido)
                    End If

                    Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tProducto='" & sProducto & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "' and tEnlace='" & RsPropiedad!tEnlace & "'"
                End If
                If Cantidad <> 1 Then
                           lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & DescripcionPropiedad & ": (" & Cantidad & "), ", "")
                Else
                           lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & DescripcionPropiedad & ", ", "")
                End If
    '            lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & DescripcionPropiedad & ", ", "")
            End If
        Else
        
             ncantidadPropiedad = 1
            If RsPropiedad!lsolicitacantidad = 1 Or RsPropiedad!lsolicitacantidad = True Then
                sTipo = "Prepintado"
            
                sCodigo = ncantidadPropiedad
            
                frmNumPad.Show vbModal
                If wEnter And Val(sDescrip) > 0 Then
            
                            ncantidadPropiedad = sDescrip
                        
                End If
            End If
            If nOperadorPropiedad > 0 Then
                If wAgregaCombo Then
                   Isql = "SELECT COUNT(dbo.TCOMBOPROPIEDAD.tCodigoPedido) AS codigo " & _
                          "FROM dbo.TCOMBOPROPIEDAD INNER JOIN dbo.TPROPIEDAD ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TCOMBOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
                          "where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and dbo.TCOMBOPROPIEDAD.tProducto='" & sCombo & "'  and tOperador='" & RsOperador!codigo & "'"
                   If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                      MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                      Exit Sub
                   End If
                Else
                    Isql = "SELECT COUNT(dbo.TPRODUCTOPROPIEDAD.tCodigoPedido) AS codigo FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN " & _
                           "dbo.TPROPIEDAD ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
                           "where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tOperador='" & RsOperador!codigo & "'"
                    If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                       MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                       Exit Sub
                    End If
                End If
            End If
    
         '   cmdPropiedad(nBotonPropiedad).FontBold = True
            If Not RsPropiedad.EOF Then
               nInsumo = IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo)
               nGasto = IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto)
               nMObra = IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra)
               If wAgregaCombo Then
                    Cn.Execute "Insert into TCOMBOPROPIEDAD values ('" & sPedido & "', '" & sItem & "', '" & xItem & "', '" & RsPropiedad!codigo & "', '" & sCombo & "', '" & RsPropiedad!tEnlace & "', " & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nInsumo * -1, ncantidadPropiedad * nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nGasto * -1, ncantidadPropiedad * nGasto) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nMObra * -1, ncantidadPropiedad * nMObra) & ", 1," & ncantidadPropiedad & " , " & IIf(RsPropiedad!codigo = "9999", nInsumo * -1, nInsumo) & ", " & IIf(RsPropiedad!codigo = "9999", nGasto * -1, nGasto) & " , " & IIf(RsPropiedad!codigo = "9999", nMObra * -1, nMObra) & ")"
               Else
                    Cn.Execute "Insert into TPRODUCTOPROPIEDAD values ('" & sPedido & "', '" & sItem & "', '" & RsPropiedad!codigo & "', '" & sProducto & "', '" & RsPropiedad!tEnlace & "', " & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nInsumo * -1, ncantidadPropiedad * nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nGasto * -1, ncantidadPropiedad * nGasto) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nMObra * -1, ncantidadPropiedad * nMObra) & ", 1, " & ncantidadPropiedad & ",  " & IIf(RsPropiedad!codigo = "9999", nInsumo * -1, nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", nGasto * -1, nGasto) & "," & IIf(RsPropiedad!codigo = "9999", nMObra * -1, nMObra) & ")"
                    If RsPropiedad!nPrecio <> 0 Then
                        txtSuma.Caption = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad), sPedido)
                    End If
               End If
                If ncantidadPropiedad <> 1 Then
                
                      lblResumen.Text = lblResumen.Text & RsOperador!descripcion & " " & DescripcionPropiedad & ": (" & ncantidadPropiedad & "), "
                Else
                      lblResumen.Text = lblResumen.Text & RsOperador!descripcion & " " & DescripcionPropiedad & ", "
                End If
            End If
    End If
    
    If wAgregaCombo Then
       RsComboPropiedad.Requery
    Else
       RsProductoPropiedad.Requery
    End If
    
   ' HabilitaTimerColor (True)
End Sub



Private Function Obtener_ProductosNoImpresosPorPedido(ByVal tCodigoPedido) As Recordset
    If wAgregaCombo Then
        Isql = "Select CPedido.*,TProducto.tDetallado from CPedido Left Join TProducto On CPedido.tProductoCombo = TProducto.tCodigoProducto " & _
               "Where CPedido.lImprime = '0' And tItem = '" & sItem & "' And CPedido.tCodigoPedido = '" & tCodigoPedido & "' "
    Else
        Isql = "Select DPedido.*,TProducto.tDetallado from DPedido Left Join TProducto On DPedido.tCodigoProducto = TProducto.tCodigoProducto Where DPedido.tCodigoPedido = '" & tCodigoPedido & "' And DPedido.lImprime = '0'"
    End If
    Set Obtener_ProductosNoImpresosPorPedido = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_OperadoresObligatoriosPorProducto(ByVal tProducto) As Recordset
    Isql = "select * from vOperador where lActivo = 1 AND lObligaPropiedad = 1" & _
       "AND ((select Count(tCodigoPropiedad) " & _
       "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
       "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 " & _
       "And TPROPIEDAD.tOperador = vOperador.Codigo and tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) " & _
       "order by Codigo"
    Set Obtener_OperadoresObligatoriosPorProducto = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_PropiedadesSeleccionadasPorProducto(ByVal tCodigoPedido, ByVal tItem, ByVal tProducto, ByVal tOperador) As Recordset
    If wAgregaCombo Then
        Isql = "Select * From TComboPropiedad " & _
            "Where tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And tCodigoPropiedad In ( " & _
            "Select tCodigoPropiedad from TPropiedad " & _
            "Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    Else
        Isql = "Select * From TProductoPropiedad " & _
            "Where tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And tCodigoPropiedad In ( " & _
            "Select tCodigoPropiedad from TPropiedad " & _
            "Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    End If
    Set Obtener_PropiedadesSeleccionadasPorProducto = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_ProductoDeCombo(ByVal tCombo As String, ByVal tCodigoProducto As String) As Recordset
    Isql = "Select lUnico,tEtiqueta From TCOMBO Where tCombo = '" & tCombo & "' And tCodigoProducto = '" & tCodigoProducto & "'"
    Set Obtener_ProductoDeCombo = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function ObtenerSumaCantidadesEnElCombo(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal tEtiqueta As String) As Double
    Isql = "Select ISNULL(Sum(nCantidad),0) as nCantidad from CPEDIDO " & _
           "Where   tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And tProductoCombo IN ( " & _
           "Select tCodigoProducto From TCOMBO Where tCombo = CPEDIDO.tProducto And tEtiqueta = '" & tEtiqueta & "' And lUnico = '1') "
    Dim Suma As Double
    Dim oRsResultado As Recordset
    Set oRsResultado = Lib.OpenRecordset(Isql, Cn)
    Suma = oRsResultado!nCantidad
    ObtenerSumaCantidadesEnElCombo = Suma
End Function

Private Function ObtenerSumaCantidadesEnElComboSinEsteProducto(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal tItemCombo As String, ByVal tEtiqueta As String) As Double
    Isql = "Select ISNULL(Sum(nCantidad),0) as nCantidad from CPEDIDO " & _
           "Where   tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And tItemCombo <> '" & tItemCombo & "' And tProductoCombo IN ( " & _
           "Select tCodigoProducto From TCOMBO Where tCombo = CPEDIDO.tProducto And tEtiqueta = '" & tEtiqueta & "' And lUnico = '1') "
    Dim Suma As Double
    Dim oRsResultado As Recordset
    Set oRsResultado = Lib.OpenRecordset(Isql, Cn)
    Suma = oRsResultado!nCantidad
    ObtenerSumaCantidadesEnElComboSinEsteProducto = Suma
End Function

Private Function ObtenerDetalleProducto(ByVal tCodigoPedido As String, ByVal tItem As String) As Recordset
    ' oo Isql = "Select D.* ,P.nPrecioVenta as 'PrecioProducto', P.tMoneda as 'tMonedaProducto' from DPedido As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tCodigoPedido='" & tCodigoPedido & "' and D.tItem ='" & tItem & "'"
   '00 28/10/2011
   
   'Isql = " Select D.* ,   CASE (D.tTipoPedido)  WHEN '01' THEN P.nprecioventa when '02' then p.npreciodelivery when '03' then p.npreciollevar when '04' then p.npreciocanal4 when '05' then p.npreciocanal5 END as 'PrecioProducto',P.tMoneda as 'tMonedaProducto' from DPedido As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tCodigoPedido='" & tCodigoPedido & "' and D.tItem ='" & tItem & "'"
    Isql = " Select D.* , " & _
           " case when ( " & _
           " CASE (D.tTipoPedido)  WHEN '01' THEN P.nprecioventa " & _
           " when '02' then p.npreciodelivery " & _
           "            when '03' then p.npreciollevar " & _
           "            when '04' then p.npreciocanal4 " & _
           "            when '05' then p.npreciocanal5 " & _
           " END)=0 then p.nprecioventa else " & _
           " ( CASE (D.tTipoPedido)  WHEN '01' THEN P.nprecioventa " & _
           "            when '02' then p.npreciodelivery " & _
           "            when '03' then p.npreciollevar " & _
           "            when '04' then p.npreciocanal4 " & _
           "            when '05' then p.npreciocanal5 End) end as 'PrecioProducto',P.tMoneda as 'tMonedaProducto' from DPedido As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tCodigoPedido='" & tCodigoPedido & "' and D.tItem ='" & tItem & "'"
'    Isql = " Select D.* ,d.nprecioventa   as 'PrecioProducto',P.tMoneda as 'tMonedaProducto' from DPedido As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tCodigoPedido='" & tCodigoPedido & "' and D.tItem ='" & tItem & "'"
      
    Set ObtenerDetalleProducto = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_CantidadMaximaDeUnicoEtiqueta(ByVal tCodigoPedido As String, ByVal tItem As String, cantidadActual As Double) As Double
    Isql = "Select Sum(P.nCantidad) as Cantidad " & _
           "from CPedido as P Left Join TCOMBO as C On P.tProductoCombo = C.tCodigoProducto " & _
           "where P.tCodigoPedido='" & tCodigoPedido & "' and P.tItem ='" & tItem & "' And C.lUnico = '1' " & _
           "Group By C.tEtiqueta"
     '06/02
    Isql = "Select Sum(P.nCantidad) as Cantidad " & _
           "from CPedido as P Left Join TCOMBO as C ON P.tProducto = C.tCombo AND P.tProductoCombo = C.tCodigoProducto " & _
           "where P.tCodigoPedido='" & tCodigoPedido & "' and P.tItem ='" & tItem & "' And C.lUnico = '1' " & _
           "Group By C.tEtiqueta,c.tcodigoproducto  order by 1 desc"
    Dim oRsResultado As Recordset
    Dim oi As Integer
    Dim CantMax As Double
    CantMax = 0
    CantMax = cantidadActual
    Set oRsResultado = Lib.OpenRecordset(Isql, Cn)
    If oRsResultado.RecordCount > 0 Then
        For oi = 1 To oRsResultado.RecordCount
            If oRsResultado!Cantidad <> CantMax Then
                CantMax = oRsResultado!Cantidad
            End If
            oRsResultado.MoveNext
        Next oi
        Obtener_CantidadMaximaDeUnicoEtiqueta = CantMax
    Else
        Obtener_CantidadMaximaDeUnicoEtiqueta = 0
    End If
End Function
Private Function Obtener_PropiedadesPendientesPorItems(ByVal tProducto As String, ByVal tItem As String) As Boolean
    Dim oj As Integer
    Dim oflag As Boolean
    oflag = True
    Dim oRsOperadoresObligados As Recordset ' Lista de Operadores obligados de un producto
    Set oRsOperadoresObligados = Obtener_OperadoresObligatoriosPorProducto(tProducto)
    For oj = 1 To oRsOperadoresObligados.RecordCount
        Dim oRsPropiedadesDeOperador As Recordset ' Lista de Propiedades por Operador
        Set oRsPropiedadesDeOperador = Obtener_PropiedadesSeleccionadasPorProducto(sPedido, tItem, tProducto, oRsOperadoresObligados!codigo)
        If oRsPropiedadesDeOperador.RecordCount = 0 Then
            oflag = False
           oj = oRsOperadoresObligados.RecordCount
        End If
        oRsOperadoresObligados.MoveNext
    Next oj
    Obtener_PropiedadesPendientesPorItems = oflag
End Function

Private Function ExistenPropiedadesPendientesEnCombos(ByVal tCodigoPedido As String) As Boolean
    Dim oRsCombosPedido As Recordset
    Dim oi, oj, ok As Integer
    Dim oMensaje As String
    oMensaje = "Combo(s) con productos con propiedades obligatorias: "
    Set oRsCombosPedido = Obtener_TodosLosCombosDelPedido(tCodigoPedido)
    Dim oflag As Boolean
    oflag = True
    For oi = 1 To oRsCombosPedido.RecordCount
        Dim oRsProductoNoImpCombo As Recordset
        Set oRsProductoNoImpCombo = Obtener_LosProductosNoImpDelCombo(tCodigoPedido, oRsCombosPedido!tItem)
        For oj = 1 To oRsProductoNoImpCombo.RecordCount
            Dim oRsOperadores As Recordset
            Set oRsOperadores = Obtener_OperadoresObligatoriosDeUnProducto(oRsProductoNoImpCombo!tProductoCombo)
            For ok = 1 To oRsOperadores.RecordCount
                Dim oRsPropiedadesDeOperador As Recordset
                Set oRsPropiedadesDeOperador = Obtener_PropiedadesSeleccionadasPorProductoDeCombo(tCodigoPedido, oRsCombosPedido!tItem, oRsProductoNoImpCombo!tItemCombo, oRsProductoNoImpCombo!tProductoCombo, oRsOperadores!codigo)
                If oRsPropiedadesDeOperador.RecordCount = 0 Then
                    oflag = False
                    oMensaje = oMensaje + "(" + oRsCombosPedido!tDetallado + ")"
                    oj = oRsProductoNoImpCombo.RecordCount
                    ok = oRsOperadores.RecordCount
                End If
                oRsOperadores.MoveNext
            Next ok
            oRsProductoNoImpCombo.MoveNext
        Next oj
        oRsCombosPedido.MoveNext
    Next oi
    
    If (oflag = False) Then
        MsgBox (oMensaje)
    End If
    ExistenPropiedadesPendientesEnCombos = oflag
End Function

Private Function Obtener_TodosLosCombosDelPedido(ByVal tCodigoPedido As String) As Recordset
    Isql = "Select  DPedido.*,TProducto.tDetallado From DPedido Left Join TProducto On DPedido.tCodigoProducto = TProducto.tCodigoProducto " & _
           "Where   DPedido.tCodigoPedido = '" & tCodigoPedido & "' And DPedido.lCombinacion = '1'"
    Set Obtener_TodosLosCombosDelPedido = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_LosProductosNoImpDelCombo(ByVal tCodigoPedido As String, ByVal tItem As String) As Recordset
    Isql = "Select * from CPedido Where tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And lImprime = '0'"
    Set Obtener_LosProductosNoImpDelCombo = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_OperadoresObligatoriosDeUnProducto(ByVal tProducto As String) As Recordset
    Isql = "select * from vOperador where lActivo = 1 AND lObligaPropiedad = 1 AND ( " & _
           "(select Count(tCodigoPropiedad) FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
           "Where TPROPIEDAD.lActivo = 1 And IsNull(tOperador.lStockMenos, 0) <> 1 And TPROPIEDAD.tOperador = vOperador.Codigo " & _
           "And tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) order by Codigo"
    Set Obtener_OperadoresObligatoriosDeUnProducto = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_PropiedadesSeleccionadasPorProductoDeCombo(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal tItemCombo As String, ByVal tProducto As String, ByVal tOperador As String) As Recordset
    Isql = "Select * From TComboPropiedad " & _
           "Where tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And tItemCombo = '" & tItemCombo & "' And tCodigoPropiedad  In ( " & _
           "Select tCodigoPropiedad from TPropiedad Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    Set Obtener_PropiedadesSeleccionadasPorProductoDeCombo = Lib.OpenRecordset(Isql, Cn)
End Function
'---------------------------------------------
Private Function ExistenPropiedadesPendientesEnPedido(ByVal tCodigoPedido As String) As Boolean
    Dim oRsDPedidoNoImp As Recordset 'Lista de Productos no impresos
    Set oRsDPedidoNoImp = Obtener_ProductosNoImpresosPorPedido(sPedido)
    Dim oi As Integer
    Dim oj As Integer
    Dim oflag As Boolean
    Dim oMensaje As String
    oMensaje = "Item(s) con obligatoriedad de propiedad: "
    oflag = True
    'Para cada`producto de DPedido cual lImprime = '0'
    For oi = 1 To oRsDPedidoNoImp.RecordCount
    'Obtener Operadores Obligatorios Filtrados
        Dim oRsOperadoresObligados As Recordset ' Lista de Operadores obligados de un producto
        Set oRsOperadoresObligados = Obtener_OperadoresObligatoriosPorProducto(oRsDPedidoNoImp!tCodigoProducto)
        'Para cada operador Obtener la lista de propiedades
        For oj = 1 To oRsOperadoresObligados.RecordCount
            Dim oRsPropiedadesDeOperador As Recordset ' Lista de Propiedades por Operador
            Set oRsPropiedadesDeOperador = Obtener_PropiedadesSeleccionadasPorProducto(sPedido, oRsDPedidoNoImp!tItem, oRsDPedidoNoImp!tCodigoProducto, oRsOperadoresObligados!codigo)
            'Verificar si en TProductoPropiedad existe para el item y tCodigoPropiedad IN (lista de propiedades)
            If oRsPropiedadesDeOperador.RecordCount = 0 Then
                oflag = False
                oMensaje = oMensaje + "(" + oRsDPedidoNoImp!tDetallado + ")"
                oj = oRsOperadoresObligados.RecordCount
            End If
            oRsOperadoresObligados.MoveNext
        Next oj
        oRsDPedidoNoImp.MoveNext
    Next oi
    If oflag = False Then
        MsgBox (oMensaje)
    End If
    ExistenPropiedadesPendientesEnPedido = oflag
End Function

Private Function ObtenerCodigoMozo(ByVal tResumido As String) As String
    Dim RsMozo As Recordset
    Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by nBoton"
    Set RsMozo = Lib.OpenRecordset(Isql, Cn)
    RsMozo.Filter = "tResumido = '" & tResumido & "'"
    If RsMozo.RecordCount = 0 Then
        txtMozo.Caption = "Sin " & Mesero
        ObtenerCodigoMozo = "0000"
    Else
        txtMozo.Caption = tResumido
        ObtenerCodigoMozo = RsMozo!codigo
    End If
End Function
'Fin Oscar Ortega-----------------------------
Public Function verificaCantidadDeItemsCombos(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal numeroDeItemCombos As Double, ByVal ncantidadNueva As Double) As Boolean
    Dim cantidadMaximoNueva As Double
    Dim cantidadMaximoPosible As Double
    Dim X As Integer
    Dim oRstRecorriendoCombo As New Recordset
    Dim oRstCantidadNUnicos As New Recordset
    
    verificaCantidadDeItemsCombos = False
    cantidadMaximoNueva = numeroDeItemCombos * ncantidadNueva
    cantidadMaximoPosible = 0
    Set oRstRecorriendoCombo = Lib.OpenRecordset("select tcombo.ncantidad ,cpedido.tproductocombo from cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where cpedido.tcodigopedido='" & tCodigoPedido & "' and cpedido.titem='" & tItem & "' and lfijo=1", Cn)
    If Not (oRstRecorriendoCombo.EOF Or oRstRecorriendoCombo.BOF) Then
        oRstRecorriendoCombo.MoveFirst
        For X = 0 To oRstRecorriendoCombo.RecordCount - 1
            cantidadMaximoPosible = cantidadMaximoPosible + (oRstRecorriendoCombo!nCantidad * ncantidadNueva)
            oRstRecorriendoCombo.MoveNext
        Next X
    End If
    
    Set oRstCantidadNUnicos = Lib.OpenRecordset("select isnull(sum(cpedido.ncantidad),0) from cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where tcodigopedido='" & tCodigoPedido & "' and titem='" & tItem & "' and tcombo.lfijo=0 ", Cn)
    If Not (oRstCantidadNUnicos.EOF Or oRstCantidadNUnicos.BOF) Then
        cantidadMaximoPosible = cantidadMaximoPosible + oRstCantidadNUnicos.Fields(0)
    End If
    
    If cantidadMaximoNueva >= cantidadMaximoPosible Then
        verificaCantidadDeItemsCombos = True
    End If
    
End Function



'luchiinsumo
Public Sub verificatitulo()
                    
                    'INSUMOCRITICO23
                    
                    Dim rsInsumo As New ADODB.Recordset
                    If Calcular("select isnull(lControlInsumoCritico,0) as codigo from tproducto  INNER JOIN Dpedido on tproducto.tcodigoproducto=dpedido.tcodigoproducto where dpedido.tcodigopedido='" & sPedido & "' and titem='" & sItem & "'", Cn) = True Then
                                    Set rsInsumo = Lib.OpenRecordset("select isnull(tcodigoinsumo,'') tcodigoinsumo , isnull(tinsumo.descripcion,'') ,isnull(nstock,0) , dpedido.ncantidad from tproducto inner join tinsumo on tproducto.tcodigoinsumo =tinsumo.tcodigo inner join dpedido on tproducto.tcodigoproducto=dpedido.tcodigoproducto  where    dpedido.tcodigopedido='" & sPedido & "' and titem='" & sItem & "' and tinsumo.lactivo=1", Cn)
                                    If Not (rsInsumo.EOF Or rsInsumo.BOF) Then
                                            Me.Caption = muestra & "   Insumo Crítico --->   " & rsInsumo.Fields(1) & " =  Stock: " & str(rsInsumo.Fields(2)) & " / Solicitado: " + str(rsInsumo.Fields(3))
                                    End If
                            Else
                                    Me.Caption = muestra
                    End If
                    'INSUMOCRITICO
End Sub
'luchoinsumo


'diaContable
Public Function obtieneDiaContable() As Date
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@lDiaContable", adBoolean, adParamInput, 1, lDiaContable
                  oComando.CreateParameter "@sHoraCierre", adVarChar, adParamInput, 5, tHoraCierreDiaContable
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamOutput, 10, DiaContable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneDiaContable = oComando.GetParameterValue("@fDiaContable")
End Function
'diaContable


Private Sub CalculaAplicaTope(nTope As Double)
    Dim sCriterio As String
    Dim lAcumulable As Boolean
    Dim nOferta As Double
    Dim nSuma As Double
    
    nSuma = Calcular("SELECT sum(nPrecioOficial*nCantidad) as Codigo FROM dbo.DPEDIDO LEFT OUTER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where tCodigoPedido='" & sPedido & "' and lDescuento=1", Cn)

If RsDetalle.RecordCount <> 0 Then
   RsDetalle.MoveFirst
   
   Do While Not RsDetalle.EOF
      'Busca Oferta
      nPVenta = 0
      sCriterio = "tCodigoProducto ='" & RsDetalle!tCodigoProducto & "' and lActivo=1"
      sCriterio = sCriterio & " and (tFrecuencia='00' or tFrecuencia='0" & Weekday(FechaServidor(), vbMonday) & "' or (tFrecuencia='99' and fFecha='" & Format(FechaServidor(), "yyyy/MM/dd 00:00") & "') and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "')"
      sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "'))"
        
      Isql = "select * from TOFERTA where " & sCriterio
      Set RsOferta = Lib.OpenRecordset(Isql, Cn)
      
      lAcumulable = True
      nOferta = 0
      Acumulado = 0
      
      If RsOferta.RecordCount > 0 Then
         RsOferta.MoveFirst
         lAcumulable = RsOferta!lAcumulable
         nOferta = RsDetalle!nPrecioOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
      End If
      
      If RsDetalle!lDescuento And lAcumulable = True Then
         
         xDescuento = (RsDetalle!nPrecioOficial - nOferta) * (RsDetalle!nCantidad * 100 / nSuma)
         nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((nTope * xDescuento / 100) / RsDetalle!nCantidad)
         
          Select Case pais ' ok
            Case "001" 'Bolivia
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                         
            Case Else 'Peru, Ecuador
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = 1 + (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
          
          End Select
         Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
                "nDescuento = " & RsDetalle!nPrecioOficial - nPVenta & ", " & _
                "nRecargo = " & nRecargo & ", " & _
                "nPrecioOficial = " & RsDetalle!nPrecioOficial & ", " & _
                "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
                "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
                "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
                "nPrecioVenta = " & nPVenta & ", " & _
                "nventa = " & nPVenta * RsDetalle!nCantidad & ", " & _
                "nCantidad = " & RsDetalle!nCantidad & ", " & _
                "nImpuesto1 = " & nImpuesto1 * RsDetalle!nCantidad & ", " & _
                "nImpuesto2 = " & nImpuesto2 * RsDetalle!nCantidad & ", " & _
                "nImpuesto3 = " & nImpuesto3 * RsDetalle!nCantidad & ", " & _
                "tCortesia = '" & sCortesia & "' " & _
                "where tItem = '" & RsDetalle!tItem & "' and tCodigoPedido ='" & sPedido & "'"
                Cn.Execute Isql
      End If
   RsDetalle.MoveNext
   Loop
End If

End Sub


Public Sub InicializaVariableCabecera()
    wDetalle = False
    wAgrega = True
    sPedido = ""
    sHabitacion = ""
    sReserva = ""
    sCodigoDescuento = ""
    tAutorizaDescuento = ""
    sPuntoVenta = ""
    txtPuntoVenta.Caption = ""
    sMesa = ""
    nTiempoDelivery = 0
    xDescuento = 0
    sAutoriza = ""
    frmVenta.txtAdulto.Text = 0
    frmVenta.txtNino.Text = 0
    frmVenta.txtJuntar.Text = 0
    lPrecuenta = False
    
    txtEnvioAntes.Text = 0
    vOrigenVentas = ""
    'pariente2013
    sCodigoInvitado = ""
    sCodigoParienteSeleccionado = ""
    'pariente2013
End Sub

Public Sub InicializaVariableCanales()
    'txtMozo.Caption = ""
    If wAgrega Then
    txtMotorizado.Caption = ""
    txtDescuento.Caption = "0.00"
    txtCliente.Caption = ""
    txtDireccion.Caption = ""
    txtTelefono.Caption = ""
    txtFecha.Caption = FechaServidor()
    txtFechaProg.Caption = ""
    txtObservacion.Caption = ""
    txtMontoMax.Caption = "0"
          'entregarA
    txtEntregar.Caption = ""
    sTipoAtencion = "01"
    'sMozo = "0000"
    sMotorizado = "0000"
    sCliente = ""
        'pariente2013
    sCodigoInvitado = ""
    sCodigoParienteSeleccionado = ""
    End If
    'pariente2013
    'Canales de Venta
    RsCanalesVenta.MoveFirst
    RsCanalesVenta.Filter = "Codigo = '" & sTipoPedido & "'"
    
    If RsCanalesVenta.RecordCount = 0 Then
        Set RsCanalesVenta = Lib.OpenRecordset("select * from vTipoPedido", Cn)
        MsgBox "No se puede mostrar la información, el canal de venta predeterminado para esta caja esta Inactivo", vbCritical, sMensaje
        cmdOpcion_Click (2)
        Exit Sub
    End If

    'lActivaMozo = IIf(IsNull(RsCanalesVenta!lActivaMozo), False, RsCanalesVenta!lActivaMozo)
    lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
    lCanalDelivery = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
    lCanalCentralPedidos = IIf(IsNull(RsCanalesVenta!lCanalCentralPedidos), False, RsCanalesVenta!lCanalCentralPedidos)
    lObligaMotorizado = IIf(IsNull(RsCanalesVenta!lObligaMotorizado), False, RsCanalesVenta!lObligaMotorizado)
    lObligaMesa = IIf(IsNull(RsCanalesVenta!lObligaMesa), False, RsCanalesVenta!lObligaMesa)
    lObligaPax = IIf(IsNull(RsCanalesVenta!lObligaPax), False, RsCanalesVenta!lObligaPax)
    lObligaFechaEntrega = IIf(IsNull(RsCanalesVenta!lObligaIngresoFechaEntrega), False, RsCanalesVenta!lObligaIngresoFechaEntrega)
    lObligaClienteFrecuente = IIf(IsNull(RsCanalesVenta!lObligaClienteFrecuente), False, RsCanalesVenta!lObligaClienteFrecuente)
    lObligaOrigenVenta = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
    'entregarA
    lObligaEntregarA = IIf(IsNull(RsCanalesVenta!lObligaEntregarA), False, RsCanalesVenta!lObligaEntregarA)
    'origen de ventas
    lOrigenVentas = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
        
    fraMotorizado.Visible = False
    
    'Origen de ventas
    Call ActivarOrigenVentas
    
End Sub
Private Sub ActivarOrigenVentas()
On Error GoTo fin
       Me.fraOrigenVentas.Visible = True
       Isql = "select * from vOrigenVenta where Activo = 1 and Visible = 1 and CodCanalVenta='" & sTipoPedido & "' and CodOrigenVenta not in (select CodigoOrigenVenta from TCAJAORIGEN_BLOQUEO  where tcaja='" & sCaja & "' ) Order by Boton"
       Set RsOrigenVentas = Lib.OpenRecordset(Isql, Cn)
       
       If RsOrigenVentas.RecordCount > 0 Then
            AsignaBotonOrigenVentas 19, RsOrigenVentas, Me.cmdOrigen()
        Else
            AsignaBotonOrigenVentas 19, RsOrigenVentas, Me.cmdOrigen()
            Me.fraOrigenVentas.Visible = False
            If lObligaOrigenVenta = True Then
                MsgBox "Obligatoriedad de origen de ventas:" & vbNewLine & "No hay Origenes de venta visibles configuradas, " & vbNewLine & "Consultar con su administrador!!!", vbInformation
                Me.fraOrigenVentas.Visible = True
            End If
            
            If lActivaMotorizado Then
                fraMotorizado.Visible = True
            Else
                fraMotorizado.Visible = False
            End If
       End If

Exit Sub
fin:
End Sub





Private Sub Timer_Timer()
   'Pedido Timer
   RsTimer.Requery
   Dim i As Integer
   If RsTimer.RecordCount <> 0 Then
      RsTimer.MoveFirst
      While Not RsTimer.EOF()
            If Format(RsTimer!fProgramacion, "dd/mm/yyyy") = Format(FechaServidor(), "dd/mm/yyyy") And RsTimer!nMinutosAntesEnvio >= DateDiff("n", Format(Time, "hh:mm"), Format(RsTimer!fProgramacion, "hh:mm")) Then
            
               'Screen.MousePointer = vbHourglass
               Isql = "select * from vPedido Where lImprime = 0 And lImprimeArea = 1 and Codigo = '" & RsTimer!codigo & "' " & _
                      "ORDER BY tItem,tetiqueta,combo"
                
               Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                       
               If RsImpresion.RecordCount = 0 Then
                  LimpiaRs
                  'MsgBox "No existen Productos Nuevos a Imprimir", vbExclamation, sMensaje
               Else
                  RsArea.MoveFirst
                  For i = 1 To RsArea.RecordCount
                      RsImpresion.Filter = "Codigo= '" & RsTimer!codigo & "' and tArea = '" & RsArea!tArea & "'"
                      
                      If RsImpresion.RecordCount <> 0 Then
                         RsImpresion.MoveFirst
                         ImprimePedido RsImpresion, "N", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Inforest"
                      End If
                      RsArea.MoveNext
                  Next i
                  Cn.Execute "Update DPEDIDO Set lImprime = 1 where tCodigoPedido = '" & RsTimer!codigo & "'"
                  RsDetalle.Requery
               End If
               LimpiaRs
            
               'Isql = "Update MPEDIDO Set tTipoAtencion = '01' where tCodigoPedido = '" & RsTimer!codigo & "'"
               'Cn.Execute Isql
               'RsCabecera.Requery
            End If
            RsTimer.MoveNext
      Wend
      
    
   End If
   
End Sub
  '----------------- WEB APP -- PARA PINTAR EL COLOR
Private Sub TimerMESA247_Timer()
    On Error GoTo fin
       If MESA247 Then
            Isql = "select tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND isnull(idPedidoMesa,'')<>'' "
            Set RsPedido = Lib.OpenRecordset(Isql, Cn)
            If RsPedido.RecordCount > 0 Then
                Me.cmdMesa247.Caption = "(" + CStr(RsPedido.RecordCount) + ") MESA 24/7"
                Me.cmdMesa247.Visible = True
            Else
                Me.cmdMesa247.Visible = False
                
            End If
        End If
        contadormesa = contadormesa + 1
        If MESA247 And Me.cmdMesa247.Visible = True Then
            If Me.cmdMesa247.backColor = &HFF00& Then
               Me.cmdMesa247.backColor = &H8000000F
            Else
               Me.cmdMesa247.backColor = &HFF00&
            End If

        End If

    
   Exit Sub
fin:
End Sub


Private Sub cmdMesa247_Click()
     '----------------
     Dim CodPedido1, NomMesa As String
        sTemp = ""
      Isql = "exec usp_WS_MESA_Inforest '','',2"
      Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Cliente", 2, "Descripcion", 5800, 0, 0, "")
      frmBusquedaRapida.nPredeterm = 1
      frmBusquedaRapida.Show vbModal
    '-------------
        CodPedido1 = sCodigo
        NomMesa = Calcular("select tDetallado as codigo from TMESA where tCodigoMesa='" & sMesa & "'", Cn)
    If CodPedido1 <> "" Then
   
            Dim value As String
               'value = MsgBox("¿Enlazar Pedido:( " & NomMesa & " / " & sPedido & " )a usuario Mesa24/7?", vbYesNoCancel)
            If RsCabecera.RecordCount > 0 Then
               'value = MsgBox("¿Enlazar Pedido:( " & NomMesa & " / " & sPedido & " )a usuario Mesa24/7?", vbYesNoCancel)
               value = MsgBox("¿Enlazar Pedido:( " & NomMesa & " / " & sPedido & " )a usuario Mesa24/7?" + vbNewLine + "Si elige No Se generara un nuevo Pedido", vbYesNoCancel)
               'value = vbNo
            Else
               value = vbNo
            End If
            
            If value = vbYes Then
                 If NomMesa = "0" Then
                    If lObligaMesa And sMesa = "" Then
                        MsgBox "Asigne una Mesa", vbExclamation, sMensaje
                        cmdCabecera_Click (13)
                        'Exit Sub
                    End If
                    If sMesa = "" Then
                        If MsgBox("No hay Mesa seleccionada, ¿Desea continuar?", vbYesNo) = vbNo Then
                          Exit Sub
                        End If
                    Else
                        Cn.Execute "update mpedido set tmesa='" & sMesa & "' where tcodigopedido='" & sPedido & "'"
                        Cn.Execute "update TMESA set tEstadoMesa ='02' where tCodigoMesa='" & sMesa & "'"
                        sMesa = ""
                    End If
                 End If
                 
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tTurno='MOZO' and isnull(idPedidoMesa,'')=''  and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                    MsgBox "El Pedido ya fue Facturado ó Ya esta Enlazado A Usuario de Mesa24/7", vbExclamation, sMensaje
                     RsCabecera.Requery
                     RsDetalle.Requery
                    Exit Sub
                End If
                       Isql = "exec usp_WS_MESA_Inforest '" & sPedido & "', '" & CodPedido1 & "',1"
                        Cn.Execute Isql
                    GoTo fin
            Else
                    If value = vbCancel Then
                        Exit Sub
                    End If
            
            End If

     ''''' OBLIGATORIEDAD DE MESA
        sMesa = ""
        If lObligaMesa And sMesa = "" Then
            MsgBox "Asigne una Mesa", vbExclamation, sMensaje
            cmdCabecera_Click (13)
            'Exit Sub
        End If
        If sMesa = "" Then
            If MsgBox("No hay Mesa seleccionada, ¿Desea continuar?", vbYesNo) = vbNo Then
              GoTo fin:
            End If
        End If

        Dim delmax As Integer
        Dim cantPed As Integer
        cantPed = 1
        delmax = 0
        Isql = ""
        
        Isql = "select tCodigoPedido from MPEDIDO where tTurno<>'' and tMozo<>'0000' and tCaja<>'' and idPedidoMesa<>'' and tCodigoPedido= '" + sCodigo + "'"
        Set RsPedido = Lib.OpenRecordset(Isql, Cn)

        If RsPedido.RecordCount = 0 And CodPedido1 <> "" Then
            Isql = "Update Mpedido Set tmesa=  '" + sMesa + "', tTurno = '" + sTurno + "', TCaja = '" + sCaja + "' ,tMozo = '" + sMozo + "' Where tCodigoPedido = '" + CodPedido1 + "'"
            Cn.Execute Isql
            Cn.Execute "update TMESA set tEstadoMesa ='02' where tCodigoMesa='" & sMesa & "'"
            sMesa = ""
        Else
            
        End If
  End If
fin:
        Me.cmdMesa247.Visible = False
        RsCabecera.Requery
        RsDetalle.Requery
        RsCombo.Requery
        RsComboPropiedad.Requery
        RsProductoPropiedad.Requery
        sMesa = ""
        Me.cmdOpcion(10).Visible = True
    End Sub
  

Private Sub ActivaMesa247(dat As Boolean)
    If MESA247 Then
        If dat Then
            TimerMESA247.Enabled = True
            TimerMESA247.Interval = 500
            cmdMesa247.Visible = True
        Else
            TimerMESA247.Enabled = False
            TimerMESA247.Interval = 0
            cmdMesa247.Visible = False
        End If
    End If
End Sub

Private Function VerificaComboMaximo() As Boolean
   'Miguel Mendoza - 10062019 XD XD XD
   Dim nPos As Integer
   VerificaComboMaximo = True
   With RsDetalle
      If .RecordCount > 0 Then
         .MoveFirst

         Do While Not .EOF
            iProductoVerificaComboMaximo = 0
            iProductoVerificaComboMinimo = 0
            iProductoVerificaComboUsado = 0
            iProductoVerificaComboMaximo = _
            Calcular("select isnull(ncombinacion,0) codigo FROM TPRODUCTO WHERE tCodigoProducto='" & Trim(.Fields(3)) & "'", Cn)

            iProductoVerificaComboMinimo = _
            Calcular("select isnull(ncombinacionminima,0) codigo FROM TPRODUCTO WHERE tCodigoProducto='" & Trim(.Fields(3)) & "'", Cn)
            
            If iProductoVerificaComboMaximo > 0 Then
               sProductoVerificaComboUsado = _
               Calcular("select isnull(tdetallado,'') codigo FROM tproducto WHERE tCodigoProducto='" & Trim(.Fields(3)) & "'", Cn)

               iProductoVerificaComboUsado = _
               Calcular("select isnull(sum(nCantidad),0) codigo FROM CPEDIDO WHERE tCodigoPedido='" & sPedido & "' " & _
               "and tproducto='" & Trim(.Fields(3)) & "' and titem='" & RsDetalle!tItem & "'", Cn)

               If iProductoVerificaComboUsado > iProductoVerificaComboMaximo * CInt(.Fields(15)) Then
                  MsgBox "La cantidad máxima para el Combo " & UCase(Trim(sProductoVerificaComboUsado)) & " debe ser de " & _
                  Trim(CStr(iProductoVerificaComboMaximo * CInt(.Fields(15)))), vbCritical + vbOKOnly, "Inforest"
                  VerificaComboMaximo = False
                  'RsDetalle.Requery
                    nPos = RsDetalle.AbsolutePosition
                    RsDetalle.Requery
                    RsDetalle.AbsolutePosition = nPos
                  Exit Function
               End If
               
               If iProductoVerificaComboUsado < iProductoVerificaComboMinimo * CInt(.Fields(15)) Then
                  MsgBox "La cantidad mínima para el Combo " & UCase(Trim(sProductoVerificaComboUsado)) & " debe ser de " & _
                  Trim(CStr(iProductoVerificaComboMinimo * CInt(.Fields(15)))), vbCritical + vbOKOnly, "Inforest"
                  VerificaComboMaximo = False
                  'RsDetalle.Requery
                    nPos = RsDetalle.AbsolutePosition
                    RsDetalle.Requery
                    RsDetalle.AbsolutePosition = nPos
                  Exit Function
               End If
           End If
            .MoveNext
         Loop
      End If
   End With
End Function ' lBuscarClienteBarraMesero


Private Sub txtbuscarpedidobarra_KeyPress(KeyAscii As Integer)
On Error GoTo fin
    Dim rsPed As Recordset
    lEstadoBuscarPedido = False
    If lBuscarPedidoBarraMesero Then
        If KeyAscii = 13 And Trim(txtbuscarpedidobarra.Text) <> "" Then
          cmdOpcion_Click (2)
          
          Isql = ""
          Isql = "select * from mpedido where tcodigopedido = '" & Trim(Mid(txtbuscarpedidobarra, 1, 10)) & "'"
          Set rsPed = Lib.OpenRecordset(Isql, Cn)
    
          If rsPed.RecordCount = 0 Then MsgBox "Pedido no encontrado : " & Trim(txtbuscarpedidobarra.Text): txtbuscarpedidobarra.Text = "": Exit Sub
    
          Isql = ""
          Isql = "select * from mpedido c inner join dpedido d on c.tcodigopedido = d.tcodigopedido " & _
          "and c.tcodigopedido = '" & Trim(Mid(txtbuscarpedidobarra, 1, 10)) & "' " & _
          "and d.limprime = 0"
    
          Set rsPed = Lib.OpenRecordset(Isql, Cn)
          If Not (Not rsPed.BOF And Not rsPed.EOF) Then
             MsgBox "No se puede mostrar Pedido: " & Trim(Mid(txtbuscarpedidobarra, 1, 10)) & Chr(10) & Chr(13) & _
             "Pedido enviado a las areas de producción en su totalidad."
             txtbuscarpedidobarra.Text = ""
             Exit Sub
          End If
          
          Cn.Execute "update mpedido set tmozo='" & sMozo & "' where tcodigopedido = '" & Trim(Mid(txtbuscarpedidobarra, 1, 10)) & "'"
          Isql = ""
          Isql = "select * from vPedidoCabecera " & _
                 "where tEstadoPedido <> '03' and " & _
                 "codigo = '" & Trim(Mid(txtbuscarpedidobarra, 1, 10)) & "'"
    
          Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
          Set grdCabecera.DataSource = RsCabecera
    
          Isql = ""
          Isql = "select * from vPedidoDetalle " & _
                 "where tEstadoItem ='N' and " & _
                 "tcodigopedido = '" & Trim(Mid(txtbuscarpedidobarra, 1, 10)) & "'"
       
          Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
          Set grdDetalle.DataSource = RsDetalle
          
          Call ProcesoBuscarPedidoBarra(False)
          lEstadoBuscarPedido = True
       End If
    End If
    Exit Sub
fin:
     MsgBox "Error al buscar : " & Error
End Sub

Private Sub ProcesoBuscarPedidoBarra(ByVal Valor As Boolean)
On Error GoTo fin
    Me.fraPlato.Enabled = Valor
    'Me.fraCombo.Enabled = False
    'Me.fraEliminacion.Enabled = False
    'Me.fraPedido.Enabled = False
    
    Me.cmdCabecera(0).Enabled = Valor
    Me.cmdCabecera(1).Enabled = Valor
    Me.cmdCabecera(2).Enabled = Valor
    'Me.cmdCabecera(3).Enabled = False
    Me.cmdCabecera(4).Enabled = Valor
    Me.cmdCabecera(5).Enabled = Valor
    
    
    Me.cmdDetalle(0).Enabled = Valor
    Me.cmdDetalle(1).Enabled = Valor
    Me.cmdDetalle(2).Enabled = Valor
    Me.cmdDetalle(3).Enabled = Valor
    Me.cmdDetalle(4).Enabled = Valor
    Me.cmdDetalle(7).Enabled = Valor
    Me.cmdDetalle(8).Enabled = Valor
    Me.cmdDetalle(6).Enabled = Valor
    
    Me.cmdNavegar(1).Enabled = Valor
    Me.cmdNavegar(2).Enabled = Valor
    Me.cmdNavegar(3).Enabled = Valor
    Me.cmdNavegar(4).Enabled = Valor
    Me.cmdNavegar(5).Enabled = Valor
    Me.cmdNavegar(6).Enabled = Valor
    Me.cmdNavegar(7).Enabled = Valor
    Me.cmdNavegar(8).Enabled = Valor
    Me.cmdNavegar(9).Enabled = Valor
    Me.cmdNavegar(10).Enabled = Valor
    Me.cmdNavegar(11).Enabled = Valor
    
    cmdOpcion(13).Enabled = Valor
    cmdOpcion(12).Enabled = Valor
    cmdOpcion(11).Enabled = Valor
    cmdOpcion(10).Enabled = Valor
    cmdOpcion(5).Enabled = Valor
    cmdOpcion(4).Enabled = Valor
    cmdOpcion(1).Enabled = Valor
    'cmdOpcion(0).Enabled = Valor
    'Me.cmdMesa247.Visible = Valor
    
    cmdOpcion(2).Enabled = Not Valor
    cmdOpcion(2).Visible = Not Valor
Exit Sub
fin:
MsgBox "Error: " & Error
End Sub
Private Sub TerminaprocesoBarra()
    
   grdCabecera.DataSource = Nothing
   grdDetalle.DataSource = Nothing
    
   Call ProcesoBuscarPedidoBarra(True)
   lEstadoBuscarPedido = False
   
   If lSomelier Then
      Isql = "SELECT * from vPedidoCabecera " & _
             "where tEstadoPedido = '01' and tTurno = 'MOZO' order by Codigo"
   Else
      Isql = "SELECT * from vPedidoCabecera " & _
             "where tEstadoPedido = '01' and tTurno = 'MOZO' and tMozo='" & sMozo & "' order by Codigo"
   End If
   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
   grdCabecera.DataSource = RsCabecera
   
   If lSomelier Then
      Isql = "select * from vPedidoDetalle " & _
             "where tEstadoItem ='N' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tTurno = 'MOZO') " & _
             "ORDER BY tCodigoPedido, tItem"
   Else
      Isql = "select * from vPedidoDetalle " & _
             "where tEstadoItem ='N' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tTurno = 'MOZO' and tMozo='" & sMozo & "') " & _
             "ORDER BY tCodigoPedido, tItem"
   End If
   Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
   grdDetalle.DataSource = RsDetalle
   txtbuscarpedidobarra.Text = ""
   AsignaPedido
End Sub
Private Sub txtBarra_KeyPress(KeyAscii As Integer)
    ' Productos por Lectora de barras
    If KeyAscii = 13 And Trim(txtBarra.Text) <> "" Then
        
'        If lRotulado = True Then
'              'CESAR ROTULADO
'              Dim rCodigoEtiqueta As String
'              Dim rCodigoProducto As String
'              Dim rLenBarra As String
'              Dim rCantidad As Double
'              Dim X As Integer
'
'              rLenBarra = Len(Trim(txtBarra.Text))
'              X = rLenBarra - 31
'
'              If nLongitudBarra <> 0 Then
'                rCodigoProducto = Mid(txtBarra.Text, nLongitudBarra + 1, 7)
'                rCodigoEtiqueta = Mid(txtBarra.Text, 1, nLongitudBarra)
'
'                If lCapturaPeso Then
'                   rCantidad = Val(Mid(txtBarra.Text, 31 + 1, X))
'                   InsertaProductoRotulado rCodigoProducto, rCantidad, rCodigoEtiqueta
'                Else
'                   rCantidad = 1
'                   InsertaProductoRotulado rCodigoProducto, rCantidad, rCodigoEtiqueta
'                End If
'              Else
'                  MsgBox "Error: Longitud de barra no registrada", vbCritical, sMensaje
'              End If
'              txtBarra.Text = ""
'
'        Else
              If RsCabecera.RecordCount <> 0 Then
                 RsProducto.Filter = adFilterNone
                 RsProducto.MoveFirst
                 RsProducto.Find "tbarra = '" & Trim(txtBarra.Text) & "'"
                  
                  If Not RsProducto.EOF Then
                    'INSUMOCRITICO23
                            If validadIngresoProducto(RsProducto!codigo) = False Then
                                Exit Sub
                            End If
                   'INSUMOCRITICO23
                  End If
             
                  If Not RsProducto.EOF Then
                    sProducto = RsProducto!codigo
                    sComanda = IIf(sTemp = "", "", Mid("0000000000", 1, 10 - Len(Trim(sTemp))) & Trim(sTemp))
        
                    If lBal And RsProducto!lBalanza Then
                       Dim nResultado As Double
                       nResultado = Pesar(nBalanzaPuerto)
                       If (lActivaTresDecimales) Then
'                        nResultado = ""
                        nResultado = Format(nResultado, "#,##0.000")
                       Else
'                        nResultado = ""
                        nResultado = Format(nResultado, "#,##0.00")
                       End If
                       If nResultado > 0 Then
                          InsertaProducto nResultado
                       End If
                    Else
                    nCantidad = 1
                       InsertaProducto 1
                    End If
        
                    If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                       lPropiedad = True
                    End If
                  Else
                    If nLongitudBarra > 0 Then
                       RsProducto.MoveFirst
                       RsProducto.Find "tbarra = '" & Trim(Mid(txtBarra.Text, 1, nLongitudBarra)) & "'"
                       If Not RsProducto.EOF Then
                          sProducto = RsProducto!codigo
                          sComanda = IIf(sTemp = "", "", Mid("0000000000", 1, 10 - Len(Trim(sTemp))) & Trim(sTemp))
                          Dim nCantidadBarra As Double
        
                          If lCapturaPeso Then
                          
                            If EAN13 Then
                                nCantidadBarra = Val(Mid(txtBarra.Text, nLongitudBarra + 1, 1) + "." + Mid(txtBarra.Text, nLongitudBarra + 2, 3))
                             Else
                                nCantidadBarra = Val(Mid(txtBarra.Text, nLongitudBarra + 1, 2) + "." + Mid(txtBarra.Text, nLongitudBarra + 3, 4))
                             End If
                             
                             InsertaProducto nCantidadBarra
                          Else
                             nCantidadBarra = Val(Mid(txtBarra.Text, nLongitudBarra + 1, 3) + "." + Mid(txtBarra.Text, nLongitudBarra + 4, 3))
                             InsertaProducto CalculaCantidad(nCantidadBarra)
                          End If
                          
                          If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                             lPropiedad = True
                          End If
                       Else
                          MsgBox "Producto no encontrado", vbCritical, sMensaje
                       End If
                    Else
                          MsgBox "Producto no encontrado", vbCritical, sMensaje
                    End If
                 End If
                 RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
              End If
              txtBarra.Text = ""
              sProducto = ""
        
        'End If
    Else
        'txtBarra.Text = ""
    End If
     
End Sub
Public Function CalculaCantidad(nPrecio As Double) As Double
    Select Case sTipoPedido
           Case "01"
                CalculaCantidad = nPrecio / RsProducto!nprecioVenta
           Case "02"
                If RsProducto!nPrecioDelivery > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPrecioDelivery
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
           
           Case "03"
                If RsProducto!nPreciollevar > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPreciollevar
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
                
           Case "04"
                If RsProducto!nPrecioCanal4 > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPrecioCanal4
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
                   
           Case "05"
                If RsProducto!nPrecioCanal5 > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPrecioCanal5
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
    End Select
      
End Function
