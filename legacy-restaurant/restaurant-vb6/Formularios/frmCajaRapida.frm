VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{B52C1CDE-38E9-11D5-98EC-00C0F01C6C81}#1.0#0"; "ifepson.ocx"
Begin VB.Form frmCajaRapida 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Caja Rápida"
   ClientHeight    =   11490
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   15450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11490
   ScaleMode       =   0  'User
   ScaleWidth      =   15452.96
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraProductoCombo 
      BackColor       =   &H00E2DEDB&
      BorderStyle     =   0  'None
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
      ForeColor       =   &H0073B43C&
      Height          =   3285
      Left            =   15480
      TabIndex        =   126
      Top             =   11760
      Width           =   2505
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame11"
         Height          =   1095
         Left            =   0
         TabIndex        =   407
         Top             =   10200
         Width           =   6255
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAE7E4&
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   0
         TabIndex        =   386
         Top             =   0
         Width           =   6255
         Begin VB.TextBox txtPorductoComboSeleccionado 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   388
            Top             =   600
            Width           =   5895
         End
         Begin VB.Label Label8 
            BackColor       =   &H00EAE7E4&
            Caption         =   "Producto seleccionado:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0073B43C&
            Height          =   255
            Left            =   120
            TabIndex        =   387
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.PictureBox cmdBuscarPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   2
         Left            =   5400
         Picture         =   "frmCajaRapida.frx":0000
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   384
         Top             =   8040
         Width           =   660
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   43
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   6975
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   44
         Left            =   1033
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   6975
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   45
         Left            =   1886
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   6975
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   46
         Left            =   2749
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   6975
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   47
         Left            =   3602
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   6975
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   48
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   6975
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   42
         Left            =   5308
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   6105
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   41
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   6105
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   40
         Left            =   3602
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   6105
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   39
         Left            =   2749
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   6105
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   38
         Left            =   1886
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   6105
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   37
         Left            =   1033
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   6105
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   36
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   6105
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   35
         Left            =   5308
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   5235
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   34
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   5235
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   33
         Left            =   3602
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   5235
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   32
         Left            =   2749
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   5235
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   31
         Left            =   1886
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   5235
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   30
         Left            =   1033
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   5235
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   29
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   5235
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   28
         Left            =   5308
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   4365
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   27
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   4365
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   26
         Left            =   3602
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   4365
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   25
         Left            =   2749
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   4365
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   24
         Left            =   1886
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   4365
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   23
         Left            =   1033
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   4365
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   22
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   4365
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   21
         Left            =   5308
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   3495
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   20
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   3495
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   19
         Left            =   3602
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   3495
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   18
         Left            =   2749
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   3495
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   17
         Left            =   1886
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   3495
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   16
         Left            =   1033
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   3495
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   15
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   3495
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   14
         Left            =   5308
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   2610
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   13
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   2610
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   12
         Left            =   3602
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   2610
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   11
         Left            =   2749
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   2610
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   10
         Left            =   1886
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   2610
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   9
         Left            =   1033
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   2610
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   8
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   2610
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   7
         Left            =   5308
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   1740
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   6
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   1740
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   5
         Left            =   3602
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   1740
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   4
         Left            =   2749
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   1740
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   3
         Left            =   1886
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   1740
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   2
         Left            =   1033
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   1740
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   1
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   1740
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   49
         Left            =   5308
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   6975
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   50
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   7860
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   51
         Left            =   1033
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   7860
         Width           =   823
      End
      Begin VB.CommandButton cmdProductoCombo 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   823
         Index           =   52
         Left            =   1886
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   7860
         Width           =   823
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E2DEDB&
         Caption         =   "Productos del Combo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0073B43C&
         Height          =   255
         Left            =   240
         TabIndex        =   385
         Top             =   1320
         Width           =   2535
      End
   End
   Begin VB.Frame FrameFeSpring 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   1200
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   6315
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmCajaRapida.frx":0B39
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label lblPaso2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Obteniendo codigo XXXX almacenado."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1185
         TabIndex        =   4
         Top             =   1155
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.Label lblPaso1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Enviando información de documento a XXXX."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1170
         TabIndex        =   3
         Top             =   870
         Visible         =   0   'False
         Width           =   3660
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Facturación Electronica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   2490
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmCajaRapida.frx":0D4C
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmCajaRapida.frx":0F5F
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmCajaRapida.frx":12A1
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "   Proceso de envio de documento a XXXXX......."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1110
         Left            =   210
         TabIndex        =   5
         Top             =   435
         Width           =   5910
      End
   End
   Begin VB.Frame fraCombo 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
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
      Height          =   1695
      Left            =   17160
      TabIndex        =   119
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.PictureBox cmdComboPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   4
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":15E3
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   408
         Top             =   1845
         Width           =   600
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         Height          =   1095
         Left            =   0
         TabIndex        =   404
         Top             =   8880
         Width           =   9855
         Begin VB.PictureBox cmdComboPicture 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   735
            Index           =   0
            Left            =   5760
            Picture         =   "frmCajaRapida.frx":224D
            ScaleHeight     =   735
            ScaleWidth      =   3705
            TabIndex        =   406
            Top             =   240
            Width           =   3705
         End
         Begin VB.PictureBox cmdComboPicture 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   735
            Index           =   99
            Left            =   120
            Picture         =   "frmCajaRapida.frx":4297
            ScaleHeight     =   735
            ScaleWidth      =   1815
            TabIndex        =   405
            Top             =   240
            Width           =   1820
         End
      End
      Begin VB.PictureBox cmdComboPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   5
         Left            =   120
         Picture         =   "frmCajaRapida.frx":57E7
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   396
         Top             =   5160
         Width           =   2200
      End
      Begin VB.PictureBox cmdComboPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   8
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":7779
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   395
         Top             =   4920
         Width           =   600
      End
      Begin VB.PictureBox cmdNavegarPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   16
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":803F
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   394
         Top             =   4305
         Width           =   600
      End
      Begin VB.PictureBox cmdNavegarPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   14
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":8BFC
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   393
         Top             =   3690
         Width           =   600
      End
      Begin VB.PictureBox cmdNavegarPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   13
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":954C
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   392
         Top             =   3075
         Width           =   600
      End
      Begin VB.PictureBox cmdNavegarPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   17
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":9E9F
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   391
         Top             =   2460
         Width           =   600
      End
      Begin VB.PictureBox cmdComboPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   3
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":AA56
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   390
         Top             =   1215
         Width           =   600
      End
      Begin VB.PictureBox cmdComboPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   2
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":B0AE
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   389
         Top             =   600
         Width           =   600
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   600
         Index           =   12
         Left            =   7800
         Picture         =   "frmCajaRapida.frx":B96D
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   6960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   600
         Index           =   15
         Left            =   7680
         Picture         =   "frmCajaRapida.frx":C237
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   7755
         Visible         =   0   'False
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
         Height          =   630
         Index           =   1
         Left            =   5160
         Picture         =   "frmCajaRapida.frx":CB01
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   6960
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
         BackColor       =   &H00F2FAED&
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
         Height          =   630
         Index           =   6
         Left            =   6870
         Picture         =   "frmCajaRapida.frx":CC03
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   5160
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdCombo 
         BackColor       =   &H00F2FAED&
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
         Height          =   630
         Index           =   7
         Left            =   7650
         Picture         =   "frmCajaRapida.frx":CD05
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   5160
         Visible         =   0   'False
         Width           =   780
      End
      Begin TrueOleDBGrid80.TDBGrid grdCombo 
         Height          =   5055
         Left            =   120
         TabIndex        =   125
         Top             =   0
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   8916
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
         Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(6)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(7)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   7
         Caption         =   "Detalle del combo"
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
         RowDividerColor =   15263202
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.locked=0,.bold=0"
         _StyleDefs(7)   =   ":id=1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H4D7C2A&,.fgcolor=&HFFFFFF&"
         _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Arial"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H70655D&,.fgcolor=&H80000005&"
         _StyleDefs(13)  =   ":id=2,.borderSize=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(14)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=2,.fontname=Arial"
         _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HF2FAED&,.fgcolor=&H232323&"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(41)  =   "Named:id=33:Normal"
         _StyleDefs(42)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
         _StyleDefs(43)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
         _StyleDefs(44)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(45)  =   ":id=33,.fontname=Small Fonts"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(48)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(49)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(50)  =   ":id=34,.fontname=Arial"
         _StyleDefs(51)  =   "Named:id=35:Footing"
         _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   "Named:id=36:Selected"
         _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(55)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(56)  =   ":id=36,.fontname=Arial"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(61)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(62)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame fraPuntoVenta 
      BackColor       =   &H00E2DEDB&
      Caption         =   " Punto de Venta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E28448&
      Height          =   975
      Left            =   8640
      TabIndex        =   109
      Top             =   13200
      Width           =   2265
      Begin VB.PictureBox cmdBuscarPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   1
         Left            =   5400
         Picture         =   "frmCajaRapida.frx":CE07
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   383
         Top             =   1560
         Width           =   660
      End
      Begin VB.CommandButton cmdPunto 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   300
         Width           =   1180
      End
      Begin VB.CommandButton cmdPunto 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   2
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   300
         Width           =   1180
      End
      Begin VB.CommandButton cmdPunto 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   3
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   300
         Width           =   1180
      End
      Begin VB.CommandButton cmdPunto 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   4
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   300
         Width           =   1180
      End
      Begin VB.CommandButton cmdPunto 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   5
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   300
         Width           =   1180
      End
      Begin VB.CommandButton cmdPunto 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   1290
         Width           =   1180
      End
      Begin VB.CommandButton cmdPunto 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   7
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   1290
         Width           =   1180
      End
      Begin VB.CommandButton cmdPunto 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   8
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   1290
         Width           =   1180
      End
      Begin VB.CommandButton cmdPunto 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   9
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   1290
         Width           =   1180
      End
   End
   Begin VB.Frame fraOrigenVentas 
      BackColor       =   &H00E2DEDB&
      Caption         =   "Origen de Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E28448&
      Height          =   1215
      Left            =   8640
      TabIndex        =   6
      Top             =   11760
      Width           =   1875
      Begin VB.PictureBox cmdBuscarPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   3
         Left            =   5280
         Picture         =   "frmCajaRapida.frx":D940
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   382
         Top             =   3000
         Width           =   660
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   17
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3000
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   19
         Left            =   3080
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3000
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   10
         Left            =   1600
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1635
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   18
         Left            =   1600
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3000
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   16
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2310
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   15
         Left            =   3080
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2310
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   14
         Left            =   1600
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2310
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   13
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2310
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   12
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1635
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   11
         Left            =   3080
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1635
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   8
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   7
         Left            =   3080
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   6
         Left            =   1600
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   9
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1635
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   5
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   3080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   300
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   1600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   1450
      End
      Begin VB.CommandButton cmdOrigen 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   300
         Width           =   1450
      End
   End
   Begin VB.Frame fraMozo 
      BackColor       =   &H00E2DEDB&
      Caption         =   "Mesero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E28448&
      Height          =   1095
      Left            =   11400
      TabIndex        =   89
      Top             =   11760
      Width           =   2265
      Begin VB.PictureBox cmdBuscarPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   0
         Left            =   5280
         Picture         =   "frmCajaRapida.frx":E479
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   381
         Top             =   2880
         Width           =   660
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   19
         Left            =   3680
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   2760
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   18
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   2760
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   17
         Left            =   1300
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   2760
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   16
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   2760
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   15
         Left            =   4870
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   1900
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   14
         Left            =   3680
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   1900
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   13
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   1900
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   12
         Left            =   1300
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   1900
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   11
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   1900
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   10
         Left            =   4870
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   1050
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   9
         Left            =   3680
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1050
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   8
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   1050
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   7
         Left            =   1300
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   1050
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   1050
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   5
         Left            =   4870
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   200
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   4
         Left            =   3680
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   200
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   3
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   200
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   2
         Left            =   1300
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   200
         Width           =   1150
      End
      Begin VB.CommandButton cmdMozo 
         BackColor       =   &H00F7F6F5&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   200
         Width           =   1150
      End
   End
   Begin VB.Frame fraMorotizado 
      BackColor       =   &H00E2DEDB&
      Caption         =   "Motorizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E28448&
      Height          =   1215
      Left            =   5880
      TabIndex        =   26
      Top             =   12720
      Width           =   2505
      Begin VB.PictureBox cmdBuscarPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   4
         Left            =   5280
         Picture         =   "frmCajaRapida.frx":EFB2
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   380
         Top             =   2880
         Width           =   660
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   17
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2920
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   19
         Left            =   3080
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2920
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   10
         Left            =   1600
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1620
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   18
         Left            =   1600
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2920
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   16
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2270
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   15
         Left            =   3080
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2270
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   14
         Left            =   1600
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2270
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   13
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2270
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   12
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1620
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   11
         Left            =   3080
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1620
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   8
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   960
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   7
         Left            =   3080
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   960
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   6
         Left            =   1600
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   960
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   9
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1620
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   5
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   960
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   300
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   3080
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   300
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   1600
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   300
         Width           =   1450
      End
      Begin VB.CommandButton cmdMotorizado 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   300
         Width           =   1450
      End
   End
   Begin VB.Frame fraPropiedad 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
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
      Height          =   3615
      Left            =   15720
      TabIndex        =   262
      Top             =   5040
      Width           =   5385
      Begin VB.PictureBox cmdBusca 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   5310
         Picture         =   "frmCajaRapida.frx":FAEB
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   378
         Top             =   2910
         Width           =   660
      End
      Begin VB.TextBox txtProductoSeleccionado 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   377
         Top             =   600
         Width           =   5895
      End
      Begin VB.PictureBox cmdOpcionPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   99
         Left            =   3120
         Picture         =   "frmCajaRapida.frx":10624
         ScaleHeight     =   660
         ScaleWidth      =   2880
         TabIndex        =   375
         Top             =   9480
         Width           =   2880
      End
      Begin VB.PictureBox cmdOpcionPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   6
         Left            =   120
         Picture         =   "frmCajaRapida.frx":12089
         ScaleHeight     =   660
         ScaleWidth      =   2880
         TabIndex        =   374
         Top             =   9480
         Width           =   2880
      End
      Begin VB.PictureBox cmdOpcionPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   7
         Left            =   3120
         Picture         =   "frmCajaRapida.frx":1364A
         ScaleHeight     =   495
         ScaleWidth      =   2895
         TabIndex        =   373
         Top             =   7200
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2DEDB&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5295
         Left            =   0
         TabIndex        =   337
         Top             =   1200
         Width           =   6225
         Begin VB.PictureBox cmdBuscarPicture 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   660
            Index           =   8
            Left            =   5280
            Picture         =   "frmCajaRapida.frx":153D6
            ScaleHeight     =   660
            ScaleWidth      =   660
            TabIndex        =   379
            Top             =   4440
            Width           =   660
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   19
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   371
            Top             =   4320
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   18
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   370
            Top             =   4320
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   16
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   369
            Top             =   4320
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   17
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   368
            Top             =   4320
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   15
            Left            =   105
            Style           =   1  'Graphical
            TabIndex        =   367
            Top             =   4320
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   13
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   366
            Top             =   3480
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   14
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   365
            Top             =   3480
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   12
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   364
            Top             =   3480
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   10
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   363
            Top             =   3480
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   11
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   362
            Top             =   3480
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   9
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   361
            Top             =   3480
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   7
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   360
            Top             =   2640
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   8
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   359
            Top             =   3480
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   6
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   358
            Top             =   2640
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   357
            Top             =   2640
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   2
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   356
            Top             =   2640
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   3
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   355
            Top             =   2640
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   4
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   354
            Top             =   2640
            Width           =   823
         End
         Begin VB.CommandButton cmdPropiedad 
            BackColor       =   &H00F2FAED&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   5
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   353
            Top             =   2640
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   7
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   350
            Top             =   1320
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   6
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   349
            Top             =   1320
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   13
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   348
            Top             =   1320
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   5
            Left            =   975
            Style           =   1  'Graphical
            TabIndex        =   347
            Top             =   1320
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   12
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   346
            Top             =   1320
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   4
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   345
            Top             =   500
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   11
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   344
            Top             =   1320
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   3
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   343
            Top             =   500
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   10
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   342
            Top             =   480
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H008BDD79&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   1
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   341
            Top             =   500
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   8
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   340
            Top             =   500
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   2
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   339
            Top             =   500
            Width           =   823
         End
         Begin VB.CommandButton cmdOperador 
            BackColor       =   &H00D1F2CA&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   823
            Index           =   9
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   338
            Top             =   500
            Width           =   823
         End
         Begin VB.Label lblPropiedad 
            AutoSize        =   -1  'True
            BackColor       =   &H00E2DEDB&
            Caption         =   "1. Operador   "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0073B43C&
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   352
            Top             =   195
            Width           =   1305
         End
         Begin VB.Label lblPropiedad 
            AutoSize        =   -1  'True
            BackColor       =   &H00E2DEDB&
            Caption         =   "2. Propiedad "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0073B43C&
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   351
            Top             =   2280
            Width           =   1275
         End
      End
      Begin VB.TextBox lblObservacion 
         Appearance      =   0  'Flat
         Height          =   1515
         Left            =   3120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   264
         Top             =   7800
         Width           =   2900
      End
      Begin VB.TextBox lblResumen 
         Appearance      =   0  'Flat
         Height          =   2100
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   263
         Top             =   7200
         Width           =   2880
      End
      Begin VB.Label lblPropiedad 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         Caption         =   "Producto Seleccionado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0073B43C&
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   376
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblPropiedad 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         Caption         =   "Detalle de Propiedades"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0073B43C&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   372
         Top             =   6840
         Width           =   2235
      End
   End
   Begin VB.Frame tabProducto 
      BackColor       =   &H00E2DEDB&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   9000
      Left            =   9240
      TabIndex        =   281
      Top             =   1440
      Width           =   6225
      Begin VB.CommandButton cmdEtiqueta 
         BackColor       =   &H00F2FAED&
         Caption         =   "Command1"
         Height          =   645
         Index           =   3
         Left            =   4125
         Style           =   1  'Graphical
         TabIndex        =   336
         Top             =   240
         Width           =   1952
      End
      Begin VB.CommandButton cmdEtiqueta 
         BackColor       =   &H00F2FAED&
         Caption         =   "Command1"
         Height          =   645
         Index           =   2
         Left            =   2130
         Style           =   1  'Graphical
         TabIndex        =   335
         Top             =   240
         Width           =   1952
      End
      Begin VB.CommandButton cmdEtiqueta 
         BackColor       =   &H008BDD79&
         Caption         =   "Command1"
         Height          =   645
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   334
         Top             =   240
         Width           =   1952
      End
      Begin VB.CommandButton cmdAgrupacion 
         BackColor       =   &H00E7ECD5&
         Height          =   671
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   333
         Top             =   1080
         Width           =   1952
      End
      Begin VB.CommandButton cmdAgrupacion 
         BackColor       =   &H00F7F6F5&
         Caption         =   "Command1"
         Height          =   671
         Index           =   2
         Left            =   2130
         Style           =   1  'Graphical
         TabIndex        =   332
         Top             =   1080
         Width           =   1952
      End
      Begin VB.CommandButton cmdAgrupacion 
         BackColor       =   &H00F7F6F5&
         Caption         =   "Command1"
         Height          =   671
         Index           =   3
         Left            =   4125
         Style           =   1  'Graphical
         TabIndex        =   331
         Top             =   1080
         Width           =   1952
      End
      Begin VB.CommandButton cmdAgrupacion 
         BackColor       =   &H00F7F6F5&
         Caption         =   "Command1"
         Height          =   671
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   330
         Top             =   1845
         Width           =   1952
      End
      Begin VB.CommandButton cmdAgrupacion 
         BackColor       =   &H00F7F6F5&
         Caption         =   "Command1"
         Height          =   671
         Index           =   5
         Left            =   2130
         Style           =   1  'Graphical
         TabIndex        =   329
         Top             =   1845
         Width           =   1952
      End
      Begin VB.CommandButton cmdAgrupacion 
         BackColor       =   &H00F7F6F5&
         Caption         =   "Command1"
         Height          =   671
         Index           =   6
         Left            =   4125
         Style           =   1  'Graphical
         TabIndex        =   328
         Top             =   1845
         Width           =   1952
      End
      Begin VB.CommandButton cmdAgrupacion 
         BackColor       =   &H00F7F6F5&
         Caption         =   "Command1"
         Height          =   671
         Index           =   7
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   327
         Top             =   2625
         Width           =   1952
      End
      Begin VB.CommandButton cmdAgrupacion 
         BackColor       =   &H00F7F6F5&
         Caption         =   "Command1"
         Height          =   671
         Index           =   8
         Left            =   2130
         Style           =   1  'Graphical
         TabIndex        =   326
         Top             =   2625
         Width           =   1952
      End
      Begin VB.CommandButton cmdAgrupacion 
         BackColor       =   &H00F7F6F5&
         Caption         =   "Command1"
         Height          =   671
         Index           =   9
         Left            =   4125
         Style           =   1  'Graphical
         TabIndex        =   325
         Top             =   2625
         Width           =   1952
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2DEDB&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   0
         TabIndex        =   282
         Top             =   3360
         Width           =   6255
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   324
            TabStop         =   0   'False
            Top             =   240
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   2
            Left            =   970
            Style           =   1  'Graphical
            TabIndex        =   323
            TabStop         =   0   'False
            Top             =   240
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   3
            Left            =   1830
            Style           =   1  'Graphical
            TabIndex        =   322
            TabStop         =   0   'False
            Top             =   240
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   4
            Left            =   2680
            Style           =   1  'Graphical
            TabIndex        =   321
            TabStop         =   0   'False
            Top             =   240
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   5
            Left            =   3540
            Style           =   1  'Graphical
            TabIndex        =   320
            TabStop         =   0   'False
            Top             =   240
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   6
            Left            =   4400
            Style           =   1  'Graphical
            TabIndex        =   319
            TabStop         =   0   'False
            Top             =   240
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   7
            Left            =   5250
            Style           =   1  'Graphical
            TabIndex        =   318
            TabStop         =   0   'False
            Top             =   240
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   8
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   317
            TabStop         =   0   'False
            Top             =   1100
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   9
            Left            =   970
            Style           =   1  'Graphical
            TabIndex        =   316
            TabStop         =   0   'False
            Top             =   1100
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   10
            Left            =   1830
            Style           =   1  'Graphical
            TabIndex        =   315
            TabStop         =   0   'False
            Top             =   1100
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   11
            Left            =   2680
            Style           =   1  'Graphical
            TabIndex        =   314
            TabStop         =   0   'False
            Top             =   1100
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "12"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   12
            Left            =   3540
            Style           =   1  'Graphical
            TabIndex        =   313
            TabStop         =   0   'False
            Top             =   1100
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "13"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   13
            Left            =   4400
            Style           =   1  'Graphical
            TabIndex        =   312
            TabStop         =   0   'False
            Top             =   1100
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "14"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   14
            Left            =   5250
            Style           =   1  'Graphical
            TabIndex        =   311
            TabStop         =   0   'False
            Top             =   1100
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   15
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   310
            TabStop         =   0   'False
            Top             =   1950
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   16
            Left            =   970
            Style           =   1  'Graphical
            TabIndex        =   309
            TabStop         =   0   'False
            Top             =   1950
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "17"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   17
            Left            =   1830
            Style           =   1  'Graphical
            TabIndex        =   308
            TabStop         =   0   'False
            Top             =   1950
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "18"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   18
            Left            =   2680
            Style           =   1  'Graphical
            TabIndex        =   307
            TabStop         =   0   'False
            Top             =   1950
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "19"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   19
            Left            =   3540
            Style           =   1  'Graphical
            TabIndex        =   306
            TabStop         =   0   'False
            Top             =   1950
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   20
            Left            =   4400
            Style           =   1  'Graphical
            TabIndex        =   305
            TabStop         =   0   'False
            Top             =   1950
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "21"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   21
            Left            =   5250
            Style           =   1  'Graphical
            TabIndex        =   304
            TabStop         =   0   'False
            Top             =   1950
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "22"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   22
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   303
            TabStop         =   0   'False
            Top             =   2800
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "23"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   23
            Left            =   970
            Style           =   1  'Graphical
            TabIndex        =   302
            TabStop         =   0   'False
            Top             =   2800
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "24"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   24
            Left            =   1830
            Style           =   1  'Graphical
            TabIndex        =   301
            TabStop         =   0   'False
            Top             =   2800
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "25"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   25
            Left            =   2680
            Style           =   1  'Graphical
            TabIndex        =   300
            TabStop         =   0   'False
            Top             =   2800
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "26"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   26
            Left            =   3540
            Style           =   1  'Graphical
            TabIndex        =   299
            TabStop         =   0   'False
            Top             =   2800
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "27"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   27
            Left            =   4400
            Style           =   1  'Graphical
            TabIndex        =   298
            TabStop         =   0   'False
            Top             =   2800
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "28"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   28
            Left            =   5250
            Style           =   1  'Graphical
            TabIndex        =   297
            TabStop         =   0   'False
            Top             =   2800
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "29"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   29
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   296
            TabStop         =   0   'False
            Top             =   3660
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   30
            Left            =   970
            Style           =   1  'Graphical
            TabIndex        =   295
            TabStop         =   0   'False
            Top             =   3660
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   31
            Left            =   1830
            Style           =   1  'Graphical
            TabIndex        =   294
            TabStop         =   0   'False
            Top             =   3660
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "32"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   32
            Left            =   2680
            Style           =   1  'Graphical
            TabIndex        =   293
            TabStop         =   0   'False
            Top             =   3660
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "33"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   33
            Left            =   3540
            Style           =   1  'Graphical
            TabIndex        =   292
            TabStop         =   0   'False
            Top             =   3660
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "34"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   34
            Left            =   4400
            Style           =   1  'Graphical
            TabIndex        =   291
            TabStop         =   0   'False
            Top             =   3660
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "35"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   35
            Left            =   5250
            Style           =   1  'Graphical
            TabIndex        =   290
            TabStop         =   0   'False
            Top             =   3660
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "36"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   36
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   289
            TabStop         =   0   'False
            Top             =   4520
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   37
            Left            =   970
            Style           =   1  'Graphical
            TabIndex        =   288
            TabStop         =   0   'False
            Top             =   4520
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "38"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   38
            Left            =   1830
            Style           =   1  'Graphical
            TabIndex        =   287
            TabStop         =   0   'False
            Top             =   4520
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "39"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   39
            Left            =   2680
            Style           =   1  'Graphical
            TabIndex        =   286
            TabStop         =   0   'False
            Top             =   4520
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "40"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   40
            Left            =   3540
            Style           =   1  'Graphical
            TabIndex        =   285
            TabStop         =   0   'False
            Top             =   4520
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "41"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   41
            Left            =   4400
            Style           =   1  'Graphical
            TabIndex        =   284
            TabStop         =   0   'False
            Top             =   4520
            Width           =   824
         End
         Begin VB.CommandButton cmdBoton 
            BackColor       =   &H00E5E4DF&
            Caption         =   "42"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   824
            Index           =   42
            Left            =   5250
            Style           =   1  'Graphical
            TabIndex        =   283
            TabStop         =   0   'False
            Top             =   4520
            Width           =   824
         End
      End
   End
   Begin VB.Frame fraDetalle 
      Caption         =   " Detalle del Plato "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2850
      Left            =   20160
      TabIndex        =   219
      Top             =   12360
      Width           =   5100
      Begin VB.Frame fraDescuento 
         Caption         =   " Descuentos / Recargos "
         ForeColor       =   &H00800080&
         Height          =   1605
         Left            =   1665
         TabIndex        =   257
         Top             =   180
         Width           =   3090
         Begin VB.CommandButton cmdDescuento 
            Caption         =   "Monto del Descuento"
            Height          =   555
            Index           =   0
            Left            =   195
            TabIndex        =   261
            Top             =   270
            Width           =   1245
         End
         Begin VB.CommandButton cmdDescuento 
            Caption         =   "Monto del Recargo"
            Height          =   555
            Index           =   2
            Left            =   195
            TabIndex        =   260
            Top             =   915
            Width           =   1245
         End
         Begin VB.CommandButton cmdDescuento 
            Caption         =   "( % ) del Descuento"
            Height          =   555
            Index           =   1
            Left            =   1650
            TabIndex        =   259
            Top             =   270
            Width           =   1245
         End
         Begin VB.CommandButton cmdDescuento 
            Caption         =   "( % ) del Recargo"
            Height          =   555
            Index           =   3
            Left            =   1650
            TabIndex        =   258
            Top             =   915
            Width           =   1245
         End
      End
      Begin VB.Frame fraPrecio 
         Caption         =   " Precio de Venta "
         ForeColor       =   &H00800080&
         Height          =   3840
         Left            =   120
         TabIndex        =   228
         Top             =   2880
         Width           =   4890
         Begin VB.Label txtObserva 
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
            Height          =   510
            Left            =   1065
            TabIndex        =   256
            Top             =   3165
            Width           =   2895
         End
         Begin VB.Label txtCortesia 
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
            Height          =   540
            Left            =   2475
            TabIndex        =   255
            Top             =   2550
            Width           =   1485
         End
         Begin VB.Label txtVenta 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   254
            Top             =   2850
            Width           =   1365
         End
         Begin VB.Label txtOficial 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   253
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label txtPVenta 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   252
            Top             =   2310
            Width           =   1365
         End
         Begin VB.Label txtNeto 
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
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   1065
            TabIndex        =   251
            Top             =   1140
            Width           =   1365
         End
         Begin VB.Label txtDImporte 
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
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1785
            TabIndex        =   250
            Top             =   540
            Width           =   645
         End
         Begin VB.Label txtRImporte 
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
            Left            =   1785
            TabIndex        =   249
            Top             =   810
            Width           =   645
         End
         Begin VB.Label txtRPorcentaje 
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
            Left            =   1065
            TabIndex        =   248
            Top             =   810
            Width           =   510
         End
         Begin VB.Label txtDPorcentaje 
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
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1065
            TabIndex        =   247
            Top             =   540
            Width           =   510
         End
         Begin VB.Label txtImpuesto3 
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
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   1065
            TabIndex        =   246
            Top             =   1965
            Width           =   1365
         End
         Begin VB.Label txtImpuesto2 
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
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   1065
            TabIndex        =   245
            Top             =   1695
            Width           =   1365
         End
         Begin VB.Label txtImpuesto1 
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
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   1065
            TabIndex        =   244
            Top             =   1410
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cortesía"
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
            Index           =   21
            Left            =   2490
            TabIndex        =   243
            Top             =   2325
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   20
            Left            =   1590
            TabIndex        =   242
            Top             =   840
            Width           =   150
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   19
            Left            =   1590
            TabIndex        =   241
            Top             =   570
            Width           =   150
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Neto :"
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
            TabIndex        =   240
            Top             =   1185
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Observación :"
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
            Index           =   22
            Left            =   180
            TabIndex        =   239
            Top             =   3165
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total :"
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
            Index           =   18
            Left            =   600
            TabIndex        =   238
            Top             =   2925
            Width           =   390
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 3 :"
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
            Index           =   12
            Left            =   240
            TabIndex        =   237
            Top             =   2010
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 1 :"
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
            Left            =   240
            TabIndex        =   236
            Top             =   1455
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 2 :"
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
            Index           =   11
            Left            =   240
            TabIndex        =   235
            Top             =   1725
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Oficial :"
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
            Index           =   15
            Left            =   90
            TabIndex        =   234
            Top             =   300
            Width           =   900
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
            Index           =   13
            Left            =   240
            TabIndex        =   233
            Top             =   585
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Venta :"
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
            Index           =   16
            Left            =   150
            TabIndex        =   232
            Top             =   2340
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Recargo :"
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
            Index           =   14
            Left            =   390
            TabIndex        =   231
            Top             =   855
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cantidad :"
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
            Index           =   17
            Left            =   375
            TabIndex        =   230
            Top             =   2625
            Width           =   615
         End
         Begin VB.Label txtCantidad 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   229
            Top             =   2580
            Width           =   1365
         End
      End
      Begin VB.Frame fraImpuesto 
         Caption         =   " Impuestos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   1065
         Left            =   120
         TabIndex        =   224
         Top             =   1785
         Width           =   4905
         Begin VB.CommandButton cmdImpuesto 
            Height          =   630
            Index           =   0
            Left            =   150
            TabIndex        =   227
            Top             =   270
            Width           =   1245
         End
         Begin VB.CommandButton cmdImpuesto 
            Height          =   630
            Index           =   1
            Left            =   1725
            TabIndex        =   226
            Top             =   270
            Width           =   1245
         End
         Begin VB.CommandButton cmdImpuesto 
            Height          =   630
            Index           =   2
            Left            =   3315
            TabIndex        =   225
            Top             =   270
            Width           =   1245
         End
      End
      Begin VB.CommandButton cmdPrecio 
         Caption         =   "Precio"
         Height          =   585
         Left            =   285
         TabIndex        =   223
         Top             =   300
         Width           =   1200
      End
      Begin VB.CommandButton cmdCortesia 
         Caption         =   "Cortesía"
         Height          =   585
         Left            =   285
         TabIndex        =   222
         Top             =   990
         Width           =   1200
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   3
         Left            =   3585
         Picture         =   "frmCajaRapida.frx":15F0F
         Style           =   1  'Graphical
         TabIndex        =   221
         Top             =   6870
         Width           =   1410
      End
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
         Height          =   630
         Index           =   4
         Left            =   1995
         Picture         =   "frmCajaRapida.frx":16011
         Style           =   1  'Graphical
         TabIndex        =   220
         Top             =   6870
         Width           =   1410
      End
   End
   Begin VB.Frame fraEliminacion 
      BackColor       =   &H00E2DEDB&
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
      ForeColor       =   &H005A65FF&
      Height          =   1815
      Left            =   15720
      TabIndex        =   179
      Top             =   3120
      Width           =   6825
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   38
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   218
         Top             =   7710
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   37
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   217
         Top             =   7710
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   36
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   216
         Top             =   7710
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   35
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   6650
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   31
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   6650
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   32
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   6650
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   33
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   6650
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   34
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   6650
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   30
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   5600
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   26
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   209
         Top             =   5600
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   27
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   208
         Top             =   5600
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   28
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   207
         Top             =   5600
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   29
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   206
         Top             =   5600
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   25
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   205
         Top             =   4530
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   21
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   204
         Top             =   4530
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   22
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   4530
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   23
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   4530
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   24
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   201
         Top             =   4530
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   20
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   3500
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   16
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   3500
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   17
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   198
         Top             =   3500
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   18
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   197
         Top             =   3500
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   19
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   196
         Top             =   3500
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   15
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   2450
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   11
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   2450
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   12
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   2450
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   13
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   192
         Top             =   2450
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   14
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   2450
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   10
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   190
         Top             =   1380
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   1
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   330
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   2
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   330
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   3
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   330
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   4
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   186
         Top             =   330
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   5
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   330
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   6
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   1380
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   7
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   183
         Top             =   1380
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   8
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   182
         Top             =   1380
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminacion 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   9
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   181
         Top             =   1380
         Width           =   1100
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00FFFFFF&
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
         Height          =   1000
         Index           =   17
         Left            =   3780
         Picture         =   "frmCajaRapida.frx":16113
         Style           =   1  'Graphical
         TabIndex        =   180
         Top             =   7680
         Width           =   2310
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "OTROS NO BORRAR"
      Height          =   2295
      Left            =   0
      TabIndex        =   47
      Top             =   12120
      Width           =   5655
      Begin VB.PictureBox ImageBarra 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3960
         ScaleHeight     =   465
         ScaleWidth      =   1065
         TabIndex        =   88
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin EPSON_Impresora_Fiscal.PrinterFiscal PrinterFiscalArgentina 
         Left            =   840
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin VB.Image imageHash 
         Height          =   735
         Left            =   2400
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Image ImageIntegracion 
         Height          =   615
         Left            =   1560
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image ImageMercadoPago 
         Height          =   735
         Left            =   3360
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label txtTipoDocumento 
         Caption         =   "TipoDocumento"
         Height          =   240
         Left            =   840
         TabIndex        =   48
         Top             =   240
         Width           =   1005
      End
      Begin VB.Image imageCab 
         Height          =   375
         Left            =   120
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imagepIE 
         Height          =   375
         Left            =   120
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FRMA ANTERIOR"
      Height          =   495
      Left            =   480
      TabIndex        =   46
      Top             =   11760
      Width           =   1575
   End
   Begin VB.Frame FrameCanalVenta 
      BackColor       =   &H00EAE7E4&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EC642&
      Height          =   1215
      Left            =   9240
      TabIndex        =   274
      Top             =   240
      Width           =   6225
      Begin VB.CommandButton cmdCabecera 
         BackColor       =   &H00F2FAED&
         Caption         =   "En el &Local"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   671
         Index           =   5
         Left            =   4920
         MaskColor       =   &H00F2FAED&
         Style           =   1  'Graphical
         TabIndex        =   279
         Top             =   480
         Width           =   1161
      End
      Begin VB.CommandButton cmdCabecera 
         BackColor       =   &H00F2FAED&
         Caption         =   "En el &Local"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   671
         Index           =   4
         Left            =   3720
         MaskColor       =   &H00F2FAED&
         Style           =   1  'Graphical
         TabIndex        =   278
         Top             =   480
         Width           =   1161
      End
      Begin VB.CommandButton cmdCabecera 
         BackColor       =   &H00F2FAED&
         Caption         =   "En el &Local"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   671
         Index           =   3
         Left            =   2520
         MaskColor       =   &H00F2FAED&
         Style           =   1  'Graphical
         TabIndex        =   277
         Top             =   480
         Width           =   1161
      End
      Begin VB.CommandButton cmdCabecera 
         BackColor       =   &H007EC642&
         Caption         =   "&Para Llevar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   671
         Index           =   1
         Left            =   120
         MaskColor       =   &H00F2FAED&
         Style           =   1  'Graphical
         TabIndex        =   276
         Top             =   480
         Width           =   1161
      End
      Begin VB.CommandButton cmdCabecera 
         BackColor       =   &H00F2FAED&
         Caption         =   "En el &Local"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   671
         Index           =   2
         Left            =   1320
         MaskColor       =   &H00F2FAED&
         Style           =   1  'Graphical
         TabIndex        =   275
         Top             =   480
         Width           =   1161
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Insumos Críticos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   1800
         TabIndex        =   412
         Top             =   0
         Width           =   4395
      End
      Begin VB.Label Label6 
         BackColor       =   &H00EAE7E4&
         Caption         =   "Canal de Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0073B43C&
         Height          =   255
         Left            =   120
         TabIndex        =   280
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   6975
      Left            =   0
      TabIndex        =   49
      Top             =   240
      Width           =   9255
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   7
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":1669D
         ScaleHeight     =   615
         ScaleWidth      =   645
         TabIndex        =   410
         Top             =   1320
         Width           =   645
      End
      Begin VB.TextBox txtcantproductos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   400
         Text            =   "Productos"
         Top             =   6480
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007EC642&
         Height          =   375
         Left            =   4200
         TabIndex        =   399
         Text            =   "Descuento"
         Top             =   6480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007EC642&
         Height          =   375
         Left            =   6480
         TabIndex        =   398
         Text            =   "Total"
         Top             =   6480
         Width           =   615
      End
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   1
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":17279
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   397
         Top             =   3160
         Width           =   600
      End
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Index           =   14
         Left            =   4080
         Picture         =   "frmCajaRapida.frx":17EE3
         ScaleHeight     =   495
         ScaleWidth      =   1845
         TabIndex        =   62
         Top             =   770
         Width           =   1845
      End
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Index           =   6
         Left            =   4080
         Picture         =   "frmCajaRapida.frx":19950
         ScaleHeight     =   495
         ScaleWidth      =   1845
         TabIndex        =   61
         Top             =   230
         Width           =   1845
      End
      Begin VB.PictureBox cmdOpcionPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1035
         Index           =   9
         Left            =   120
         Picture         =   "frmCajaRapida.frx":1B26C
         ScaleHeight     =   1035
         ScaleWidth      =   885
         TabIndex        =   60
         Top             =   240
         Width           =   885
      End
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   2
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":1C63D
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   58
         Top             =   1930
         Width           =   600
      End
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   3
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":1CEFC
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   57
         Top             =   2550
         Width           =   600
      End
      Begin VB.PictureBox cmdNavegarPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   2
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":1D554
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   56
         Top             =   4370
         Width           =   600
      End
      Begin VB.PictureBox cmdNavegarPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   3
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":1DF29
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   55
         Top             =   4995
         Width           =   600
      End
      Begin VB.PictureBox cmdNavegarPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   4
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":1E671
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   54
         Top             =   5630
         Width           =   600
      End
      Begin VB.PictureBox cmdNavegarPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   5
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":1EDB7
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   53
         Top             =   6250
         Width           =   600
      End
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   8
         Left            =   8520
         Picture         =   "frmCajaRapida.frx":1F78B
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   52
         Top             =   3720
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   580
         Index           =   1
         Left            =   8640
         Picture         =   "frmCajaRapida.frx":20051
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   580
         Index           =   6
         Left            =   8640
         Picture         =   "frmCajaRapida.frx":2091B
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle 
         Height          =   5055
         Left            =   120
         TabIndex        =   63
         Top             =   1320
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   8916
         _LayoutType     =   4
         _RowHeight      =   26
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   0
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerColor=   16777215
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
         Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   7
         Caption         =   "Detalle a Facturar"
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
         RowDividerColor =   15263202
         RowSubDividerColor=   16777215
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.locked=0,.bold=0"
         _StyleDefs(7)   =   ":id=1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H4D7C2A&,.fgcolor=&HFFFFFF&"
         _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Arial"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H70655D&,.fgcolor=&H80000005&"
         _StyleDefs(13)  =   ":id=2,.borderSize=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(14)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=2,.fontname=Arial"
         _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HF2FAED&,.fgcolor=&H232323&,.bold=0"
         _StyleDefs(19)  =   ":id=6,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(20)  =   ":id=6,.fontname=Arial"
         _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
         _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1"
         _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1"
         _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1"
         _StyleDefs(27)  =   "Splits(0).Style:id=87,.parent=1"
         _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
         _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=89,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=91,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8,.namedParent=38"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=95,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=106,.parent=87"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=88"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=89"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=91"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
         _StyleDefs(45)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
         _StyleDefs(46)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=33,.fontname=Small Fonts"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(50)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(51)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(52)  =   ":id=34,.fontname=Arial"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H808080&,.locked=-1,.appearance=0"
         _StyleDefs(57)  =   ":id=36,.borderColor=&H808000&,.bold=0,.fontsize=900,.italic=0,.underline=0"
         _StyleDefs(58)  =   ":id=36,.strikethrough=0,.charset=0"
         _StyleDefs(59)  =   ":id=36,.fontname=Arial"
         _StyleDefs(60)  =   "Named:id=37:Caption"
         _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(62)  =   "Named:id=38:HighlightRow"
         _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.appearance=0"
         _StyleDefs(64)  =   ":id=38,.borderSize=1,.bold=-1,.fontsize=675,.italic=0,.underline=0"
         _StyleDefs(65)  =   ":id=38,.strikethrough=0,.charset=0"
         _StyleDefs(66)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(67)  =   "Named:id=39:EvenRow"
         _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(69)  =   "Named:id=40:OddRow"
         _StyleDefs(70)  =   ":id=40,.parent=33"
         _StyleDefs(71)  =   "Named:id=41:RecordSelector"
         _StyleDefs(72)  =   ":id=41,.parent=34"
         _StyleDefs(73)  =   "Named:id=42:FilterBar"
         _StyleDefs(74)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtBarra 
         Height          =   315
         Left            =   360
         TabIndex        =   59
         Top             =   6480
         Width           =   1635
      End
      Begin VB.Label txtObservacion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6000
         TabIndex        =   71
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label txtEntregar 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6000
         TabIndex        =   70
         Top             =   780
         Width           =   3135
      End
      Begin VB.Label txtTelefono 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1080
         TabIndex        =   69
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label txtCliente 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1080
         TabIndex        =   68
         Top             =   780
         Width           =   2895
      End
      Begin VB.Label lblLey1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ICBPER 0.10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007EC642&
         Height          =   375
         Left            =   2520
         TabIndex        =   67
         Top             =   6480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label txtDescuento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,500.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5400
         TabIndex        =   66
         Top             =   6480
         Width           =   975
      End
      Begin VB.Label txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "250,500.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7080
         TabIndex        =   65
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label txtMontoLetras 
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
         Height          =   285
         Left            =   8640
         TabIndex        =   64
         Top             =   6840
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   3255
      Left            =   0
      TabIndex        =   72
      Top             =   7200
      Width           =   9255
      Begin VB.CommandButton cmdTipoDocumento 
         BackColor       =   &H008BDD79&
         Caption         =   "BOLETA SIN DATOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   0
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   411
         Top             =   2400
         Width           =   1769
      End
      Begin VB.PictureBox cmdOpcionPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   10
         Left            =   2370
         Picture         =   "frmCajaRapida.frx":211E5
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   401
         Top             =   840
         Width           =   2200
      End
      Begin VB.CommandButton cmdTipoDocumento 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   3
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   2400
         Width           =   1769
      End
      Begin VB.CommandButton cmdTipoDocumento 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   2
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   2400
         Width           =   1769
      End
      Begin VB.CommandButton cmdTipoDocumento 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   2400
         Width           =   1769
      End
      Begin VB.CommandButton cmdTipoDocumento 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   4
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2400
         Width           =   1769
      End
      Begin VB.PictureBox cmdOpcionPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   8
         Left            =   6890
         Picture         =   "frmCajaRapida.frx":22CFA
         ScaleHeight     =   675
         ScaleWidth      =   2235
         TabIndex        =   83
         Top             =   120
         Width           =   2235
      End
      Begin VB.PictureBox cmdOpcionPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   0
         Left            =   105
         Picture         =   "frmCajaRapida.frx":24CCF
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   82
         Top             =   1545
         Width           =   2200
      End
      Begin VB.PictureBox cmdOpcionPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   1
         Left            =   105
         Picture         =   "frmCajaRapida.frx":26B7D
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   81
         Top             =   840
         Width           =   2200
      End
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   9
         Left            =   6890
         Picture         =   "frmCajaRapida.frx":288A3
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   80
         Top             =   840
         Width           =   2200
      End
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   5
         Left            =   4630
         Picture         =   "frmCajaRapida.frx":2A21E
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   79
         Top             =   120
         Width           =   2200
      End
      Begin VB.PictureBox cmdCabeceraPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   7
         Left            =   2370
         Picture         =   "frmCajaRapida.frx":2C1B0
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   78
         Top             =   135
         Width           =   2200
      End
      Begin VB.PictureBox cmdOpcionPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   2
         Left            =   6890
         Picture         =   "frmCajaRapida.frx":2D9C8
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   76
         Top             =   1545
         Width           =   2200
      End
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   13
         Left            =   4630
         Picture         =   "frmCajaRapida.frx":2EF15
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   75
         Top             =   1560
         Width           =   2200
      End
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   10
         Left            =   4630
         Picture         =   "frmCajaRapida.frx":306DD
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   74
         Top             =   840
         Width           =   2200
      End
      Begin VB.PictureBox cmdDetallePicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   12
         Left            =   2370
         Picture         =   "frmCajaRapida.frx":3202B
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   73
         Top             =   1545
         Width           =   2200
      End
      Begin VB.PictureBox cmdCabeceraPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   680
         Index           =   6
         Left            =   120
         Picture         =   "frmCajaRapida.frx":33CCC
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   77
         Top             =   135
         Width           =   2200
      End
      Begin VB.Label txtFechaEntrega 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   409
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame FramePieForm 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame9"
      Height          =   1095
      Left            =   0
      TabIndex        =   265
      Top             =   10440
      Width           =   15465
      Begin VB.PictureBox cmdCabeceraPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   0
         Left            =   3880
         Picture         =   "frmCajaRapida.frx":356B4
         ScaleHeight     =   720
         ScaleWidth      =   1815
         TabIndex        =   403
         Top             =   240
         Visible         =   0   'False
         Width           =   1820
      End
      Begin VB.PictureBox cmdOpcionPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   14
         Left            =   2040
         Picture         =   "frmCajaRapida.frx":3752A
         ScaleHeight     =   720
         ScaleWidth      =   1815
         TabIndex        =   402
         Top             =   240
         Visible         =   0   'False
         Width           =   1820
      End
      Begin VB.PictureBox cmdTipoDocumentoPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   6
         Left            =   14400
         Picture         =   "frmCajaRapida.frx":38F78
         ScaleHeight     =   720
         ScaleWidth      =   900
         TabIndex        =   273
         Top             =   240
         Width           =   900
      End
      Begin VB.PictureBox cmdTipoDocumentoPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   5
         Left            =   13440
         Picture         =   "frmCajaRapida.frx":39FF2
         ScaleHeight     =   720
         ScaleWidth      =   900
         TabIndex        =   272
         Top             =   240
         Width           =   900
      End
      Begin VB.PictureBox cmdTipoDocumentoPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   4
         Left            =   12480
         Picture         =   "frmCajaRapida.frx":3B318
         ScaleHeight     =   720
         ScaleWidth      =   900
         TabIndex        =   271
         Top             =   240
         Width           =   900
      End
      Begin VB.PictureBox cmdTipoDocumentoPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   3
         Left            =   11520
         Picture         =   "frmCajaRapida.frx":3C6B3
         ScaleHeight     =   720
         ScaleWidth      =   900
         TabIndex        =   270
         Top             =   240
         Width           =   900
      End
      Begin VB.PictureBox cmdTipoDocumentoPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   2
         Left            =   10560
         Picture         =   "frmCajaRapida.frx":3DA57
         ScaleHeight     =   720
         ScaleWidth      =   900
         TabIndex        =   269
         Top             =   240
         Width           =   900
      End
      Begin VB.PictureBox cmdTipoDocumentoPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   1
         Left            =   9600
         Picture         =   "frmCajaRapida.frx":3ED5F
         ScaleHeight     =   720
         ScaleWidth      =   900
         TabIndex        =   268
         Top             =   240
         Width           =   900
      End
      Begin VB.PictureBox cmdOpcionPicture 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   5
         Left            =   120
         Picture         =   "frmCajaRapida.frx":3FEC5
         ScaleHeight     =   720
         ScaleWidth      =   1815
         TabIndex        =   267
         Top             =   240
         Width           =   1820
      End
      Begin VB.PictureBox cmdSinBoton 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   5880
         Picture         =   "frmCajaRapida.frx":41415
         ScaleHeight     =   720
         ScaleWidth      =   3615
         TabIndex        =   266
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Label txtTitulo 
      BackColor       =   &H004D7C2A&
      Caption         =   " Caja Rápida 001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15465
   End
   Begin VB.Line Line2 
      X1              =   120.023
      X2              =   15602.99
      Y1              =   11640
      Y2              =   11640
   End
   Begin VB.Line Line1 
      X1              =   15602.99
      X2              =   15602.99
      Y1              =   120
      Y2              =   11640
   End
End
Attribute VB_Name = "frmCajaRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hwnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

Dim nMontoPedidoFacturar As Double
' anulacion por nota de credito
Dim RsTparametro As Recordset
'--------------------------------
Dim xOperador As String
Dim wAgregarPropiedad As Boolean
Dim numeroSerieImpresora As String 'SUNAT
Dim codigoImpresora As String 'SUNAT
Dim rstFuente As ADODB.Recordset
Dim RsDetalle As Recordset
Dim RsProducto As Recordset
Dim RsPropiedad As Recordset
Dim RsArea As Recordset
Dim RsTipoDocumento As Recordset
Dim RsImpresion As Recordset
Dim RsCombo As Recordset
Dim RsMozo As Recordset
Dim RsMotivoEliminacion As Recordset
Dim RsOferta As Recordset
Dim RsProductoPropiedad As Recordset
Dim RsComboPropiedad As Recordset
Dim RsOperador As Recordset
Dim rsPuntoVenta As Recordset
Dim RsProductoCombo As Recordset
Dim RsCajaRapida As Recordset
Dim RsProductoX As Recordset
'Canales de Venta
Dim lActivaMozo As Boolean
Dim lActivaMotorizado As Boolean
Dim lObligaMesa As Boolean
'Origen de ventas
Dim lOrigenVentas As Boolean
'Public lObligaPax As Boolean
Dim lObligaMotorizado As Boolean
Dim lObligaMozo As Boolean
Dim lObligaFechaEntrega As Boolean
Dim lObligaClienteFrecuente As Boolean
Dim lCanalDelivery As Boolean
Dim lCanalCentralPedidos As Boolean
Dim Tienda As String
Dim RsMotorizado As Recordset
Dim lObligaOrigenVenta As Boolean

'origen de ventas
Dim RsOrigenVentas As Recordset
Dim RscanalOrigenVentas As Recordset
Dim vOrigenVentas As String
'Dim lOrigenVentas As Boolean
Dim RsCanalesVenta As Recordset
'-----------------------------
Dim ValProductoAsociado As Recordset


Dim sProducto As String
Dim sProductoCombo As String

Dim nMonto As Double

Public sDetalle As String
Public sProductoPropiedad As String
'validacionMontoMInimo
Dim nMontoPedidoFacturarMInimo As Double
'validacionMontoMInimo

'JSJR 09/01/2025
Public wAnulaDocumento As Boolean

Dim i As Integer
Dim sCortesia As String
Dim sCombo As String
Public sComboDetalle As String
Public sComboPropiedad As String
Dim sDetalleConsumo As String
Dim Index As Integer
'Dim sTipoPedido As String      se comento para el requerimiento REST-047-C-2024  JAPN
Dim sMonedaBase As String
Dim sPuntoVenta As String
Dim sComandaInfhotel As String

'Variables para oferta
Dim ValorOferta As Recordset
Dim respuestaOferta As String

'Variables Combo
Dim wCombo As Boolean
Dim wAgregaCombo As Boolean
Dim nCombo As Integer


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
Dim sMotorizado As String
Dim sObser As String
Dim sSubGrupo As String
Dim sGrupo As String
Dim nPos As Integer
Dim nCCombo As Double
Dim sUsuarioAutoriza As String
Dim xDescuento As Double
Public Pedido As String
Dim nOperadorPropiedad As Integer
Dim nOrden As Integer
Dim lPropiedad As Boolean
Dim lComboActivo As Boolean
Dim nRet As Integer
Dim sOperacion As String
Dim sRetorno As String * 512
Dim sClave As String
Dim sMonto As String
Dim xError As String
Dim sRefer As String
Dim nCorrela As String
Dim lEmisor As Boolean
Dim lLoop As Boolean
Dim nContador As Integer
Dim sPrefijo As String

Dim sTD As String
Dim xSuma As Double
Dim sCompania As String
Dim sContacto As String
Dim UltimaComanda As String

Dim tAutorizaDescuento As String
Dim sCodigoDescuento As String
Dim sDescripcionDescuento As String
Dim sClienteFrecuente As String

Dim ltope As Boolean
Dim nTope As Double
Dim lRatio As Boolean
Dim Acumulado As Double
Dim lImprimeAlternativa As Boolean
Dim lAplicablePedido As Boolean

'============================================= extranjero bolivia
Dim tAutorizacion As String
Dim tCodigoControl As String
Dim tDosificacion As String
Dim tIdentidadNIT As String

Dim muestra As String
Dim variableEmite As Boolean
Dim nTotalDescuento As Double
Dim sXML As String

'insumocombo
Dim sInsumoCombo As String


'------VALIDA CORREO----------
Dim sTipoDocum As String
Dim lValidaEmail As Boolean
Dim sEmail As String


'FACTURACION_E_PERU
'Dim RsImpDocumentoE As New Recordset
Dim RsCodigoHash As New ADODB.Recordset
Dim fDocumento As String
Dim xMontoTexto As String
Dim iImagenCab As Boolean
Dim xImpresionFE As String
Dim xImpresioDE As String
    
Dim RsPedido As New Recordset

'
Dim tUsuActua As String
Dim TimpresionDolaresDelivery  As Boolean

'caja contingencia
Dim sfemision As Date
Dim Rscajacontigencia As Recordset
Dim RsExisteOferta As Recordset

'Verificacion de Combo
Dim iProductoVerificaComboMaximo As Integer
Dim iProductoVerificaComboMinimo As Integer
Dim iProductoVerificaComboUsado As Integer
Dim sProductoVerificaComboUsado As String

Dim EliminarProductoCom As Integer

Private Sub ImportarPedido(tempPedido As String)
    
         Isql = "insert into [" & sProductoPropiedad & "] " & _
               "(tItem,tCodigoPropiedad,tProducto,tEnlace,nInsumo,nGasto,nManoObra,nCantidad,nInsumoUnitario,nGastounitario,nManoObraUnitario) " & _
               "select tItem,tCodigoPropiedad,tProducto,tEnlace,nInsumo,nGasto,nManoObra,nCantidad,nInsumoUnitario,nGastounitario,nManoObraUnitario from TPRODUCTOPROPIEDAD where tcodigopedido='" & tempPedido & "' "
        
         Cn.Execute Isql
         RsProductoPropiedad.Requery
         
         Isql = "insert into [" & sDetalle & "] " & _
                "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
                "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
                "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem,tsubalmacen,toferta,tCajaD,tObservacion) " & _
                "select tCodigoPedido,tTipoPedido,tItem,tCodigoProducto, tCodigoGrupo,tCodigoSubGrupo,nPrecioNeto,nRecargo,nDescuento, " & _
                "nPrecioOficial,nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta,nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
                "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem,tsubalmacen,toferta,tCajaD,tObservacion from dpedido where tcodigopedido='" & tempPedido & "' "
         Cn.Execute Isql
         RsDetalle.Requery
                             
         Isql = "insert into [" & sComboPropiedad & "] " & _
               "(tItem,tItemCombo,tCodigoPropiedad,tProducto,tEnlace,nInsumo,nGasto,nManoObra,nCantidad,nInsumoUnitario,nGastoUnitario,nManoObraUnitario) " & _
               "select tItem,tItemCombo,tCodigoPropiedad,tProducto,tEnlace,nInsumo,nGasto,nManoObra,nCantidad,nInsumoUnitario,nGastoUnitario,nManoObraUnitario from TCOMBOPROPIEDAD where tcodigopedido='" & tempPedido & "' "

         Cn.Execute Isql
         RsComboPropiedad.Requery
         
         Isql = "insert into [" & sComboDetalle & "] " & _
               "(tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, tObservacion, lCorte) " & _
               "select tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, tObservacion, lCorte from CPEDIDO where tcodigopedido='" & tempPedido & "' "
        
         Cn.Execute Isql
         RsCombo.Requery
        
         cargarDatosCabecera (tempPedido)
         
         nMonto = Format(Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn), "#,###,##0.00")
         VisualizaMonto
End Sub

Private Sub LimpiarData()
        Cn.Execute "delete " & sDetalle
        Cn.Execute "delete " & sComboDetalle
        Cn.Execute "delete " & sComboPropiedad
        Cn.Execute "delete " & sProductoPropiedad

        RsDetalle.Requery
        RsComboPropiedad.Requery
        RsProductoPropiedad.Requery
        Inicializar
        Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAgrupacion_Click(Index As Integer)
   For i = 1 To 9
        If i = Index Then
           cmdAgrupacion(i).backColor = &HE7ECD5
           cmdAgrupacion(i).FontBold = True
        Else
            cmdAgrupacion(i).backColor = &HF7F6F5
            cmdAgrupacion(i).FontBold = False
        End If
    Next i
   
   For i = 1 To 42
       'cmdBoton(i).Visible = False
       cmdBoton(i).backColor = &HE5E4DF
       cmdBoton(i).Enabled = False
       cmdBoton(i).Caption = "-"
   Next i
   RsProducto.Filter = "tCajaRapida=" & sPrefijo & Index
 
   If Not RsProducto.EOF Then
      RsProducto.MoveFirst
      Do While Not RsProducto.EOF
         If RsProducto!nBotonRapido > 0 Then
            cmdBoton(RsProducto!nBotonRapido).Visible = True
            cmdBoton(RsProducto!nBotonRapido).Enabled = True
            cmdBoton(RsProducto!nBotonRapido).backColor = cmdAgrupacion(Index).backColor
            cmdBoton(RsProducto!nBotonRapido).Caption = RsProducto!tResumido
            If ((sTipoPedido = "01" And RsProducto!lLocal = False) Or (sTipoPedido = "02" And RsProducto!lDelivery = False) Or (sTipoPedido = "03" And RsProducto!lLlevar = False) Or (RsProducto!tUnidadNegocio <> sUnidadNegocio)) Then
               cmdBoton(RsProducto!nBotonRapido).Enabled = False
            End If
         End If
         RsProducto.MoveNext
      Loop
   End If
   If txtBarra.Visible = True Then
      txtBarra.SetFocus
   End If
'   sCajaRapida = sPrefijo & Index
End Sub

Private Sub cmdBoton_Click(Index As Integer)
   txtBarra.SetFocus
      
   RsProducto.MoveFirst
   RsProducto.Find "nbotonRapido = " & Trim(str(Index))
   sProducto = RsProducto!codigo
   
     'INSUMOCRITICO23
        If validadIngresoProducto(sProducto) = False Then
            Exit Sub
        End If
    'INSUMOCRITICO23
   
    If lBal And RsProducto!lBalanza Then
       Dim nResultado As Double
       nResultado = Pesar(nBalanzaPuerto)
       nResultado = Format(nResultado, "#,##0.000")
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
End Sub

Private Sub cmdBusca_Click()
    sTipo = ""
    sTemp = ""
    'Isql = "select * from vOperador where lActivo = 1 Order by Descripcion "
    ListarOperadoresConFiltro (sProducto)
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    
    If wEnter = True Then
        Screen.MousePointer = vbHourglass
        For i = 1 To 13
            cmdOperador(i).backColor = &HD1F2CA
        Next i
        RsOperador.MoveFirst
        RsOperador.Find "Codigo='" & sCodigo & "'"
        xOperador = sCodigo
        If Not RsOperador.EOF And RsOperador!nBoton > 0 Then
           cmdOperador(RsOperador!nBoton).backColor = &H8BDD79
        End If
        AsignaPropiedad
        Screen.MousePointer = vbDefault
    End If
    txtBarra.SetFocus
End Sub

Private Sub cmdBuscar_Click(Index As Integer)
   Select Case Index
   
   Case Is = 0
      sTemp = ""
      Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by Descripcion"
      Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                      "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                        
      frmBusquedaRapida.nPredeterm = 1
      frmBusquedaRapida.Show vbModal
      If wEnter = True Then
         sMozo = sCodigo
         txtTitulo.Caption = " Caja Rápida : " & sCaja & " Mesero: " & sDescrip & IIf(sPedido = "", "", " - " & sPedido)
      End If
      tabProducto.Visible = True
      fraMozo.Visible = False
   
   Case Is = 1
      sTemp = ""
      sTipo = "Infhotel"
      Isql = "Select tPuntoVenta as Codigo, tDescripcion as Descripcion, nUltimoComanda, tmoneda" & _
             " From tPuntoVenta " & _
             " where tHotel='" & sHotel & "' AND lActivo=1 and lInforest=1"
      Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                      "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                         
      frmBusquedaRapida.nPredeterm = 1
      frmBusquedaRapida.Show vbModal
      If wEnter = True Then
         sPuntoVenta = sCodigo
         rsPuntoVenta.MoveFirst
         rsPuntoVenta.Find "Codigo= '" & sCodigo & "'"
         'cmdCabecera(0).Caption = rsPuntoVenta!Descripcion
      End If
      
      tabProducto.Visible = True
      fraPuntoVenta.Visible = False
      
   Case Is = 2 'ProductoCombo
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
            
            nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                               "FROM " & sComboDetalle & "  WHERE    tItem='" & sItem & "'", Cn)
            If nCCombo < nCombo * RsDetalle!nCantidad Then
                            'Oscar Ortega----------------------------------------------
                            Dim oRsProductoDeCombo As Recordset
                            Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sProducto)
                            If oRsProductoDeCombo.RecordCount > 0 Then
                                If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                                    'Obtener Suma de cantidades
                                    Dim nCantidadEnElCombo As Integer
                                    nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sItem, oRsProductoDeCombo!tEtiqueta)
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

            Else
               MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
            End If
            RsProducto.Filter = IIf(xxx = "0", "", xxx)
          End If

    Case 8
                sTipo = ""
                sTemp = ""
                'Isql = "select * from vOperador where lActivo = 1 Order by Descripcion "
                'ListarOperadoresConFiltro (sProducto)
                Dim sPropiedad As String
                
                
                sPropiedad = dbTemporal(sCaja, 11, "Codigo", "nVarChar(20)", _
                                                    "tProducto", "nVarChar(10)", _
                                                    "Operador", "nVarChar(150)", _
                                                    "Descripcion", "nVarChar(150)", _
                                                    "tOperador", "nVarChar(2)", _
                                                    "nPrecio", "float", _
                                                    "tEnlace", "nVarChar(15)", _
                                                    "nInsumo", "float", _
                                                    "nGasto", "float", _
                                                    "nManoObra", "float", _
                                                    "tEstado", "nvarchar(50)")
                                                    
             If wAgregaCombo = False Then
                    Isql = " insert into " & sPropiedad & " select tCodigoPropiedad as Codigo,tProducto, " & _
                           " tOperador.tDetallado AS Operador, TPROPIEDAD.tDetallado as Propiedad, " & _
                           " TPROPIEDAD.tOperador, nPrecio, tEnlace, nInsumo, nGasto, nManoObra, 'Agregar' " & _
                           " FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador =  " & _
                           " dbo.TOPERADOR.tOperador " & _
                           " Where TOPERADOR.tOperador='" & xOperador & "' AND TPROPIEDAD.tProducto='" & sProducto & "' AND TPROPIEDAD.lActivo = 1 And IsNull(tOperador.lStockMenos, 0) <> 1"
                    
                    Cn.Execute Isql
                    If lAlmacen = True Then
                    'CASO CORP.FERNANDEZ
                        If BDLink = "1" Then
                             If Calcular("select count(*) as codigo from vOperador where lStockMenos=1  and Codigo='" & xOperador & "'", Cn) > 0 Then
                                     Isql = "  insert into " & sPropiedad & " select '9999' as Codigo,  tCodigoPlato as tProducto, 'Sin' as Operador,  " & _
                                             "  tDetallado as Propiedad,    '" & xOperador & "' as tOperador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, " & _
                                             "  0, 0 ,  'Agregar' FROM " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.DRECETAVENTA t1x INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta  " & _
                                             " INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto   " & _
                                             " Where lNoDescargo = 1 and " & "t1x.tLocal='" & sLocal & "' and tCodigoPlato='" & sProducto & "'"
                
                                     Cn.Execute Isql
                             End If
                        Else
                            If Calcular("select count(*) as codigo from vOperador where lStockMenos=1  and Codigo='" & xOperador & "'", Cn) > 0 Then
                                     Isql = "  insert into " & sPropiedad & " select '9999' as Codigo,  tCodigoPlato as tProducto, 'Sin' as Operador,  " & _
                                             "  tDetallado as Propiedad,    '" & xOperador & "' as tOperador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, " & _
                                             "  0, 0 ,  'Agregar' FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA t1x INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta  " & _
                                             " INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto   " & _
                                             " Where lNoDescargo = 1 and " & "t1x.tLocal='" & sLocal & "' and tCodigoPlato='" & sProducto & "'"
                
                                     Cn.Execute Isql
                             End If
                        End If
                        
                    
                    End If
                    
                    Isql = "  update " & sPropiedad & " SET tEstado='Quitar' from " & sPropiedad & " inner join (SELECT " & sProductoPropiedad & ".tCodigoPropiedad, " & sProductoPropiedad & ".tProducto , " & sProductoPropiedad & ".tEnlace     " & _
                           "  FROM  " & sProductoPropiedad & "   where  tItem='" & sItem & "' and  " & sProductoPropiedad & ".TPRODUCTO='" & sProducto & "') t1 on " & sPropiedad & ".Codigo=t1.tCodigoPropiedad and " & sPropiedad & ".tProducto=t1.tProducto and " & sPropiedad & ".tEnlace=t1.tEnlace "
 
                    Cn.Execute "delete from " & sPropiedad & " where Codigo in (select tcodigopropiedad from VISIBILIDADPROPIEDADXCANAL where isnull(visibilidad,0)=1 and  isnull(canal" & Replace(sTipoPedido, "0", "") & ",0)=0)"
'
'                    Isql = " update " & sPropiedad & " SET tEstado='Quitar' from " & sPropiedad & " inner join (SELECT " & sProductoPropiedad & ".tCodigoPropiedad , " & sProductoPropiedad & ".tProducto   From  dbo.TOPERADOR INNER JOIN " & sProductoPropiedad & " INNER JOIN (  select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion,   tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, nInsumo,    nGasto, nManoObra, tOperador.tDetallado AS Operador    FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador =     dbo.TOPERADOR.tOperador Where TPROPIEDAD.lActivo = 1 And      IsNull(TOPERADOR.lStockMenos, 0) <> 1) " & _
'                           " T1 ON  " & sProductoPropiedad & ".tCodigoPropiedad = T1.Codigo AND " & sProductoPropiedad & ".tProducto = T1.tProducto AND " & sProductoPropiedad & ".tEnlace = T1.tEnlace ON dbo.tOperador.tOperador = T1.tOperador        COLLATE Modern_Spanish_CI_AS LEFT OUTER JOIN dbo.TPROPIEDAD ON dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador AND " & sProductoPropiedad & ".tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND " & sProductoPropiedad & ".tProducto = dbo.TPROPIEDAD.tProducto where   tItem='" & sitem & "' and " & sProductoPropiedad & ".TPRODUCTO='" & sProducto & "' and TOPERADOR.tOperador='" & xOperador & "' ) t1 on " & sPropiedad & ".Codigo=t1.tCodigoPropiedad and " & sPropiedad & ".tProducto=t1.tProducto"
                           
                    Cn.Execute Isql
                    Isql = "SELECT * FROM " & sPropiedad
             Else
             
                    Isql = " insert into " & sPropiedad & " select tCodigoPropiedad as Codigo,tProducto,tOperador.tDetallado AS Operador, TPROPIEDAD.tDetallado as Propiedad, TPROPIEDAD.tOperador, nPrecio, tEnlace, nInsumo, nGasto, " & _
                           " nManoObra, 'Agregar' FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador Where TOPERADOR.tOperador='" & xOperador & "' AND TPROPIEDAD.tProducto='" & sCombo & "' AND TPROPIEDAD.lActivo = 1 And IsNull(tOperador.lStockMenos, 0) <> 1 "
                    Cn.Execute Isql


                    If lAlmacen = True Then
                        'CASO CORP.FERNANDEZ
                        If BDLink = "1" Then
                             If Calcular("select count(*) as codigo from vOperador where lStockMenos=1  and Codigo='" & xOperador & "'", Cn) > 0 Then
                                     Isql = "  insert into " & sPropiedad & " select '9999' as Codigo,  tCodigoPlato as tProducto, 'Sin' as Operador,  " & _
                                             "  tDetallado as Propiedad,    '" & xOperador & "' as tOperador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, " & _
                                             "  0, 0 ,  'Agregar' FROM " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.DRECETAVENTA t1x INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.MRECETAVENTA  t2X ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta  " & _
                                             " INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto   " & _
                                             " Where lNoDescargo = 1 and " & "t1x.tLocal='" & sLocal & "' and tCodigoPlato='" & sCombo & "'"
                
                                     Cn.Execute Isql
                             End If
                        Else
                            If Calcular("select count(*) as codigo from vOperador where lStockMenos=1  and Codigo='" & xOperador & "'", Cn) > 0 Then
                                     Isql = "  insert into " & sPropiedad & " select '9999' as Codigo,  tCodigoPlato as tProducto, 'Sin' as Operador,  " & _
                                             "  tDetallado as Propiedad,    '" & xOperador & "' as tOperador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, " & _
                                             "  0, 0 ,  'Agregar' FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA t1x INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA  t2X ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta  " & _
                                             " INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto   " & _
                                             " Where lNoDescargo = 1 and " & "t1x.tLocal='" & sLocal & "' and tCodigoPlato='" & sCombo & "'"
                
                                     Cn.Execute Isql
                             End If
                        End If
                    
                    End If

                    Isql = "  update " & sPropiedad & " SET tEstado='Quitar' from " & sPropiedad & " inner join (SELECT " & sComboPropiedad & ".tCodigoPropiedad, " & sComboPropiedad & ".tProducto , " & sComboPropiedad & ".tEnlace     " & _
                           "  FROM  " & sComboPropiedad & "   where  " & sComboPropiedad & ".tItem='" & sItem & "' and  " & sComboPropiedad & ".tItemcombo='" & xItem & "' and  " & sComboPropiedad & ".TPRODUCTO='" & sCombo & " ') t1 on " & sPropiedad & ".Codigo=t1.tCodigoPropiedad and " & sPropiedad & ".tProducto=t1.tProducto and " & sPropiedad & ".tEnlace=t1.tEnlace  "
 
                    Cn.Execute "delete from " & sPropiedad & " where Codigo in (select tcodigopropiedad from VISIBILIDADPROPIEDADXCANAL where isnull(visibilidad,0)=1 and  isnull(canal" & Replace(sTipoPedido, "0", "") & ",0)=0)"

'                    Isql = "update " & sPropiedad & " SET tEstado='Quitar' from " & sPropiedad & " inner join (SELECT   " & sComboPropiedad & ".tItem, " & sComboPropiedad & ".tItemCombo,tpropiedad.tCodigoPropiedad, " & sComboPropiedad & ".TPRODUCTO,  T1.Descripcion, T1.Operador  FROM         dbo.TOPERADOR INNER JOIN dbo.TPROPIEDAD ON   dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador RIGHT OUTER JOIN " & sComboPropiedad & " INNER JOIN    (select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, nInsumo, nGasto, nManoObra, tOperador.tDetallado AS Operador FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.tOperador.tOperador Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1) " & _
'                            " T1 ON " & sComboPropiedad & ".tCodigoPropiedad = T1.Codigo AND " & sComboPropiedad & ".tProducto =  T1.tProducto AND " & sComboPropiedad & ".tEnlace = T1.tEnlace ON dbo.TOPERADOR.tOperador = T1.tOperador COLLATE Modern_Spanish_CI_AS AND dbo.TPROPIEDAD.tCodigoPropiedad =" & sComboPropiedad & ".tCodigoPropiedad AND dbo.TPROPIEDAD.tProducto = " & sComboPropiedad & ".tProducto
                        'where TOPERADOR.tOperador='" & xOperador & "' and  " & sComboPropiedad & ".tItem='" & sitem & "' and " & sComboPropiedad & ".tItemCombo='" & xItem & "' and " & sComboPropiedad & ".tProducto='" & sCombo & "' ) t1 on " & sPropiedad & ".Codigo=t1.tCodigoPropiedad and " & sPropiedad & ".tProducto=t1.tProducto"

                    Cn.Execute Isql
             
                    Isql = "SELECT * FROM " & sPropiedad
             End If

                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Estado", 2, "tEstado", 1500, 0, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 5500, 0, 0, "")
                                                                
'                frmBusquedaRapida.nPredeterm = 1
'                 If tipoBusqueda = "BusquedaCajaRapida" Then
'                    tenlacebusqueda = IIf(RsGrilla.EOF = True, "", RsGrilla!tenlace)
'                    tipoBusqueda = ""
'                End If
                frmBusquedaRapida.tipoBusqueda = "BusquedaCajaRapida"
                
                frmBusquedaRapida.Show vbModal
                
                If wEnter = True Then
                  '  Screen.MousePointer = vbHourglass
                    If wAgregaCombo = False Then
                            If Calcular("SELECT COUNT(*) AS CODIGO FROM " & sProductoPropiedad & " WHERE tItem='" & sItem & "' AND tCodigoPropiedad='" & sCodigo & "' AND TPRODUCTO='" & sProducto & "' and tEnlace='" & tenlacebusqueda & "'", Cn) = 0 Then
                                    wAgregarPropiedad = True
                            Else
                                    wAgregarPropiedad = False
                            End If
                             AgregarPropiedadBusqueda sCodigo, sDescrip
                    Else
                    
                           If Calcular("SELECT COUNT(*) AS CODIGO FROM " & sComboPropiedad & " WHERE tItem='" & sItem & "' AND tCodigoPropiedad='" & sCodigo & "' AND TPRODUCTO='" & sCombo & "' and titemcombo='" & xItem & "'", Cn) = 0 Then
                                    wAgregarPropiedad = True
                            Else
                                    wAgregarPropiedad = False
                            End If
                           AgregarPropiedadBusqueda sCodigo, sDescrip
                    End If
                    
                    
                   ' Screen.MousePointer = vbDefault
                End If
              '  txtBarra.SetFocus
        
    Case Is = 3 'Origen de ventas
        sTemp = ""
        Isql = "select * from vOrigenVenta where Activo = 1 and codCanalVenta='" & sTipoPedido & "' and CodOrigenVenta not in (select CodigoOrigenVenta from TCAJAORIGEN_BLOQUEO  where tcaja='" & sCaja & "' ) Order by Descripcion"
        Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "CodOrigenVenta", 2, "CodOrigenVenta", 1200, 2, 0, "", _
                                                        "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
        frmBusquedaRapida.nPredeterm = 1
        frmBusquedaRapida.tipoBusqueda = "OrigenVentas"
        frmBusquedaRapida.Show vbModal
        If wEnter = True Then
           vOrigenVentas = sCodigo
        Else
            Exit Sub
        End If
        validarOrigenVentas
        txtBarra.SetFocus

    Case Is = 4 'Motorizado
          'MsgBox "aqui estoy"
        sTemp = ""
        Isql = "select * from vMotorizado where lActivo = 1 Order by Descripcion"
        Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 2, "Codigo", 1200, 2, 0, "", _
                                                        "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
        frmBusquedaRapida.nPredeterm = 1
        frmBusquedaRapida.tipoBusqueda = "Motorizado"
        frmBusquedaRapida.Show vbModal
        If wEnter = True Then
           sMotorizado = sCodigo
           Me.fraMorotizado.Visible = False
        Else
            Exit Sub
        End If
'        validarOrigenVentas
        txtBarra.SetFocus
   End Select
   
   txtBarra.SetFocus
End Sub
Public Sub AgregarPropiedadBusqueda(ByVal CodigoPropiedad As String, ByVal DescripcionPropiedad As String)
 

   '  HabilitaTimerColor (False)
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
    For k = 1 To 19
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
           ' cmdPropiedad(nBotonPropiedad).FontBold = False
            If Not RsPropiedad.EOF Then
                If wAgregaCombo Then
                    Cantidad = Calcular("select isnull(ncantidad,1) as codigo from " & sComboPropiedad & " where   titem='" & sItem & "' and titemcombo='" & xItem & "' and  tproducto='" & sCombo & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' ", Cn)

                    Cn.Execute "delete " & sComboPropiedad & " where tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and tProducto='" & sCombo & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "'"
                Else
                    Cantidad = Calcular("select isnull(ncantidad,1) as codigo from " & sProductoPropiedad & " where   titem='" & sItem & "' and tproducto='" & sProducto & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' and tenlace='" & RsPropiedad!tEnlace & "'", Cn)
                
                    Cn.Execute "delete " & sProductoPropiedad & "  where tItem = '" & sItem & "' and tProducto='" & sProducto & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "' and tEnlace='" & RsPropiedad!tEnlace & "'"
                     If RsPropiedad!nPrecio <> 0 Then
                            nMonto = CambiaPrecio(nPVenta - RsPropiedad!nPrecio)
                            txtMonto.Caption = Format(nMonto, "###,##0.00")
                     End If
                End If
                If Cantidad <> 1 Then
                           lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & DescripcionPropiedad & ": (" & Cantidad & "), ", "")
                Else
                           lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & DescripcionPropiedad & ", ", "")
                End If
                
'                lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & DescripcionPropiedad & ", ", "")
                
            End If
        Else
        
            ncantidadPropiedad = 1
            If RsPropiedad!lsolicitacantidad = 1 Or RsPropiedad!lsolicitacantidad = True Then
                sTipo = "Prepintado"
            
                sCodigo = ncantidadPropiedad
            
                frmNumPad.Show vbModal
                If wEnter And val(sDescrip) > 0 Then
            
                            ncantidadPropiedad = sDescrip
                        
                End If
            End If
        
            If nOperadorPropiedad > 0 Then
                If wAgregaCombo Then
                       Isql = "SELECT COUNT(" & sComboPropiedad & ".tCodigoPropiedad) AS codigo " & _
                              "FROM " & sComboPropiedad & " INNER JOIN dbo.TPROPIEDAD ON " & sComboPropiedad & ".tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND " & sComboPropiedad & ".tProducto = dbo.TPROPIEDAD.tProducto " & _
                              "where tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and " & sComboPropiedad & ".tProducto='" & sCombo & "'  and tOperador='" & RsOperador!codigo & "'"
                   If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                      MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                      Exit Sub
                   End If
                Else
                    Isql = "SELECT COUNT(" & sProductoPropiedad & ".tCodigoPropiedad) AS codigo FROM " & sProductoPropiedad & " INNER JOIN " & _
                    "dbo.TPROPIEDAD ON " & sProductoPropiedad & ".tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND " & sProductoPropiedad & ".tProducto = dbo.TPROPIEDAD.tProducto " & _
                    "where tItem = '" & sItem & "' and tOperador='" & RsOperador!codigo & "'"
                    If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                       MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                       Exit Sub
                    End If
                End If
            End If
       
    
          '  cmdPropiedad(nBotonPropiedad).FontBold = True
            If Not RsPropiedad.EOF Then
               nInsumo = IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo)
               nGasto = IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto)
               nMObra = IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra)
               If wAgregaCombo Then
                    Cn.Execute "Insert into " & sComboPropiedad & " values ('" & sItem & "', '" & xItem & "', '" & RsPropiedad!codigo & "', '" & sCombo & "', '" & RsPropiedad!tEnlace & "', " & IIf(IsNull(RsPropiedad!nInsumo), 0, ncantidadPropiedad * RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, ncantidadPropiedad * RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, ncantidadPropiedad * RsPropiedad!nManoObra) & ", " & ncantidadPropiedad & ", " & IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra) & ") "
                    
                Else
                    Cn.Execute "Insert into " & sProductoPropiedad & " values ('" & sItem & "', '" & RsPropiedad!codigo & "', '" & sProducto & "', '" & RsPropiedad!tEnlace & "', " & IIf(IsNull(RsPropiedad!nInsumo), 0, ncantidadPropiedad * RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, ncantidadPropiedad * RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, ncantidadPropiedad * RsPropiedad!nManoObra) & ", " & ncantidadPropiedad & "," & IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra) & " )"
                    If RsPropiedad!nPrecio <> 0 Then
                       nMonto = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad))
                       txtMonto.Caption = Format(nMonto, "###,##0.00")
                    End If

                End If
          End If
  '        lblResumen.Text = lblResumen.Text & RsOperador!Descripcion & " " & DescripcionPropiedad & ", "
        If ncantidadPropiedad <> 1 Then
          
                lblResumen.Text = lblResumen.Text & RsOperador!descripcion & " " & DescripcionPropiedad & ": (" & ncantidadPropiedad & "), "
          Else
                lblResumen.Text = lblResumen.Text & RsOperador!descripcion & " " & DescripcionPropiedad & ", "
          End If
          
    End If
    
    
    If wAgregaCombo Then
       RsComboPropiedad.Requery
    Else
       RsProductoPropiedad.Requery
    End If
    
 '  HabilitaTimerColor (True)
End Sub
Private Sub ActivarOrigenVentas()
On Error GoTo fin:
'ELDCQ
       Me.fraOrigenVentas.Visible = True
       Isql = "select * from vOrigenVenta where Activo = 1 and Visible = 1 and CodCanalVenta='" & sTipoPedido & "' and CodOrigenVenta not in (select CodigoOrigenVenta from TCAJAORIGEN_BLOQUEO  where tcaja='" & sCaja & "' ) Order by Boton"
       Set RsOrigenVentas = Lib.OpenRecordset(Isql, Cn)
       
       If RsOrigenVentas.RecordCount > 0 Then
            AsignaBotonOrigenVentas 19, RsOrigenVentas, Me.cmdOrigen()
        Else
            AsignaBotonOrigenVentas 19, RsOrigenVentas, Me.cmdOrigen()
            Me.fraOrigenVentas.Visible = False
            If lObligaOrigenVenta = True Then
                Me.fraOrigenVentas.Visible = True
                MsgBox "Obligatoriedad de origen de ventas:" & vbNewLine & "No hay Origenes de venta visibles configuradas, " & vbNewLine & "Consultar con su administrador!!!", vbInformation
            End If
            
            If lActivaMotorizado Then
                fraMorotizado.Visible = True
            Else
                fraMorotizado.Visible = False
            End If
       End If
Exit Sub
fin:
End Sub

Private Sub LlenaColorCanalVenta(Canal As Integer)
    For i = 1 To 5
        If i = Canal Then
            cmdCabecera(i).FontBold = True
            cmdCabecera(i).backColor = &H7EC642
        Else
            cmdCabecera(i).FontBold = False
            cmdCabecera(i).backColor = &HF2FAED
        End If
    Next i
End Sub

Private Sub cmdBuscarPicture_Click(Index As Integer)
    cmdBuscar_Click Index
End Sub

Private Sub cmdCabecera_Click(Index As Integer)
   Select Case Index
      Case Is = 5 'Canal5
           If RsDetalle.RecordCount > 0 Then
                MsgBox "No se puede cambiar el canal de venta", vbExclamation, sMensaje
                Exit Sub
           Else
                sTipoPedido = "05"
                sMotorizado = "0000"
                vOrigenVentas = ""
                LlenaColorCanalVenta Index
                
                'origen de ventas
                RsCanalesVenta.Filter = "tCodigoCanalVenta = '" & sTipoPedido & "'"
                lOrigenVentas = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
                'GCAA 23032021
                lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
                lObligaOrigenVenta = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
                Call ActivarOrigenVentas
                
                'GCAA 23032021
                If lActivaMotorizado And RsCanalesVenta.RecordCount > 0 Then
                    Me.fraMorotizado.Visible = True
                Else
                    Me.fraMorotizado.Visible = False
                End If
                
           End If

      Case Is = 4 'Canal 4
           If RsDetalle.RecordCount > 0 Then
                MsgBox "No se puede cambiar el canal de venta", vbExclamation, sMensaje
                Exit Sub
           Else
                sTipoPedido = "04"
                sMotorizado = "0000"
                vOrigenVentas = ""

                LlenaColorCanalVenta Index
                
                'origen de ventas
                RsCanalesVenta.Filter = "tCodigoCanalVenta = '" & sTipoPedido & "'"
                lOrigenVentas = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
                'GCAA 23032021
                lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
                lObligaOrigenVenta = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
                
                Call ActivarOrigenVentas
                
                'GCAA 23032021
                If lActivaMotorizado And RsCanalesVenta.RecordCount > 0 Then
                    Me.fraMorotizado.Visible = True
                Else
                    Me.fraMorotizado.Visible = False
                End If

           End If
   
      Case Is = 3 'Para llevar
      
            'origen de ventas
            Me.fraOrigenVentas.Visible = False
            '--------------------------------
      
           If RsDetalle.RecordCount > 0 Then
                MsgBox "No se puede cambiar el canal de venta", vbExclamation, sMensaje
                Exit Sub
           Else
               sTipoPedido = "03"
               sMotorizado = "0000"
               vOrigenVentas = ""

               LlenaColorCanalVenta Index
               
                'origen de ventas
                RsCanalesVenta.Filter = "tCodigoCanalVenta = '" & sTipoPedido & "'"
                lOrigenVentas = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
                'GCAA 23032021
                lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
                lObligaOrigenVenta = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
                Call ActivarOrigenVentas
                'GCAA 23032021
                If lActivaMotorizado And RsCanalesVenta.RecordCount > 0 Then
                    Me.fraMorotizado.Visible = True
                Else
                    Me.fraMorotizado.Visible = False
                End If
           End If
      Case Is = 1 'En el Local
            
            'origen de ventas
            Me.fraOrigenVentas.Visible = False
            '--------------------------------
           If RsDetalle.RecordCount > 0 Then
                MsgBox "No se puede cambiar el canal de venta", vbExclamation, sMensaje
                Exit Sub
           Else
                sTipoPedido = "01"
                sMotorizado = "0000"
                vOrigenVentas = ""
                
                LlenaColorCanalVenta Index
                
                'origen de ventas
                RsCanalesVenta.Filter = "tCodigoCanalVenta = '" & sTipoPedido & "'"
                lOrigenVentas = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
                'GCAA 23032021
                lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
                lObligaOrigenVenta = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
                Call ActivarOrigenVentas

                'GCAA 23032021
                If lActivaMotorizado And RsCanalesVenta.RecordCount > 0 Then
                    Me.fraMorotizado.Visible = True
                Else
                    Me.fraMorotizado.Visible = False
                End If
           End If
                        
      Case Is = 2 'En Delivery
           If RsDetalle.RecordCount > 0 Then
                MsgBox "No se puede cambiar el canal de venta", vbExclamation, sMensaje
                Exit Sub
           Else
                sTipoPedido = "02"
                sMotorizado = "0000"
                vOrigenVentas = ""
 
                LlenaColorCanalVenta Index
                
                'origen de ventas
                RsCanalesVenta.Filter = "tCodigoCanalVenta = '" & sTipoPedido & "'"
                lOrigenVentas = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
                lObligaOrigenVenta = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
                'GCAA 23032021
'                lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
                Call ActivarOrigenVentas
           End If

           
      Case Is = 0 'Punto de Venta
            tabProducto.Visible = False
            fraPuntoVenta.Visible = True
            
      Case Is = 6
            frmPrograma.Show vbModal
            If wEnter = True Then
                txtFechaEntrega.Caption = sCodigo
                cmdCabeceraPicture(6).Visible = False
            Else
                txtFechaEntrega.Caption = ""
                cmdCabeceraPicture(6).Visible = True
            End If
            
      Case Is = 7
      
            If RsDetalle.RecordCount > 0 And Pedido = "" Then
                Exit Sub
            End If
      
            If lPasswordImportarPedido Then
               If Supervisor("15") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
            
            If lFiltroTipoPedido Then
                Isql = "select *, Caso = case vpedidoGrilla.tCaja when '" & sCaja & "' then 'Exportar' ELSE 'Importar' END " & _
                       "from vPedidoGrilla INNER JOIN dbo.MPEDIDO ON dbo.vPedidoGrilla.Codigo = dbo.MPEDIDO.tCodigoPedido " & _
                       "where MPEDIDO.tTipoPedido='" & sTipoPedidoPD & "' and tCodigoPedido not in (select distinct dbo.DPEDIDO.tcodigopedido FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido where  tEstadoPedido='01' and (tFacturado = 'F' or tFacturado='P')) and " & _
                       "vpedidoGrilla.tEstadoPedido ='01' and " & _
                       "(vpedidoGrilla.tCaja <>'" & sCaja & "' or (vpedidoGrilla.tCaja='" & sCaja & "' and len(ltrim(tCajaAnterior))<>0 )) " & _
                       "order by Mesa, vPedidoGrilla.tObservacion"
            Else
                If lMCPV Then
                    Isql = "select *, 'Importar' as Caso " & _
                           "from vPedidoGrilla INNER JOIN dbo.MPEDIDO ON dbo.vPedidoGrilla.Codigo = dbo.MPEDIDO.tCodigoPedido " & _
                           "where tCodigoPedido not in (select distinct dbo.DPEDIDO.tcodigopedido FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido where  tEstadoPedido='01' and (tFacturado = 'F' or tFacturado='P')) and " & _
                           "vpedidoGrilla.tEstadoPedido ='01' and " & _
                           "vpedidoGrilla.tUsuario <>'" & sUsuario & "' " & _
                           "order by Mesa, vPedidoGrilla.tObservacion"
                Else
                    Isql = "select *, Caso = case vpedidoGrilla.tCaja when '" & sCaja & "' then 'Exportar' ELSE 'Importar' END " & _
                           "from vPedidoGrilla INNER JOIN dbo.MPEDIDO ON dbo.vPedidoGrilla.Codigo = dbo.MPEDIDO.tCodigoPedido " & _
                           "where tCodigoPedido not in (select distinct dbo.DPEDIDO.tcodigopedido FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido where  tEstadoPedido='01' and (tFacturado = 'F' or tFacturado='P')) and " & _
                           "vpedidoGrilla.tEstadoPedido ='01' and " & _
                           "(vpedidoGrilla.tCaja <>'" & sCaja & "' or (vpedidoGrilla.tCaja='" & sCaja & "' and len(ltrim(tCajaAnterior))<>0 )) " & _
                           "order by Mesa, vPedidoGrilla.tObservacion"
                End If
            End If
            
            Call ConfGrilla(8, frmBusquedaRapida.grdGrilla, "Función", 2, "Caso", 800, 0, 0, "", _
                                                            "Caja", 2, "tCaja", 550, 2, 0, "", _
                                                            "Mesa", 2, "Mesa", 1000, 0, 0, "", _
                                                            "Observacion", 2, "tObservacion", 1800, 0, 0, "", _
                                                            "Pedido", 2, "Descripcion", 1000, 0, 0, "", _
                                                            "Monto", 2, "Suma", 900, 1, 0, "###,##0.00", _
                                                            Mesero, 2, "Mozo", 1100, 0, 0, "", _
                                                            "Usuario", 2, "tUsuario", 1100, 0, 0, "")
                                                            
            If lBuscaPedidoNumero = True Then
                frmBusquedaRapida.nPredeterm = 4
            Else
                frmBusquedaRapida.nPredeterm = 2
            End If
                
            
            frmBusquedaRapida.Show vbModal
            
            If wEnter Then
                                                     
               sPedido = sCodigo
               
               'Refrescar
               Dim RsRefresca As Recordset
               Set RsRefresca = Lib.OpenRecordset("select tTurno, tCaja, tCajaAnterior, tTurnoAnterior from MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
  
               If lMCPV Then
                  ActualizaPedido
                  Cn.Execute "update MPEDIDO set tUsuario = '" & sUsuario & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
                  LimpiarData
                  
               ElseIf RsRefresca!tTurno = "MOZO" Then 'Importar desde Mozos
                  If lInfhotel Then
                     Cn.Execute "update MPEDIDO set tPuntoVenta='" & sPuntoVentaInfhotel & "', tCaja = '" & sCaja & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
                  Else
                     Cn.Execute "update MPEDIDO set tCaja = '" & sCaja & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
                  End If
                  
                  If RsDetalle.RecordCount > 0 And Pedido <> sPedido Then
                     Exit Sub
                  End If
                  
                  If Pedido <> "" Then
                     Exit Sub
                  End If
                         
                  ActualizaPedido
                  Cn.Execute "update TCAJA set lRefresca=1 where tCaja='" & RsRefresca!tCaja & "'"
                  
                  Pedido = sPedido
                        
                  ImportarPedido Pedido
                  
                  txtTitulo.Caption = " Caja Rápida : " & sCaja & " " & Mesero & ": Sin " & Mesero & IIf(sPedido = "", "", " - Pedido:" & sPedido)
                  
               ElseIf RsRefresca!tTurno = sTurno And Not IsNull(RsRefresca!tTurnoAnterior) Then  'Exportar Mozo
                    
                  If sCaja = RsRefresca!tCajaAnterior Then
                        Exit Sub
                  End If
                  
                  If RsDetalle.RecordCount > 0 And Pedido <> sPedido Then
                        Exit Sub
                  End If

                  ActualizaPedido
                  Cn.Execute "update MPEDIDO set tTurno = '" & RsRefresca!tTurnoAnterior & "', tCaja='" & RsRefresca!tCajaAnterior & "' where tCodigoPedido='" & sPedido & "' "
               
                  LimpiarData
                    
               ElseIf RsRefresca!tTurno = sTurno And IsNull(RsRefresca!tTurnoAnterior) Then 'Exportar
                      
                  If sCaja = RsRefresca!tCajaAnterior Then
                        Exit Sub
                  End If
                  
                  If RsDetalle.RecordCount > 0 And Pedido <> sPedido Then
                      Exit Sub
                  End If
                  
                  ActualizaPedido
               
                  Cn.Execute "update TCAJA set lRefresca=1 where tCaja='" & RsRefresca!tCajaAnterior & "'"
                  Cn.Execute "update MPEDIDO set tTurno = 'MOZO', tCaja='" & RsRefresca!tCajaAnterior & "' where tCodigoPedido='" & sPedido & "' "
                  
                  LimpiarData

               Else  'Importar
                      If Calcular("select count(ddocumento.tDocumento) as Codigo from DDOCUMENTO inner join mdocumento on ddocumento.tdocumento= mdocumento.tdocumento where tCodigoPedido='" & sPedido & "' and mdocumento.testadodocumento<>'04'", Cn) > 0 Then
                         MsgBox "Error: No se puede importar pedido con Documentos", vbExclamation, sMensaje
                         Exit Sub
                      Else
                                 
                         If RsDetalle.RecordCount > 0 And Pedido <> sPedido Then
                            Exit Sub
                         End If
                         
                         If Pedido <> "" Then
                            Exit Sub
                         End If
                            
                         Cn.Execute "update MPEDIDO set tCaja = '" & sCaja & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
                         Pedido = sPedido
                        
                         ImportarPedido Pedido
                        
                      End If
                End If
            
            End If
            
    
   End Select
   cmdEtiqueta_Click (1)
   txtBarra.SetFocus
End Sub


Private Sub cargarDatosCabecera(Pedido As String)

    Dim RsCabeceraPedido As Recordset
    Dim xTipoPedido As String
    
        Isql = "SELECT M.tCodigoPedido,T.tCodigoDelivery,T.tNombre + ' ' + t.tApellido as cliente,TM.tDetallado,M.fProgramacion,M.tEntregarA, M.tObservacion, M.tTipoPedido, M.nDescuento, M.tMozo, M.tDescuento, M.tClienteDelivery, M.tMesa " & _
           ", M.CodigoOrigenVenta FROM MPEDIDO M LEFT JOIN TDELIVERY T ON M.tClienteDelivery = T.tCodigoDelivery " & _
           "LEFT JOIN TMESA TM ON M.tMesa = TM.tCodigoMesa " & _
           "WHERE M.tCodigoPedido = '" & Pedido & "'"
    Set RsCabeceraPedido = Lib.OpenRecordset(Isql, Cn)
    
    txtTelefono.Caption = IIf(IsNull(RsCabeceraPedido!tCodigoDelivery), "", RsCabeceraPedido!tCodigoDelivery)
    txtCliente.Caption = IIf(IsNull(RsCabeceraPedido!Cliente), "", RsCabeceraPedido!Cliente)
    txtObservacion.Caption = IIf(IsNull(RsCabeceraPedido!tObservacion), "", RsCabeceraPedido!tObservacion)
    txtEntregar.Caption = IIf(IsNull(RsCabeceraPedido!tentregara), "", RsCabeceraPedido!tentregara)
    txtFechaEntrega.Caption = IIf(IsNull(RsCabeceraPedido!fProgramacion), "", RsCabeceraPedido!fProgramacion)
    xTipoPedido = IIf(IsNull(RsCabeceraPedido!tTipoPedido), "", RsCabeceraPedido!tTipoPedido)
    sMesa = IIf(IsNull(RsCabeceraPedido!tMesa), "", RsCabeceraPedido!tMesa)
    xDescuento = IIf(IsNull(RsCabeceraPedido!nDescuento), 0, RsCabeceraPedido!nDescuento)
    sCodigoDescuento = IIf(IsNull(RsCabeceraPedido!tDescuento), 0, RsCabeceraPedido!tDescuento)
    sMozo = IIf(IsNull(RsCabeceraPedido!tMozo), "", RsCabeceraPedido!tMozo)
    sClienteFrecuente = IIf(IsNull(RsCabeceraPedido!tClienteDelivery), "", RsCabeceraPedido!tClienteDelivery)
    vOrigenVentas = IIf(IsNull(RsCabeceraPedido!tClienteDelivery), "", RsCabeceraPedido!CodigoOrigenVenta)
    
    Select Case xTipoPedido
        Case "01"
            sTipoPedido = "01"
            LlenaColorCanalVenta 1
        Case "02"
            sTipoPedido = "02"
            LlenaColorCanalVenta 2
        Case "03"
            sTipoPedido = "03"
            LlenaColorCanalVenta 3
        Case "04"
            sTipoPedido = "04"
            LlenaColorCanalVenta 4
        Case "05"
            sTipoPedido = "05"
            LlenaColorCanalVenta 5
    End Select
    
    
End Sub

Private Sub cmdCabeceraPicture_Click(Index As Integer)
    cmdCabecera_Click Index
End Sub

Private Sub cmdCombo_Click(Index As Integer)
   txtBarra.SetFocus
   Dim nPos As Integer
   Dim nOrd As Integer
   Select Case Index
          Case Is = 0 ' Salir
               fraCombo.Visible = False
               fraProductoCombo.Visible = False
               wAgregaCombo = False
               ActivaCabecera True
               
                If fraPropiedad.Visible = True Then
                  cmdOpcion_Click (6)
               End If
               AsignaProducto
               RsDetalle.Requery
          
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
                        tUsuActua = sUsuario
                     If Supervisor("02") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                     End If
                     sUsuarioAutoriza = sVar1
                     sUsuario = tUsuActua
                  End If
                                                                            
                  If lPrinter = False Or (lPrinter = True And RsCombo!lImprime) Then
                     'Impresion del Pedidos Anulados
                     Isql = "SELECT TPRODUCTO_1.tDetallado AS Producto, dbo.vSalon.tResumido + ' - ' + dbo.TMESA.tResumido AS Mesa, dbo.TPRODUCTOAREA.tArea, dbo.MPEDIDO.tTipoPedido AS TipoPedido, dbo.MPEDIDO.nAdulto, dbo.MPEDIDO.lPrioridad AS Prioridad, dbo.MPEDIDO.tObservacion AS Observacion, dbo.vMozo.Descripcion AS Mozo, dbo.CPEDIDO.nCantidad AS nCombo, dbo.CPEDIDO.tItem, dbo.CPEDIDO.tItemCombo, dbo.CPEDIDO.tObservacion AS tObservacionCombo, TPRODUCTO_2.tDetallado AS Combo, dbo.vDelivery.Cliente " & _
                            "FROM dbo.TPRODUCTO TPRODUCTO_2 LEFT OUTER JOIN dbo.TPRODUCTOAREA ON TPRODUCTO_2.tCodigoProducto = dbo.TPRODUCTOAREA.tCodigoProducto RIGHT OUTER JOIN dbo.TMESA LEFT OUTER JOIN dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo RIGHT OUTER JOIN dbo.vMozo RIGHT OUTER JOIN dbo.vDelivery RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vDelivery.Codigo = dbo.MPEDIDO.tClienteDelivery LEFT OUTER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo ON " & _
                            "dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.tMesa ON TPRODUCTO_2.tCodigoProducto = dbo.CPEDIDO.tProductoCombo LEFT OUTER JOIN dbo.TPRODUCTO TPRODUCTO_1 ON dbo.CPEDIDO.tProducto = TPRODUCTO_1.tCodigoProducto " & _
                            "Where dbo.CPEDIDO.lImprime = 1 And dbo.CPEDIDO.lImprimeArea = 1 and dbo.CPEDIDO.tCodigoPedido = '" & Pedido & "' and dbo.CPEDIDO.tItem ='" & RsCombo!tItem & "' and dbo.CPEDIDO.tItemCombo='" & RsCombo!tItemCombo & "' ORDER BY dbo.CPEDIDO.tItem"
              
                     Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                     Dim i As Integer
                     If RsImpresion.RecordCount = 0 Then
                        LimpiaRs
                     Else
                     
                        If Not RsImpresion.EOF Then
                           RsArea.MoveFirst
                           For i = 1 To RsArea.RecordCount
                               RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
                               If RsArea!tIcono = "" Or RsArea!nValor = 1 Then
                                  If RsImpresion.RecordCount <> 0 Then
                                     RsImpresion.MoveFirst
                                     sPedido = Pedido
                                     ImprimePedido RsImpresion, "A", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Rapido"
                                     sPedido = ""
                                  End If
                               End If
                               RsArea.MoveNext
                           Next i
                        End If
                      End If
                      LimpiaRs
                   End If
                        
                        'Oscar Ortega----------------------------------------------
                        Dim RstCombo2 As Recordset
                        Isql = "Select c.nCantidad, t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & Pedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & xItem & "'"
                         Isql = "Select t.nAumento From [" & sComboDetalle & "] As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tItem = '" & sItem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
                        Set RstCombo2 = Lib.OpenRecordset(Isql, Cn)
                        If RstCombo2.RecordCount > 0 Then
                                If IIf(IsNull(RstCombo2!nAumento), 0, RstCombo2!nAumento) > 0 Then
                                        txtMonto.Caption = Format(CambiaPrecio(nPVenta - ((RstCombo2!nAumento / nCantidad) * RsCombo!nCantidad)), "#,###,##0.00")
                                End If
                        End If
                        'Fin Oscar Ortega------------------------------------------
                        
                        'KDS2
                        If lKDS Then
                            Dim kdsRsCabecera As Recordset
                            Isql = "SELECT * From vPedidoCabecera Where Codigo = '" & Pedido & "' Order By codigo "
                            Set kdsRsCabecera = Lib.OpenRecordset(Isql, Cn)
                            Call KDS_EliminarProductoDeCombo(kdsRsCabecera, sItem, xItem)
                        End If
                        
                                  
                        
                        'insumoCOMBO2013
                        'INSUMOCRITICO23
                    Dim rstItems As New ADODB.Recordset
                    Set rstItems = New ADODB.Recordset
                    Set rstItems = Lib.OpenRecordset("SELECT     dbo.TPRODUCTO.tCodigoInsumo, " & sDetalle & ".nCantidad * " & sComboDetalle & ".nCantidad AS nCantidad FROM         " & sDetalle & " INNER JOIN     " & sComboDetalle & " ON   " & sDetalle & ".tItem = " & sComboDetalle & ".tItem INNER JOIN                       dbo.TPRODUCTO ON " & sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto WHERE      " & sComboDetalle & ".tItem ='" & sItem & "' and tItemCombo='" & xItem & "' and  (dbo.TPRODUCTO.lControlInsumoCritico = 1) AND (ISNULL(dbo.TPRODUCTO.tCodigoInsumo, '') <> '') AND (ISNULL(" & sComboDetalle & ".lImprime, 0) = 1) ", Cn)
                
                    If Not (rstItems.EOF Or rstItems.BOF) Then
                        modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
                    End If
                    
                    
                   Dim cMax As String
                   cMax = Calcular("select max(tItem) as Codigo from APEDIDO where tCodigoPedido='" & Pedido & "'", Cn)
                   cMax = Lib.Correlativo(cMax, 3)
                   Isql = "insert into APEDIDO (tCodigoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                          "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
                          "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, " & _
                          "tComanda, lImprime, tUsuario, fRegistro, tUsuarioAnulado, fRegistroAnulado, " & _
                          "tObservacion, tObservacionAnulado, tEstadoItem, lImprimeArea, tArea, tMotivoEliminacion, tTurnoAnulado,fDiaContable) " & _
                          "select '" & Pedido & "' as tCodigoPedido, '" & cMax & "' as tItem, cpedido.tProductocombo, cpedido.tCodigoGRupo, cpedido.tCodigoSubGrupo, " & _
                          "cpedido.nPrecioNeto, cpedido.nImpuesto1/cpedido.ncantidad, cpedido.nImpuesto2/cpedido.ncantidad, cpedido.nImpuesto3/cpedido.ncantidad, cpedido.nVenta/cpedido.ncantidad, " & _
                          "0, cpedido.nPrecioNeto, cpedido.nCantidad, cpedido.nImpuesto1, cpedido.nImpuesto2, cpedido.nImpuesto3, cpedido.nVenta, '', cpedido.lImprime, " & _
                          "'" & sUsuario & "' as tUsuario, dpedido.fregistro as fRegistro, " & _
                          "'" & sUsuarioAutoriza & "' as tUsuarioAnulado, getDate() as fRegistroAnulado, " & _
                          "'Anulado de Combo' as tObservacion, 'Anul. de Combo:' + t.tResumido as tObservacionAnulado, 'N', cpedido.lImprimeArea, '', '000', '" & sTurno & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
                          "from  dpedido inner join cpedido on dpedido.tCodigoPedido=cpedido.tcodigopedido and  dpedido.tItem = CPEDIDO.tItem inner join tproducto t on t.tcodigoproducto = dpedido.tCodigoProducto " & _
                          "where cpedido.tCodigoPedido = '" & Pedido & "' and cpedido.tItem ='" & sItem & "' and cpedido.tItemCombo='" & xItem & "'"
                    Cn.Execute Isql
    
                  
                        
                   Cn.Execute "DELETE " & sComboPropiedad & " where tItem ='" & sItem & "' and tItemCombo='" & xItem & "'"
                   Cn.Execute "DELETE " & sComboDetalle & " where tItem ='" & sItem & "' and tItemCombo='" & xItem & "'"
                
                   Cn.Execute "delete from CPEDIDO where tCodigoPedido ='" & Pedido & "' and tItem ='" & sItem & "' and tItemCombo='" & xItem & "'"
                   Cn.Execute "delete from TCOMBOPROPIEDAD where tCodigoPedido='" & Pedido & "' and tItem ='" & sItem & "' and tItemCombo='" & xItem & "'"
                   
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
                                  "FROM " & sComboDetalle & " WHERE tItem='" & sItem & "'", Cn)
               If nCCombo + 1 > nCombo * RsDetalle!nCantidad Then
                  MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
                  Exit Sub
               End If
                'OO----------------------------------------------
                Dim oRsProductoDeCombo As Recordset
                Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sCombo)
                If oRsProductoDeCombo.RecordCount > 0 Then
                    If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                         Dim nCantidadEnElCombo As Integer
                         nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sItem, oRsProductoDeCombo!tEtiqueta)
                         If nCantidadEnElCombo >= RsDetalle!nCantidad Then
                             MsgBox "Solo es permitido " & nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                             Exit Sub
                         End If
                    End If
                End If
                '----------------------------------------------------------
               nPos = RsCombo.AbsolutePosition
               Cn.Execute "update " & sComboDetalle & " set nCantidad = " & RsCombo!nCantidad + 1 & " where tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               'OO------------------------------------------------------------
               Isql = "Select t.nAumento From [" & sComboDetalle & "] As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tItem = '" & sItem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
               Dim RstCombo As Recordset
               Set RstCombo = Lib.OpenRecordset(Isql, Cn)
               If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                    txtMonto.Caption = Format(CambiaPrecio(nPVenta + RstCombo!nAumento / nCantidad), "#,###,##0.00")
               End If
               '--------------------------------------------------------
               RsCombo.Requery
               RsCombo.AbsolutePosition = nPos
               
          Case Is = 3 'Disminuir
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
          
               If grdCombo.Columns(2).Text > 1 Then
                  nPos = RsCombo.AbsolutePosition
                  Cn.Execute "update " & sComboDetalle & " set nCantidad = " & RsCombo!nCantidad - 1 & " where tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                  'Oscar Ortega------------------------------------------------------------
                  Isql = "Select t.nAumento From [" & sComboDetalle & "] As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & Pedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
                  Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                  If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                    txtMonto.Caption = Format(CambiaPrecio(nPVenta - RstCombo!nAumento / nCantidad), "#,###,##0.00")
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
               If wEnter And val(sDescrip) > 0 Then
                  nCantidad = val(sDescrip)
                  nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                                     "FROM " & sComboDetalle & " WHERE tItem='" & sItem & "'", Cn)
                  If nCCombo + nCantidad - RsCombo!nCantidad > nCombo * RsDetalle!nCantidad Then
                     MsgBox "La cantidad máxima de items para este producto es de " & nCombo, vbExclamation, sMensaje
                     nCantidad = 1
                     Exit Sub
                  End If
                  
                  'Oscar Ortega----------------------------------------------
                  Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sCombo)
                  If oRsProductoDeCombo.RecordCount > 0 Then
                     If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                         nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sItem, oRsProductoDeCombo!tEtiqueta)
                         If nCantidad > RsDetalle!nCantidad Then
                             MsgBox "Solo es permitido " & RsDetalle!nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                             nCantidad = 1
                             Exit Sub
                         End If
                     End If
                  End If
                  '----------------------------------------------------------
                  'OO----------------------------------------------
                  nCantidad = RsDetalle!nCantidad
                  Isql = "Select c.nCantidad, t.nAumento From [" & sComboDetalle & "] As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & xItem & "'"
                  Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                  If RstCombo.RecordCount > 0 Then
                    If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                        txtMonto.Caption = Format(CambiaPrecio(nPVenta - ((RstCombo!nAumento / nCantidad) * RstCombo!nCantidad)), "#,###,##0.00")
                        txtMonto.Caption = Format(CambiaPrecio(nPVenta + ((RstCombo!nAumento / nCantidad) * val(sDescrip))), "#,###,##0.00")
                    End If
                    
                  End If
                  '------------------------------------------
                  
                  nPos = RsDetalle.AbsolutePosition
                  Cn.Execute "update " & sComboDetalle & " set nCantidad = " & val(sDescrip) & " where tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
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
               fraProductoCombo.Visible = False
               fraPropiedad.Visible = True
               
               ListarOperadoresConFiltro (sCombo)
   
           Case Is = 6  'Orden +
                If RsCombo.RecordCount = 0 Then
                   Exit Sub
                End If
               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
                nPos = RsCombo.AbsolutePosition
                nOrd = IIf(IsNull(RsCombo!nOrden), 0, RsCombo!nOrden)
                Cn.Execute "update " & sComboDetalle & " set nOrden = " & nOrd + 1 & " where tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                RsCombo.Requery
                RsCombo.AbsolutePosition = nPos
                
           Case Is = 7  'Orden -
                If RsCombo.RecordCount = 0 Then
                   Exit Sub
                End If
                If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
                nPos = RsCombo.AbsolutePosition
                nOrd = IIf(IsNull(RsCombo!nOrden), 0, RsCombo!nOrden)
                If nOrd > 1 Then
                   Cn.Execute "update " & sComboDetalle & " set nOrden = " & nOrd - 1 & " where tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                   RsCombo.Requery
                   RsCombo.AbsolutePosition = nPos
                End If
               
          Case Is = 8  'Linea Corte
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
          
               nPos = RsCombo.AbsolutePosition
               If IIf(IsNull(RsCombo!lCorte), False, RsCombo!lCorte) Then
                  Cn.Execute "update  " & sComboDetalle & " set lCorte = 0 where tCodigoPedido='" & Pedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               Else
                  Cn.Execute "update  " & sComboDetalle & " set lCorte = 1 where tCodigoPedido='" & Pedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               End If
               RsCombo.Requery
               RsCombo.AbsolutePosition = nPos
               
   End Select
End Sub

Private Sub cmdComboPicture_Click(Index As Integer)
    If Index = 99 Then
        cmdCombo_Click 0
    Else
       cmdCombo_Click Index
    End If
    
End Sub

Private Sub cmdCortesia_Click()
    sTemp = ""
    Isql = "select * from vCortesia where lActivo = 1 Order by Descripcion"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                         
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
       sCortesia = sCodigo
       txtCortesia.Caption = sDescrip
       
       txtDImporte.Caption = "0.00"
       txtRImporte.Caption = "0.00"
       txtImpuesto1.Caption = "0.00"
       txtImpuesto2.Caption = "0.00"
       txtImpuesto3.Caption = "0.00"
       txtPVenta.Caption = "0.00"
       txtVenta.Caption = "0.00"
       txtDPorcentaje.Caption = "0.00"
       txtRPorcentaje.Caption = "0.00"
       
       nPBase = 0
       nRecargo = 0
       nDescuento = 0
       nPVenta = 0
       nImpuesto1 = 0
       nImpuesto2 = 0
       nImpuesto3 = 0
    Else
       sCortesia = ""
       txtCortesia.Caption = ""
    End If
    txtBarra.SetFocus
End Sub

Private Sub cmdDescuento_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Dscto. Monto
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     nDescuento = val(sDescrip)
                     txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
          Case Is = 1 ' Dscto. Porcentaje
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     txtDPorcentaje.Caption = Format(sDescrip, "###,###,###,##0.00")
                     nDescuento = nOficial * val(sDescrip) / 100
                     txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
          Case Is = 2 ' Recargo Monto
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     nRecargo = val(sDescrip)
                     txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
          Case Is = 3 ' Recargo Porcentaje
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     txtRPorcentaje.Caption = Format(sDescrip, "###,###,###,##0.00")
                     nRecargo = nOficial * val(sDescrip) / 100
                     txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
    End Select
    txtBarra.SetFocus
End Sub

Private Sub cmdDetalle_Click(Index As Integer)
   txtBarra.SetFocus
   wAgregaCombo = False
   fraCombo.Visible = False
   Select Case Index
          Case Is = 0 'Eliminar
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
               
                'GCAA 10/07/2023 - REST-012-C-2023
               If lBloqueoAnulaItemsPedido And RsDetalle!lImprime Then
                   MsgBox "Esta opcion se encuentra Bloqueada, favor de comunicar al Supervisor", vbExclamation, sMensaje
                   Exit Sub
               End If
               
               If MsgBox("Seguro de Eliminar el Producto?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                  sUsuarioAutoriza = sUsuario
                  If (lPassword And lPrinter = False) Or (lPassword And lPrinter And RsDetalle!lImprime) Then
                        tUsuActua = sUsuario
                     If Supervisor("02") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                     End If
                     sUsuario = tUsuActua
                     sUsuarioAutoriza = sVar1
                  End If
                                                                            
                  If (lElimina And lPrinter = False) Or (lElimina And lPrinter = True And RsDetalle!lImprime) Then
                     fraEliminacion.Visible = True
                     tabProducto.Visible = False
                     ActivaCabecera False
                     EliminarProductoCom = 1
                  Else
                     sCodigo = ""
                     sDescrip = ""
                     SoloEliminaItem
                  End If
               End If
                                                                                                                                                                                
          Case Is = 1 ' Cantidad
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
          
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               'GCAA
               If Calcular("select count(*) as codigo FROM [" & sComboDetalle & "] cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where isnull(tcombo.lfijo,0)=0 and titem='" & sItem & "' and isnull(tcombo.naumento,0)>0", Cn) > 0 Then
                    MsgBox ("No es posible modificar la cantidad!!! " & vbNewLine & "Hay al menos 1 item combo con aumenta precio no fijo que restringue el proceso!!!"), vbInformation
                    Exit Sub
                End If
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter And val(sDescrip) > 0 Then
                    'Oscar Ortega------------
                    Dim oRsDetalleProducto As Recordset
                    Set oRsDetalleProducto = ObtenerDetalleProducto(sItem)
                    
                    'CESAR ROTULADO
                    If oRsDetalleProducto!tCodigoEtiqueta <> "" Then
                       MsgBox "No es posible aplicar los cambios", vbCritical + vbInformation
                       Exit Sub
                    End If
                    
                    Dim Xse As Integer
                    Dim ix As Integer
                    
                    If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                       If verificaCantidadDeItemsCombos(sItem, oRsDetalleProducto!nCombinacion, val(sDescrip)) = False Then
                           MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                           Exit Sub
                        End If
                        
                        If oRsDetalleProducto.Fields("NCANTIDAD") > val(sDescrip) Then 'DISMINUYE
                           Dim nCantidadMax As Double
                           nCantidadMax = Obtener_CantidadMaximaDeUnicoEtiqueta(sItem, nCantidad)
                           If nCantidad > nCantidadMax Then
                              'nCantidad = Val(sDescrip)
                              'Cn.Execute "update [" & sDetalle & "] set nCantidad = " & Val(sDescrip) & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1 * " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2 * " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3 * " & nCantidad & " where tItem ='" & sitem & "'"
                              Xse = nCantidad - val(sDescrip)
                              For i = 1 To Xse
                                    nPos = RsDetalle.AbsolutePosition
                                    nCantidad = nCantidad - 1
                                    Set oRsDetalleProducto = ObtenerDetalleProducto(sItem)
                                    Dim DDcombo As String
                                    DDcombo = CambiaPrecioCombo((oRsDetalleProducto!nVenta - (oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento)) / nCantidad)
                              Next i
                           Else
                              MsgBox ("No puedes reducir la cantidad de combos con elementos únicos" & Chr(13) & "Disminuya primero la cantidad de productos dentro del combo"), vbExclamation
                              Exit Sub
                           End If
                           
                        Else 'AUMENTA
                           'nCantidad = Val(sDescrip)
                           'Cn.Execute "update [" & sDetalle & "] set nCantidad = " & Val(sDescrip) & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1 * " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2 * " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3 * " & nCantidad & " where tItem ='" & sitem & "'"
                           Xse = val(sDescrip) - nCantidad
                           For i = 1 To Xse
                                nPos = RsDetalle.AbsolutePosition
                                nCantidad = nCantidad + 1
                                Set oRsDetalleProducto = ObtenerDetalleProducto(sItem)
                                Dim AAcombo As String
                                AAcombo = CambiaPrecioCombo(((oRsDetalleProducto!nVenta - oRsDetalleProducto!nDescuento) + oRsDetalleProducto!nPrecioOficial) / nCantidad)
                           Next i
                        End If
                    Else
                        nCantidad = val(sDescrip)
                        nPos = RsDetalle.AbsolutePosition
                        Cn.Execute "update [" & sDetalle & "] set nCantidad = " & val(sDescrip) & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1 * " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2 * " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3 * " & nCantidad & " where tItem ='" & sItem & "'"
                    End If
                  
                    RsDetalle.Requery
                    RsDetalle.MoveFirst
                    RsDetalle.Find "tItem = '" & sItem & "'"
                    fxCombo "M", nCantidad, sProducto
                    nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
                                    
                    verificatitulo
               End If
               
          Case Is = 2 ' Aumentar
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
          
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               'GCAA
                If Calcular("select count(*) as codigo FROM [" & sComboDetalle & "] cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where isnull(tcombo.lfijo,0)=0 and titem='" & sItem & "' and isnull(tcombo.naumento,0)>0", Cn) > 0 Then
                    MsgBox ("No es posible modificar la cantidad!!! " & vbNewLine & "Hay al menos 1 item combo con aumenta precio no fijo que restringue el proceso!!!"), vbInformation
                    Exit Sub
                End If
               
               nCantidad = nCantidad + 1
               'Oscar Ortega Aumenta Combo---------------
               Set oRsDetalleProducto = ObtenerDetalleProducto(sItem)
               If (oRsDetalleProducto!lCombinacion) Then
                  If verificaCantidadDeItemsCombos(sItem, oRsDetalleProducto!nCombinacion, nCantidad) = False Then
                     MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                     Exit Sub
                  End If
                  'txtMonto.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad - 1))
                  'Cn.Execute "update [" & sDetalle & "] set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tItem ='" & sitem & "'"
                  Dim Acombo As String
                  Acombo = CambiaPrecioCombo(((oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento) + oRsDetalleProducto!nVenta) / nCantidad)
               Else
                  'CESAR ROTULADO
                  If oRsDetalleProducto!tCodigoEtiqueta <> "" Then
                       MsgBox "No es posible aplicar los cambios", vbCritical + vbInformation
                    Else
                  Cn.Execute "update [" & sDetalle & "] set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tItem ='" & sItem & "'"
                  End If
                  
               End If
               RsDetalle.Requery
               RsDetalle.MoveFirst
               RsDetalle.Find "tItem = '" & sItem & "'"
               fxCombo "M", nCantidad, sProducto
               nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
               verificatitulo
                    
          Case Is = 3 ' Disminuir
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
          
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               'GCAA
                If Calcular("select count(*) as codigo FROM [" & sComboDetalle & "] cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where isnull(tcombo.lfijo,0)=0 and titem='" & sItem & "' and isnull(tcombo.naumento,0)>0 ", Cn) > 0 Then
                    MsgBox ("No es posible modificar la cantidad!!! " & vbNewLine & "Hay al menos 1 item combo con aumenta precio no fijo que restringue el proceso!!!"), vbInformation
                    Exit Sub
                End If
                
          
               If grdDetalle.Columns(4).Text > 1 Then
               
                  nCantidad = nCantidad - 1
                  'Oscar Ortega Disminuir Combo---------------
                  Set oRsDetalleProducto = ObtenerDetalleProducto(sItem)
                  If (oRsDetalleProducto!lCombinacion) Then
                            If verificaCantidadDeItemsCombos(sItem, oRsDetalleProducto!nCombinacion, nCantidad) = False Then
                                MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                                Exit Sub
                            End If
                            
                            'Dim nCantidadMax As Double
                            nCantidadMax = Obtener_CantidadMaximaDeUnicoEtiqueta(sItem, nCantidad + 1)
                            If nCantidad + 1 > nCantidadMax Then
                               'Disminuir Combo
                               'txtMonto.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad + 1))
                               'Cn.Execute "update [" & sDetalle & "] set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tItem ='" & sitem & "'"
                               Dim Dcombo As String
                               Dcombo = CambiaPrecioCombo((oRsDetalleProducto!nVenta - (oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento)) / nCantidad)
                            Else
                                MsgBox ("El combo tiene demasiados elementos únicos.")
                                nCantidad = nCantidad + 1
                                Exit Sub
                            End If
                  Else
                        'CESAR ROTULADO
                        If oRsDetalleProducto!tCodigoEtiqueta <> "" Then
                           MsgBox "No es posible aplicar los cambios", vbCritical + vbInformation
                        Else
                           Cn.Execute "update [" & sDetalle & "] set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tItem ='" & sItem & "'"
                        End If
                  
                  End If
                  RsDetalle.Requery
                  RsDetalle.MoveFirst
                  RsDetalle.Find "tItem = '" & sItem & "'"
                  fxCombo "M", nCantidad, sProducto
                  nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
                  verificatitulo
               End If
               
          Case Is = 4 ' Precios
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
                tUsuActua = sUsuario
               If Supervisor("03") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               sUsuario = tUsuActua
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               tabProducto.Visible = False
               fraDetalle.Visible = True
               ActivaCabecera False
               
          Case Is = 5 ' Propiedades
          
            'origen de ventas
            Me.fraOrigenVentas.Visible = False
            '--------------------------------
            
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               ActivaCabecera False
               tabProducto.Visible = False
               ListarOperadoresConFiltro (sProducto)
               AsignaPropiedad
               fraPropiedad.Visible = True
          
          Case Is = 6 ' Observacion
               
                'origen de ventas
                Me.fraOrigenVentas.Visible = False
                '--------------------------------
               
               
               frmKeyBoard.Caption = "Nombre / Observación"
               frmKeyBoard.txtResultado.Text = sObser
               frmKeyBoard.Show vbModal
               
               If wEnter = True Then
                  sObser = sDescrip
                  txtObservacion.Caption = sObser
               End If
                    
          Case Is = 7 ' Visualizacion de Pedido
               If RsDetalle.RecordCount = 0 Then
                  MsgBox "No existen Datos a Visualizar", vbExclamation, sMensaje
                  Exit Sub
               End If
               sTipo = "CajaRapida"
               frmPedido.Show vbModal
               If wEnter Then
                  RsDetalle.Requery
               End If
          
          Case Is = 8
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
          
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               nPos = RsDetalle.AbsolutePosition
               If IIf(IsNull(RsDetalle!lCorte), False, RsDetalle!lCorte) Then
                  Cn.Execute "update " & sDetalle & " set lCorte = 0 where tCodigoPedido='" & Pedido & "' and tItem ='" & sItem & "'"
               Else
                  Cn.Execute "update " & sDetalle & " set lCorte = 1 where tCodigoPedido='" & Pedido & "' and tItem ='" & sItem & "'"
               End If
               RsDetalle.Requery
               RsDetalle.AbsolutePosition = nPos
          
          Case Is = 9 ' Mozos
          
                'origen de ventas
                Me.fraOrigenVentas.Visible = False
                '--------------------------------
          
               tabProducto.Visible = False
               fraMozo.Visible = True
               Me.fraMozo.Caption = Mesero
          
          Case Is = 10  'cuenta Corriente
          
            If RsDetalle.RecordCount = 0 Then
               Exit Sub
            End If
            
            ' Miguel Mendoza - 10062019
            If Not VerificaComboMaximo() Then: RsDetalle.Requery: Exit Sub
            ' *************************
            
               'VALIDACION CANAL DE VENTA
                Dim rsCanalVentas As Recordset
                Dim lObligaMozo As Boolean
                Dim lObligaMotorizado As Boolean
                Dim lObligaClienteFrecuente As Boolean
                Dim lObligaFechaEntrega As Boolean
                Dim lObligaEntregarA As Boolean
                
                Set rsCanalVentas = Lib.OpenRecordset("select * from vTipoPedido", Cn)
                rsCanalVentas.Filter = "Codigo = '" & sTipoPedido & "'"
                
                lObligaMozo = IIf(IsNull(rsCanalVentas!lObligaMozo), False, rsCanalVentas!lObligaMozo)
                lObligaMotorizado = IIf(IsNull(rsCanalVentas!lObligaMotorizado), False, rsCanalVentas!lObligaMotorizado)
                lObligaClienteFrecuente = IIf(IsNull(rsCanalVentas!lObligaClienteFrecuente), False, rsCanalVentas!lObligaClienteFrecuente)
                lObligaFechaEntrega = IIf(IsNull(rsCanalVentas!lObligaIngresoFechaEntrega), False, rsCanalVentas!lObligaIngresoFechaEntrega)
                lObligaEntregarA = IIf(IsNull(rsCanalVentas!lObligaEntregarA), False, rsCanalVentas!lObligaEntregarA)
   
               'Obligatoriedad de Mozo
               If lObligaMozo Then
                  If lMCPV Then
                      sMozo = ObtenerCodigoMozo(sVar1)
                      If sMozo = "" Then
                         MsgBox "Configure un Cajero-Mesero", vbExclamation, sMensaje
                         RsDetalle.Requery
                         Exit Sub
                      End If
                  Else
                      If sMozo = "" Or sMozo = "0000" Then
                         MsgBox "Asigne al " & Mesero, vbExclamation, sMensaje
                         RsDetalle.Requery
                         Exit Sub
                      End If
                  End If
               End If
               
               'Obligatoriedad de Motorizado
'               If lObligaMotorizado Then
'                  If sMotorizado = "" Or sMotorizado = "0000" Then
'                     MsgBox "Asigne al Motorizado", vbExclamation, sMensaje
'                     Exit Sub
'                  End If
'               End If
               
               'Obligatoriedad de Mesa
'               If lObligaMesa And sMesa = "" And txtObservacion.Caption = "" Then
'                  MsgBox "Asigne una Mesa", vbExclamation, sMensaje
'                  cmdCabecera_Click (13)
'                  Exit Sub
'               End If
               
               'Obligatoriedad de Cliente Frecuente
               If sClienteFrecuente = "" And lObligaClienteFrecuente Then
                  MsgBox "Asigne el Cliente Delivery", vbExclamation, sMensaje
                  cmdOpcion_Click (9)
                  RsDetalle.Requery
                  Exit Sub
               End If
               
               'Obligatoriedad de Fecha de Entrega
               If Me.txtFechaEntrega.Caption = "" And lObligaFechaEntrega Then
                  MsgBox "Asigne la Fecha de Entrega", vbExclamation, sMensaje
                  cmdCabecera_Click (6)
                  RsDetalle.Requery
                  Exit Sub
               End If
               
               'Entregar A
               If lObligaEntregarA = True And Me.txtEntregar.Caption = "" Then
                  MsgBox "Asigne información en Entregar A", vbExclamation, sMensaje
                  cmdDetalle_Click (14)
                  RsDetalle.Requery
                  Exit Sub
               End If
               
            
            If lMCPV Then
                If Not ValidaExistenciaProducto() Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsDetalle.Requery
                    Exit Sub
                End If
            End If
            variableEmite = False
                   
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=0", Cn)
               If i > 0 Then
                  cmdOpcion_Click (8)
                        'insumocritico
                           If variableEmite = False Then: RsDetalle.Requery: Exit Sub
                        'insumocritico
                   
               End If
            End If
                        
                   'insumocritico
                        variableEmite = False
                   'insumocritico
                        
            If Calcular("select count(tFacturado) as Codigo from " & sDetalle & " where isnull(tFacturado,'0') <> '0' and len(ltrim(tFacturado)) <> 0", Cn) > 0 Then
               MsgBox "Imposible pasar el pedido a Cuenta Corrientes, pedidos con items Facturados", vbExclamation, sMensaje
               RsDetalle.Requery
               Exit Sub
            End If
            tUsuActua = sUsuario
            sUsuarioAutoriza = sUsuario
            If Supervisor("09") = False Then
               MsgBox "Clave no permitida", vbExclamation, sMensaje
               RsDetalle.Requery
               Exit Sub
            End If
            sUsuarioAutoriza = sVar1
            sUsuario = tUsuActua
            'Chequea si existe platos a facturar
            sTD = "N"
            RsDetalle.MoveFirst
            Do While Not RsDetalle.EOF
               If (Len(Trim(RsDetalle!tFacturado)) = 0 Or IsNull(RsDetalle!tFacturado)) Then
                  sTD = "S"
                  Exit Do
               End If
               RsDetalle.MoveNext
            Loop
    
            If sTD <> "S" Then
               MsgBox "Error: No existen Productos a Facturar", vbCritical, sMensaje
               RsDetalle.Requery
               Exit Sub
            End If
                                                                        
            sTemp = ""
            Isql = "select * from vCompania where lActivo = 1 order by Descripcion"
            Call ConfGrilla(6, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 800, 2, 0, "", _
                                                            "Identidad", 2, "Identidad", 1500, 0, 0, "", _
                                                            "Cliente", 2, "Descripcion", 2800, 0, 0, "", _
                                                            "Linea", 2, "nLinea", 1050, 1, 0, "##,##0.00", _
                                                            "Consumo", 2, "nConsumo", 1050, 1, 0, "##,##0.00", _
                                                            "Fecha Venc", 2, "fFechaVence", 1050, 0, 0, "dd/mm/yyyy")
            frmBusquedaRapida.nPredeterm = 2
            frmBusquedaRapida.Show vbModal
                                                                         
            If Not wEnter Or sCodigo = "" Then
                RsDetalle.Requery
               Exit Sub
            End If

            sCliente = sCodigo
            xSuma = Calcular("select sum(nVenta) as Codigo FROM " & sDetalle & " where tEstadoItem = 'N' and isnull(tFacturado,'') = ''", Cn)
            
            Dim vley1 As Double
            vley1 = CDbl(sValorLey1) * (Calcular(" select isnull(sum(" & sDetalle & ".ncantidad),0) as codigo from " & sDetalle & " inner join vproducto on " & sDetalle & ".tcodigoproducto=vproducto.codigo where vproducto.lley1=1 and " & sDetalle & ".tEstadoItem = 'N' and isnull(" & sDetalle & ".tFacturado,'') = '' ", Cn))
            xSuma = xSuma + vley1
            
            'Validacion de escoger segun estadoFrecuente
            Dim lValidaEstado As Boolean
            lValidaEstado = False
            lValidaEstado = Calcular("select ISNULL(tb.nValor,0) as codigo from TDELIVERY t INNER JOIN TTABLA tb on t.tEstadoFrecuente = tb.TCODIGO where  t.tCodigoDelivery='" & sCliente & "' and tb.TTABLA='ESTADOFRECUENTE'", Cn)
            If lValidaEstado Then
                MsgBox "No es posible seleccionar al cliente, estado no permitido", vbCritical, sMensaje
                RsDetalle.Requery
                Exit Sub
            End If
            
            
            'centralizada
            Dim xLinea As Double
            Dim xConsumo As Double
            
'            If lCentral = False Then
'                    xLinea = Calcular("select nLinea as Codigo FROM TCOMPANIA where tCodigoCliente = '" & sCliente & "'", Cn)
'                    xConsumo = Calcular("select nConsumo as Codigo FROM TCOMPANIA where tCodigoCliente = '" & sCliente & "'", Cn)
                    xLinea = Calcular("select nLinea as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", Cn)
                    xConsumo = Calcular("select nConsumo as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", Cn)
'            Else
'                Dim conServidor As ADODB.Connection
'                Set conServidor = devuelveConexionCentral(sServidorCentral, bdInforestCentral)
'                If conServidor.State Then
''                       xLinea = Calcular("select isnull(nLinea,0) as Codigo FROM TCOMPANIA where tCodigoCliente = '" & sCliente & "'", conServidor)
''                       xConsumo = Calcular("select isnull(nConsumo,0) as Codigo FROM TCOMPANIA where tCodigoCliente = '" & sCliente & "'", conServidor)
'                        xLinea = Calcular("select isnull(nLinea,0) as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", conServidor)
'                        xConsumo = Calcular("select isnull(nConsumo,0) as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", conServidor)
'                 Else
'                        MsgBox "No es posible conectar con el servidor central" & vbCrLf & "No se puede trabajar con la cuenta corriente", vbCritical, sMensaje
'                        Exit Sub
'
'                End If
'
'
'            End If
            
            If xSuma > xLinea - xConsumo Then
               MsgBox "El Cliente no tiene linea suficiente " & Chr(13) & _
                      "Linea : " & Format(xLinea, "###,##0.00") & "  Consumo : " & Format(xConsumo, "###,##0.00") & Chr(13) & _
                      "Saldo : " & Format(xLinea - xConsumo, "###,##0.00"), vbCritical, sMensaje
                RsDetalle.Requery
               Exit Sub
            End If

            If MsgBox("Esta seguro de Enviar el Pedido Nro: " & Pedido & _
               Chr(13) & "a Cuentas Corrientes al Cliente " & sDescrip & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
               RsDetalle.Requery
               Exit Sub
            End If
                        
            Cn.Execute "Update MPEDIDO set tClienteCtaCte ='" & sCliente & "', tEstadoPedido = '04' where tCodigoPedido='" & Pedido & "'"
'            If lCentral = False Then
            
                Cn.Execute "Update TDELIVERY set nConsumo = " & xConsumo + xSuma & " where TCODIGODELIVERY ='" & sCliente & "'"
'            Else
'                conServidor.Execute "Update TDELIVERY set nConsumo = " & xConsumo + xSuma & " where TCODIGODELIVERY ='" & sCliente & "'"
'            End If
            
            Isql = "select * from vCtaCte " & _
                   "WHERE Codigo='" & Pedido & "'"
            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
            
            If RsImpresion.RecordCount = 0 Then
               LimpiaRs
               MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
            Else
                'GCAA REQ REST-C009-2021 - PASTELERIA TENTACIONES
               If (lActivarMensajeCtsCorrientes) Then
                     If MsgBox("Desea Imprimir? ", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                       
                     Else
                        ImprimeCtaCte RsImpresion
                     End If
                Else
                    ImprimeCtaCte RsImpresion
                End If
            End If
            
            LimpiaRs
                        
            Cn.Execute "delete " & sDetalle
            Cn.Execute "delete " & sComboDetalle
            Cn.Execute "delete " & sComboPropiedad
            Cn.Execute "delete " & sProductoPropiedad
            
            RsDetalle.Requery
            RsComboPropiedad.Requery
            RsProductoPropiedad.Requery
            Inicializar
            Screen.MousePointer = vbDefault
                                                                           
         Case Is = 12 '
                Dim wCalcula As Boolean
                Dim sDescripcionDescuento  As String
                Dim ltope As Boolean
                Dim procedeDescuento As Boolean
              
                If RsDetalle.RecordCount > 0 Then
                         tUsuActua = sUsuario
                         If Supervisor("10") = False Then
                            MsgBox "Clave no permitida", vbExclamation, sMensaje
                            Exit Sub
                         End If
                         sUsuario = tUsuActua
                         sUsuarioAutoriza = sVar1
                         tAutorizaDescuento = sUsuarioAutoriza
                         sTemp = ""
                         
                         Isql = "SELECT Codigo, LTRIM(RTRIM(Descripcion)) as Descripcion, case lRatio when 1 then nRatio else 0 END as nRatio, case lRatio when 0 then nRatio else 0 END as nMonto FROM vMotivoDescuento WHERE lActivo = '1' and AplicaAnticipo=0 ORDER BY Descripcion"
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
                            
                            If sCodigo = "000" Then
                                If xDescuento > 99 Then
                                   xDescuento = 0
                                   MsgBox "Porcentaje no válido", vbCritical, sMensaje
                                   Exit Sub
                                 End If
                      
                               frmKeyBoard.Caption = "Descripcion del Descuento"
                               frmKeyBoard.Show vbModal
                               sDescripcionDescuento = sDescrip
                            End If
                         Else
                            Exit Sub
                         End If
                         
                         Dim SumTotalPedido As Double
                         If Pedido = "" Then
                            SumTotalPedido = Calcular("select sum(d.nventa) as codigo from " & sDetalle & " d inner join TPRODUCTO p on d.tCodigoProducto = p.tCodigoProducto where p.lDescuento = 1", Cn)
                         Else
                            SumTotalPedido = Calcular("select sum(d.nventa) as codigo from DPEDIDO d inner join TPRODUCTO p on d.tCodigoProducto = p.tCodigoProducto where d.tCodigoPedido='" & Pedido & "' and p.lDescuento = 1", Cn)
                         End If
                         
                         'CDbl(txtMonto.Caption)
                         If Not RsDesc!lRatio And (RsDesc!nRatio > SumTotalPedido) Then
                            sCodigoDescuento = ""
                            xDescuento = 0
                            MsgBox "Descuento mayor al Pedido", vbCritical, sMensaje
                            Exit Sub
                         End If
                         
                         If RsDesc!lBloqueo Then
                            sTipo = "Prepintado"
                            sCodigo = xDescuento
                            frmNumPad.Show vbModal
                         Else
                            wEnter = False
                         End If
                                        
                         If wEnter Then
                            xDescuento = val(sDescrip)
                            
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
                         End If
                         
                         CalculaDescuento
                         RsDetalle.Requery
                         nMonto = Calcular("select sum(nventa) as codigo from " & sDetalle & "", Cn)
                         VisualizaMonto
                End If
         
         Case Is = 13  ' ofertaaaaaaaaaaaaa dic 2010
               tUsuActua = sUsuario
               If Supervisor("10") = False Then
                   MsgBox "Clave no permitida", vbExclamation, sMensaje
                   Exit Sub
                End If
                sUsuario = tUsuActua
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
                    Select Case pais ' ok
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
                    
                    Isql = "Update " & sDetalle & " Set nPrecioNeto = " & nPBase & ", " & _
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
                    "where tItem = '" & sItem & "'"
                    Cn.Execute Isql
                    
                    nPos = RsDetalle.AbsolutePosition
                    RsDetalle.Requery
                    RsDetalle.AbsolutePosition = nPos
                   
                    nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
                   
                End If
         '====================
         Case Is = 14 'entregar a
                frmKeyBoard.Caption = "Entregar A"
                frmKeyBoard.txtResultado.Text = txtEntregar.Caption
                frmKeyBoard.Show vbModal
                If wEnter = True Then
                    txtEntregar.Caption = sDescrip
                End If
                              
   End Select
   
End Sub

Private Sub cmdDirecto_Click(Index As Integer)
    sTemp = ""
    Isql = "select * from vProducto where lActivo = 1 and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " Order by Descripcion"
    Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                    "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                    "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                    "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                    "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    
    If wEnter Then
       sProducto = sCodigo
       
        'INSUMOCRITICO23
        If validadIngresoProducto(sProducto) = False Then
            Exit Sub
        End If
      'INSUMOCRITICO23
    
       Dim xxx As String
       xxx = RsProducto.Filter
       RsProducto.Filter = adFilterNone
       RsProducto.MoveFirst
       RsProducto.Find ("Codigo='" & sProducto & "'")
    
       If Not RsProducto.EOF() Then
          If wAgregaCombo Then
             nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                               "FROM " & sComboDetalle & " WHERE tItem='" & sItem & "'", Cn)
        
             If nCCombo < nCombo * RsDetalle!nCantidad Then
                InsertaCombo sProducto
             Else
                MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
             End If
          Else
             If lBal And RsProducto!lBalanza Then
                Dim nResultado As Double
                nResultado = Pesar(nBalanzaPuerto)
                nResultado = Format(nResultado, "#,##0.00")
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
          End If
       End If
       RsProducto.Filter = IIf(xxx = "0", "", xxx)
    End If
    txtBarra.SetFocus
End Sub

Private Sub cmdDetallePicture_Click(Index As Integer)
    cmdDetalle_Click Index
End Sub

Private Sub cmdEtiqueta_Click(Index As Integer)
    sPrefijo = Index
    RsCajaRapida.Filter = "Prefijo='" & sPrefijo & "'"
    RsCajaRapida.MoveFirst
    
    cmdEtiqueta(1).FontBold = IIf(Index = 1, True, False)
    cmdEtiqueta(1).backColor = IIf(Index = 1, &H8BDD79, &HF2FAED)
    
    cmdEtiqueta(2).FontBold = IIf(Index = 2, True, False)
    cmdEtiqueta(2).backColor = IIf(Index = 2, &H8BDD79, &HF2FAED)
    
    cmdEtiqueta(3).FontBold = IIf(Index = 3, True, False)
    cmdEtiqueta(3).backColor = IIf(Index = 3, &H8BDD79, &HF2FAED)
    
    For i = 1 To 9
        
        'cmdAgrupacion(i).backColor = IIf(RsCajaRapida!nValor = 0, -2147483633, RsCajaRapida!nValor)
        'cmdAgrupacion(i).backColor = IIf(RsCajaRapida!nValor = 0, -2147483633, RsCajaRapida!nValor)
        
        If LTrim(RsCajaRapida!tDetallado) = "" Then
           'cmdAgrupacion(i).Caption = "(no utilizado)"
           cmdAgrupacion(i).Visible = False
        Else
           cmdAgrupacion(i).Caption = RsCajaRapida!tDetallado
           cmdAgrupacion(i).Visible = True
        End If
        RsCajaRapida.MoveNext
    Next i
    cmdAgrupacion_Click (1)
    If txtBarra.Visible = True Then
        txtBarra.SetFocus
    End If
End Sub

Private Sub cmdImpuesto_Click(Index As Integer)
    Select Case Index
        Case Is = 0
             nImpuesto1 = IIf(nImpuesto1 = 0, nPBase * nPorcentaje1 / 100, 0)
             txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
        Case Is = 1
             nImpuesto2 = IIf(nImpuesto2 = 0, nPBase * nPorcentaje2 / 100, 0)
             txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
        Case Is = 2
             nImpuesto3 = IIf(nImpuesto3 = 0, nPBase * nPorcentaje3 / 100, 0)
             txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
    End Select
    nPVenta = nPBase + nImpuesto1 + nImpuesto2 + nImpuesto3
    txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
    txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
    txtBarra.SetFocus
End Sub

Private Sub cmdMotorizado_Click(Index As Integer)
   'origen de ventas
    'HabilitaTimerColor (False)
    
   RsMotorizado.MoveFirst
   RsMotorizado.Find "nboton = " & Trim(str(Index))
   'frmVenta.txtMotorizado.Caption = RsMotorizado!descripcion
   sMotorizado = RsMotorizado!codigo
   Me.fraMorotizado.Visible = False
   'HabilitaTimerColor (True)
End Sub
Private Sub cmdMozo_Click(Index As Integer)
   RsMozo.MoveFirst
   RsMozo.Find "nboton = " & Trim(str(Index))
   txtTitulo.Caption = " Caja Rápida : " & sCaja & " " & Mesero & ": " & RsMozo!descripcion & IIf(sPedido = "", "", " - " & sPedido)
   sMozo = RsMozo!codigo
   tabProducto.Visible = True
   fraMozo.Visible = False
   txtBarra.SetFocus
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 1 'Primero
                MoverPuntero Primero, grdDetalle
           Case Is = 2 'PgUp
                MoverPuntero pgup, grdDetalle
           Case Is = 3 'Previo
                MoverPuntero previo, grdDetalle
           Case Is = 4 'Siguiente
                MoverPuntero siguiente, grdDetalle
           Case Is = 5 'PgDn
                MoverPuntero pgdn, grdDetalle
           Case Is = 6 'Ultimo
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
    txtBarra.SetFocus
End Sub

Private Sub cmdNavegarPicture_Click(Index As Integer)
    cmdNavegar_Click Index
End Sub


Private Sub cmdOpcion_Click(Index As Integer)
   txtBarra.SetFocus
   Select Case Index
   
          Case Is = 0  'Anulacion Documento
          
                'anulacion por nota de credito
                Isql = "SELECT * FROM TPARAMETRO"
                Set RsTparametro = Lib.OpenRecordset(Isql, Cn)
                '------------------------------------------------------------------------------
                   'anulacion de documentos por nota de credito
                
                If RsTparametro!lanula = True Then
                    Sw = True
                       If Supervisor("27") = False Then
                          MsgBox "Clave no permitida", vbExclamation, sMensaje
                          Exit Sub
                       End If
                      modProcedimiento.aNotaCredito ("documento")
                      frmNotaCreditoDetalle.Show vbModal
                   Exit Sub
                End If
                '--------------------------------------------
         
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=1", Cn)
               If i > 0 Then
                  MsgBox "El pedido actual esta activo", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If Pedido <> "" Then
                  MsgBox "El pedido actual esta activo, cancele y vuelva a entrar", vbExclamation, sMensaje
                  Exit Sub
               End If
               tUsuActua = sUsuario
               If Supervisor("05") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               sUsuario = tUsuActua
               Isql = "SELECT MAX(dbo.DDOCUMENTO.tCodigoPedido) AS Descripcion, dbo.MDOCUMENTO.tDocumento AS Codigo, MAX(dbo.MDOCUMENTO.fRegistro) AS fFecha, dbo.MDOCUMENTO.tUsuario, dbo.TCLIENTE.tEmpresa AS Cliente, dbo.MDOCUMENTO.nVenta, MAX(dbo.MPEDIDO.tObservacion) As tObservacion " & _
                      "FROM dbo.MPEDIDO RIGHT OUTER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido RIGHT OUTER JOIN dbo.MDOCUMENTO LEFT OUTER JOIN dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
                      "where dbo.MDOCUMENTO.tTurno='" & sTurno & "' and tEstadoDocumento = '02' " & _
                      "GROUP BY dbo.MDOCUMENTO.tDocumento, dbo.TCLIENTE.tEmpresa, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.tUsuario " & _
                      "ORDER BY dbo.MDOCUMENTO.tDocumento"
                      
               Call ConfGrilla(6, frmBusquedaRapida.grdGrilla, "Documento", 2, "Codigo", 1500, 0, 0, "", _
                                                               "Fec.Emis", 2, "fFecha", 1000, 0, 0, "", _
                                                               "Monto", 2, "nVenta", 1000, 1, 0, "###,###,##0.00", _
                                                               "Observacion", 2, "tObservacion", 2000, 0, 0, "", _
                                                               "Cliente", 2, "Cliente", 1700, 0, 0, "", _
                                                               "Usu.Emis", 2, "tUsuario", 1000, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 0
               frmBusquedaRapida.Show vbModal
               
               If wEnter Then
               
                If lFEBiz Then
                    Dim DiaDoc As Long
                    DiaDoc = Calcular("SELECT DATEDIFF(DAY,fRegistro,GETDATE()) as codigo FROM MDOCUMENTO where tdocumento='" & sCodigo & "'", Cn)
                    If DiaDoc > 7 Then
                        MsgBox "No se puede eliminar el documento,Tiempo maximo de eliminacion superada: 7 Dias, Emita una Nota de Credito", vbInformation, sMensaje
                        Exit Sub
                    End If
                End If
               
               
                  Dim lContinua As Boolean
                  lContinua = True
                  'Pin Pad
                  Dim RsPinPad As Recordset
                  Set RsPinPad = Lib.OpenRecordset("select nMonto, tReferencia from DPAGOTARJETA where tDocumento='" & sCodigo & "'", Cn)
                  If RsPinPad.RecordCount > 0 Then
                     Dim sMonto As String
                     lContinua = False
                     sMonto = Format(str(RsPinPad!nMonto), "0000000000.00")
                     sMonto = Mid(sMonto, 1, 10) & Mid(sMonto, 12, 2)
                     sOperacion = OP_FINANCIERA & "A" & sMonto & Chr$(FS) & _
                                                  "B" & "000000000000" & Chr$(FS) & _
                                                  "C" & "0" & Chr$(FS) & _
                                                  "D" & sEmpresa & Chr$(FS) & _
                                                  "E" & sCaja
                     nRet = fiStartOperation(sOperacion, 2, sMensaje)
                                                                        
                     If nRet = RET_OK Or nRet = RET_RUNNING Then
                        If Not Imprimir(sPreCuenta) Then
                           Exit Sub
                        End If
                        Printer.FontName = sFont
                        Printer.FontBold = False
                        sClave = ""
                        nContador = 0
                        lEmisor = True
                        lLoop = True
                        Do
                          sRetorno = ""
                          nRet = fiGetStatus(sRetorno, 512)
                          lEmisor = ImprimeCabecera(sRetorno, lEmisor)
                          sClave = MensajePinPad(sRetorno)
                          If Mid(sClave, 1, 3) = "A00" Or Mid(sClave, 1, 3) = "A11" Then
                             sRefer = BuscaRetornoPinPad(sClave, "E")
                             Cn.Execute "update DPAGOTARJETA set tEstadoDocumento='04' where tDocumento='" & sCodigo & "' and tReferencia='" & sRefer & "'"
                             lContinua = True
                          Else
                             xError = BuscaRetornoPinPad(sClave, "B")
                             If Len(xError) > 0 Then
                                Mensaje xError, "VisaNet", 1000
                             End If
                          End If
                          
                          Mensaje "PinPad Listo. Esperando...", "PinPad", 500
                          nContador = nContador + 1
                          If nContador >= nTimeOut Then
                             If MsgBox("Tiempo de espera agotado, deseas mas tiempo?", vbExclamation + vbOKCancel, "VisaNet") = vbOK Then
                                lLoop = True
                                nContador = nTimeOut / 2
                             Else
                                 lLoop = False
                             End If
                          End If
                           
                           If nRet <> "0" Then
                              nContador = 0
                           End If
                        Loop While (Mid$(sRetorno, 5, 2) <> "C1") And lLoop
                     Else
                        MsgBox "Error de conectividad", vbCritical, sMensaje
                        Exit Sub
                     End If
                  End If
                                    
                  If Not lContinua Then
                     Exit Sub
                  End If
                  
                  Pedido = sDescrip
                  sPedido = Pedido
                  Cn.Execute "delete from " & sDetalle
                  Cn.Execute "UPDATE DPEDIDO SET TDOCUMENTO='' , TFACTURADO='' WHERE TCODIGOPEDIDO='" & Pedido & "'"
                  Cn.Execute "Insert into " & sDetalle & " (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                  "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                  "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,tsubalmacen) " & _
                  "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                  "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                  "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,tsubalmacen " & _
                  "From DPEDIDO where tEstadoItem='N' and tCodigoPedido='" & Pedido & "'"
                  
                  Cn.Execute "delete from " & sComboDetalle
                  Cn.Execute "insert into " & sComboDetalle & "(tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, tObservacion, lCorte) " & _
                  "select tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, tObservacion, lCorte " & _
                  "From CPEDIDO where tCodigoPedido='" & Pedido & "'"

                  Cn.Execute "delete from " & sProductoPropiedad
                  Cn.Execute "insert into " & sProductoPropiedad & "(tItem, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra) " & _
                  "select tItem, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra " & _
                  "From TPRODUCTOPROPIEDAD where tCodigoPedido='" & Pedido & "'"
                  
                  Cn.Execute "delete from " & sComboPropiedad
                  Cn.Execute "insert into " & sComboPropiedad & "(tItem, tItemCombo, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra) " & _
                  "select tItem, tItemCombo, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra " & _
                  "From TCOMBOPROPIEDAD where tCodigoPedido='" & Pedido & "'"
                             
                  Cn.Execute "Update MDOCUMENTO set tObservacion ='ANULACION RAPIDA' where tDocumento ='" & sDocumento & "'"
                  If Not lFECarbajal Then
                    Cn.Execute "delete from DPAGODOCUMENTO where tDocumento='" & sCodigo & "'"
                    Cn.Execute "update MDOCUMENTO set tEstadoDocumento='04', tUsuarioAnulado='" & sVar1 & "', fRegistroAnulado=getdate(), tObservacion='ANULACION RAPIDA' ,lreplica=1  where tDocumento='" & sCodigo & "'"
                    Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1, tEstadoPedido='01'  where tCodigoPedido='" & Pedido & "'"
                    RsDetalle.Requery
                  End If
                  'FACTURACION_E_PERU
                  If pais = "000" Then
                    If lFacturacionE Then
                    
                           If lFEOfisis Then 'OFISIS
                                 Dim lDocElecOfisis As Boolean
                                 lDocElecOfisis = Calcular("select isnull(tdi.lDocumentoElectronicoOfisis,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MDOCUMENTO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tDocumento= '" & sCodigo & "'", Cn)
                                 
                                 If lDocElecOfisis Then 'DOC ELECTRONICO OFISIS
                                        Dim xCDROfisis As String
                                        Dim RsDocumentoOfisis As Recordset
                                        Dim xContOfisis As Integer
                                        
                                        fDocumento = Mid(sCodigo, 1, 1) + Mid(sCodigo, 4, 3) + "-" + CStr(CLng(Mid(sCodigo, 8, 8)))
                                        Isql = "Select * From dbo.TCFACT_ELEC where NU_DOCU='" & fDocumento & "'"
                                        Set RsDocumentoOfisis = Lib.OpenRecordset(Isql, CnFE)
                                        
                                        If RsDocumentoOfisis.RecordCount > 0 Then
                                            CnFE.Execute "Update TCFACT_ELEC set CO_ESTA_DOCU = 'ANU' Where NU_DOCU = '" & fDocumento & "' and TI_DOCU <> 'D'"
                                        End If
                                 End If
                           ElseIf lFESpring Then 'SPRING
                           
                           ElseIf lFEpape Then 'PAPERLESS
                           
                           ElseIf lFEBiz Then ' SPRING
                           
                           ElseIf lFECarbajal Then 'CARBAJAL
                                Dim lDocElec As Boolean
                                Dim sImporteLetra As String
                                lDocElec = Calcular("select isnull(tdi.lFacturacionElectronica,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MDOCUMENTO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tDocumento= '" & sDocumento & "'", Cn)
                                If lDocElec Then 'DOC ELECTRONICO INFOFACT
                                    sImporteLetra = NumeroCadena(str(Calcular("select isnull(nVenta,0) as Codigo from mDocumento where tDocumento='" & sDocumento & "'", Cn))) + " " + sMonedaN
                                    If Not INSERTAFE_CARVAJAL(sDocumento, sImporteLetra, 0, 1) Then '----CABECERA
                                        Cn.Execute "Update MDOCUMENTO set tObservacion ='' where tDocumento ='" & sDocumento & "'"
                                        Exit Sub
                                    End If
                                End If
                                Cn.Execute "delete from DPAGODOCUMENTO where tDocumento='" & sCodigo & "'"
                                Cn.Execute "update MDOCUMENTO set tEstadoDocumento='04', tUsuarioAnulado='" & sVar1 & "', fRegistroAnulado=getdate(), tObservacion='ANULACION RAPIDA' ,lreplica=1  where tDocumento='" & sCodigo & "'"
                                Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1, tEstadoPedido='01'  where tCodigoPedido='" & Pedido & "'"
                                RsDetalle.Requery
                           Else ' INFOFACT
                           
                                Dim lDocElecInfofact As Boolean
                                lDocElecInfofact = Calcular("select isnull(tdi.lFacturacionElectronica,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MDOCUMENTO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tDocumento= '" & sCodigo & "'", Cn)
                                
                                If lDocElecInfofact Then 'DOC ELECTRONICO INFOFACT
                                    If Not INSERTAFE(Format(FechaServidor() - 1, "yyyymmdd"), "", 3, "") Then '
                                    End If
                                End If
                           
                           End If
                    End If
                  End If
                  
                  If RsDetalle.RecordCount = 0 Then
                     nMonto = 0
                  Else
                     nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
                  End If
                  VisualizaMonto
                  
                  Dim RsTemp As Recordset
                  Set RsTemp = Lib.OpenRecordset("select tTipoPedido, tMozo, tObservacion, vMozo.Descripcion as Mozo FROM dbo.MPEDIDO LEFT OUTER JOIN dbo.vMozo ON dbo.MPEDIDO.tMozo = dbo.vMozo.Codigo Where tcodigopedido='" & Pedido & "'", Cn)
                                    
                  sTipoPedido = IIf(IsNull(RsTemp!tTipoPedido), "01", RsTemp!tTipoPedido)
                  If sTipoPedido = "01" Then
                     'cmdCabecera_Click (2)
                  Else
                     'cmdCabecera_Click (1)
                  End If
                  sMozo = IIf(IsNull(RsTemp!tMozo), "0000", RsTemp!tMozo)
                  txtTitulo.Caption = " Caja Rápida : " & sCaja & " " & Mesero & ": " & IIf(IsNull(RsTemp!Mozo), "", RsTemp!Mozo) & IIf(sPedido = "", "", " - " & sPedido)
                  sObser = IIf(IsNull(RsTemp!tObservacion), "", RsTemp!tObservacion)
                  txtObservacion.Caption = sObser
                  Set RsTemp = Nothing
                  
                  If lInfhotel Then
                     sComandaInfhotel = Calcular("select tComanda as Codigo From MPEDIDO where tCodigoPedido='" & Pedido & "'", Cn)
                     CnInfhotel.Execute "update MCOMANDA set TESTADO='04' where tComanda ='" & sComandaInfhotel & "'  and tPuntoVenta='" & sPuntoVenta & "'"
                     CnInfhotel.Execute "delete from DCOMANDA where tcomanda='" & sComandaInfhotel & "' and tcodigoitem='100000' and tPuntoVenta='" & sPuntoVenta & "'"
                     CnInfhotel.Execute "delete from WMCOMANDA where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
                     CnInfhotel.Execute "delete from WDCOMANDA where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
                  End If
                  
                  'JSJR 09/01/2025
                  wAnulaDocumento = True
                  
               End If
          
          Case Is = 1  'Cancelar
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
                  
               If MsgBox("Seguro de Cancelar el Pedido?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
               
                  If Pedido <> "" Then
                     sUsuarioAutoriza = sUsuario
                     tUsuActua = sUsuario
                      If lPasswordC Then
                         If Supervisor("01") = False Then
                            MsgBox "Clave no permitida", vbExclamation, sMensaje
                            RsDetalle.Requery
                            Exit Sub
                         End If
                         sUsuarioAutoriza = sVar1
                      End If
                      sUsuario = tUsuActua
                      Sw = True
                      If lEliminaC Then
                         'JSJR 09/01/2025
                         wAnulaDocumento = True
                         fraEliminacion.Visible = True
                         tabProducto.Visible = False
                         ActivaCabecera False
                      Else
                         sCodigo = ""
                         sDescrip = ""
                         EliminaCabecera
                      End If
                       'Inicializar
                  Else
                      Cn.Execute "delete from " & sDetalle
                      Cn.Execute "delete from " & sComboDetalle
                      Cn.Execute "delete from " & sProductoPropiedad
                      Cn.Execute "delete from " & sComboPropiedad
                      'Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='" & sCodigo & "', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='" & sDescrip & "'   where tCodigoPedido ='" & Pedido & "'"
                      Cn.Execute "Update DPEDIDO Set tEstadoItem = 'A' where tCodigoPedido = '" & Pedido & "'"
                      Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido='" & Pedido & "'"
                      Cn.Execute "delete CPEDIDO where tCodigoPedido='" & Pedido & "'"
                      Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido='" & Pedido & "'"
                                         
                      RsDetalle.Requery
                      RsCombo.Requery
                      RsPropiedad.Requery
                      RsProductoPropiedad.Requery
                      RsComboPropiedad.Requery
                    
                      nMonto = 0
                      Pedido = ""
                      sProducto = ""
                      wCombo = False
                      nCombo = 0
                      sItem = ""
                      sCodigoDescuento = ""
                      tAutorizaDescuento = ""
                      ltope = False
                      nTope = 0
                      xDescuento = 0
                      Inicializar
                      VisualizaMonto
                  End If
               End If
                              
          Case Is = 2  'Aceptar
               wEnter = False
               If Pedido = "" Then
                  If nPuerto > 0 Then
                     Visor String(Int((19 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((19 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
                  End If
                  Unload Me
               Else
                  If sModulo = "INFOREST" Or sModulo = "ADICION" Then
                     Cn.Execute "delete from DPEDIDO where tCodigoPedido='" & Pedido & "'"
                     'Inserta el Detalle
                     Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                                "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                                "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,fregistro, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubAlmacen,tCodigoEtiqueta,tunidadnegocio,nenvio,fenvio,fdiacontable) " & _
                                "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo,  tMoneda, " & _
                                "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                                "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,getdate(), nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubAlmacen,tCodigoEtiqueta,'" & sUnidadNegocio & "' ,nenvio,fenvio,'" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
                                "From [" & sDetalle & "] where tEstadoItem='N'"
                                
                                
                        'ACTUALIZA  tAplicaExoneracion, tAplicaInafectacion
                       Dim oComando As clsComando
                        Set oComando = New clsComando
                        If Not oComando.CreateCmdSp("spUpdate_DPEDIDO_Ina", Cn) Then
                           Set oComando = Nothing
                           Exit Sub
                        End If
                        
                        oComando.CreateParameter "@pedido", adVarChar, adParamInput, 10, Pedido
                        
                        If Not oComando.GetParamOK Then
                           Set oComando = Nothing
                           Exit Sub
                        End If
                    
                        If Not oComando.ExecSP Then
                           Set oComando = Nothing
                           Exit Sub
                         End If
                             
                     If nPuerto > 0 Then
                        Visor String(Int((19 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((19 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
                     End If
                     Unload Me
                  Else
                     If RsDetalle.RecordCount = 0 Then
                        Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='000', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='Cancelación de un Pedido en Blanco desde Caja Rapida'  where tCodigoPedido ='" & Pedido & "'"
                        Unload Me
                     Else
                        MsgBox "No debes tener pedidos sin atender", vbCritical, sMensaje
                     End If
                  End If
               End If
               
          Case Is = 3 ' Aceptar Precios
               GrabaProducto
               tabProducto.Visible = True
               fraDetalle.Visible = False
               ActivaCabecera True
               
          Case Is = 4 ' Cancelar Precios
               tabProducto.Visible = True
               fraDetalle.Visible = False
               ActivaCabecera True
                         
          Case Is = 5 ' Salir
               wEnter = True
               
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=1", Cn)
               If i > 0 Or RsDetalle.RecordCount > 0 Then
                  MsgBox "El pedido actual esta activo", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If Pedido = "" Then
                  If nPuerto > 0 Then
                     Visor String(Int((19 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((19 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
                  End If
                  Unload Me
               Else
                  If sModulo = "INFOREST" Or sModulo = "ADICION" Then
                     Cn.Execute "delete from DPEDIDO where tCodigoPedido='" & Pedido & "'"
                     'Inserta el Detalle
                     Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                                "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                                "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,fregistro, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubAlmacen,tCodigoEtiqueta,tunidadnegocio) " & _
                                "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo,  tMoneda, " & _
                                "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                                "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,getdate(), nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubAlmacen,tCodigoEtiqueta,'" & sUnidadNegocio & "' " & _
                                " From [" & sDetalle & "] where tEstadoItem='N'"
                     
                'ACTUALIZA  tAplicaExoneracion, tAplicaInafectacion
               ' Dim oComando As clsComando
                Set oComando = New clsComando
                If Not oComando.CreateCmdSp("spUpdate_DPEDIDO_Ina", Cn) Then
                   Set oComando = Nothing
                   Exit Sub
                End If
                
                oComando.CreateParameter "@pedido", adVarChar, adParamInput, 10, Pedido
                
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Sub
                End If
            
                If Not oComando.ExecSP Then
                   Set oComando = Nothing
                   Exit Sub
                 End If


                     If nPuerto > 0 Then
                        Visor String(Int((19 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((19 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
                     End If
                     Unload Me
                  Else
                     If RsDetalle.RecordCount = 0 Then
                        Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='000', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='Cancelación de un Pedido en Blanco desde Caja Rapida'  where tCodigoPedido ='" & Pedido & "'"
                        Unload Me
                     Else
                        MsgBox "No debes tener pedidos sin atender", vbCritical, sMensaje
                     End If
                     
                  End If
               End If
                         
          Case Is = 6 ' Aceptar Propiedades
          
               If wAgregaCombo Then
               
                    If ObligaPropiedad(sCombo) = False Then
                        Exit Sub
                    Else
                        grdDetalle.Enabled = True
                    End If
                    cmdOpcionPicture(1).Enabled = False
                    RsCombo.Requery
                    If RsCombo.RecordCount > 0 Then
                        RsCombo.MoveFirst
                    End If
                    RsCombo.Find "titemCombo = '" & xItem & "'"
                    fraProductoCombo.Visible = True
                    fraPropiedad.Visible = False
                    
               Else
                    'Oscar Ortega---------------------------------
                    If ObligaPropiedad(sProducto) = False Then
                        Exit Sub
                    Else
                        grdDetalle.Enabled = True
                    End If
                    RsDetalle.Requery
                    RsDetalle.MoveFirst
                    RsDetalle.Find "titem = '" & sItem & "'"
                    ActivaCabecera True
                    tabProducto.Visible = True
                    fraPropiedad.Visible = False
               End If
               
          Case Is = 7 ' Observaciones
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
               lblObservacion.Text = sDescrip
               
'               If wEnter = True Then
'                  If wAgregaCombo Then
'                     nPos = RsCombo.AbsolutePosition
'                     Cn.Execute "Update " & sComboDetalle & " set tObservacion = '" & sDescrip & "' where tItem ='" & sitem & "' and tItemCombo='" & xItem & "'"
'                  Else
'                     nPos = RsDetalle.AbsolutePosition
'                     Cn.Execute "Update " & sDetalle & " set tObservacion = '" & sDescrip & "' where tItem ='" & sitem & "'"
'                  End If
'                  lblObservacion.Text = sDescrip
'               End If
          
       Case Is = 8  'Imp.Pedido
       
            ' Miguel Mendoza - 10062019
            If Not VerificaComboMaximo() Then Exit Sub
            ' *************************
            
           
            'If (lCanalDelivery) Then
                     If (lObligaOrigenVenta) Then
                     If sTipoPedidoPD = sTipoPedido And (vOrigenVentas = "" Or vOrigenVentas = Null) Then
                            vOrigenVentas = sDefaultOrigenVenta
                        End If
                          If (vOrigenVentas = "" Or vOrigenVentas = Null) Then
                              MsgBox "Debe Seleccionar un Origen de Venta", vbInformation, sMensaje
                              Call ActivarOrigenVentas
'                              fraOrigenVentas.Visible = True
                              Exit Sub
                          End If
                      End If
             'End If
             
            If lObservacion And Trim(txtObservacion.Caption) = "" Then
               MsgBox "Debes ingresar la Observación", vbInformation, sMensaje
               cmdDetalle_Click (6)
               If Trim(txtObservacion.Caption) = "" Then
                  Exit Sub
               End If
            End If
       
            If RsArea.RecordCount = 0 Then
               MsgBox "No existe area definida", vbInformation, sMensaje
               Exit Sub
            End If
            
            If RsDetalle.RecordCount = 0 Then
               Exit Sub
            End If
            
            'OO
            If ExistenPropiedadesPendientesEnPedido() Then
                If ExistenPropiedadesPendientesEnCombos() Then
                    Screen.MousePointer = vbHourglass
                                        
                     'insuimo2013
                    Cn.Execute "delete from " & sInsumoCombo
                    Dim X As Integer
                    Dim rstItems As New Recordset
                    Dim cadenaInsumos As String
                    Dim cadenaAEnviar As String
                    Dim cmdInsumo As New ADODB.Command
                    Dim resultado As String
                    Cn.Execute "insert into " & sInsumoCombo & " select sum(ncantidad) ncantidad, TPRODUCTO.TCODIGOINSUMO from " & sDetalle & " inner join tproducto on " & sDetalle & ".tcodigoproducto=tproducto.tcodigoproducto INNER JOIN  dbo.TINSUMO ON dbo.TPRODUCTO.tcodigoInsumo = dbo.TINSUMO.tcodigo  where lcontrolinsumocritico=1 and isnull(limprime,0)=0 AND ISNULL(TCODIGOINSUMO,''  )<>'' and (tinsumo.lactivo=1) group by tcodigoinsumo"
                    Cn.Execute "insert into " & sInsumoCombo & " sELECT     SUM(" & sDetalle & ".nCantidad * " & sComboDetalle & ".nCantidad) AS ncantidad, dbo.TPRODUCTO.tCodigoInsumo FROM  " & sComboDetalle & " INNER JOIN " & sDetalle & " ON  " & sComboDetalle & ".tItem = " & sDetalle & ".tItem INNER JOIN dbo.TINSUMO INNER JOIN dbo.TPRODUCTO ON dbo.TINSUMO.tcodigo = dbo.TPRODUCTO.tCodigoInsumo ON " & sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto WHERE     (dbo.TPRODUCTO.lControlInsumoCritico = 1) AND (dbo.TINSUMO.lactivo = 1) AND (ISNULL(" & sComboDetalle & ".lImprime, 0) = 0) AND (ISNULL(dbo.TPRODUCTO.tCodigoInsumo, N'') <> '') AND (" & sDetalle & ".lCombinacion = 1) GROUP BY dbo.TPRODUCTO.tCodigoInsumo "
                    cadenaInsumos = "select  SUM(ncantidad), tCodigoInsumo from " & sInsumoCombo & "  group by tCodigoInsumo order by 2"
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
                                            If resultado <> "1" Then:  MsgBox "No hay cantidad disponible de : " & vbNewLine & resultado, vbExclamation, sMensaje: variableEmite = False: Screen.MousePointer = vbDefault: RsDetalle.Requery: Exit Sub
                                        End If
                                        variableEmite = True
                    'InsumosCriticos
                    
                    
                    If Pedido = "" Then
                       GeneraPedido
                    Else
                       ActualizaPedido
                    End If

                                
                                
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
                        
                        oComando_.CreateParameter "@pedido", adVarChar, adParamInput, 10, Pedido
                        oComando_.CreateParameter "@tturnoKDS", adVarChar, adParamInput, 10, Pedido
                        
                        If Not oComando_.GetParamOK Then
                           Set oComando_ = Nothing
                           Exit Sub
                        End If
                    
                        If Not oComando_.ExecSP Then
                           Set oComando_ = Nothing
                           Exit Sub
                         End If
                   End If
                   
                   
                   
                   '---------------------------------
                   '---------------------------------
                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                    If nPuerto > 0 Then
                       Visor "Enviando Pedido...", "", nPuerto, "N"
                    End If
                                                        
                    If lOrden Then
                       Isql = "select * from vPedido " & _
                              "Where Codigo = '" & Pedido & "' and nOrden in (select nOrden from DPEDIDO where tCodigoPedido='" & Pedido & "' and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))) " & _
                              "ORDER BY nOrden, tItem, nOrdenCombo,tetiqueta,combo " 'tItemCombo"
                    Else
                       Isql = "select * from vPedido " & _
                              "Where Codigo = '" & Pedido & "' And lImprimeArea = 1 and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))" & _
                              "ORDER BY nOrden, tItem, nOrdenCombo, tetiqueta,combo " ' tItemCombo"
                    End If
                                                                    
                    Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                    If RsImpresion.RecordCount = 0 Then
                        MsgBox "No se encontraron datos para imprimir"
                        Screen.MousePointer = vbDefault
                        RsDetalle.Requery
                        Exit Sub
                    End If
                    If Not RsImpresion.EOF Then
                       RsArea.MoveFirst
                       For i = 1 To RsArea.RecordCount
'                           RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
'                           If RsArea!tIcono = "" Or RsArea!nValor = 1 Then
'                              If RsImpresion.RecordCount <> 0 Then
'                                 RsImpresion.MoveFirst
'                                 sPedido = Pedido
'                                 ImprimePedido RsImpresion, "N", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Rapido"
'                                 sPedido = ""
'                              End If
'                           End If

                            If lImpComandaBarra Then
                                Dim RsAreaPedidoItem As Recordset
                                Set RsAreaPedidoItem = Lib.OpenRecordset(" select isnull(titem,'000') as tItem from vPedido where Codigo = '" & sPedido & "' And lImprimeArea = 1 group by titem", Cn)
                                Dim Item As Integer
                                For Item = 1 To RsAreaPedidoItem.RecordCount
                                    RsImpresion.Filter = "tArea='" & RsArea!tArea & "' And tItem='" & RsAreaPedidoItem!tItem & "'"
                                    If RsArea!tIcono = "" Or RsArea!nValor = 1 Then
                                        If RsImpresion.RecordCount <> 0 Then
                                            RsImpresion.MoveFirst
                                            If lImpComandaBarra Then
                                                Set ImageBarra = Nothing
                                                'BarraCodigo RsAreaPedidoItem!tItem + Right(sPedido, 8), ImageBarra
                                                ImageBarra.Picture = LoadPicture(CrearImagenQR_Comanda("P|" + sPedido + "|" + RsAreaPedidoItem!tItem + "|" + sCaja + "|" + RsArea!tArea + "|" + Format(RsImpresion!fRegistro, "yyyyMMdd HH:mm:ss")))
                                                'Exit Sub
                                            End If
                                            ImprimePedido RsImpresion, "N", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Rapido", ImageBarra
                                            'If (lCompatibilidadTVS) Then 'TVS
                                            '    Call TVS_EnviarTicket(RsImpresion, RsArea!Area, "Nuevo Ticket", sPedido)
                                            'End If
                                        End If
                                    End If
                                    RsAreaPedidoItem.MoveNext
                                Next Item
                            Else
                                RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
                                If RsArea!tIcono = "" Or RsArea!nValor = 1 Then
                                    If RsImpresion.RecordCount <> 0 Then
                                        RsImpresion.MoveFirst
                                        ImprimePedido RsImpresion, "N", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Rapido"
                                        'If (lCompatibilidadTVS) Then 'TVS
                                        '    Call TVS_EnviarTicket(RsImpresion, RsArea!Area, "Nuevo Ticket", sPedido)
                                        'End If
                                    End If
                                End If
                            
                            End If
                            
                           RsArea.MoveNext
                       Next i
                    End If
                    
                    Set ImageBarra = Nothing
                    
                    If lKDS Then
                        Dim kdsRsCabecera As Recordset
                        Isql = "SELECT * From vPedidoCabecera Where Codigo = '" & Pedido & "' Order By codigo "
                        Set kdsRsCabecera = Lib.OpenRecordset(Isql, Cn)
                       Call KDS_AnadirNuevaOrden(kdsRsCabecera, 0)
                    End If
                    
                    Cn.Execute "Update CPEDIDO Set lImprime = 1 where tCodigoPedido = '" & Pedido & "'"
                    Cn.Execute "Update " & sComboDetalle & "  Set lImprime = 1"
                    RsCombo.Requery
                    
                    'CESAR----CHEF CONTROL
                    Dim ChefEnvio As Boolean
                    ChefEnvio = Calcular("select ISNULL(lEnvioChef,0) as Codigo FROM TPARAMETRO", Cn)
                    
                    Cn.Execute "Update DPEDIDO Set lNoCantado=0 where tCodigoPedido = '" & Pedido & "' and lNoCantado IS NULL"
                    
                    If ChefEnvio Then
                    Cn.Execute "Update DPEDIDO Set lCantadoc=1,fCantadoC=GetDate(), lTipoEnvio=0 where tCodigoPedido = '" & Pedido & "' and lImprime<>1"
                    End If
                    '---------------------------------
        
        
        
        
        
        
                    Cn.Execute "Update DPEDIDO Set lImprime = 1, fenvio=getdate(), nEnvio= isnull(nEnvio,0) + 1 where tCodigoPedido = '" & Pedido & "' and limprime<>1 "
                    Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1 where tCodigoPedido='" & Pedido & "'"
                    Cn.Execute "Update " & sDetalle & "  Set fenvio = getdate(), nEnvio = isnull(nEnvio,0) + 1 where limprime <> 1"
                    Cn.Execute "Update " & sDetalle & "  Set lImprime = 1"
                    sPedido = Pedido
                    'RsDetalle.Requery
                End If
            End If
            
             If sMozo = "" Or sMozo = "0000" Then
               sMozo = "0000"
               txtTitulo.Caption = " Caja Rápida : " & sCaja & " " & Mesero & ": Sin " & Mesero & IIf(sPedido = "", "", " - Pedido:" & sPedido)
            Else
               cmdDetallePicture(9).Enabled = False
               txtTitulo.Caption = " Caja Rápida : " & sCaja & " " & Mesero & ": " & Calcular("select descripcion as codigo from vMozo where Codigo='" & sMozo & "'", Cn) & IIf(sPedido = "", "", " - Pedido:" & sPedido)
            End If
            
            
            
            RsDetalle.Requery
            LimpiaRs
            Label2.Caption = muestra
            Screen.MousePointer = vbDefault
                        
       Case Is = 9  'Cliente Frecuente
                If Not Sw Then
                   'sTemp = txtTelefono.Caption
                End If

                wEnter = False
                sTipo = sTipoPedido
                sCodigo = ""
                sCodigoParienteSeleccionado = ""
                sCodigoInvitado = ""
                If lClub Then
                    frmBusquedaSocio.Show vbModal
                Else
                
                    frmBusquedaDelivery.Show vbModal
                End If
                If wEnter = True Then
                
                   sClienteFrecuente = sCodigo
                   txtCliente.Caption = sDescrip
                   txtTelefono.Caption = sClienteFrecuente
                   xDescuento = nVar1
                   sCodigoDescuento = "000"
                   sDescripcionDescuento = "Descuento por Cliente Frecuente"
                End If
               
       Case Is = 10  'Combos
            If wCombo Then
               tabProducto.Visible = False
               fraCombo.Visible = True
               fraProductoCombo.Visible = True
               wAgregaCombo = True
               ActivaCabecera False
               
               If Not RsCombo.EOF Then
                  RsCombo.MoveFirst
               End If
               AsignaProductoCombo
               
               txtBarra.SetFocus
            End If
            
       'OO--------------------------------------------------------------------------------------------------------
       Case Is = 14  'Cargos
            If RsDetalle.RecordCount = 0 Then
               Exit Sub
            End If
            
            ' Miguel Mendoza - 10062019
            If Not VerificaComboMaximo() Then Exit Sub
            ' *************************
            
            If sPuntoVenta = "" Then
               MsgBox "Falta ingresar el punto de venta", vbExclamation, sMensaje
               Exit Sub
            End If
            
            If Calcular("select count(tFacturado) as Codigo from DPEDIDO where tCodigoPedido ='" & sPedido & "' and isnull(tFacturado,'0') <> '0' and len(ltrim(tFacturado)) <> 0", Cn) > 0 Then
               MsgBox "Imposible pasar el pedido a Cargos, pedidos con items Facturados", vbExclamation, sMensaje
               Exit Sub
            End If
            
            'Clave de Multi Cajero
            tUsuActua = sUsuario
            If lMultiCajero Then
               If Supervisor("16") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               sUsuario = sVar1
            End If
            sUsuario = tUsuActua
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=0", Cn)
               If i > 0 Then
                  cmdOpcion_Click (8)
               End If
            End If
            sPedido = Pedido
                                                                  
            'Chequea si existe platos a facturar
            sTD = "N"
            RsDetalle.MoveFirst
            Do While Not RsDetalle.EOF
               If (Len(Trim(RsDetalle!tFacturado)) = 0 Or IsNull(RsDetalle!tFacturado)) Then
                  sTD = "S"
                  Exit Do
               End If
               RsDetalle.MoveNext
            Loop
    
            If sTD <> "S" Then
               MsgBox "Error: No existen Productos a Facturar", vbCritical, sMensaje
               Exit Sub
            End If

            frmCargo.Show vbModal

            If Not wEnter Then
               Exit Sub
            End If
           
            Dim tItem As Integer
            Dim Correlativo As Integer
            Dim CorrelaProp As Integer
            Dim nMovimiento As Integer
            Dim sAsignado   As String
            Dim MonPuntoventa As String
            
            sCliente = sCodigo
            xSuma = Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido='" & sPedido & "'", Cn)
            
            If sDescrip = "Infhotel" Then
               If MsgBox("Esta seguro de Generar el Pedido Nro: " & sPedido & " en Infhotel?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               sHabitacion = ""
               sReserva = ""
               sPasajero = ""
            Else
               If MsgBox("Esta seguro de Enviar el Pedido Nro: " & sPedido & _
                  Chr(13) & "a la " & Trim(sDescrip) & " " & IIf(sDescrip = "Reserva", sReserva, sHabitacion) & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
            End If
            
            MonPuntoventa = Calcular("select max(tmoneda) as codigo from vpuntoventa where tpuntoventa='" & sPuntoVenta & "'", CnInfhotel)
               Call PUltimaComanda
               sComandaInfhotel = Calcular("select left(MAX(tComanda),8) as Codigo from MCOMANDA where tPuntoVenta='" & sPuntoVenta & "'", CnInfhotel)
               sComandaInfhotel = Lib.Correlativo(sComandaInfhotel, 8)
               CnInfhotel.Execute "Update TPUNTOVENTA Set nUltimoComanda = '" & sComandaInfhotel & "' where tPuntoVenta='" & sPuntoVenta & "'"
               sComandaInfhotel = sComandaInfhotel & "-" & UCase(Mid(rsPuntoVenta!descripcion, 1, 3))
               rsPuntoVenta.Requery
                              
               'Genero las comandas en Infhotel
               'Cabecera
               If sDescrip = "Infhotel" Then sAsignado = "01"
               If sDescrip = "Reserva" Then sAsignado = "03"
               If sDescrip = "Habitacion" Then sAsignado = "02"
               If sTipoComanda = "01" Then
                    Isql = "Insert into MCOMANDA " & _
                           "(tComanda, tPuntoVenta, tHotel, nMovimiento, fFecha, hHora, nTotal, tEstado, " & _
                           "tEmitido, tAsignacion, tCodigoReserva, tNumeroHabitacion, tCodigoFuncionario, " & _
                           "tCaja, tDocumento, tUsuario, nTCambio, tCodigoCompania, tCliente, tMoneda, fFechaE, hHoraE, tUsuarioE, tIncluido, nRoomSer, nDescuento, tNotaPedido, tCompania, tContacto, lInforest ) " & _
                           "values('" & sComandaInfhotel & "', '" & sPuntoVenta & "', '" & sHotel & "', 1,  getdate(), getdate(), " & IIf(MonPuntoventa = "01", xSuma, xSuma / nTC) & ", '01', " & _
                           "1,'" & sAsignado & "', '" & sReserva & "', '" & sHabitacion & "', '', " & _
                           "'', '', '" & xUsuario & "', " & nTC & ", '', '" & sPasajero & "', '" & MonPuntoventa & "', getdate(), getdate(), '" & xUsuario & "','" & sTipoComanda & "', " & IIf(sTipoPedido = "03", nLlevar, 0) & ", " & xDescuento & ", '" & sPedido & "', '" & sCompania & "', '" & sContacto & "', 1)"
                    CnInfhotel.Execute Isql
               Else
                    Isql = "Insert into MCOMANDA " & _
                          "(tComanda, tPuntoVenta, tHotel, nMovimiento, fFecha, hHora, nTotal, tEstado, " & _
                          "tEmitido, tAsignacion, tCodigoReserva, tNumeroHabitacion, tCodigoFuncionario, " & _
                          "tCaja, tDocumento, tUsuario, nTCambio, tCodigoCompania, tCliente, tMoneda, fFechaE, hHoraE, tUsuarioE, tIncluido, nRoomSer, nDescuento, tNotaPedido, tCompania, tContacto, lInforest) " & _
                          "values('" & sComandaInfhotel & "', '" & sPuntoVenta & "', '" & sHotel & "', 1,  getdate(), getdate(), 0, '01', " & _
                          "1,'" & sAsignado & "', '" & sReserva & "', '" & sHabitacion & "', '', " & _
                          "'', '', '" & xUsuario & "', " & nTC & ", '', '" & sPasajero & "', '" & MonPuntoventa & "', getdate(), getdate(), '" & xUsuario & "','" & sTipoComanda & "', " & IIf(sTipoPedido = "03", nLlevar, 0) & ", " & xDescuento & ", '" & sPedido & "', '" & sCompania & "', '" & sContacto & "', 1)"
                    CnInfhotel.Execute Isql
               End If
            
            If sDescrip = "Habitacion" Then
               txtObservacion.Caption = "Hab: " & sHabitacion
            ElseIf sDescrip = "Reserva" Then
               txtObservacion.Caption = "Res: " & sReserva
               sHabitacion = ""
            End If
            Isql = "Update MPEDIDO Set " & _
                    "tEstadoPedido ='05', " & _
                    "tComanda ='" & sComandaInfhotel & "', " & _
                    "tPuntoVenta ='" & sPuntoVenta & "', " & _
                    "tReserva ='" & sReserva & "', " & _
                    "tHabitacion ='" & sHabitacion & "', " & _
                    "tObservacion='" & txtObservacion.Caption & "', " & _
                    "tFichaPasajero ='" & sFichaPasajero & "', " & _
                    "tTipoComanda ='" & sTipoComanda & "', " & _
                    "tPasajero ='" & sPasajero & "' " & _
                    "  where tCodigoPedido = '" & sPedido & "'"
            Cn.Execute Isql
            
            'Detalle
''            CnInfhotel.Execute "delete from DCOMANDA where tComanda ='" & RsCabecera!tComanda & "'"
            nMovimiento = Calcular("select max(nmovimiento) as codigo from dcomanda where tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
            If sTipoComanda = "01" Then
                Isql = "Insert into DCOMANDA " & _
                       "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                       "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                       "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                       "select '" & sComandaInfhotel & "' as tComanda, '" & sPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem , " & nMovimiento & " as nMovimiento, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                       IIf(MonPuntoventa = "01", "T1.nPrecioVenta", "T1.nPrecioVenta / " & nTC) & " as nPrecioUnitario, nCantidad, " & IIf(MonPuntoventa = "01", "nVenta", "nVenta / " & nTC) & " as nTotal, T1.nInsumo+T1.nGasto+T1.nManoObra as nPrecioCos, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                       "'' as tCuenta, '' as tCaja, '' as tDocuemento, '" & sAsignado & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                       "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                       "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
                CnInfhotel.Execute Isql
            Else
                Isql = "Insert into DCOMANDA " & _
                       "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                       "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                       "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                       "select '" & sComandaInfhotel & "' as tComanda, '" & sPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem ," & nMovimiento & " as nMovimiento, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                       " 0 as nPrecioUnitario, nCantidad, 0 as nTotal, T1.nInsumo+T1.nGasto+T1.nManoObra as nPrecioCos, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                       "'' as tCuenta, '' as tCaja, '' as tDocuemento, '" & sAsignado & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                       "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                       "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
                CnInfhotel.Execute Isql
            End If
            
            'Propinas
            If val(sPropina) > 0 Then
                'Grabo la propina en el Mpropina del Inforest
                Isql = "insert into MPROPINA " & _
                      "(tcodigopedido,fregistro,tmoneda,nmonto,tusuario, tComanda) " & _
                      "values('" & sPedido & "',getdate(),'" & sMonPropina & "'," & sPropina & ",'" & xUsuario & "', '" & sComandaInfhotel & "')"
                Cn.Execute Isql
                
                'Grabo la propina en el Detalle de la comanda
                tItem = Calcular("select max(titem) as codigo from dcomanda where tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
                Isql = "Insert Into dcomanda " & _
                       "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                       "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                       "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                       "values('" & sComandaInfhotel & "','" & sPuntoVenta & "','00','" & tItem & "','1','" & sPedido & "', " & _
                       "'100000'," & sPropina & ",'1'," & sPropina & ",'','" & sReserva & "','" & sHabitacion & "','', " & _
                       " '','','" & sAsignado & "','" & xUsuario & "',getdate(),getdate())"
                CnInfhotel.Execute Isql
               
                'Grabo la propina en el Mpropina del Infhotel
                CorrelaProp = Calcular("select max(ncorrela) as codigo from mpropina", CnInfhotel) + 1
            
                Isql = "Insert Into MPROPINA " & _
                       "(ncorrela,tcodigoreserva,tnumerohabitacion,tcomanda,tcodigoitem,tmoneda,nmonto,tdocumento,tresponsable,testado,ffecha,tusuario, tPuntoVenta) " & _
                       "values(" & CorrelaProp & ",'" & sReserva & "','" & sHabitacion & "','" & sComandaInfhotel & "','100000','" & sMonPropina & "'," & sPropina & ", " & _
                       "'','" & Mid(sMozo, 2, 3) & "','01',getdate(),'" & xUsuario & "','" & sPuntoVenta & "')"
                CnInfhotel.Execute Isql
                
                Isql = "Update MCOMANDA set ncorrelaprop=" & CorrelaProp & " where tcomanda='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
                CnInfhotel.Execute Isql
            End If
                        
            'Actualiza las Cuentas Corrientes Infhotel
            If sDescrip = "Reserva" Then
               i = Calcular("select max(tNumeroCorrelativo) as Codigo from TCUENTARESERVA where tCodigoReserva='" & sReserva & "'", CnInfhotel)
               Isql = "Insert into TCUENTARESERVA " & _
                      "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad,testado, ttipo,tHotel, " & _
                      " nTotal, tNumeroCorrelativo,tpuntoventa, tItem, tUsuario) " & _
                      "select '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion , getdate() as fFecha, getdate() as hHoraMovimiento, '" & sComandaInfhotel & "' as tComanda, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                      IIf(sMonedaBase = "01", "T1.nPrecioVenta", "T1.nPrecioVenta / " & nTC) & " as nPrecioUnitario, nCantidad,'' as testado,'' as ttipo,'" & sHotel & "' as tHotel," & IIf(sMonedaBase = "01", "nVenta", "nVenta / " & nTC) & " as nTotal, tItem + " & i & ",'" & sPuntoVenta & "' as tpuntoventa ,tItem, '" & xUsuario & "' as tUsuario " & _
                      "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                      "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
               CnInfhotel.Execute Isql
                             
               If val(sPropina) > 0 Then
                    'Graba la propina en la Cta Cte de la Reserva
                    i = Calcular("select max(titem) as Codigo from TCUENTARESERVA where tCodigoReserva='" & sReserva & "'and tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
                    Correlativo = Calcular("select max(tnumerocorrelativo) as Codigo from TCUENTARESERVA where tCodigoReserva='" & sReserva & "'", CnInfhotel) + 1
                    Isql = "insert into tcuentareserva " & _
                          "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad,testado,ttipo,tHotel, " & _
                          " nTotal, tNumeroCorrelativo,tpuntoventa ,tItem, tUsuario,ncorrelaprop) " & _
                          " values('" & sReserva & "','" & sHabitacion & "',getdate(),getdate(),'" & sComandaInfhotel & "','" & sPedido & "','100000'," & sPropina & ",'1','','','00', " & _
                          " " & sPropina & "," & Correlativo & ",'" & sPuntoVenta & "'," & i & ",'" & xUsuario & "'," & CorrelaProp & ")"
                    CnInfhotel.Execute Isql
                    
                    If sMonedaBase = sMonPropina Then
                        Isql = "UPDATE tCuentaReserva " & _
                              "SET nPrecioUnitario=" & CDbl(sPropina) & "," & _
                              "nCantidad=1," & _
                              "nTotal=" & CDbl(sPropina) & " " & _
                              "WHERE ncorrelaprop='" & CorrelaProp & "'"
                         CnInfhotel.Execute Isql
                    Else
                        If (sMonedaBase = "02" And sMonPropina = "01") Then
                        'La moneda base esta en $ y la propina esta en S/.
                            Isql = "UPDATE tcuentareserva " & _
                                 "SET nPrecioUnitario=" & CDbl(sPropina) / nTC & "," & _
                                 "nCantidad=1," & _
                                 "ntotal=" & CDbl(sPropina) / nTC & "" & _
                                 "WHERE nCorrelaProp='" & CorrelaProp & "'"
                            CnInfhotel.Execute Isql
                         Else
                            'La moneda base esta en S/. y la propina esta en $
                            Isql = "UPDATE tcuentaReserva" & _
                                 "SET nPrecioUnitario=" & CDbl(sPropina) * nTC & "," & _
                                 "nCantidad=1," & _
                                 "ntotal=" & CDbl(sPropina) * nTC & "" & _
                                 "WHERE nCorrelaProp='" & CorrelaProp & "'"
                            CnInfhotel.Execute Isql
                         End If
                    End If
                End If

            ElseIf sDescrip = "Habitacion" Then
               i = Calcular("select max(tNumeroCorrelativo) as Codigo from TCUENTAHABITACION where tNumeroHabitacion='" & sHabitacion & "' and tCodigoReserva='" & sReserva & "'", CnInfhotel)
               Isql = "Insert into TCUENTAHABITACION " & _
                      "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento,testado,ttipo, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad, tHotel, " & _
                      " nTotal, tNumeroCorrelativo,tpuntoventa, tItem, tUsuario) " & _
                      "select '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion , getdate() as fFecha, getdate() as hHoraMovimiento,'' as testado,'' as ttipo ,'" & sComandaInfhotel & "' as tComanda, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                      IIf(sMonedaBase = "01", "T1.nPrecioVenta", "T1.nPrecioVenta / " & nTC) & " as nPrecioUnitario, nCantidad, '" & sHotel & "' as tHotel," & IIf(sMonedaBase = "01", "nVenta", "nVenta / " & nTC) & " as nTotal, tItem + " & i & ",'" & sPuntoVenta & "' as tpuntoventa ,tItem, '" & xUsuario & "' as tUsuario " & _
                      "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                      "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
               CnInfhotel.Execute Isql
               
               If val(sPropina) > 0 Then
                  'Graba la propina en la Cta Cte de la Habitacion
                  i = Calcular("select max(titem) as Codigo from TCUENTAHABITACION where tCodigoReserva='" & sReserva & "'and tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
                  Correlativo = Calcular("select max(tnumerocorrelativo) as Codigo from TCUENTAHABITACION where tCodigoReserva='" & sReserva & "'", CnInfhotel) + 1
                  Isql = "Insert into TCUENTAHABITACION " & _
                         "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento,testado,ttipo, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad, tHotel, " & _
                         " nTotal, tNumeroCorrelativo,tpuntoventa ,tItem, tUsuario,ncorrelaprop) " & _
                         " Values('" & sReserva & "','" & sHabitacion & "',getdate(),getdate(),'','','" & sComandaInfhotel & "','" & sPedido & "','100000'," & sPropina & ",'1','00', " & _
                         " " & sPropina & "," & Correlativo & ",'" & sPuntoVenta & "'," & i & ",'" & xUsuario & "'," & CorrelaProp & ")"
                  CnInfhotel.Execute Isql
                  
               If sMonedaBase = sMonPropina Then
                   Isql = "UPDATE tCuentaHabitacion " & _
                         "SET nPrecioUnitario=" & CDbl(sPropina) & "," & _
                         "nCantidad=1," & _
                         "nTotal=" & CDbl(sPropina) & " " & _
                         "WHERE ncorrelaprop='" & CorrelaProp & "'"
                    CnInfhotel.Execute Isql
                  Else
                    If (sMonedaBase = "02" And sMonPropina = "01") Then
                    'La moneda base esta en $ y la propina esta en S/.
                        Isql = "UPDATE tcuentahabitacion " & _
                             "SET nPrecioUnitario=" & CDbl(sPropina) / nTC & "," & _
                             "nCantidad=1," & _
                             "ntotal=" & CDbl(sPropina) / nTC & "" & _
                             "WHERE nCorrelaProp='" & CorrelaProp & "'"
                        CnInfhotel.Execute Isql
                    Else
                        'La moneda base esta en S/. y la propina esta en $
                        Isql = "UPDATE tcuentahabitacion " & _
                             "SET nPrecioUnitario=" & CDbl(sPropina) * nTC & "," & _
                             "nCantidad=1," & _
                             "ntotal=" & CDbl(sPropina) * nTC & "" & _
                             "WHERE nCorrelaProp='" & CorrelaProp & "'"
                        CnInfhotel.Execute Isql
                    End If
                  End If
               End If
            End If

           'CarlosD 13/11/2006
            CorrelativoC = 0
            CorrelativoC = val(Calcular("select max(nmovimiento) as codigo from wmcomanda where tcomanda='" & sComandaInfhotel & "' and tpuntoventa='" & sPuntoVenta & "' and thotel='" & sHotel & "'", CnInfhotel)) + 1
            
            Isql = "INSERT INTO WMCOMANDA([tComanda],[tPuntoVenta],[fFecha],[hHora],[tMoneda],[nTotal],[tEstado],[tEmitido], [tAsignacion],[tCodigoReserva],[tNumeroHabitacion],[tUsuario],[tCodigoCompania],[tCliente],[tCodigoFuncionario],[tHotel],[NTCAMBIO],[NMOVIMIENTO],[NDESCUENTO],[NROOMSER],[TMOZO],[tIncluido],[tMesa], tNotaPedido, tCompania, tContacto ) " & _
                   "Values (" & _
                    "'" & sComandaInfhotel & "', " & _
                    "'" & sPuntoVenta & "', " & _
                    "Getdate(), " & _
                    "GetDate(), " & _
                    "'" & MonPuntoventa & "'," & _
                    "" & IIf(MonPuntoventa = "01", xSuma, xSuma / nTC) & ", " & _
                    "'01',0, " & _
                    "'" & sAsignado & "', " & _
                    "'" & sReserva & "', " & _
                    "'" & sHabitacion & "', " & _
                    "'" & xUsuario & "', " & _
                    "'', " & _
                    "'" & sCliente & "', " & _
                    "'', " & _
                    "'" & sHotel & "', " & _
                    "'" & nTC & "', '" & CorrelativoC & "', " & nDescuento & ",'','','" & sTipoComanda & "', '" & sMesa & "', '" & sPedido & "', '" & sCompania & "', '" & sContacto & "')"
                CnInfhotel.Execute Isql

            Isql = "INSERT INTO WDCOMANDA SELECT * From dbo.DComanda where tComanda='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
            CnInfhotel.Execute Isql
            
            'Libera la Mesa
            Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa ='" & sMesa & "'"
            'Juntar Mesa
            Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
            
            wEnter = False
            
            Dim xPrecuenta As Boolean
            xPrecuenta = False
            If lObligaPrecuenta Then
               xPrecuenta = True
            Else
               If MsgBox("Deseas imprimir la Precuenta?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                  xPrecuenta = True
               End If
            End If
            
            If xPrecuenta Then
                Screen.MousePointer = vbHourglass
                'Imprime Precuenta
                If lPrecuentaImpresora Then
                   frmPrecuentaImpresora.Show vbModal
                   If Not wEnter Then
                      sPropina = ""
                      sTipoComanda = ""
                      Exit Sub
                   End If
                Else
                   sCodigo = sPreCuenta
                End If
                                        
                If lPrecuentaAgrupada Then
                   Isql = "select * from vPrecuentaAgrupada WHERE Codigo='" & sPedido & "' order by tItem"
                Else
                   Isql = "select * from vPrecuenta WHERE Codigo='" & sPedido & "' order by tItem"
                End If
    
                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                ImprimeInfhotel RsImpresion, sCodigo
                Cn.Execute "update MPEDIDO set fRegCuenta = getDate() where tCodigoPedido='" & sPedido & "'"
                LimpiaRs
            End If
            
            sPropina = ""
            sTipoComanda = ""
            
            RsDetalle.Requery
            RsCombo.Requery
            
            sHabitacion = ""
            sReserva = ""
            sPasajero = ""
            Pedido = ""
            sPedido = ""

            Cn.Execute "delete " & sDetalle
            Cn.Execute "delete " & sComboDetalle
            Cn.Execute "delete " & sComboPropiedad
            Cn.Execute "delete " & sProductoPropiedad
    
            RsDetalle.Requery
            RsComboPropiedad.Requery
            RsProductoPropiedad.Requery
            Inicializar
            Screen.MousePointer = vbDefault

        'OO Fin--------------------------------------------------------------------------------------------------------
                        
       Case Is = 17  'Cancelacion
       
            'JSJR 09/01/2025
            wAnulaDocumento = False
            
            sCodigo = ""
            sDescrip = ""
            fraEliminacion.Visible = False
            tabProducto.Visible = True
            ActivaCabecera True
            EliminarProductoCom = 0
   End Select
End Sub

Private Sub cmdOpcionPicture_Click(Index As Integer)
    If Index = 99 Or Index = 6 Then
        If wAgregaCombo Then
            nPos = RsCombo.AbsolutePosition
            Cn.Execute "Update " & sComboDetalle & " set tObservacion = '" & Trim(lblObservacion.Text) & "' where tItem ='" & sItem & "' and tItemCombo='" & xItem & "'"
        Else
            nPos = RsDetalle.AbsolutePosition
            Cn.Execute "Update " & sDetalle & " set tObservacion = '" & Trim(lblObservacion.Text) & "' where tItem ='" & sItem & "'"
        End If
    
        cmdOpcion_Click 6
    Else
        cmdOpcion_Click Index
    End If
    
End Sub

Private Sub cmdOperador_Click(Index As Integer)
   Dim i As Integer
   Screen.MousePointer = vbHourglass
   For i = 1 To 13
       cmdOperador(i).backColor = &HD1F2CA
   Next i
   RsOperador.MoveFirst
   RsOperador.Find "nboton = " & Trim(str(Index))
   nOperadorPropiedad = RsOperador!nControl
    xOperador = RsOperador!codigo

   cmdOperador(Index).backColor = &H8BDD79
   If wAgregaCombo Then
      AsignaComboPropiedad
   Else
      AsignaPropiedad
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOrigen_Click(Index As Integer)
' origen de ventas
   'HabilitaTimerColor (False)
   RsOrigenVentas.MoveFirst
   RsOrigenVentas.Find "boton = " & Trim(str(Index))
   vOrigenVentas = RsOrigenVentas!CodOrigenVenta
   'sMotorizado = RsMotorizado!codigo
   'HabilitaTimerColor (True)
   validarOrigenVentas
End Sub
Private Sub cmdPrecio_Click()
   Dim Acumulado As Double
   sTipo = ""
   frmNumPad.Show vbModal
   If wEnter And val(nPVenta) > 0 Then
      nPVenta = val(sDescrip)
      nOficial = nPVenta
      nDescuento = 0
      nRecargo = 0
      txtDPorcentaje.Caption = Format(0, "###,###,###,##0.00")
      txtRPorcentaje.Caption = Format(0, "###,###,###,##0.00")
      txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
      txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
      
       'extranjero bolivia
      Select Case pais ' ok
        Case "001" 'Bolivia
                Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = (Acumulado / 100)
                nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta * nPorcentaje3 / 100, 0)
        
        Case Else 'Peru, Ecuador
                Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = 1 + (Acumulado / 100)
                nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
      
      End Select
      nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
      txtNeto.Caption = Format(nPBase, "###,###,###,##0.00")
      txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
      txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
      txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
      txtOficial.Caption = Format(nOficial, "###,###,##0.00")
      txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
      txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
   End If
   txtBarra.SetFocus
End Sub

Private Sub cmdProducto_Click(Index As Integer)
   txtBarra.SetFocus
      
   RsProducto.MoveFirst
   RsProducto.Find "nbotonRapido = " & Trim(str(Index))
   sProducto = RsProducto!codigo
   
   If validadIngresoProducto(sProducto) = False Then
      Exit Sub
   End If
    
   If wAgregaCombo Then
      nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                        "FROM " & sComboDetalle & " WHERE tItem='" & sItem & "'", Cn)

      If nCCombo < nCombo * RsDetalle!nCantidad Then
         InsertaCombo sProducto
      Else
         MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
      End If
   Else
        If lBal And RsProducto!lBalanza Then
           Dim nResultado As Double
           nResultado = Pesar(nBalanzaPuerto)
           nResultado = Format(nResultado, "#,##0.00")
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
   End If
End Sub

Private Sub cmdProductoCombo_Click(Index As Integer)
    Dim xFiltro As String
    xFiltro = RsProducto.Filter
    RsProducto.Filter = adFilterNone
    txtBarra.SetFocus
    RsProducto.MoveFirst
    RsProducto.Find "tResumido = '" & cmdProductoCombo(Index).Caption & "'"
    sProducto = RsProducto!codigo
    
    nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                       "FROM " & sComboDetalle & " WHERE tItem='" & sItem & "'", Cn)
 
    If nCCombo < nCombo * RsDetalle!nCantidad Then
        'Oscar Ortega----------------------------------------------
        Dim oRsProductoDeCombo As Recordset
        Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sProducto)
        If oRsProductoDeCombo.RecordCount > 0 Then
            If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                'Obtener Suma de cantidades
                Dim nCantidadEnElCombo As Integer
                nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sItem, oRsProductoDeCombo!tEtiqueta)
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
    RsProducto.Filter = xFiltro
End Sub

Private Sub cmdSinBoton_Click()
    sTemp = ""
    If Calcular("Select count(*) as Codigo from tclienteproducto where tcodigodelivery='" & sCliente & "' ", Cn) = "0" Then
        Isql = "select * from vProducto where lActivo = 1 and (vProducto.tUnidadNegocio='" & sUnidadNegocio & "' )  and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " Order by Descripcion"
        frmBusquedaRapida.cmdOpcion(4).Visible = False
    Else
        Isql = "select vProducto.Grupo, vProducto.Descripcion , tclienteproducto.nprecio As nPrecioVenta , vProducto.nBoton, vProducto.SubGrupo, vProducto.Codigo from vProducto inner join tclienteproducto on vProducto.codigo=tclienteproducto.tcodigoproducto where vProducto.lActivo = 1 and (vProducto.tUnidadNegocio='" & sUnidadNegocio & "' ) Order by vProducto.Descripcion"
        frmBusquedaRapida.cmdOpcion(4).Visible = True
        frmBusquedaRapida.cmdOpcion(4).FontBold = True
    End If
    
    'Isql = "select * from vProducto where lActivo = 1 and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " Order by Descripcion"
    Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                    "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                    "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                    "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                    "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.cmdOpcion(0).Visible = True
    frmBusquedaRapida.Show vbModal
    
    If wEnter Then
       sProducto = sCodigo
        'INSUMOCRITICO23
        If validadIngresoProducto(sProducto) = False Then
            Exit Sub
        End If
        'INSUMOCRITICO23
 
       Dim xxx As String
       xxx = RsProducto.Filter
       RsProducto.Filter = adFilterNone
       RsProducto.MoveFirst
       RsProducto.Find ("Codigo='" & sProducto & "'")
    
       If Not RsProducto.EOF() Then
          If wAgregaCombo Then
             nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                               "FROM " & sComboDetalle & " WHERE tItem='" & sItem & "'", Cn)
        
             If nCCombo < nCombo * RsDetalle!nCantidad Then
                InsertaCombo sProducto
             Else
                MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
             End If
          Else
                          
             If lBal And RsProducto!lBalanza Then
                Dim nResultado As Double
                nResultado = Pesar(nBalanzaPuerto)
                nResultado = Format(nResultado, "#,##0.00")
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
          End If
       End If
       RsProducto.Filter = IIf(xxx = "0", "", xxx)
    End If
    txtBarra.SetFocus

End Sub

Private Sub cmdTipoDocumento_Click(Index As Integer)
   sCliente = ""
   Dim lprocesoSinDatos As Boolean
   lprocesoSinDatos = False
    If Index < 100 Then
        txtTipoDocumento.Caption = cmdTipoDocumento(Index).Caption
        cmdTipoDocumento(0).backColor = &HF2FAED
        cmdTipoDocumento(1).backColor = &HF2FAED
        cmdTipoDocumento(2).backColor = &HF2FAED
        cmdTipoDocumento(3).backColor = &HF2FAED
        cmdTipoDocumento(4).backColor = &HF2FAED
        cmdTipoDocumento(Index).backColor = &H8BDD79
        If Calcular(" select TOP 1  isnull(tcodigoexterno,'') Codigo from vTipodocumentoImpresora where tCaja = '" & sCaja & "' and Descripcion = '" & cmdTipoDocumento(Index).Caption & "' ", Cn) <> "" Then
            Exit Sub
        End If
    Else
         
        If cmdTipoDocumento(0).backColor = &H8BDD79 Then
            lprocesoSinDatos = True
            If pais = "000" Then
                txtTipoDocumento.Caption = Calcular(" select TOP 1  isnull(Descripcion,'') Codigo from vTipodocumentoImpresora where tCaja = '" & sCaja & "' and Prefijo ='B' and tCodigoExterno ='03' ", Cn)
            Else
                txtTipoDocumento.Caption = Calcular(" select TOP 1  isnull(Descripcion,'') Codigo from vTipodocumentoImpresora where tCaja = '" & sCaja & "' and Prefijo ='F' ", Cn)
            End If
            
            If sCodigoClienteGeneral = "" Or sCodigoClienteGeneral = "0" Then
                MsgBox "Codigo de cliente general no configurado en el INI.", vbInformation, sMensaje
                Exit Sub
            End If
            If txtTipoDocumento.Caption = "" Then
                MsgBox "No se encontro el documento configurado con los parametros para emision sin documento!!!.", vbInformation, sMensaje
                Exit Sub
            End If
        End If
    End If
    
    
   txtBarra.SetFocus
   
   nTotalPR = txtMonto.Caption

   If lObservacion And Trim(txtObservacion.Caption) = "" Then
      MsgBox "Debes ingresar la Observación", vbInformation, sMensaje
      cmdDetalle_Click (6)
      If Trim(txtObservacion.Caption) = "" Then
         Exit Sub
      End If
   End If
   
    'If (lCanalDelivery) Then
              If (lObligaOrigenVenta) Then
                          If (vOrigenVentas = "" Or vOrigenVentas = Null) Then
                              MsgBox "Debe Seleccionar un Origen de Venta", vbInformation, sMensaje
                              Call ActivarOrigenVentas
'                              fraOrigenVentas.Visible = True
                              Exit Sub
                          End If
                      End If
      'End If
             
       ' Miguel Mendoza - 10062019
    If Not VerificaComboMaximo() Then Exit Sub
    ' *************************
   
                Screen.MousePointer = vbHourglass
                                        
                     'insuimo2013
                    Cn.Execute "delete from " & sInsumoCombo
                    Dim X As Integer
                    Dim rstItems As New Recordset
                    Dim cadenaInsumos As String
                    Dim cadenaAEnviar As String
                    Dim cmdInsumo As New ADODB.Command
                    Dim resultado As String
                    Cn.Execute "insert into " & sInsumoCombo & " select sum(ncantidad) ncantidad, TPRODUCTO.TCODIGOINSUMO from " & sDetalle & " inner join tproducto on " & sDetalle & ".tcodigoproducto=tproducto.tcodigoproducto INNER JOIN  dbo.TINSUMO ON dbo.TPRODUCTO.tcodigoInsumo = dbo.TINSUMO.tcodigo  where lcontrolinsumocritico=1 and isnull(limprime,0)=0 AND ISNULL(TCODIGOINSUMO,''  )<>'' and (tinsumo.lactivo=1) group by tcodigoinsumo"
                    Cn.Execute "insert into " & sInsumoCombo & " sELECT     SUM(" & sDetalle & ".nCantidad * " & sComboDetalle & ".nCantidad) AS ncantidad, dbo.TPRODUCTO.tCodigoInsumo FROM  " & sComboDetalle & " INNER JOIN " & sDetalle & " ON  " & sComboDetalle & ".tItem = " & sDetalle & ".tItem INNER JOIN dbo.TINSUMO INNER JOIN dbo.TPRODUCTO ON dbo.TINSUMO.tcodigo = dbo.TPRODUCTO.tCodigoInsumo ON " & sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto WHERE     (dbo.TPRODUCTO.lControlInsumoCritico = 1) AND (dbo.TINSUMO.lactivo = 1) AND (ISNULL(" & sComboDetalle & ".lImprime, 0) = 0) AND (ISNULL(dbo.TPRODUCTO.tCodigoInsumo, N'') <> '') AND (" & sDetalle & ".lCombinacion = 1) GROUP BY dbo.TPRODUCTO.tCodigoInsumo "
                    cadenaInsumos = "select  SUM(ncantidad), tCodigoInsumo from " & sInsumoCombo & "  group by tCodigoInsumo order by 2"
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
                        If resultado <> "1" Then:  MsgBox "No hay cantidad disponible de : " & vbNewLine & resultado, vbExclamation, sMensaje: variableEmite = False: Screen.MousePointer = vbDefault: RsDetalle.Requery: Exit Sub
                    End If
                    variableEmite = True
                    'InsumosCriticos
                    
                    
   
   
   
   Label2.Caption = muestra
   Label2.Visible = False
   'Clave de Multi Cajero
    tUsuActua = sUsuario
    If lMultiCajero Then
      If Supervisor("16") = False Then
         MsgBox "Clave no permitida", vbExclamation, sMensaje
         Exit Sub
      End If
      sUsuario = sVar1
    End If
    sUsuario = tUsuActua
    'Chequea Descuento
    nTotalDescuento = CDbl(Calcular("select sum(nDescuento*nCantidad) as Codigo from " & sDetalle, Cn))
    If nTotalDescuento > 0 Then
       Dim nTope As Double
       Dim nTotalMes As Double
       Dim nConsumo As Double
       Dim aplicaTope As Boolean
       
       lAplicablePedido = Calcular("select lAplicablePedido as Codigo FROM vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
       nTope = Calcular("select nTope as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
       
         
       If nTope > 0 Then
          If Calcular("select lTopePedido as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn) Then
             If nTotalDescuento > nTope Then
                If MsgBox("El Descuento a aplicar Supera El Tope Registrado por Pedido" & Chr(13) & "¿Desea aplicar el Tope de " & sMonN & " " & nTope & "?", vbQuestion + vbYesNo) = vbYes Then
                   CalculaAplicaTope (nTope)
                Else
                   Exit Sub
                End If
             End If
          Else
             nTotalMes = Calcular("select sum(DPEDIDO.nDescuento*nCantidad) as Codigo FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                                  "WHERE month(MPEDIDO.fFecha) = month(getdate()) and year(MPEDIDO.fFecha)=year(getdate()) and mPedido.tDescuento='" & sCodigoDescuento & "' and tEstadoPedido<>'01' and tEstadoPedido<>'03'", Cn)
             
             If nTotalDescuento + nTotalMes > nTope Then
                If nTotalDescuento < nTope Then
                   If MsgBox("El Descuento a aplicar Supera El Tope Registrado dentro de un mes" & Chr(13) & "¿Desea aplicar el saldo " & sMonN & " " & nTope - nTotalMes & "?", vbQuestion + vbYesNo) = vbYes Then
                      CalculaAplicaTope (nTope - nTotalMes)
                   Else
                      Exit Sub
                   End If
                Else
                    MsgBox "El Descuento a aplicar Supera El Tope Registrado dentro de un mes", vbExclamation
                    RsDetalle.Requery
                    Exit Sub
                End If
             End If
          End If
       End If
    End If
    sCodigoDescuento = IIf(lAplicablePedido, "", sCodigoDescuento)
    nMonto = Calcular("select sum(nventa) as codigo from " & sDetalle & "", Cn)
    
    If Index > 100 And Index < 500 Then
        If nMonto > CDbl(Index - 100) Then
            MsgBox "El monto del pedido es mayor al pago seleccionado!!!", vbExclamation
            RsDetalle.Requery
            Exit Sub
        End If
    End If
    
    VisualizaMonto
    variableEmite = False
   
   'VALIDACION CANAL DE VENTA
   Dim rsCanalVentas As Recordset
   Dim lObligaMozo As Boolean
   Dim lObligaMotorizado As Boolean
   Dim lObligaClienteFrecuente As Boolean
   Dim lObligaFechaEntrega As Boolean
   Dim lObligaEntregarA As Boolean
   
   Set rsCanalVentas = Lib.OpenRecordset("select * from vTipoPedido", Cn)
   rsCanalVentas.Filter = "Codigo = '" & sTipoPedido & "'"
   
   lObligaMozo = IIf(IsNull(rsCanalVentas!lObligaMozo), False, rsCanalVentas!lObligaMozo)
   lObligaMotorizado = IIf(IsNull(rsCanalVentas!lObligaMotorizado), False, rsCanalVentas!lObligaMotorizado)
   lObligaClienteFrecuente = IIf(IsNull(rsCanalVentas!lObligaClienteFrecuente), False, rsCanalVentas!lObligaClienteFrecuente)
   lObligaFechaEntrega = IIf(IsNull(rsCanalVentas!lObligaIngresoFechaEntrega), False, rsCanalVentas!lObligaIngresoFechaEntrega)
   lObligaEntregarA = IIf(IsNull(rsCanalVentas!lObligaEntregarA), False, rsCanalVentas!lObligaEntregarA)
   
               'Obligatoriedad de Mozo
               If lObligaMozo Then
                  If lMCPV Then
                      sMozo = ObtenerCodigoMozo(sVar1)
                      If sMozo = "" Then
                         MsgBox "Configure un Cajero-Mesero", vbExclamation, sMensaje
                         Exit Sub
                      End If
'                  Else
'                      If sMozo = "" Or sMozo = "0000" Then
'                         MsgBox "Asigne al " & Mesero, vbExclamation, sMensaje
'                         RsDetalle.Requery
'                         Exit Sub
'                      End If
                  End If
               End If

               'Obligatoriedad de Cliente Frecuente
               If sClienteFrecuente = "" And lObligaClienteFrecuente Then
                  MsgBox "Asigne el Cliente Delivery", vbExclamation, sMensaje
                  Screen.MousePointer = vbDefault
                  cmdOpcion_Click (9)
                  Exit Sub
               End If
               
               'Obligatoriedad de Fecha de Entrega
               If Me.txtFechaEntrega.Caption = "" And lObligaFechaEntrega Then
                  MsgBox "Asigne la Fecha de Entrega", vbExclamation, sMensaje
                  Screen.MousePointer = vbDefault
                  cmdCabecera_Click (6)
                  Exit Sub
               End If
               
               'Entregar A
               If lObligaEntregarA = True And Me.txtEntregar.Caption = "" Then
                  MsgBox "Asigne información en Entregar A", vbExclamation, sMensaje
                  Screen.MousePointer = vbDefault
                  cmdDetalle_Click (14)
                  Exit Sub
               End If
               
    'CAJA CONTINGENCIA
        Dim scajacontingencia As Boolean
        Isql = "select isnull(lCajaContingencia,0) as lCajaContingencia From TCAJA Where tCaja = '" & sCaja & "'"
        Set Rscajacontigencia = Lib.OpenRecordset(Isql, Cn)
        scajacontingencia = Rscajacontigencia!lCajaContingencia
        If scajacontingencia = True Then
            frmFechaEmision.Show vbModal
            sfemision = frmFechaEmision.femision
        Else
            sfemision = FechaServidor()
        End If

    'FIN DE CAJA CONTINGENCIA

   Call Facturar(Index, lprocesoSinDatos)
   
        'INICIO GCAA 06072022 - CREDITO CORPORATIVO
        If (lActivaCreditoCoorporativo) Then
'             Dim xSaberCuentasxCobrar As String
'
'            xSaberCuentasxCobrar = Calcular("Select tEstadoDocumento as codigo from MDOCUMENTO Where tdocumento='" & sDocumento & "'", Cn)
'
'            If xSaberCuentasxCobrar = "03" Then
'               If Not INSERTAFE_CREDITO_CORPO(sDocumento) Then '----CABECERA
'               End If
'            End If
            If Not INSERTAFE_CREDITO_CORPO(sDocumento) Then '----CABECERA
            End If
        End If
        'FIN GCAA 06072022 - CREDITO CORPORATIVO
        
    
   If wEnter = True Then
        variableEmite = False
        Inicializar
        sPedido = ""
        If nPuerto > 0 Then
           Visor String(Int((19 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((19 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
        End If
   End If
   
End Sub

Private Sub cmdTipoDocumentoPicture_Click(Index As Integer)
    If Index = 1 Then
        cmdTipoDocumento_Click 105
    End If
    If Index = 2 Then
        cmdTipoDocumento_Click 110
    End If
    If Index = 3 Then
        cmdTipoDocumento_Click 120
    End If
    If Index = 4 Then
        cmdTipoDocumento_Click 150
    End If
    If Index = 5 Then
        cmdTipoDocumento_Click 200
    End If
    If Index = 6 Then
        cmdTipoDocumento_Click 999
    End If
End Sub


Private Sub Form_Activate()
   If txtBarra.Enabled = True Then
      txtBarra.SetFocus
   End If
End Sub
Private Function validarOrigenVentas()
' origen de ventas
    'lActivaMozo = IIf(IsNull(RsCanalesVenta!lActivaMozo), False, RsCanalesVenta!lActivaMozo)
    lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
    lCanalDelivery = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
    lCanalCentralPedidos = IIf(IsNull(RsCanalesVenta!lCanalCentralPedidos), False, RsCanalesVenta!lCanalCentralPedidos)
    'entregarA
    'lObligaEntregarA = IIf(IsNull(RsCanalesVenta!lObligaEntregarA), False, RsCanalesVenta!lObligaEntregarA)
    
    'origen de ventas
    'GCAA 16032021
     lOrigenVentas = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
 
    lObligaMotorizado = IIf(IsNull(RsCanalesVenta!lObligaMotorizado), False, RsCanalesVenta!lObligaMotorizado)
         
    If lActivaMotorizado Then
        Me.fraMorotizado.Visible = True
    Else
        Me.fraMorotizado.Visible = False
    End If
    
'    If lOrigenVentas Then
'        Me.fraOrigenVentas.Visible = True
'    Else
        'GCAA 16032021
        Me.fraOrigenVentas.Visible = False
'    End If
    
    
    
    
End Function

Private Sub OrganizaComponentes()
    On Error GoTo fin
    
        frmCajaRapida.Width = 15585
        frmCajaRapida.Height = 11540
        
        'FRAME ELIMINACION
        fraEliminacion.Height = 9015
        fraEliminacion.Width = 6226.194
        fraEliminacion.Top = 1440
        fraEliminacion.Left = 9241.771
        'FRAME ELIMINACION
        
        'FRAME PROPIEDADES
        fraPropiedad.Height = 10215
        fraPropiedad.Width = 6226.194
        fraPropiedad.Top = 240
        fraPropiedad.Left = 9241.771
        'FRAME PROPIEDADES
        
        'FRAME MOTORIZADO
        fraMorotizado.Height = 9015
        fraMorotizado.Width = 6226.194
        fraMorotizado.Top = 1440
        fraMorotizado.Left = 9241.771
        'FRAME MOTORIZADO
        
        'FRAME MESERO
        fraMozo.Height = 9015
        fraMozo.Width = 6226.194
        fraMozo.Top = 1440
        fraMozo.Left = 9241.771
        'FRAME MESERO
        
        'FRAME ORIGEN DE VENTAS
        fraOrigenVentas.Height = 9015
        fraOrigenVentas.Width = 6226.194
        fraOrigenVentas.Top = 1440
        fraOrigenVentas.Left = 9241.771
        'FRAME ORIGEN DE VENTAS
        
        'FRAME PUNTO VENTA
        fraPuntoVenta.Height = 9015
        fraPuntoVenta.Width = 6226.194
        fraPuntoVenta.Top = 1440
        fraPuntoVenta.Left = 9241.771
        'FRAME PUNTO VENTA
        
        'FRAME PRODUCTO COMBO
        fraProductoCombo.Height = 11325
        fraProductoCombo.Width = 6226.193
        fraProductoCombo.Top = 240
        fraProductoCombo.Left = 9241.771
        'FRAME PRODUCTO COMBO
        
        'FRAME DET COMBO
        fraCombo.Height = 9975
        fraCombo.Width = 9256.773
        fraCombo.Top = 1560
        fraCombo.Left = 0
        'FRAME DET COMBO
    Exit Sub
fin:
    MsgBox "Error" & Error
End Sub
Private Sub Form_Load()
    
    Call OrganizaComponentes
    
    If lOCR Then
        cmdOpcion(2).Enabled = False
        cmdOpcion(2).Visible = False
    End If
    
    sPedido = ""
    'origen de ventas
    
    Me.fraOrigenVentas.Visible = False
    Me.fraMorotizado.Visible = False
    
    '--------------------------------------------

    'InsumosCriticos ' 23 2013
   sInsumoCombo = dbTemporal(sCaja, 2, "nCantidad", "float", _
                                    "tCodigoInsumo", "nVarChar(20)")
   
   sDetalle = dbTemporal(sCaja, 45, "tCodigoPedido", "nVarChar(10)", _
                                    "tItem", "nVarChar(3)", _
                                    "tTipoPedido", "nVarChar(2)", _
                                    "tCodigoProducto", "nVarChar(7)", _
                                    "tCodigoGrupo", "nVarChar(2)", _
                                    "tCodigoSubGrupo", "nVarChar(4)", _
                                    "tMoneda", "nVarChar(3)", _
                                    "nPrecioNeto", "Float", _
                                    "nPrecioImpuesto1", "Float", "nPrecioImpuesto2", "Float", "nPrecioImpuesto3", "Float", _
                                    "nPrecioVenta", "Float", _
                                    "nRecargo", "Float", "nDescuento", "Float", _
                                    "nPrecioOficial", "Float", _
                                    "nCantidad", "Float", _
                                    "nImpuesto1", "Float", "nImpuesto2", "Float", "nImpuesto3", "Float", _
                                    "nVenta", "Float", _
                                    "tObservacion", "nVarChar(255)", _
                                    "tCortesia", "nVarChar(4)", _
                                    "lImprime", "Bit", _
                                    "tEstadoItem", "nVarChar(3)", _
                                    "tArea", "nVarChar(3)", _
                                    "lCombinacion", "Bit", "nCombinacion", "Smallint", "lImprimeArea", "Bit", _
                                    "tFacturado", "nVarChar(1)", "tDocumento", "nVarChar(20)", "lTransferido", "Bit", "tComanda", "nVarChar(10)", _
                                    "nInsumo", "Float", "nGasto", "Float", "nManoObra", "Float", "nOrden", "int", "lCorte", "bit", "Estado", "nVarChar(1)", "toferta", "nvarchar(5)", "tautorizaoferta", "nvarchar(15)", "tSubAlmacen", "nvarchar(6)", "tCodigoEtiqueta", "nvarchar(50)", "fenvio", "datetime", "nenvio", "int", "tCajaD", "nvarchar(3)")
      
   Centrar Me
   Dim sTemp1 As String
   Dim sTemp2 As String
   Dim sTemp3 As String
      
   muestra = Label2.Caption
   Label2.Visible = False
      
   cmdCabecera(1).Caption = sBoton1
   cmdCabecera(2).Caption = sBoton2
   cmdCabecera(3).Caption = sBoton3
   cmdCabecera(4).Caption = sBoton4
   cmdCabecera(5).Caption = sBoton5
      
    If sBoton1 = "" Then
        cmdCabecera(1).Enabled = False
        cmdCabecera(1).Caption = "N/D"
    End If
    If sBoton2 = "" Then
        cmdCabecera(2).Enabled = False
        cmdCabecera(2).Caption = "N/D"
    End If
    If sBoton3 = "" Then
        cmdCabecera(3).Enabled = False
        cmdCabecera(3).Caption = "N/D"
    End If
    If sBoton4 = "" Then
        cmdCabecera(4).Enabled = False
        cmdCabecera(4).Caption = "N/D"
    End If
    If sBoton5 = "" Then
        cmdCabecera(5).Enabled = False
        cmdCabecera(5).Caption = "N/D"
    End If

   If lBloqueoCanal1 Then
    cmdCabecera(1).Enabled = False
   End If
    If lBloqueoCanal2 Then
    cmdCabecera(2).Enabled = False
   End If
    If lBloqueoCanal3 Then
    cmdCabecera(3).Enabled = False
   End If
    If lBloqueoCanal4 Then
    cmdCabecera(4).Enabled = False
   End If
   If lBloqueoCanal5 Then
    cmdCabecera(5).Enabled = False
   End If



   Select Case sTipoPedidoPD
        Case Is = "01"
            sTipoPedido = "01"
            If sBoton1 = "" Then
                cmdCabecera(1).Enabled = False
                cmdCabecera(1).Caption = "N/D"
            Else
                LlenaColorCanalVenta 1
            End If
            vOrigenVentas = sDefaultOrigenVenta

        Case Is = "02"
            sTipoPedido = "02"
            If sBoton2 = "" Then
                cmdCabecera(2).Enabled = False
                cmdCabecera(2).Caption = "N/D"
            Else
                LlenaColorCanalVenta 2
            End If
            vOrigenVentas = sDefaultOrigenVenta
        Case Is = "03"
            sTipoPedido = "03"
            If sBoton3 = "" Then
                cmdCabecera(3).Enabled = False
                cmdCabecera(3).Caption = "N/D"
            Else
                LlenaColorCanalVenta 3
            End If
            vOrigenVentas = sDefaultOrigenVenta
        Case Is = "04"
            sTipoPedido = "04"
            If sBoton4 = "" Then
                cmdCabecera(4).Enabled = False
                cmdCabecera(4).Caption = "N/D"
            Else
                LlenaColorCanalVenta 4
            End If
            vOrigenVentas = sDefaultOrigenVenta
        Case Is = "05"
            sTipoPedido = "05"
            If sBoton5 = "" Then
                cmdCabecera(5).Enabled = False
                cmdCabecera(5).Caption = "N/D"
            Else
                LlenaColorCanalVenta 5
            End If
            vOrigenVentas = sDefaultOrigenVenta
        Case Else
   End Select

   nOperadorPropiedad = 0
            
   sTemp1 = Calcular("select tDetallado as Codigo from TTABLA where tTabla='ETIQUETA' and tCodigo='01'", Cn)
   sTemp2 = Calcular("select tDetallado as Codigo from TTABLA where tTabla='ETIQUETA' and tCodigo='02'", Cn)
   sTemp3 = Calcular("select tDetallado as Codigo from TTABLA where tTabla='ETIQUETA' and tCodigo='03'", Cn)
   cmdEtiqueta(1).Caption = IIf(sTemp1 = "0", "", sTemp1)
   cmdEtiqueta(2).Caption = IIf(sTemp2 = "0", "", sTemp2)
   cmdEtiqueta(3).Caption = IIf(sTemp3 = "0", "", sTemp3)
   sMozo = ""
  
   'Operador
   Isql = "select * from vOperador where lActivo = 1 order by Codigo"
   Set RsOperador = Lib.OpenRecordset(Isql, Cn)
      
   'Propiedades
   sProductoPropiedad = dbTemporal(sCaja, 11, "tItem", "nVarChar(3)", _
                                             "tCodigoPropiedad", "nVarChar(4)", _
                                             "tProducto", "nVarChar(7)", _
                                             "tEnlace", "nVarChar(7)", _
                                             "nInsumo", "float", _
                                             "nGasto", "float", _
                                             "nManoObra", "float", _
                                             "nCantidad", "float", _
                                             "nInsumoUnitario", "float", _
                                             "nGastounitario", "float", _
                                             "nManoObraUnitario", "float")
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
                      "FROM " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.DRECETAVENTA  t1x INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto " & _
                      "Where lNoDescargo = 1"
            Else
                xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                      "nInsumo, nGasto, nManoObra, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                      "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                      "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                      "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!descripcion & "' as Operador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0,0 " & _
                      "FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA  t1x INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto " & _
                      "Where lNoDescargo = 1"
            End If
      Else
         xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
      End If
   Else
         xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
   End If
   
   
    If BDLink = "1" Then
        Set RsPropiedad = Lib.OpenRecordset("usp_Inforest_ObtienePropiedad '" & IIf(lAlmacen = True, 1, 0) & "','" & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & "','" & sLocal & "'", Cn)
    Else
        Set RsPropiedad = Lib.OpenRecordset("usp_Inforest_ObtienePropiedad '" & IIf(lAlmacen = True, 1, 0) & "','" & sAlmacenMDB & "','" & sLocal & "'", Cn)
    End If
   
   'Set RsPropiedad = Lib.OpenRecordset(xSql, Cn)
                      
   
   Isql = "SELECT [" & sDetalle & "].*, dbo.TPRODUCTO.tDetallado AS Producto, dbo.vCortesia.Descripcion AS Cortesia, dbo.TPRODUCTO.lDescuento AS lDescuento, CASE [" & sDetalle & "].nDescuento WHEN 0 THEN 0 ELSE [" & sDetalle & "].nDescuento * 100 / [" & sDetalle & "].nPrecioOficial END AS Descuento, " & _
          "dbo.TPRODUCTO.lModificable AS lModificable, CONVERT(bit, ISNULL(DATALENGTH([" & sDetalle & "].tObservacion), 0)) AS lObservacion, ISNULL(T1.nPropiedad, 0) AS lPropiedad, 'X' as Eliminar " & _
          "FROM [" & sDetalle & "] LEFT OUTER JOIN (SELECT tItem, CASE WHEN COUNT(tProducto) > 0 THEN 1 ELSE 0 END AS nPropiedad FROM [" & sProductoPropiedad & "] Group by tItem) T1 " & _
          "ON [" & sDetalle & "].tItem = T1.tItem LEFT OUTER JOIN dbo.vCortesia ON [" & sDetalle & "].tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN " & _
          "dbo.TPRODUCTO ON [" & sDetalle & "].tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
          "ORDER BY [" & sDetalle & "].tItem"
   Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
            
   'Producto Propiedad
   
   Isql = "SELECT " & sProductoPropiedad & ".tItem, " & sProductoPropiedad & ".tCodigoPropiedad, " & sProductoPropiedad & ".tProducto, t1.Descripcion AS Descripcion, Operador, " & sProductoPropiedad & ".ncantidad  " & _
          "FROM " & sProductoPropiedad & " INNER JOIN (" & xSql & ") T1 ON " & sProductoPropiedad & ".tCodigoPropiedad = T1.Codigo and " & sProductoPropiedad & ".tProducto = T1.tProducto and " & sProductoPropiedad & ".tenlace= t1.tenlace "
   Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
            
   'Combo Propiedad
    sComboPropiedad = dbTemporal(sCaja, 12, "tItem", "nVarChar(3)", _
                                          "tItemCombo", "nVarChar(3)", _
                                          "tCodigoPropiedad", "nVarChar(4)", _
                                          "tProducto", "nVarChar(7)", _
                                          "tEnlace", "nVarChar(7)", _
                                          "nInsumo", "float", _
                                          "nGasto", "float", _
                                          "nManoObra", "float", _
                                          "nCantidad", "float", _
                                          "nInsumoUnitario", "float", _
                                          "nGastoUnitario", "float", _
                                          "nManoObraUnitario", "float")

   Isql = "SELECT " & sComboPropiedad & ".tItem, " & sComboPropiedad & ".tItemCombo, T1.Descripcion, T1.Operador , " & sComboPropiedad & ".ncantidad " & _
          "FROM " & sComboPropiedad & " INNER JOIN (" & xSql & ") T1 ON " & sComboPropiedad & ".tCodigoPropiedad = T1.Codigo AND " & sComboPropiedad & ".tProducto = T1.tProducto AND " & sComboPropiedad & ".tEnlace = T1.tEnlace "
   Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
         
   'Combos
   Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto AS Codigo, dbo.TPRODUCTO.tResumido AS Descripcion ,ISNULL(TCOMBO.NVALOR,-2147483633) NVALOR " & _
          "FROM dbo.TCOMBO INNER JOIN dbo.TPRODUCTO ON dbo.TCOMBO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
          "where lActivo=1 ORDER BY TCOMBO.TETIQUETA,dbo.TPRODUCTO.tResumido  "
   Set RsProductoCombo = Lib.OpenRecordset(Isql, Cn)
         
   'Producto
   Isql = "select * from vProducto where lActivo = 1 Order by nBoton"
   Set RsProducto = Lib.OpenRecordset(Isql, Cn)

   'Areas
   Set RsArea = Lib.OpenRecordset("select * from vAreaImpresora where tCaja ='" & sCaja & "'", Cn)
                                    
   'Mozos
   Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by nBoton"
   Set RsMozo = Lib.OpenRecordset(Isql, Cn)
   AsignaBoton 19, RsMozo, cmdMozo()
   
   
      'Origen de ventas
   Isql = "select * from vOrigenVenta where Activo = 1 and Visible = 1 and CodOrigenVenta not in (select CodigoOrigenVenta from TCAJAORIGEN_BLOQUEO  where tcaja='" & sCaja & "' ) Order by Boton"
   Set RsOrigenVentas = Lib.OpenRecordset(Isql, Cn)
   
   Isql = "select * from vTipoPedido where Codigo = '02'"
   Set RscanalOrigenVentas = Lib.OpenRecordset(Isql, Cn)
   
   AsignaBotonOrigenVentas 19, RsOrigenVentas, Me.cmdOrigen()
   Set RsCanalesVenta = Lib.OpenRecordset("select * from TCANALVENTA", Cn)
   
   'Motorizado
   Isql = "select * from vMotorizado where lActivo = 1 Order by nBoton"
   Set RsMotorizado = Lib.OpenRecordset(Isql, Cn)
   AsignaBoton 19, RsMotorizado, cmdMotorizado()
   '-----------------------------------------------------
   
   
   'Motivo de Eliminacion
   Isql = "select * from vMotivoEliminacion where lActivo = 1 order by Codigo"
   Set RsMotivoEliminacion = Lib.OpenRecordset(Isql, Cn)
   AsignaComando 38, RsMotivoEliminacion, cmdEliminacion()
   
   'Tipo de Documentos
'   If pais = "002" Then 'Ecuador
'      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' order by tTipoEmision", Cn)
'   Else
'      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 order by tTipoEmision", Cn)
'   End If

   If pais = "002" Then 'Ecuador
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' And lNotaCredito = 0 And lActivo = 1 UNION Select * From vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lFacturacionElectronica=1 and lActivo =1 order by tTipoEmision", Cn)
   Else
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lActivo = 1 order by tTipoEmision", Cn)
   End If
   
   'LlenaProducto "01"
   If (lActivaTresDecimales) Then
    Call ConfGrilla(10, grdDetalle, "Or", 2, "nOrden", 300, 1, 0, "#0", _
                                   "-", 2, "lCorte", 250, 0, 4, "", _
                                   "Producto", 2, "Producto", 3370, 0, 0, "", _
                                   "Precio", 2, "nPrecioVenta", 800, 1, 0, "###,###,##0.00", _
                                   "Cant.", 2, "nCantidad", 750, 1, 0, "#,##0.000", _
                                   "SubTotal", 2, "nVenta", 1000, 1, 0, "###,###,##0.00", _
                                   "Eliminar", 2, "Eliminar", 900, 2, 0, "###,###,##0.00", _
                                   "E", 2, "lImprime", 250, 0, 4, "", _
                                   "P", 2, "lPropiedad", 250, 0, 4, "", _
                                   "O", 2, "lObservacion", 250, 0, 4, "")
   Else
      Call ConfGrilla(10, grdDetalle, "Or", 2, "nOrden", 300, 1, 0, "#0", _
                                  "-", 2, "lCorte", 250, 0, 4, "", _
                                  "Producto", 2, "Producto", 3370, 0, 0, "", _
                                  "Precio", 2, "nPrecioVenta", 800, 1, 0, "###,###,##0.00", _
                                  "Cant.", 2, "nCantidad", 750, 1, 0, "#,##0.00", _
                                  "SubTotal", 2, "nVenta", 1000, 1, 0, "###,###,##0.00", _
                                  "Eliminar", 2, "Eliminar", 900, 2, 0, "###,###,##0.00", _
                                  "E", 2, "lImprime", 250, 0, 4, "", _
                                  "P", 2, "lPropiedad", 250, 0, 4, "", _
                                  "O", 2, "lObservacion", 250, 0, 4, "")
  
   End If
   
   grdDetalle.Columns(0).DividerStyle = dbgNoDividers
   grdDetalle.Columns(1).DividerStyle = dbgNoDividers
   grdDetalle.Columns(2).DividerStyle = dbgNoDividers
   grdDetalle.Columns(3).DividerStyle = dbgNoDividers
   grdDetalle.Columns(4).DividerStyle = dbgNoDividers
   grdDetalle.Columns(5).DividerStyle = dbgNoDividers
   grdDetalle.Columns(6).DividerStyle = dbgNoDividers
   grdDetalle.Columns(7).DividerStyle = dbgNoDividers
   grdDetalle.Columns(8).DividerStyle = dbgNoDividers
   grdDetalle.Columns(9).DividerStyle = dbgNoDividers
   
   grdDetalle.Columns(6).ForeColor = vbRed
   grdDetalle.Columns(6).Font.Bold = True
   grdDetalle.Columns(6).Font.Size = 12
   
   Set grdDetalle.DataSource = RsDetalle
   nMonto = 0
   txtMonto.Caption = Format(nMonto, "###,##0.00")
         
   'Asigna Operador
   AsignaBoton 13, RsOperador, cmdOperador()
   If RsOperador.RecordCount > 0 Then
      RsOperador.MoveFirst
      If Not IsNull(RsOperador!nBoton) And RsOperador!nBoton > 0 Then
         cmdOperador_Click (RsOperador!nBoton)
      End If
   End If
         
   With RsTipoDocumento
       If .RecordCount > 0 Then
          .MoveFirst
          For i = 1 To IIf(.RecordCount >= 4, 4, .RecordCount)
              cmdTipoDocumento(i).Visible = True
              cmdTipoDocumento(i).Caption = !descripcion
              .MoveNext
          Next i

          For i = .RecordCount + 1 To 4
              cmdTipoDocumento(i).Visible = False
          Next i
       Else
          For i = 1 To 4
              cmdTipoDocumento(i).Visible = False
          Next i
       End If
  End With
  
  If Not lPrinter Then
     cmdOpcionPicture(8).Visible = False
  End If
  
   'Combo
   Call ConfGrilla(8, grdCombo, "-", 2, "lCorte", 350, 0, 4, "", _
                                "Producto", 2, "Producto", 4250, 0, 0, "", _
                                "Cant.", 2, "nCantidad", 850, 1, 0, "#,##0.00", _
                                "Eliminar", 2, "Eliminar", 1200, 2, 0, "#,##0.00", _
                                "E", 2, "lImprime", 350, 0, 4, "", _
                                "P", 2, "lPropiedad", 350, 0, 4, "", _
                                "O", 2, "lObservacion", 350, 0, 4, "", _
                                "Ord", 2, "nOrden", 400, 2, 0, "#0")

   grdCombo.Columns(0).DividerStyle = dbgNoDividers
   grdCombo.Columns(1).DividerStyle = dbgNoDividers
   grdCombo.Columns(2).DividerStyle = dbgNoDividers
   grdCombo.Columns(3).DividerStyle = dbgNoDividers
   grdCombo.Columns(4).DividerStyle = dbgNoDividers
   grdCombo.Columns(5).DividerStyle = dbgNoDividers
   grdCombo.Columns(6).DividerStyle = dbgNoDividers
   grdCombo.Columns(7).DividerStyle = dbgNoDividers
   
   grdCombo.Columns(3).ForeColor = vbRed
   grdCombo.Columns(3).Font.Bold = True
   grdCombo.Columns(3).Font.Size = 12
   
   sComboDetalle = dbTemporal(sCaja, 24, "tCodigoPedido", "nVarchar(10)", "tItem", "nVarchar(3)", "tItemCombo", "nVarchar(3)", "tProducto", "nVarchar(7)", "tProductoCombo", "nVarchar(7)", "nCantidad", "float", "tCodigoGrupo", "nVarchar(2)", "tCodigoSubGrupo", "nVarchar(4)", _
                                         "nPrecioNeto", "float", "nImpuesto1", "float", "nImpuesto2", "float", "nImpuesto3", "float", "nVenta", "float", "nInsumo", "float", "nGasto", "float", "nManoObra", "float", "lImprimeArea", "bit", "lImprime", "bit", "nOrden", "int", "tObservacion", "nVarchar(250)", "lCorte", "bit", "lAtendidoC", "BIT", "fAtendidoC", "DATETIME", "tUsuarioAtendio", "nvarchar(15)")
   
   Isql = "SELECT dbo." & sComboDetalle & ".tProducto, dbo." & sComboDetalle & ".tItem, dbo." & sComboDetalle & ".tItemCombo, dbo." & sComboDetalle & ".tProductoCombo, dbo." & sComboDetalle & ".nCantidad, dbo." & sComboDetalle & ".tCodigoGrupo, dbo." & sComboDetalle & ".tCodigoSubGrupo, dbo.TPRODUCTO.tDetallado AS Producto, " & _
          "dbo.MPEDIDO.tEstadoPedido, dbo.MPEDIDO.tCaja, dbo." & sComboDetalle & ".lImprimeArea, dbo." & sComboDetalle & ".lImprime, dbo." & sComboDetalle & ".nOrden, CONVERT(bit,ISNULL(DATALENGTH(dbo." & sComboDetalle & ".tObservacion), 0)) AS lObservacion, ISNULL(T1.nPropiedad, 0) AS lPropiedad, dbo." & sComboDetalle & ".tObservacion, dbo." & sComboDetalle & ".lCorte , 'X' as Eliminar " & _
          "FROM dbo." & sComboDetalle & " LEFT OUTER JOIN (SELECT tItem, tItemCombo, CASE WHEN COUNT(tProducto) > 0 THEN 1 ELSE 0 END AS nPropiedad From " & sComboPropiedad & " " & _
          "GROUP BY tItem, tItemCombo) AS T1 ON dbo." & sComboDetalle & ".tItemCombo = T1.tItemCombo AND dbo." & sComboDetalle & ".tItem = T1.tItem LEFT OUTER JOIN dbo.TPRODUCTO ON dbo." & sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto LEFT OUTER JOIN dbo.MPEDIDO ON dbo." & sComboDetalle & ".tCodigoPedido = dbo.MPEDIDO.tCodigoPedido"
                
   Set RsCombo = Lib.OpenRecordset(Isql, Cn)
   Set grdCombo.DataSource = RsCombo
  
  Impuesto
  
  If sMozo = "" Then
     sMozo = "0000"
     txtTitulo.Caption = " Caja Rápida : " & sCaja & " " & Mesero & ": Sin " & Mesero & IIf(sPedido = "", "", " - " & sPedido)
  Else
     cmdDetallePicture(9).Enabled = False
     'txtTitulo.Caption = " Caja Rápida : " & sCaja & " Mesero: " & Calcular("select descripcion as codigo from vMozo where Codigo='" & sMozo & "'", Cn) & IIf(sPedido = "", "", " - " & sPedido)
     txtTitulo.Caption = " Caja Rápida : " & sCaja & Mesero & ": " & Calcular("select descripcion as codigo from vMozo where Codigo='" & sMozo & "'", Cn) & IIf(sPedido = "", "", " - " & sPedido)
  End If
  
  fraMozo.Visible = False
  fraDetalle.Visible = False
  fraPropiedad.Visible = False
  fraEliminacion.Visible = False
  fraPuntoVenta.Visible = False
  fraProductoCombo.Visible = False
  Pedido = ""
  sObser = ""
  txtObservacion.Caption = sObser
  wCombo = False
  wAgregaCombo = False
  nCombo = 0
  sComandaInfhotel = ""
  Sw = False
  lPropiedad = False
  Set RsCajaRapida = Lib.OpenRecordset("select tCodigo, substring(tCodigo,1,1) as Prefijo, tDetallado, nValor from TTABLA where tTabla='CAJARAPIDA'", Cn)

  If lInfhotel Then
     'Moneda Base
     sMonedaBase = Calcular("select tMoneda as Codigo from TPARAMETRO", CnInfhotel)
     
     'Punto de Venta
     Isql = "Select tPuntoVenta as Codigo, tDescripcion as Descripcion, nUltimoComanda, tmoneda" & _
            " From tPuntoVenta " & _
            " where tHotel='" & sHotel & "' AND lActivo=1 and lInforest=1"
     Set rsPuntoVenta = Lib.OpenRecordset(Isql, CnInfhotel)
     AsignaComando 9, rsPuntoVenta, cmdPunto()
     cmdCabeceraPicture(0).Visible = True
     
     For i = 1 To 9
         cmdPunto(i).FontBold = False
     Next i
    
     For i = 1 To 9
         If cmdPunto(i).Caption = Calcular("select tDescripcion as codigo from tPuntoventa where tPuntoVenta='" & sPuntoVentaInfhotel & "'", CnInfhotel) Then
           cmdPunto_Click (i)
        End If
    Next i
    sPuntoVenta = sPuntoVentaInfhotel
     
  End If
  
  If lMultiCajero Then
      cmdOpcionPicture(2).Enabled = False
  End If
  
  If lInfhotel Then
      cmdOpcionPicture(14).Visible = True
  End If
  
  Dim RsEtiquetaDefault As Recordset
  Set RsEtiquetaDefault = Lib.OpenRecordset("select top 1 isnull(tCodigoAgrupadorCajaRapida,'') as tCodigoAgrupadorCajaRapida,isnull(lActivaAgrupadorCajaRapida,0) as lActivaAgrupadorCajaRapida from TCAJA where tCaja = '" & sCaja & "'", Cn)
  
  If RsEtiquetaDefault.RecordCount = 1 Then
    If RsEtiquetaDefault!lActivaAgrupadorCajaRapida Then
        Select Case RsEtiquetaDefault!tCodigoAgrupadorCajaRapida
            Case "01"
                cmdEtiqueta_Click (1)
            Case "02"
                cmdEtiqueta_Click (2)
            Case "03"
                cmdEtiqueta_Click (3)
            Case Else
                cmdEtiqueta_Click (1)
        End Select
    Else
        cmdEtiqueta_Click (1)
    End If
  Else
    cmdEtiqueta_Click (1)
  End If
  
  
  TxtDescuento.Caption = "0.00"
  If sLey1 <> "" Then
    lblLey1.Visible = True
    lblLey1.Caption = sLey1 & " 0.00"
  End If
  Screen.MousePointer = vbDefault
  
 
End Sub

Public Sub InsertaProducto(xCantidad As Double)
    Screen.MousePointer = vbHourglass
    Dim nValor As Double
    Dim precioventa As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    Dim nOrden As Integer
    Dim RsOrd As Recordset
        
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
    
'    If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
'        MsgBox "Se ha aplicado Anticipo al Pedido!!!, no se puede Ingresar mas productos!!", vbInformation, sMensaje
'         Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    
    
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
                        Label2.Caption = " Insumo Crítico--> " & rsInsumo.Fields(1) & " =  Stock: " & str(rsInsumo.Fields(2)) & "     Solicitado: " + str(xCantidad)
                        Label2.Visible = True
                    End If
    Else
        Label2.Caption = muestra
        Label2.Visible = False
    End If
    'INSUMOCRITICO
    
    sItem = Lib.Correlativo(Calcular("select max(tItem) as codigo from [" & sDetalle & "]", Cn), 3)
    If RsDetalle.RecordCount = 0 Then
       'sitem = "001"
       nOrden = 1
    Else
       'sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from [" & sDetalle & "]", Cn), 3)
       If lOrden Then
          Set RsOrd = Lib.OpenRecordset("select nOrden, lImprime from " & sDetalle & " Order by nOrden DESC", Cn)
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
    nPVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
    
    'CPvalicacion LG
    If lClienteControlaProducto Then
    'gcaa 28/02/2020
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
    sCriterio = "tCodigoProducto ='" & sProducto & "' and lActivo=1"
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
       
       
    Isql = "select * from TOFERTA where tCodigoProducto ='" & sProducto & "' and lActivo=1  "
    Set RsExisteOferta = Lib.OpenRecordset(Isql, Cn)
    
       
       
    Isql = "select * from TOFERTA where " & sCriterio
    Set RsOferta = Lib.OpenRecordset(Isql, Cn)
        
    'inserta descto
'    nOferta = 0
'    lAcumulable = False
'    If RsOferta.RecordCount > 0 And RsProducto!lDescuento Then
'       RsOferta.MoveFirst
'       If RsOferta!nRatio > 0 Then
'          nOferta = nOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
'       Else
'          nOferta = nOficial - IIf(IsNull(RsOferta!nMonto), 0, RsOferta!nMonto)
'       End If
'    End If
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
'                 Else
'                    MsgBox "Fuera de rango" & Chr(13) & "Consulte con el Manager", vbCritical, sMensaje
'                    Exit Sub
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
        'CPvalicacion LG
    
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
'
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
    lProductoMultiArea = Calcular("select isnull(lmultiarea,0) as codigo from tproducto where tcodigoproducto='" & RsProducto.Fields("codigo") & "'", Cn)
    
    If lProductoMultiArea = False Then
        tsubalmacen = ""
    Else
        tsubalmacen = ""
        If lMultiAreaSubGrupo = True Then
            tAreaProduccion = Calcular("select isnull(tarea,'') codigo from TAREASUBGRUPO where tcaja='" & sCaja & "' and tSubGrupo='" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "'", Cn)
            tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
        End If
        If lMultiAreaCaja = True Then
            tAreaProduccion = Calcular("select isnull(tsubalmacen,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
            tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
        End If
        
    
        
        If tsubalmacen = "0" Then
            tsubalmacen = ""
        End If
    
    End If
    
    
   '' llena " & sDetalle & " 07102020
    fxCombo "A", 1, sProducto
    Isql = "insert into [" & sDetalle & "] " & _
           "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
           "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
           "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem,tsubalmacen,toferta,tCajaD) " & _
           "Values( '" & Pedido & "', '" & sTipoPedido & "', " _
                   & "'" & sItem & "', " _
                   & "'" & sProducto & "', " _
                   & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                   & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                   & nPBase & ", " & nRecargo & ", " _
                   & nDescuento & ", " _
                   & nOficial & ", " _
                   & nImpuesto1 & ", " & nImpuesto2 & ", " & nImpuesto3 & ", " _
                   & nPVenta & ", " & xCantidad & ", " _
                   & nPVenta * xCantidad & ", " _
                   & nImpuesto1 * xCantidad & ", " & nImpuesto2 * xCantidad & ", " & nImpuesto3 * xCantidad & ", " _
                   & "0, '" & RsProducto!tArea & "', " & IIf(RsProducto!lImprimeArea, -1, 0) & "," _
                   & IIf(RsProducto!lCombinacion, 1, 0) & ", " & RsProducto!nCombinacion & ", " _
                   & nInsumo & ", " _
                   & nGasto & ", " _
                   & nMObra & ", " _
                   & nOrden & ", " _
                   & "'N','" & tsubalmacen & "','" & tOferta & "', '" & sCaja & "') "
    Cn.Execute Isql
    RsDetalle.Requery
    nMonto = Format(Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn), "#,###,##0.00")
    RsDetalle.MoveLast
    txtcantproductos.Text = RsDetalle.RecordCount & " Productos"
    Screen.MousePointer = vbDefault

    If (RsProducto!lCombinacion) And (RsProducto!lComboVisible) Then
        lComboActivo = True
    End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If lBal Then
'       frmMsComm.MSCommBalanza.PortOpen = False
'    End If
   
   Cn.Execute "drop table " & sDetalle
   Cn.Execute "drop table " & sComboDetalle
   Cn.Execute "drop table " & sComboPropiedad
   Cn.Execute "drop table " & sProductoPropiedad
   Set frmCajaRapida = Nothing
End Sub

Public Sub AsignaProducto()
   If Not RsDetalle.EOF Then
      cmdPrecio.Enabled = IIf(IsNull(RsDetalle!lModificable), False, RsDetalle!lModificable)
      sProducto = IIf(IsNull(RsDetalle!tCodigoProducto), "", RsDetalle!tCodigoProducto)
      sSubGrupo = IIf(IsNull(RsDetalle!tCodigoSubGrupo), "", RsDetalle!tCodigoSubGrupo)
      sGrupo = IIf(IsNull(RsDetalle!tCodigoGrupo), "", RsDetalle!tCodigoGrupo)
      sItem = IIf(IsNull(RsDetalle!tItem), "001", RsDetalle!tItem)
      nOrden = IIf(IsNull(RsDetalle!nOrden), 0, RsDetalle!nOrden)
      sCortesia = IIf(IsNull(RsDetalle!tCortesia), "", RsDetalle!tCortesia)
      
      nOficial = IIf(IsNull(RsDetalle!nPrecioOficial), 0, RsDetalle!nPrecioOficial)
      nDescuento = IIf(IsNull(RsDetalle!nDescuento), 0, RsDetalle!nDescuento)
      nRecargo = IIf(IsNull(RsDetalle!nRecargo), 0, RsDetalle!nRecargo)
      nPBase = IIf(IsNull(RsDetalle!nPrecioNeto), 0, RsDetalle!nPrecioNeto)
      nImpuesto1 = IIf(IsNull(RsDetalle!nprecioImpuesto1), 0, RsDetalle!nprecioImpuesto1)
      nImpuesto2 = IIf(IsNull(RsDetalle!nprecioImpuesto2), 0, RsDetalle!nprecioImpuesto2)
      nImpuesto3 = IIf(IsNull(RsDetalle!nprecioImpuesto3), 0, RsDetalle!nprecioImpuesto3)
      nPVenta = IIf(IsNull(RsDetalle!nprecioVenta), 0, RsDetalle!nprecioVenta)
      nCantidad = IIf(IsNull(RsDetalle!nCantidad), 0, RsDetalle!nCantidad)
    
      txtOficial.Caption = Format(nOficial, "###,###,###,##0.00")
      txtNeto.Caption = Format(nPBase, "###,###,###,##0.00")
      txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
      txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
      txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
      txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
      txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
      txtPVenta.Caption = Format(nPVenta, "###,###,###,##0.00")
      txtCantidad.Caption = Format(nCantidad, "##,##0.00")
      txtVenta.Caption = Format(nPVenta * nCantidad, "###,###,###,##0.00")
      lblObservacion.Text = IIf(IsNull(RsDetalle!tObservacion), "", RsDetalle!tObservacion)
        If IIf(IsNull(RsDetalle!lImprime), False, RsDetalle!lImprime) = False Then
            'luchoinsumos
             verificatitulo
             'luchoinsumos
        Else
            Label2.Caption = muestra
            Label2.Visible = False
        End If
                                        
                    
      If nOficial = 0 Then
         txtDPorcentaje.Caption = "0.00"
         txtRPorcentaje.Caption = "0.00"
      Else
         txtDPorcentaje.Caption = Format(nDescuento * 100 / nOficial, "###,###,###,##0.00")
         txtRPorcentaje.Caption = Format(nRecargo * 100 / nOficial, "###,###,###,##0.00")
      End If
           
           
      'Llena el Combo
      'fraCombo.Caption = IIf(IsNull(RsDetalle!Producto), "", " " & RsDetalle!Producto & " ")
      txtPorductoComboSeleccionado.Text = IIf(IsNull(RsDetalle!Producto), "", " " & RsDetalle!Producto & " ")
      txtProductoSeleccionado.Text = IIf(IsNull(RsDetalle!Producto), "", " " & RsDetalle!Producto & " ")
      wCombo = IIf(IsNull(RsDetalle!lCombinacion), False, RsDetalle!lCombinacion)
      nCombo = IIf(IsNull(RsDetalle!nCombinacion), 1, RsDetalle!nCombinacion)
      RsCombo.Filter = "[tItem]='" & sItem & "'"
      fraCombo.Visible = False
      wAgregaCombo = False
      
       If wCombo = True Then
        sProductoCombo = sProducto
      End If
      
      txtCortesia.Caption = IIf(IsNull(RsDetalle!Cortesia), "", RsDetalle!Cortesia)
      txtObserva.Caption = IIf(IsNull(RsDetalle!tObservacion), "", RsDetalle!tObservacion)
      VisualizaMonto
      
      If EliminarProductoCom <> 1 Then
      tabProducto.Visible = True
      ActivaCabecera True
      
      End If
      fraPropiedad.Visible = False
      
      
        
      'ojoooooooo
      ListarOperadoresConFiltro sProducto
      AsignaPropiedad
   End If
End Sub

Public Sub Facturar(ByVal IndexProceso As Integer, ByVal lBoletaSindatos As Boolean)
On Error GoTo fin
    Dim sSerie As String
    Dim sCorrela As String
    Dim sPrefijo As String
    Dim RsSuma As Recordset
    Dim sTipoDocumento As String
    Dim sImp As String
    Dim wConsumo As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim X As Integer
    Dim nRespuesta As Integer
    Dim imprimeDolaDocumentos As String
    Dim msgError As String
    Dim RscadenaCodigoHash As Recordset
    
    Dim xUltimoCorrelativo As String
    sDetalleConsumo = ""
    
    'FACTURACION_E_PERU
    Dim cadenaCodigoHash As String

    lImprimeAlternativa = False
    sCliente = ""
    wConsumo = False
    
    tAutorizacion = ""
    tCodigoControl = ""
    tDosificacion = ""
    tIdentidadNIT = ""
    
    'FACTURACION OFISIS
    Dim oComandoCabeceraOfisis As clsComando
    Dim oComandoDetalleOfisis As clsComando
    Dim oComandoFirmaDocumentoOfisis As clsComando
    
    Dim oComandoCabeceraOfisis1 As clsComando
    Dim oComandoDetalleOfisis1 As clsComando
    Dim oComandoFirmaDocumentoOfisis1 As clsComando
    
    Dim rdi As Integer
    
    
    
    lblPaso1.Visible = False
    lblPaso2.Visible = False
    imgProceso(0).Visible = False
    imgProceso(1).Visible = False
    imgProceso(2).Visible = False
    imgProceso(3).Visible = False
    FrameFeSpring.Visible = False

    sUsuarioAutoriza = sUsuario
    If RsDetalle.RecordCount = 0 Then
        MsgBox "No hay datos a facturar"
        RsDetalle.Requery
       Exit Sub
    End If

    'Chequea Consistencia
    RsTipoDocumento.Requery
    RsTipoDocumento.MoveFirst
    RsTipoDocumento.Find ("Descripcion='" & txtTipoDocumento.Caption & "'")
    If RsTipoDocumento.EOF Then
       MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
       RsDetalle.Requery
       Exit Sub
    Else
       xlTipoDocumento = Calcular("Select lValidaRuc As Codigo From TTIPODOCUMENTO Where tCodigoTipoDocumento = '" & RsTipoDocumento!TTipoEmision & "'", Cn)
    End If
    
    If nPuerto > 0 Then
       Visor txtTipoDocumento.Caption, "", nPuerto, "N"
    End If

    X = Calcular("select count(tItem) as codigo from [" & sDetalle & "] where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0)", Cn)
    If X <= 0 Then
       MsgBox "Error: No existen Productos a Facturar", vbCritical, sMensaje
       RsDetalle.Requery
       Exit Sub
    End If
    
    If lInfhotel Then
        If rsPuntoVenta.BOF Or rsPuntoVenta.EOF Then
            MsgBox "No se ha Asignado un Punto de Venta de Infhotel, Verificar!!!"
            RsDetalle.Requery
            Exit Sub
        End If
    End If
    
     'validacionmontominimo
    nMontoPedidoFacturarMInimo = 0
       nMontoPedidoFacturarMInimo = Calcular("select isnull(nMontoMinimo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
       If nMontoPedidoFacturarMInimo > 0 Then
           If nMontoPedidoFacturarMInimo >= CDbl(txtMonto.Caption) Then
               MsgBox "El Monto a Facturar no llega al Minimo Permitido al Tipo de Documento"
               wEnter = False
               RsDetalle.Requery
               Exit Sub
           End If
       End If
           
    'validacionMontoMaximo
    nMontoPedidoFacturar = 0
    nMontoPedidoFacturar = Calcular("select isnull(nMontoMaximo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
    If nMontoPedidoFacturar > 0 Then
    If nMontoPedidoFacturar < CDbl(txtMonto.Caption) Then
        MsgBox "El Monto a Facturar supera al Máximo Permitido al Tipo de Documento"
        wEnter = False
        RsDetalle.Requery
        Exit Sub
    End If
    End If
    

    'Consistencia Cortesia
    sCortesia = ""
    If RsTipoDocumento!TTipoEmision = "00" Then
       tUsuActua = sUsuario
       If Supervisor("04") = False Then
          MsgBox "Clave no permitida", vbExclamation, sMensaje
          Exit Sub
       End If
       sUsuario = tUsuActua
       sUsuarioAutoriza = sVar1
       sTemp = ""
       Isql = "select * from vCortesia where lActivo = 1 Order by Descripcion"
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                       "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
       xtipodoc = RsTipoDocumento!prefijo
       frmBusquedaRapida.nPredeterm = 1
       frmBusquedaRapida.Show vbModal
       
       If wEnter = True Then
          sCortesia = sCodigo
          If nPuerto > 0 Then
             Visor "Cortesia", sDescrip, nPuerto, "N"
          End If
       Else
        RsDetalle.Requery
          Exit Sub
       End If
    End If

    If RsTipoDocumento!TTipoEmision = "00" And sCortesia <> "" Then
        Dim nTotalCortesiaActual As Double
        Dim nTopeCortesia As Double
        Dim nTotalDocActual As Double
        nTopeCortesia = Calcular("select isnull(tope,0) as codigo from vcortesia where codigo='" & sCortesia & "'", Cn)
        If nTopeCortesia > 0 Then
'                nTotalCortesiaActual = Calcular("select sum(isnull(nventa,0))  as codigo From mDocumento where ttipodocumento='00' and tcortesia='" & sCortesia & "' and month(fregistro)=month(getdate()) ", Cn)
                nTotalCortesiaActual = Calcular("select sum(isnull(nventa,0))  as codigo From mDocumento where ttipodocumento='00' and tcortesia='" & sCortesia & "' and YEAR(FREGISTRO)=YEAR(GETDATE()) AND month(fregistro)=month(getdate()) ", Cn)
                nTotalDocActual = val(txtMonto.Caption)
                    If nTotalCortesiaActual + nTotalDocActual > nTopeCortesia Then
                            MsgBox "Con esta Emisión se supera el Tope Mensual asignado para la Cortesia " & UCase(sDescrip) & vbCrLf & "Tope Mensual: " & nTopeCortesia & ". Ya Asignado : " & nTotalCortesiaActual, vbCritical
                            wEnter = False
                            variableEmite = False
                            RsDetalle.Requery
                            Exit Sub
                    End If
        End If
    End If

       'impresion imagen
       Set rstFuente = New ADODB.Recordset
       imageCab.Picture = Nothing
       imagepIE.Picture = Nothing
       Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
       imageCab.DataField = "foto"
       Set imageCab.DataSource = rstFuente
       imagepIE.DataField = "fotoPie"
       Set imagepIE.DataSource = rstFuente
        
    
    'Por Consumo
    If lConsumo3 = True Then
       If RsTipoDocumento!TTipoEmision <> "00" Then
          nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
          If nRespuesta = vbYes Then
             frmKeyBoard.txtResultado = "POR CONSUMO"
             frmKeyBoard.Show vbModal
             If sDescrip = "" Or Not wEnter Then
                MsgBox "Error: La descripcion no puede ser en blanco", vbCritical, sMensaje
                RsDetalle.Requery
                Exit Sub
             End If
             sDetalleConsumo = sDescrip
             wConsumo = True
          ElseIf nRespuesta = vbCancel Then
            RsDetalle.Requery
             Exit Sub
          End If
       End If
    End If
    
      '----VALIDA BLOQUEO DE DOCUMENTOS POR CONSUMO CON BONIFICACIONES / 03012020 / GCAA -------------------
            If lbloqueadocconsumo = True And wConsumo = True Then
              Dim n1 As Double
                For i = 0 To grdDetalle.ApproxCount - 1
                    grdDetalle.Row = i
                    n1 = CDbl(grdDetalle.Columns(5))
                   If n1 = 0 Then
                        MsgBox "Información: No se puede emitir el documento por consumo con bonificaciones (item con precio 0)", vbInformation, sMensaje
                       
                        Exit Sub
                   End If
                Next
            End If
    
        If lBloqueoAfectoInafecto And wConsumo Then
            Dim inafecto As Integer
            Dim afecto As Integer
            inafecto = Calcular("select count(*) as codigo from " & sDetalle & " where round(nprecioimpuesto1,2)=0", Cn)
            afecto = Calcular("select count(*) as codigo from " & sDetalle & " where round(nprecioimpuesto1,2)>0", Cn)
            
            If inafecto > 0 And afecto > 0 Then
                MsgBox "Bloqueo de emision: Documento por consumo con afectos e inafectos", vbInformation, sMensaje
                RsDetalle.Requery
                Exit Sub
            End If
            
        End If
    
    TimpresionDolaresDelivery = False
    '-------- impresion en dolares si esta activo el check en el cliente delivery.
    If Calcular("select isnull(lEmisionMonedaExtranjera,0) as codigo from tdelivery where tcodigodelivery='" & sClienteFrecuente & "'", Cn) Then
        If MsgBox("¿Desea Imprimir en " & sMonedaE & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            TimpresionDolaresDelivery = True
        Else
            TimpresionDolaresDelivery = False
        End If
    Else
        TimpresionDolaresDelivery = False
    End If
    
    
    'Tipo de Emision
    If Not wConsumo And RsTipoDocumento!tFormulario <> "01" Then
       'Factura
       'Consistencia Factura

       If RsTipoDocumento!Cliente Then
          sTemp = ""
          Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
        Isql = "exec usp_Inforest_ObtieneClientesFactura '" & sClienteFrecuente & "','" & RsTipoDocumento!TTipoEmision & "'," & CDbl(txtMonto.Caption) & ", '" & sCaja & "'"

          frmBusquedaRapida.cmdOpcion(1).Enabled = True
          frmBusquedaRapida.cmdOpcion(2).Enabled = True
          frmBusquedaRapida.cmdOpcion(3).Enabled = True
          
          Select Case pais
            Case "001" 'Bolivia
                 Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
            Case Else 'Peru, Ecuador
                        If lClub Then
                            Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 4500, 0, 0, "", _
                                                                          "Enlace", 2, "tEnlace", 1100, 0, 0, "")
                        Else
                            Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                        End If
          End Select
          
          sTipoDocum = RsTipoDocumento!TTipoEmision
          xtipodoc = RsTipoDocumento!prefijo
          frmBusquedaRapida.nPredeterm = 1
          frmBusquedaRapida.Show vbModal

          If wEnter = True And sCodigo <> "" Then
             sCliente = sCodigo
                     'imprimedni
                      Dim RsTc1 As ADODB.Recordset
                      Set RsTc1 = New ADODB.Recordset
                      Set RsTc1 = Lib.OpenRecordset("usp_Inforest_ValidaClienteSel '" & sTipoDocum & "','" & sCliente & "'", Cn)
                      If Not (RsTc1.EOF Or RsTc1.BOF) Then
                       RsTc1.MoveFirst
                       If RsTc1.Fields(0) <> "ok" Then
                           MsgBox "Error: El tipo de Identidad del Cliente no Corresponde al Tipo de Documento", vbCritical, sMensaje
                           Exit Sub
                       End If
                      End If
                 lValidaEmail = Calcular("Select lValidaEmail As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn)
                 
                 If lValidaEmail = True Then
                    sEmail = Calcular("Select ISNULL(tcorreo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                 
                    If sEmail = "" Then
                       MsgBox "El cliente no tiene Email registrado", vbCritical, sMensaje
                       Exit Sub
                    End If
                 End If
             
          Else
             RsDetalle.Requery
             Exit Sub
          End If
       End If
       
       If Pedido = "" Then
            GeneraPedido
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=0", Cn)
               If i > 0 Then
                  cmdOpcion_Click (8)
                  If variableEmite = False Then: Exit Sub
               End If
            Else
               GeneraPedido
            End If
       Else
          sPedido = Pedido
          Cn.Execute "delete from DPEDIDO where tCodigoPedido='" & Pedido & "'"
          'Inserta el Detalle
          Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                     "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                     "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,fregistro, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tsubalmacen,tCodigoEtiqueta,tunidadnegocio,fenvio,nenvio,fdiacontable) " & _
                     "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                     "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                     "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,getdate(), nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tsubalmacen,tCodigoEtiqueta,'" & sUnidadNegocio & "' ,fenvio,nenvio,'" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
                     "From [" & sDetalle & "] where tEstadoItem='N'"
          Cn.Execute "Update MPEDIDO set tObservacion='" & txtObservacion.Caption & "', tMozo='" & sMozo & "' where tCodigoPedido='" & Pedido & "'"
       End If

        'ACTUALIZA  tAplicaExoneracion, tAplicaInafectacion
         
             Dim oComando  As clsComando
            Set oComando = New clsComando
            If Not oComando.CreateCmdSp("spUpdate_DPEDIDO_Ina", Cn) Then
               Set oComando = Nothing
               Exit Sub
            End If
            
            oComando.CreateParameter "@pedido", adVarChar, adParamInput, 10, Pedido
            
            If Not oComando.GetParamOK Then
               Set oComando = Nothing
               Exit Sub
            End If
        
            If Not oComando.ExecSP Then
               Set oComando = Nothing
               Exit Sub
             End If



       'Genera y Actualiza los Numero de Documento
       RsDetalle.MoveFirst
       
       If RsTipoDocumento!tFormulario = "03" Then
          nFItem = nItemV
       End If

       For i = 1 To IIf(X Mod nFItem = 0, Int(X / nFItem), Int(X / nFItem) + 1)
           RsTipoDocumento.Requery
           RsTipoDocumento.MoveFirst
           RsTipoDocumento.Find ("Descripcion='" & txtTipoDocumento.Caption & "'")
           If RsTipoDocumento.EOF Then
              MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
              Exit Sub
           End If

           sSerie = RsTipoDocumento!tSerie
           sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
           sPrefijo = RsTipoDocumento!prefijo
           sTipoDocumento = RsTipoDocumento!TTipoEmision
           sImp = RsTipoDocumento!timpresora
           sDocumento = sPrefijo & sSerie & sCorrela
           sResumen = RsTipoDocumento!lResumen
           
             Select Case pais
                 Case "001" 'Bolivia
                         tAutorizacion = obtieneAutorizacionDosificacion(sCaja, "1")
                         tDosificacion = obtieneAutorizacionDosificacion(sCaja, "2")
                         If tAutorizacion <> "" And tDosificacion <> "" Then
                             
                         Else
                             MsgBox "Error al obtener Número de Autorización o Dosificación. Verifique.", vbCritical, sMensaje
                             Exit Sub
                         End If
                 Case "002" 'ECUADOR
                     tAutorizacion = RsTipoDocumento!tNumeroAutorizacion
                 Case Else 'Peru
                     tAutorizacion = ""
                     tCodigoControl = ""
                     tDosificacion = ""
                         
             End Select
           
           'Genera el Detalle de DDOCUMENTO
           Dim xClave As String
           For j = 1 To nFItem
               xClave = RsDetalle!tItem
               Isql = "Update DPEDIDO set tDocumento = '" & sDocumento & "' where tItem = '" & xClave & "' and tCodigoPedido = '" & Pedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado)) = 0) "
               Cn.Execute Isql
               RsDetalle.MoveFirst
               RsDetalle.Find ("tItem ='" & xClave & "'")
               RsDetalle.MoveNext
               If RsDetalle.EOF Then
                  Exit For
               End If
           Next j

           'Inserta Detalle de Documento
           Isql = "Insert into DDOCUMENTO " & _
                  "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                  "select  '" & sDocumento & "' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From DPEDIDO " & _
                  "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tDocumento ='" & sDocumento & "'"
           Cn.Execute Isql

           'Genera el Detalle de MDOCUMENTO
           Isql = "Insert into MDOCUMENTO " & _
                  "     ( tDocumento, tTipoDocumento, tCodigoCliente, tEstadoDocumento, tCaja, tSalon, tTurno, tUsuario, tUsuarioAutoriza, fRegistro, fDiaContable, tConsumo, lImpresionMonedaExtranjera, femision) " & _
                  "Values(   '" & sDocumento & "', " _
                          & "'" & sTipoDocumento & "', " _
                          & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                          & "'01', " _
                          & "'" & sCaja & "', " _
                          & "'" & sSalon & "', " _
                          & "'" & sTurno & "', " _
                          & "'" & Mid(sUsuario, 1, 15) & "', " _
                          & "'" & sUsuarioAutoriza & "', " _
                          & " getdate(), '" & Format(obtieneDiaContable, "yyyyMMdd") & "','" & sDetalleConsumo & "', " & IIf(TimpresionDolaresDelivery, 1, 0) & " , '" & Format(sfemision, "yyyyMMdd") & "') "
           Cn.Execute Isql

           'Calcula el total de la cabecera
           Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta, isnull(sum(nDescuento*nCantidad),0) nDescuento " & _
                                          " from DPEDIDO where tDocumento ='" & sDocumento & "' group by tDocumento", Cn)

           'Actualiza el Documento con el Temporal
           nCargo = Round(RsSuma!nVenta, 2)

           Select Case pais
               Case "001"
                   tCodigoControl = devuelveCodigoControl(sCaja, sCorrela, tAutorizacion, tDosificacion, sCliente, nCargo)
           End Select
           
           Isql = "Update MDOCUMENTO set nNeto = " & RsSuma!nNeto & " , " & _
                                        "nRecargo = 0, " & _
                                        "nDescuento = " & IIf(lAplicablePedido, 0, RsSuma!nDescuento) & ", " & _
                                        "nPrecioOficial = 0 , " & _
                                        "nPrecioImpuesto1 = " & RsSuma!nImpuesto1 & " , " & _
                                        "nPrecioImpuesto2 = " & RsSuma!nImpuesto2 & " , " & _
                                        "nPrecioImpuesto3 = " & RsSuma!nImpuesto3 & " , " & _
                                        "tautorizacion = '" & tAutorizacion & "' , " & _
                                        "tcodigocontrol = '" & tCodigoControl & "' , " & _
                                        "nVenta = " & RsSuma!nVenta & _
                                        " ,lreplica=1 where tDocumento = '" & sDocumento & "'"
           Cn.Execute Isql
           
           
           wEnter = True
           
           If lPagoAntesImpresion Then
                 xTipo = ""
                 sFormulario = "CajaRapida"
                
                 If sFormulario = "CajaRapida" And lPagoRapido = True Then
                    frmPagoRapido.Show vbModal
                 Else
                    frmPago.Show vbModal
                 End If
                 
                 If wEnter = False Then
                    Dim RsCantDocumentos As Recordset
                    Set RsCantDocumentos = Lib.OpenRecordset("select distinct tDocumento from DDOCUMENTO where tCodigoPedido ='" & Pedido & "'", Cn)
                    
                    For rdi = 0 To RsCantDocumentos.RecordCount - 1
                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                        Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & Pedido & "'"
                        RsCantDocumentos.MoveNext
                    Next rdi
                    
                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    
                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                    Exit Sub
                 End If
           End If
           
           
           
           Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
           
           If lPagoAntesImpresion = False Then
           
                   nMonto = RsSuma!nVenta
                   
                   'Actualiza Base de Datos Detalle del Pedido
                   Cn.Execute "Update DPEDIDO set tFacturado = 'F' where tDocumento ='" & sDocumento & "'"
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido ='" & sPedido & "'"
                   
                   'PARA NO FISCALES
                   Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tCodigoPedido ='" & Pedido & "' and tDocumento ='" & sDocumento & "'"
                   
                   'Imprime Documentos
                    If wConsumo = False And lDescripcionAlternativa = True Then
                     If validaImpresionAlternativa(sDocumento) = False Then
                            If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                                  lImprimeAlternativa = True
                            End If
                      End If
                    End If
                    '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                    ' ELDCQ 22/11/2017
                    If lImprimeAlternativa = False Then
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
                        'FACTURACION_E_PERU
                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',4"
                    Else
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"
                        'FACTURACION_E_PERU
                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',8"
                    End If
                    '-------------------------------------------------------------------------------------------------------------
                    'FACTURACION_E_PERU
                    'Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
                    xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
                    xImpresioDE = Mid(sDocumento, 1, 1)
                    '---------------------------------------
                
                    Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        
                        imprimeDolaDocumentos = Calcular("select isnull(lequivadolares,0) as codigo from vtipodocumentoimpresora where tcaja='" & sCaja & "' and ttipoemision='" & sTipoDocumento & "' ", Cn)
                        If imprimeDolaDocumentos = "Verdadero" Then
                            lDocumEquivaPrecuenta = True
                        Else
                            lDocumEquivaPrecuenta = False
                        End If
                                      
                    If RsImpresion.RecordCount = 0 Then
                       LimpiaRs
                       MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                    Else
                            'SUNAT
                            numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
                            codigoImpresora = sImp
                            'SUNAT
                            Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "

                           'FACTURACION_E_PERU
                           If pais = "000" Then
                               If lFacturacionE Then
                                    If lFEOfisis Then 'OFISIS
                        
                                            '----CABECERA
                                            Set oComandoCabeceraOfisis = New clsComando
                                            If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                        
                                            If Not oComandoCabeceraOfisis.GetParamOK Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            If Not oComandoCabeceraOfisis.ExecSP Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                        
                                            '----FIRMA DOCUMENTO OFISIS
                                            If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                                Set oComandoFirmaDocumentoOfisis = New clsComando
                                                If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                        
                                                If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                
                                                'VALIDAR RESPUESTA CODIGO DE BARRA
                                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                                Sleep 3000
                                                If lImpresionCodigoBarras Then
                                                    imageHash.DataField = "foto"
                                                    Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                                    Set imageHash.DataSource = RsCodigoHash
                                                    
                                                ElseIf lQRFE Then
                                                    Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                                Else
                                                    Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                                    If RscadenaCodigoHash.RecordCount > 0 Then
                                                        cadenaCodigoHash = RscadenaCodigoHash!codigo
                                                    End If
                                                    'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                                End If
                                            End If
                                    
                                    ElseIf lFESpring Then
                                    
                                    ElseIf lFECarbajal Then
                                    
                                    ElseIf lFEpape Then
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                            If tCodigoFE = "000" Then
                                                 If lQRFE Then
                                                     Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                                 Else
                                                     If lImpresionCodigoBarras Then
                                                         'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                     Else
                                                         cadenaCodigoHash = PapeMatricial
                                                     End If
                                                 End If
                                             End If
                                        End If
                                    ElseIf lFEGesa Then
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                            If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                                MsgBox "No se pudo enviar el documento a facturacion electronica", vbInformation
                                            End If
                                            If lQRFE Then
                                                Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                            Else
                                                If lImpresionCodigoBarras Then
                                                    Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                                Else
                                                    cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                                End If
                                            End If
                                        End If
                                    Else 'INFOFACT
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                             If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                                 'Exit Sub
                                             End If
'                                             If RsImpresion!Ruc <> "" Then
'                                                 If Not INSERTAFE(sDocumento, "", 2, RsImpresion!Ruc) Then '----CLIENTE
'                                                     Exit Sub
'                                                 End If
'                                             End If
                                             'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                             fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                             If tCodigoFE = "000" Then
                                                 If lQRFE Then
                                                     Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                 Else
                                                     If lImpresionCodigoBarras Then
                                                         Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sDocumento))
                                                     Else
                                                         cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sDocumento)
                                                     End If
                                                 End If
                                             End If
                                         End If
                                    End If
                               End If
                                                   
                        End If
                        '---------------------------------------
        
                        'Configura la Impresora
                        Imprimir (sImp)
                        Printer.FontName = sFont
                        Printer.FontBold = False
        
                       'TVARIABLE CESAR
                       'FORMATO TICKET VARIABLE
                       If RsTipoDocumento!tFormulario = "03" Then
                               If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                               
                                     ImprimeFacturaVariable RsImpresion, sEmpresa
                                                   
                                  NFactura = sCorrela
                                  'frmVenta.lblFactura.Caption = NFactura
                               ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                  If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                     ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                  End If
                               Else
                                  ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                               End If
                              
                       Else
                       
                               'FORMATO VARIABLE
                               If lFacturacionE And lFEOfisis = False Then
                                       'FACTURACION_E_PERU
                                       'FORMATO A4
                                       If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                              ImprimeFormatoA sDocumento
                                              Kill App.path & "\fact.bmp"
                                       Else
                                              ImprimeFormatoA sDocumento
                                       End If
                                                                                 
                               Else
                                           If sTipoDocumento = "01" Then
                                              If wConsumo Then
                                                 ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                              Else
                                                 ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                                              End If
                                              NFactura = sCorrela
                                              'frmVenta.lblFactura.Caption = NFactura
                                           Else
                                              If wConsumo Then
                                                 ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                              Else
                                                 ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                              End If
                                           
                                           End If
                               
                               End If
            
            
                        End If
                      
                   End If
                   
'                   If pais = "002" Then
'                        sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
'                        GeneraFacturaElectronica sXML, sDocumento
'                   End If
                    If pais = "002" And lFEEcuador = False Then
                       sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
                       GeneraFacturaElectronica sXML, sDocumento
                    End If
                    
                    If lFEEcuador Then
                     If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
                         MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                     End If
                    End If
                   
           End If
           
           If lPagoAntesImpresion = False Then
                xTipo = ""
                sFormulario = "CajaRapida"
                
                If sFormulario = "CajaRapida" And lPagoRapido = True Then
                   frmPagoRapido.Show vbModal
                Else
                   frmPago.Show vbModal
                End If
           End If
        Next i
        
        
        
        If lPagoAntesImpresion Then
            Set RsCantDocumentos = Lib.OpenRecordset("select distinct tDocumento from DDOCUMENTO where tCodigoPedido ='" & Pedido & "'", Cn)
            For rdi = 0 To RsCantDocumentos.RecordCount - 1
                sDocumento = RsCantDocumentos!tDocumento
                
                   'nMonto = RsSuma!nVenta
                   
                   'Actualiza Base de Datos Detalle del Pedido
                   Cn.Execute "Update DPEDIDO set tFacturado = 'F' where tDocumento ='" & sDocumento & "'"
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido ='" & sPedido & "'"
                   
                   'PARA NO FISCALES
                   Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tCodigoPedido ='" & Pedido & "' and tDocumento ='" & sDocumento & "'"
                   
                   'Imprime Documentos
                    If wConsumo = False And lDescripcionAlternativa = True Then
                     If validaImpresionAlternativa(sDocumento) = False Then
                            If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                                  lImprimeAlternativa = True
                            End If
                      End If
                    End If
                    '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                    ' ELDCQ 22/11/2017
                    If lImprimeAlternativa = False Then
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
                        'FACTURACION_E_PERU
                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',4"
                    Else
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"
                        'FACTURACION_E_PERU
                        'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',8"
                    End If
                    '----------------------------------------------------------------------------------------------------------------------
                    'FACTURACION_E_PERU
                    'Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
                    xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
                    xImpresioDE = Mid(sDocumento, 1, 1)
                    '---------------------------------------
                
                    Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        
                        imprimeDolaDocumentos = Calcular("select isnull(lequivadolares,0) as codigo from vtipodocumentoimpresora where tcaja='" & sCaja & "' and ttipoemision='" & sTipoDocumento & "' ", Cn)
                        If imprimeDolaDocumentos = "Verdadero" Then
                            lDocumEquivaPrecuenta = True
                        Else
                            lDocumEquivaPrecuenta = False
                        End If
                                      
                    If RsImpresion.RecordCount = 0 Then
                       LimpiaRs
                       MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                    Else
                   
                            'SUNAT
                            numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
                            codigoImpresora = sImp
                            'SUNAT
                            Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "
                
        
                           'FACTURACION_E_PERU
                           If pais = "000" Then
                               If lFacturacionE Then
                                    If lFEOfisis Then 'OFISIS
                        
                                            '----CABECERA
                                            Set oComandoCabeceraOfisis = New clsComando
                                            If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                        
                                            If Not oComandoCabeceraOfisis.GetParamOK Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            If Not oComandoCabeceraOfisis.ExecSP Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                           
                                            '----FIRMA DOCUMENTO OFISIS
                                            If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                                Set oComandoFirmaDocumentoOfisis = New clsComando
                                                If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                        
                                                If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                
                                                'VALIDAR RESPUESTA CODIGO DE BARRA
                                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                                Sleep 3000
                                                If lImpresionCodigoBarras Then
                                                    imageHash.DataField = "foto"
                                                    Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                                    Set imageHash.DataSource = RsCodigoHash
                                                Else
                                                    cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                                End If
                                            End If
                                     
                                    ElseIf lFESpring Then
                                    
                                    ElseIf lFECarbajal Then
                        
                                    ElseIf lFEpape Then
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                            If tCodigoFE = "000" Then
                                                 If lQRFE Then
                                                     Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                                 Else
                                                     If lImpresionCodigoBarras Then
                                                         'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                     Else
                                                         cadenaCodigoHash = PapeMatricial
                                                     End If
                                                 End If
                                             End If
                                        End If
                                    Else 'INFOFACT
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                             If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                                 'Exit Sub
                                             End If
'                                             If RsImpresion!Ruc <> "" Then
'                                                 If Not INSERTAFE(sDocumento, "", 2, RsImpresion!Ruc) Then '----CLIENTE
'                                                     Exit Sub
'                                                 End If
'                                             End If
                                             'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                             fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                             If tCodigoFE = "000" Then
                                                 If lQRFE Then
                                                     Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                 Else
                                                     If lImpresionCodigoBarras Then
                                                         Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sDocumento))
                                                     Else
                                                         cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sDocumento)
                                                     End If
                                                 End If
                                             End If
                                        End If
                                    End If
                               End If
                        End If
                        '---------------------------------------
        
                        'Configura la Impresora
                        Imprimir (sImp)
                        Printer.FontName = sFont
                        Printer.FontBold = False
        
                       'TVARIABLE CESAR
                       'FORMATO TICKET VARIABLE
                       If RsTipoDocumento!tFormulario = "03" Then
                               If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                               
                                     ImprimeFacturaVariable RsImpresion, sEmpresa
                                                   
                                  NFactura = sCorrela
                                  'frmVenta.lblFactura.Caption = NFactura
                               ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                  If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                     ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                  End If
                               Else
                                  ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                               End If
                              
                       Else
                       
                               'FORMATO VARIABLE
                               If lFacturacionE And lFEOfisis = False Then
                                       'FACTURACION_E_PERU
                                       'FORMATO A4
                                       If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                              ImprimeFormatoA sDocumento
                                              Kill App.path & "\fact.bmp"
                                       Else
                                              ImprimeFormatoA sDocumento
                                       End If
                                                                                 
                               Else
                                           If sTipoDocumento = "01" Then
                                              If wConsumo Then
                                                 ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                              Else
                                                 ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                                              End If
                                              NFactura = sCorrela
                                              'frmVenta.lblFactura.Caption = NFactura
                                           Else
                                              If wConsumo Then
                                                 ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                              Else
                                                 ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                              End If
                                           
                                           End If
                               
                               End If
            
            
                        End If
                      
                   End If
                   
'                   If pais = "002" Then
'                        sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
'                        GeneraFacturaElectronica sXML, sDocumento
'                   End If
                    If pais = "002" And lFEEcuador = False Then
                       sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
                       GeneraFacturaElectronica sXML, sDocumento
                    End If
                    
                    If lFEEcuador Then
                     If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
                         MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                     End If
                    End If

            RsCantDocumentos.MoveNext
            Next rdi
        End If

        LimpiaRs
       
       
    '------------------- FORMATO TICKET
    Else
    
        'JAPN 08/09/2023
        Dim tDocumentoVenta As String
       
       ' JAPN 8/09/2023
        tDocumentoVenta = IIf(IsNull(RsTipoDocumento!tCodigoExterno), "", RsTipoDocumento!tCodigoExterno)
    
       If RsTipoDocumento!Cliente And (RsTipoDocumento!Monto <= nMonto Or RsTipoDocumento!Monto = 0) And lBoletaSindatos = False Then
          'Factura
          sTemp = ""
          Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
          'imprimedni
          Isql = "exec usp_Inforest_ObtieneClientesFactura '" & sClienteFrecuente & "','" & RsTipoDocumento!TTipoEmision & "', " & nMonto & ", '" & sCaja & "' "

          frmBusquedaRapida.cmdOpcion(1).Enabled = True
          frmBusquedaRapida.cmdOpcion(2).Enabled = True
          frmBusquedaRapida.cmdOpcion(3).Enabled = True
          
          Select Case pais
            Case "001" 'Bolivia
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
            Case Else 'Peru, Ecuador
                        If lClub Then
                            Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 4500, 0, 0, "", _
                                                                          "Enlace", 2, "tEnlace", 1100, 0, 0, "")
                        Else
                            Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                        End If
          End Select
           '------VALIDA CORREO----------
          sTipoDocum = RsTipoDocumento!TTipoEmision
          xtipodoc = RsTipoDocumento!prefijo
          frmBusquedaRapida.nPredeterm = 1
          'frmBusquedaRapida.Show vbModal
          Dim sBoleta As Boolean
            sBoleta = Calcular("select ISNULL(lBODato,0) as Codigo FROM TPARAMETRO", Cn)
        
            If sBoleta = "Verdadero" Then
                If xtipodoc = "B" And tDocumentoVenta = "03" Then
                    'Me.FrameSolicitudDatos.Visible = True
                    frmSolicitudBoleta.Show vbModal
                Else
                    frmBusquedaRapida.Show vbModal
                End If
            Else
                 frmBusquedaRapida.Show vbModal
            End If
          
          If wEnter = True And sCodigo <> "" Then
             sCliente = sCodigo
                    Dim RsTc As ADODB.Recordset
                    Set RsTc = New ADODB.Recordset
                    Set RsTc = Lib.OpenRecordset("usp_Inforest_ValidaClienteSel '" & sTipoDocum & "','" & sCliente & "'", Cn)
                    If Not (RsTc.EOF Or RsTc.BOF) Then
                     RsTc.MoveFirst
                     If RsTc.Fields(0) <> "ok" Then
                         MsgBox "Error: El tipo de Identidad del Cliente no Corresponde al Tipo de Documento", vbCritical, sMensaje
                         RsDetalle.Requery
                         Exit Sub
                     End If
                    End If
                    
                    lValidaEmail = Calcular("Select lValidaEmail As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn)
                    
                    If lValidaEmail = True Then
                       sEmail = Calcular("Select ISNULL(tcorreo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                    
                       If sEmail = "" Then
                          MsgBox "El cliente no tiene Email registrado", vbCritical, sMensaje
                          wEnter = False
                          RsDetalle.Requery
                          Exit Sub
                       End If
                    End If
                    
                    If Calcular("Select lValidaUbigeo As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn) = True Then
                        Dim TempUbigeo As String
                        Dim TempUrbaniza As String
                        TempUbigeo = Calcular("Select ISNULL(CodigoUbigeo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                        TempUrbaniza = Calcular("Select ISNULL(Urbanizacion,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                        If Trim(TempUbigeo) = "" Or Trim(TempUrbaniza) = "" Then
                            MsgBox "El cliente no tiene Ubigeo ó Urbanizacion registrado, Favor de verificar!!!", vbCritical, sMensaje
                            wEnter = False
                            RsDetalle.Requery
                            Exit Sub
                        End If
                    End If
                    
          Else
            RsDetalle.Requery
             'MsgBox "Proceso Cancelado", vbCritical, sMensaje
             Exit Sub
          End If
       Else
            
            sCliente = sCodigoClienteGeneral
            
       End If
       
        If lImpComandaBarra And Not lPrinter Then
            i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=0", Cn)
            If i > 0 Then
               cmdOpcion_Click (8)
               If variableEmite = False Then: RsDetalle.Requery: Exit Sub
            End If
        End If
       
       If Pedido = "" Then
          'GeneraPedido
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=0", Cn)
               If i > 0 Then
                  cmdOpcion_Click (8)
                  If variableEmite = False Then: RsDetalle.Requery: Exit Sub
               End If
            Else
               GeneraPedido
            End If
          
       Else
          sPedido = Pedido

          'IMPRESION DE PRODUCTOS NO ENVIADOS
          If lPrinter And lObligaPrinter Then
            i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=0", Cn)
            If i > 0 Then
               cmdOpcion_Click (8)
            End If
          End If
          
          ActualizaPedido

          Cn.Execute "update MPEDIDO set tObservacion='" & txtObservacion.Caption & "' where tCodigoPedido='" & sPedido & "'"
          
       End If

       'impresion imagen
       Set rstFuente = New ADODB.Recordset
       imageCab.Picture = Nothing
       imagepIE.Picture = Nothing
       Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
       imageCab.DataField = "foto"
       Set imageCab.DataSource = rstFuente
       imagepIE.DataField = "fotoPie"
       Set imagepIE.DataSource = rstFuente


       'Genera y Actualiza los Numero de Documento
       sSerie = RsTipoDocumento!tSerie
       sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
       sPrefijo = RsTipoDocumento!prefijo
       sTipoDocumento = RsTipoDocumento!TTipoEmision
       sImp = RsTipoDocumento!timpresora
       sDocumento = sPrefijo & sSerie & sCorrela
       sResumen = RsTipoDocumento!lResumen
       

       'Calcula el total de la cabecera
       Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta " & _
                                      " from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & Pedido & "' group by tCodigoPedido ", Cn)

       'Inserta el Documento
       nCargo = Round(RsSuma!nVenta, 2)
       nCargo = CDbl(txtMonto.Caption)
       
       Select Case pais
            Case "001"
                    tAutorizacion = obtieneAutorizacionDosificacion(sCaja, "1")
                    tDosificacion = obtieneAutorizacionDosificacion(sCaja, "2")
                    If tAutorizacion <> "" And tDosificacion <> "" Then
                        tCodigoControl = devuelveCodigoControl(sCaja, sCorrela, tAutorizacion, tDosificacion, sCliente, nCargo)
                        If tCodigoControl = "" Then
                           MsgBox "Error al generar Código de Control", vbCritical, sMensaje
                           Exit Sub
                        End If
                    Else
                        MsgBox "Error al obtener Número de Autorización o Dosificación. Verifique.", vbCritical, sMensaje
                        Exit Sub
                    End If
                    
            Case "002" 'Ecuador
                    tAutorizacion = RsTipoDocumento!tNumeroAutorizacion

            Case Else 'Peru
                tAutorizacion = ""
                tCodigoControl = ""
                tDosificacion = ""

       End Select

       If lAplicablePedido Then
          nTotalDescuento = 0
       Else
          nTotalDescuento = Calcular("select sum(nDescuento*nCantidad) as Codigo From " & sDetalle, Cn)
       End If

        ''ELDCQ (INSERTA DOCUMENTO)
        If Not INSERTADOCUMENTO(Pedido, sDocumento, sTipoDocumento, IIf(sCliente = "", "", sCliente), "01", sCaja, sTurno, sSalon, sUsuario, sUsuarioAutoriza, Format(obtieneDiaContable, "yyyy/MM/dd"), IIf(lAplicablePedido, "", sCodigoDescuento), sDetalleConsumo, IIf(TimpresionDolaresDelivery, 1, 0), tAutorizacion, tCodigoControl, sCortesia, Format("01/01/1900", "yyyy/mm/dd"), Format("01/01/1900", "yyyy/mm/dd"), "", 1, sfemision) Then
           GoTo fin:
        End If
           
       Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
       
        If lIntVisaNetQR And sTipoDocumento <> "00" Then
            Set ImageIntegracion = Nothing
            If (EjecutaMotorIntegraciones(sDocumento, sUsuario, 1, 1) = True) Then
                        'Sleep (500)
                Dim Ob As New Scripting.FileSystemObject
                If Ob.FileExists(lRutaImagenIntegraciones + "\VISAQR" + sRUC + sCaja + sDocumento + ".bmp") Then
                    Set ImageIntegracion.Picture = LoadPicture(lRutaImagenIntegraciones + "\VISAQR" + sRUC + sCaja + sDocumento + ".bmp")
                    Kill lRutaImagenIntegraciones + "\VISAQR" + sRUC + sCaja + sDocumento + ".bmp"
                Else
                 Set ImageIntegracion = Nothing
                End If
            
            Else
                MsgBox "No se pudo generar el QR Visa Net, el sistema continuara con la operacion!!!", vbInformation
                Set ImageIntegracion = Nothing
            End If
        Else
            Set ImageIntegracion = Nothing
        End If
       
       
       ' aqui ---
       ' GCAA 08012021 PAGO ANTES DE IMPRESION
        If lPagoAntesImpresion = False Then
             If lIntMercadoPagoQR And sTipoDocumento <> "00" Then
                Set ImageMercadoPago = Nothing
                If (EjecutaMotorIntegraciones(sDocumento, sUsuario, 5, 1) = True) Then
                    'Sleep (500)
                    Dim Obx As New Scripting.FileSystemObject
                    If Obx.FileExists(lRutaImagenIntegraciones + "\MERCADOPAGOQR" + sRUC + sCaja + sDocumento + ".bmp") Then
                        Set ImageMercadoPago.Picture = LoadPicture(lRutaImagenIntegraciones + "\MERCADOPAGOQR" + sRUC + sCaja + sDocumento + ".bmp")
                        Kill lRutaImagenIntegraciones + "\MERCADOPAGOQR" + sRUC + sCaja + sDocumento + ".bmp"
                    Else
                     Set ImageMercadoPago = Nothing
                    End If
                Else
                    MsgBox "No se pudo generar el QR Mercado Pago, el sistema continuara con la operacion!!!", vbInformation
                    Set ImageMercadoPago = Nothing
                End If
            Else
                Set ImageMercadoPago = Nothing
            End If
        End If
        '-----------------
       
       
       
       wEnter = True
       If lPagoAntesImpresion Then
       
            If RsTipoDocumento!TTipoEmision <> "00" Then
               Cn.Execute "Update MPEDIDO set tEstadoPedido = '02'  where tCodigoPedido ='" & sPedido & "'"
               xTipo = ""
               sFormulario = "CajaRapida"
               If sFormulario = "CajaRapida" And lPagoRapido = True Then
                  
                    If IndexProceso > 100 And IndexProceso < 500 Then
                        frmPagoRapido.MontoBilleteRapido = CDbl(IndexProceso - 100)
                    ElseIf IndexProceso = 999 Then
                        frmPagoRapido.MontoBilleteRapido = nMonto
                    End If
                  frmPagoRapido.Show vbModal
                  
                    If wEnter = False Then
                        Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & Pedido & "'"
                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        
                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    
                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                        Exit Sub
                    End If
               Else
                    
                    If IndexProceso > 100 And IndexProceso < 500 Then
                        Dim ValorEfectivo As Integer
                        ValorEfectivo = IndexProceso - 100
                        wEnter = True
                        sDescrip = ValorEfectivo
                        frmPago.txtME(0).Caption = Format(ValorEfectivo, "###,###,###,##0.00")
                        frmPago.nAbonoN = ValorEfectivo
                        frmPago.nEN = ValorEfectivo
                        frmPago.CalculaTotal
                    ElseIf IndexProceso = 999 Then
                        'ValorEfectivo = IndexProceso - 100
                        wEnter = True
                        sDescrip = nMonto
                        frmPago.txtME(0).Caption = Format(nMonto, "###,###,###,##0.00")
                        frmPago.nAbonoN = nMonto
                        frmPago.nEN = nMonto
                        frmPago.CalculaTotal
                    End If
                    
                  frmPago.Show vbModal
                  
                    If wEnter = False Then
                        Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & Pedido & "'"
                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        
                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    
                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                        Exit Sub
                    Else
                        'JSJR 09/01/2025
                        wAnulaDocumento = False
                    End If
               End If
            End If
       
       End If
       
            '-----------------------
           If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                If Not FacturarTCPIP(2, sDocumento, 0) Then
                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    
                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                
                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                    
                    wEnter = False
                    MsgBox "Ocurrio un Problema al Procesar el Documento!!!", vbInformation, sMensaje
                   Exit Sub
                End If
           End If
           '------------------------
       'Actualiza Base de Datos Detalle del Pedido
        If sPrefijo = "0" Then
           Cn.Execute "Update DPEDIDO set tFacturado = 'C', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & Pedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tfacturado))=0)"
           Cn.Execute "Update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido = '" & sPedido & "'"
           Cn.Execute "Update MDOCUMENTO set tEstadoDocumento ='02',lreplica=1  where tDocumento = '" & sDocumento & "'"
        Else
        '// cambio realizado el 05/05/2018 ELDC - actualiza a "P" en Pago antes de Impresion
            If lPagoAntesImpresion Then
                Cn.Execute "Update DPEDIDO set tFacturado = 'P', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & Pedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado))=0)"
            Else
                Cn.Execute "Update DPEDIDO set tFacturado = 'F', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & Pedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado))=0)"
            End If
           'Cn.Execute "Update DPEDIDO set tFacturado = 'F', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & Pedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado))=0)"
           Cn.Execute "Update DPEDIDO set tFacturado = 'C' where tDocumento ='" & sDocumento & "' and len(ltrim(tCortesia)) = 4 "
           Cn.Execute "Update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido ='" & sPedido & "'"
         '// fin de cambio
        End If
                
        'LOG
        If lLogCajaRapida Then
            Cn.Execute "INSERT INTO TLOG_IMPRESION (TDOCUMENTO,TPOSICION1) VALUES('" & sDocumento & "','CREACION DOCUMENTO')"
        End If
        
        'PARA NO FISCALES
        Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tCodigoPedido ='" & Pedido & "' and tDocumento ='" & sDocumento & "'"
        
        If lInfhotel Then
           Dim xSuma As Double
           xSuma = Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tDocumento ='" & sDocumento & "' and tCodigoPedido='" & Pedido & "'", Cn)
           
           If sComandaInfhotel = "" Then
              sComandaInfhotel = Calcular("select left(MAX(tComanda),8) as Codigo from MCOMANDA where tPuntoVenta='" & sPuntoVenta & "'", CnInfhotel)
              sComandaInfhotel = Lib.Correlativo(sComandaInfhotel, 8)
              CnInfhotel.Execute "Update TPUNTOVENTA Set nUltimoComanda = '" & sComandaInfhotel & "' where tPuntoVenta='" & sPuntoVenta & "'"
              sComandaInfhotel = sComandaInfhotel & "-" & UCase(Mid(rsPuntoVenta!descripcion, 1, 3))
              rsPuntoVenta.Requery
              rsPuntoVenta.MoveFirst
              rsPuntoVenta.Find "Codigo='" & sPuntoVenta & "'"
                                          
              'Genero las comandas en Infhotel
              'Cabecera
              Isql = "Insert into MCOMANDA " & _
                     "(tComanda, tPuntoVenta, tHotel, nMovimiento, fFecha, hHora, nTotal, tEstado, " & _
                     "tEmitido, tAsignacion, tCodigoReserva, tNumeroHabitacion, tCodigoFuncionario, " & _
                     "tCaja, tDocumento, tUsuario, nTCambio, tCodigoCompania, tCliente, tMoneda, fFechaE, hHoraE, tUsuarioE,TNOTAPEDIDO) " & _
                     "values('" & sComandaInfhotel & "', '" & sPuntoVenta & "', '" & sHotel & "', 1,  getdate(), getdate(), " & xSuma & ", '01', " & _
                     "1, '" & IIf(RsTipoDocumento!TTipoEmision = "00", "05", "01") & "', '', '', '" & IIf(RsTipoDocumento!TTipoEmision = "00", Mid(sCortesia, 3, 2), "") & "', " & _
                     "'" & sCajaInfhotel & "', '" & IIf(pais = "002", Mid(sDocumento, 1, 1) + Mid(sDocumento, 3), sDocumento) & "', '" & xUsuario & "', " & nTC & ", '', '" & sPasajero & "', '01', getdate(), getdate(), '" & xUsuario & "','" & Pedido & "')"
              CnInfhotel.Execute Isql
           Else
              'sComandaInfhotel = RsCabecera!tComanda
              CnInfhotel.Execute "update MCOMANDA set TASIGNACION='" & IIf(RsTipoDocumento!TTipoEmision = "00", "05", "01") & "', TCODIGORESERVA='', TNUMEROHABITACION='', TCLIENTE='', nTotal= " & xSuma & ", tEstado='01' " & _
                                 "where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
           End If
           
           'Detalle
           Dim nMovimiento As Integer
           CnInfhotel.Execute "delete from DCOMANDA where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
           nMovimiento = Calcular("select max(nmovimiento) as codigo from dcomanda where tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
           
           Isql = "Insert into DCOMANDA " & _
                  "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                  "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                  "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                  "select '" & sComandaInfhotel & "' as tComanda, '" & sPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem , " & nMovimiento & ", '" & Pedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                  "T1.nPrecioVenta as nPrecioUnitario, nCantidad, nVenta as nTotal, T1.nInsumo+T1.nGasto+T1.nManoObra as nPrecioCos, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                  "'' as tCuenta, '" & sCajaInfhotel & "' as tCaja, '" & sDocumento & "' as tDocumento, '" & IIf(sDescrip = "Reserva", "03", "02") & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                  "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                  "where tDocumento='" & sDocumento & "' and tCodigoPedido ='" & Pedido & "'"
           CnInfhotel.Execute Isql
           Cn.Execute "update MPEDIDO set tComanda = '" & sComandaInfhotel & "', tPuntoVenta='" & sPuntoVenta & "'  where tCodigoPedido='" & Pedido & "'"
           sComandaInfhotel = ""
        End If
        
        
       'Imprime Documentos
           'Imprime documentos
           If wConsumo = False And lDescripcionAlternativa = True Then
              If validaImpresionAlternativa(sDocumento) = False Then
                    If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                          lImprimeAlternativa = True
                    End If
              End If
           End If
             '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
             ' ELDCQ 22/11/2017
             If lImprimeAlternativa = False Then
                If lDocumentoAgrupado Then
                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',1"
                    'FACTURACION_E_PERU
                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',2"
                Else
                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
                    'FACTURACION_E_PERU
                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',4"
                End If
            Else
                If lDocumentoAgrupado Then
                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',5"
                    'FACTURACION_E_PERU
                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',6"
                Else
                    Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"
                    'FACTURACION_E_PERU
                    'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',8"
                End If
                    
            End If
            '-------------------------------------------------------------------------------------------------------------
        'FACTURACION_E_PERU
        'Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
        xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
        xImpresioDE = Mid(sDocumento, 1, 1)
        '---------------------------------------
        
        Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        
        'Log
        If lLogCajaRapida Then
            Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION2='RECORDSET IMPRESION CARGADO ' + '" & RsImpresion.RecordCount & "' WHERE TDOCUMENTO='" & sDocumento & "'"
        End If
        
        imprimeDolaDocumentos = Calcular("select isnull(lequivadolares,0) as codigo from vtipodocumentoimpresora where tcaja='" & sCaja & "' and ttipoemision='" & sTipoDocumento & "' ", Cn)
        If imprimeDolaDocumentos = "Verdadero" Then
            lDocumEquivaPrecuenta = True
        Else
            lDocumEquivaPrecuenta = False
        End If
        
        
       If RsImpresion.RecordCount = 0 Then
          LimpiaRs
          MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
       Else
            'SUNAT
            numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
            codigoImpresora = sImp
            'SUNAT
            Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "

            'Log
            If lLogCajaRapida Then
            Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION3='IMPRESORA' + '" & sImp & "'+' CAJA' +'" & sCaja & "' WHERE TDOCUMENTO='" & sDocumento & "'"
            End If

               'FACTURACION_E_PERU
               If pais = "000" Then
                   If lFacturacionE Then
                   
                       If lFEOfisis Then
                                '----CABECERA
                                Set oComandoCabeceraOfisis = New clsComando
                                If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                     Set oComandoCabeceraOfisis = Nothing
                                     Exit Sub
                                End If
                                oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
    
                                If Not oComandoCabeceraOfisis.GetParamOK Then
                                     Set oComandoCabeceraOfisis = Nothing
                                     Exit Sub
                                End If
                                If Not oComandoCabeceraOfisis.ExecSP Then
                                     Set oComandoCabeceraOfisis = Nothing
                                     Exit Sub
                                End If

                                '----FIRMA DOCUMENTO OFISIS
                                If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                    Set oComandoFirmaDocumentoOfisis = New clsComando
                                    If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                         Exit Sub
                                    End If
                                    oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
        
                                    If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                         Exit Sub
                                    End If
                                    If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                         Exit Sub
                                    End If
                                    
                                    'VALIDAR RESPUESTA CODIGO DE BARRA
                                    fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                    Sleep 3000
                                    If lImpresionCodigoBarras Then
                                        imageHash.DataField = "foto"
                                        Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                        Set imageHash.DataSource = RsCodigoHash
                                        
                                    ElseIf lQRFE Then
                                        Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                    Else
                                        Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                        If RscadenaCodigoHash.RecordCount > 0 Then
                                            cadenaCodigoHash = RscadenaCodigoHash!codigo
                                        End If
                                        'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                    End If
                                End If
                               
                               
                       ElseIf lFESpring Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then

                                 frmMensajeFeSpring.sDocumento = sDocumento
                                 frmMensajeFeSpring.oVenta = 4 ' 3: "Formulario Caja Rapida"
                                 frmMensajeFeSpring.Show vbModal
                                 If frmMensajeFeSpring.lEnvio = False Then
                                    Exit Sub
                                 End If
                                                     
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         If frmMensajeFeSpring.lQrInf Then
                                            Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                         Else
                                            Set imageHash.Picture = LoadPicture(ImagenFeSpring(lQRFE, sDocumento))
                                         End If
                                     Else
                                         If lImpresionCodigoBarras Then
                                             
                                         Else
                                             cadenaCodigoHash = ImagenFeSpring(lQRFE, sDocumento)
                                         End If
                                     End If
                                 End If
                             End If
                               
                             
                       ElseIf lFECarbajal Then
                            Label5.Caption = "   Proceso de envio de documento a InfoFact......."
                            lblPaso1.Caption = "Enviando información de documento a InfoFact."
                            lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                Dim sImporteLetra As String
                                sImporteLetra = NumeroCadena(str(RsImpresion!nVenta)) + " " + sMonedaN
                                FrameFeSpring.Visible = True
                                lblPaso1.Visible = True
                                lblPaso2.Visible = True
                                imgProceso(0).Visible = False
                                imgProceso(1).Visible = False
                                imgProceso(2).Visible = False
                                imgProceso(3).Visible = False
                                Sleep 1000
                                If Not INSERTAFE_CARVAJAL(sDocumento, sImporteLetra, 0, 0) Then '----CABECERA
                                        Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                        Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica=1 where tCodigoPedido = '" & sPedido & "'"
                                        Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                        imgProceso(2).Visible = True
                                        imgProceso(3).Visible = True
                                        Sleep 1000
                                        FrameFeSpring.Visible = False
                                        Exit Sub
                                 End If
                                 imgProceso(0).Visible = True
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, sDocumento, 0))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, sDocumento, 0))
                                         Else
                                             cadenaCodigoHash = ImagenFeCarvajal(2, sDocumento, 0)
                                         End If
                                     End If
                                 End If
                                 imgProceso(1).Visible = True
                                 Sleep 1000
                                 FrameFeSpring.Visible = False
                            End If
                                
                       ElseIf lFEpape Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                         Else
                                             cadenaCodigoHash = PapeMatricial
                                         End If
                                     End If
                                 End If
                            End If
                       ElseIf lFEBiz Then
                       
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Date) Then '----CABECERA
                                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                    Cn.Execute "delete from mpedido where tcodigopedido='" & sPedido & "'"
                                    Cn.Execute "delete from dpedido where tcodigopedido='" & sPedido & "'"
                                    Cn.Execute "delete from cpedido where tcodigopedido='" & sPedido & "'"
                                    Cn.Execute "delete from tComboPropiedad where tcodigopedido='" & sPedido & "'"
                                    Cn.Execute "delete from tProductoPropiedad where tcodigopedido='" & sPedido & "'"
'                                    Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
'                                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica=1 where tCodigoPedido = '" & sPedido & "'"
'                                    Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                    If TipoFacturacion = "5" Then
                                        msgError = "Facturacion USIL"
                                    Else
                                        msgError = "Bizlink"
                                    End If
                                    GoTo fin
                                 End If
                                
                                 
                                 
                                 
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                         Else
                                             cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                         End If
                                     End If
                                 End If
                            End If
                        ElseIf lFEGesa Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                    MsgBox "No se pudo enviar el documento a facturacion electronica", vbInformation
                                End If
                                If lQRFE Then
                                    Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                Else
                                    If lImpresionCodigoBarras Then
                                        Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                    Else
                                        cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                    End If
                                End If
                            End If
                       Else 'INFOFACT
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                    'Exit Sub
                                End If
'                                If RsImpresion!Ruc <> "" Then
'                                    If Not INSERTAFE(sDocumento, "", 2, RsImpresion!Ruc) Then '----CLIENTE
'                                        Exit Sub
'                                    End If
'                                End If
                                'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                If tCodigoFE = "000" Then
                                    If lQRFE Then
                                        Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                    Else
                                        If lImpresionCodigoBarras Then
                                            Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sDocumento))
                                        Else
                                            cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sDocumento)
                                        End If
                                    End If
                                End If
                            End If
                       End If
                   End If
            End If
            '---------------------------------------


            'Configura la Impresora
            Imprimir (sImp)
            Printer.FontName = sFont
            Printer.FontBold = False
            
            If pais = "003" And Calcular("select isnull(tpuertoserial,'') as codigo from timpresora where tcaja='" & sCaja & "' and timpresora='" & sImp & "'", Cn) <> "" Then  ' Pais Argentina
                Dim RespDev As Integer
                RespDev = ImpresoraFiscalEpson(Me.PrinterFiscalArgentina, 2, RsImpresion)
                If RespDev = 0 Then
                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    xUltimoCorrelativo = Calcular("select isnull(MAX(tDocumento),'000000000') as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                    Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 1 where tCodigoPedido = '" & sPedido & "'"
                    Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                    msgError = "Error en la conexion con la impresora Fiscal" & vbNewLine & "Revisar fuente de papel, conexion de impresora"
                    GoTo fin
                ElseIf RespDev = 1 Then
                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuario & "', fRegistroAnulado = getdate() , tObservacion ='Anulacion por error en impresora fiscal', lreplica=1  where tDocumento ='" & sDocumento & "'"
'                    xUltimoCorrelativo = Calcular("select isnull(MAX(tDocumento),'000000000') as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
'                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
'                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                    Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 1 where tCodigoPedido = '" & sPedido & "'"
                    Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                    msgError = "Error en el proceso de impresion, " & vbNewLine & "Corte de energia u otro, revisar fuente de papel, conexion de impresora"
                    GoTo fin
                End If
            Else
            
            If wConsumo Then
               If RsTipoDocumento!tFormulario = "01" Then
                      If RsTipoDocumento!Cliente Then
                         'FACTURACION ELECTRONICA
                         If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                              ImprimeFacturaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion
                         Else
                              ImprimeFacturaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                         End If
                         
                         NFactura = sCorrela
                         'frmVenta.lblFactura.Caption = NFactura
                      Else
                          'FACTURACION ELECTRONICA gcaa
                          If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                              ImprimeBoletaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion
                          Else
                              ImprimeBoletaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                          End If
                      End If
               Else
                      If lFacturacionE Then
                                                                  
                              If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                  ImprimeFormatoAConsumo sDocumento
                                  Kill App.path & "\fact.bmp"
                              Else
                                  ImprimeFormatoAConsumo sDocumento
                              End If
                      Else
                      
                          If RsTipoDocumento!Cliente Then
                             ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                             NFactura = sCorrela
                             'frmVenta.lblFactura.Caption = NFactura
                          Else
                             ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                          End If
                      End If
                  
               End If
             
          Else
               If RsTipoDocumento!tFormulario = "01" Then
                   If RsTipoDocumento!Cliente Then
                      
                       'FACTURACION ELECTRONICA
                       If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                          ImprimeFacturaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion, ImageMercadoPago
                       Else
                            'Log
                            If lLogCajaRapida = True Then
                             Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION4='INIC IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
                            End If
                
                               ImprimeFacturaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                           If lLogCajaRapida = True Then
                               Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION7='FIN IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
                           End If
                       End If
                      
                      NFactura = sCorrela
                      'frmVenta.lblFactura.Caption = NFactura
                      
                   ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                      If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                         ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                      End If
                      
                   Else
                
                       'FACTURACION ELECTRONICA
                       If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                          ImprimeBoletaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion, ImageMercadoPago
                       Else
                       'Log
                       If lLogCajaRapida = True Then
                             Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION4='INIC IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
                       End If
                
                          ImprimeBoletaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                       If lLogCajaRapida = True Then
                             Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION7='FIN IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
                       End If
                          
                       End If
                  End If
                
             Else
                  If RsTipoDocumento!Cliente Then
                       ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                       NFactura = sCorrela
                       'frmVenta.lblFactura.Caption = NFactura
                    ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                       If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                          ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                       End If
                    Else
                        'Log
                        If lLogCajaRapida = True Then
                            Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION4='INIC IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
                        End If
                       ImprimeBoletaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                       If lLogCajaRapida = True Then
                            Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION7='FIN IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
    
                       End If
                  End If
                
             End If
             
          End If
        End If  ' Pais Argentina
        
        'CESAR FACTURACION ELECTRONICA
        If pais = "002" Then
             If lFacturacionE Then
              ' PARA FACTURACION DE ECUADOR
             End If
        End If
        '---------------------------------
'               If pais = "002" Then
'                   sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
'                   GeneraFacturaElectronica sXML, sDocumento
'               End If
         If pais = "002" And lFEEcuador = False Then
            sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
            GeneraFacturaElectronica sXML, sDocumento
         End If
         
         If lFEEcuador Then
          If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
              MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
          End If
         End If
          
       End If

       LimpiaRs
        If lPagoAntesImpresion = False Then
                'Cancelacion del Documento
                If RsTipoDocumento!TTipoEmision <> "00" Then
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02'  where tCodigoPedido ='" & sPedido & "'"
                   xTipo = ""
                   sFormulario = "CajaRapida"
                   If sFormulario = "CajaRapida" And lPagoRapido = True Then
                    If IndexProceso > 100 And IndexProceso < 500 Then
                        frmPagoRapido.MontoBilleteRapido = CDbl(IndexProceso - 100)
                    ElseIf IndexProceso = 999 Then
                        frmPagoRapido.MontoBilleteRapido = nMonto
                    End If
                      frmPagoRapido.Show vbModal
                   Else
                        If IndexProceso > 100 And IndexProceso < 500 Then
                            Dim ValorEfectivo2 As Integer
                            ValorEfectivo2 = IndexProceso - 100
                            wEnter = True
                            sDescrip = ValorEfectivo2
                            frmPago.txtME(0).Caption = Format(ValorEfectivo2, "###,###,###,##0.00")
                            frmPago.nAbonoN = ValorEfectivo2
                            frmPago.nEN = ValorEfectivo2
                            frmPago.CalculaTotal
                        ElseIf IndexProceso = 999 Then
                            'ValorEfectivo = IndexProceso - 100
                            wEnter = True
                            sDescrip = nMonto
                            frmPago.txtME(0).Caption = Format(nMonto, "###,###,###,##0.00")
                            frmPago.nAbonoN = nMonto
                            frmPago.nEN = nMonto
                            frmPago.CalculaTotal
                        End If
                      frmPago.Show vbModal
                   End If
                End If
        End If
        
        
        '-----------------------
        If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
             If Not FacturarTCPIP(3, sDocumento, 0) Then
                MsgBox ("La confirmacion ha fallado, reenviar Documento!!!"), vbInformation, sMensaje
             End If
        End If
        '------------------------
    End If

    '------------------------- liberacion de la mesa en caso tengan
    Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa ='" & sMesa & "'"
    Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica = 1 where tCodigoPedido ='" & sPedido & "'"
    'Juntar Mesa
    Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
    '-------------------------------------------------------------------

    Cn.Execute "delete " & sDetalle
    Cn.Execute "delete " & sComboDetalle
    Cn.Execute "delete " & sComboPropiedad
    Cn.Execute "delete " & sProductoPropiedad

    RsDetalle.Requery
    RsComboPropiedad.Requery
    RsProductoPropiedad.Requery
    Inicializar
    Screen.MousePointer = vbDefault
    Exit Sub
fin:
wEnter = False
RsDetalle.Requery
RsComboPropiedad.Requery
RsProductoPropiedad.Requery
'Inicializar
Screen.MousePointer = vbDefault
Call Log_Inforest("CAJA RAPIDA", "EMISION DE DOCUMENTO", sPedido, "", sDocumento, Error, "", "FALLA AL GENERAR DOCUMENTO CAJA RAPIDA ", sUsuario)
MsgBox "Error: Emision de Documento / " + msgError & vbNewLine & Error, vbInformation, sMensaje
    
End Sub


Private Sub ImprimeFormatoAConsumo(ByVal nDocumento As String)

                    Dim ReporteC As New dsrBoletaC
                    
                    If RsTipoDocumento!lImprimeImageCab Then
                       iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
                    End If
                    
                    ReporteC.DiscardSavedData
                    ReporteC.Database.SetDataSource RsImpresion 'RsImpDocumentoE
                    
                    If xImpresioDE = "B" Then
                       ReporteC.Text13.SetText "BOLETA DE VENTA ELECTRONICA"
                    ElseIf xImpresioDE = "F" Then
                       ReporteC.Text13.SetText "FACTURA ELECTRONICA"
                    End If
                                                        
                    ReporteC.Text8.SetText sRazonSocial
                    ReporteC.ReportTitle = sDireccion
                    ReporteC.Text15.SetText sTelefono
                    ReporteC.Text33.SetText sFax
                    ReporteC.Text16.SetText sRUC
                    ReporteC.Text50.SetText sWeb
                    
                    ReporteC.Text31.SetText sDetalleConsumo
                    
                    If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & nDocumento & "')", Cn) = 1 Then
                       ReporteC.ReportComments = tTextoAgenteRetencion
                    End If
                    
                    xMontoTexto = "SON: " & NumeroCadena(str(RsImpresion!nVenta)) & " " & sMonedaN
                    ReporteC.Text4.SetText xMontoTexto
                    ReporteC.Text32.SetText xImpresionFE

'                        frmEmite.CRViewer.DisplayGroupTree = False
'                        frmEmite.CRViewer.ReportSource = ReporteC
'                        frmEmite.CRViewer.ViewReport
'                        frmEmite.Show vbModal
                    
                    ReporteC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    ReporteC.PaperOrientation = crPortrait
                    ReporteC.PrintOut False, 1, False, 1, 1
                    '----------------
                                                      
                
                    If iImagenCab Then
                       Kill App.path & "\cliente.jpg"
                    End If
End Sub


Private Sub ImprimeFormatoA(ByVal nDocumento As String)
                    Dim Reporte As New dsrBoleta

                    If RsTipoDocumento!lImprimeImageCab Then
                       iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
                    End If
                
                    Reporte.DiscardSavedData
                    Reporte.Database.SetDataSource RsImpresion 'RsImpDocumentoE
                                                        
                    If xImpresioDE = "B" Then
                       Reporte.Text13.SetText "BOLETA DE VENTA ELECTRONICA"
                    ElseIf xImpresioDE = "F" Then
                       Reporte.Text13.SetText "FACTURA ELECTRONICA"
                    End If
                    
                    Reporte.Text8.SetText sRazonSocial
                    Reporte.ReportTitle = sDireccion
                    Reporte.Text15.SetText sTelefono
                    Reporte.Text14.SetText sFax
                    Reporte.Text16.SetText sRUC
                    Reporte.Text50.SetText sWeb
                    
                    If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & nDocumento & "')", Cn) = 1 Then
                    Reporte.ReportComments = tTextoAgenteRetencion
                    End If
                    
                    xMontoTexto = "SON: " & NumeroCadena(str(RsImpresion!nVenta)) & " " & sMonedaN
                    Reporte.Text4.SetText xMontoTexto
                    Reporte.Text31.SetText xImpresionFE

'                        frmEmite.CRViewer.DisplayGroupTree = False
'                        frmEmite.CRViewer.ReportSource = Reporte
'                        frmEmite.CRViewer.ViewReport
'                        frmEmite.Show vbModal
                    
                    Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    Reporte.PaperOrientation = crPortrait
                    Reporte.PrintOut False, 1, False, 1, 1
                    '----------------
                    
                    If iImagenCab Then
                       Kill App.path & "\cliente.jpg"
                    End If
End Sub

Public Sub GeneraPedido()
    Dim oComando As clsComando
    Set oComando = New clsComando
    ' INSERTA CABECERA DEL PEDIDO 07102020 - GCAA
    If Not oComando.CreateCmdSp("spIns_MPEDIDO", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 7, sClienteFrecuente
    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, sTipoPedido
    oComando.CreateParameter "@lPrioridad", adBoolean, adParamInput, 1, 0
    oComando.CreateParameter "@tTipoAtencion", adVarChar, adParamInput, 2, "01"
    oComando.CreateParameter "@tMesa", adVarChar, adParamInput, 3, ""
    oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, Right(sMozo, 4)
    oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 4, sMotorizado
    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
    oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 2, sSalon
    oComando.CreateParameter "@tTurno", adVarChar, adParamInput, 10, sTurno
    oComando.CreateParameter "@tObservacion", adVarChar, adParamInput, 250, sObser
    oComando.CreateParameter "@nTiempo", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, Mid(sUsuario, 1, 15)
    oComando.CreateParameter "@nAdulto", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nNino", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nMesa", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@tPuntoVenta", adVarChar, adParamInput, 2, ""
    oComando.CreateParameter "@tHabitacion", adVarChar, adParamInput, 6, ""
    oComando.CreateParameter "@tReserva", adVarChar, adParamInput, 6, ""
    oComando.CreateParameter "@tPasajero", adVarChar, adParamInput, 50, ""
    oComando.CreateParameter "@tCompania", adVarChar, adParamInput, 5, ""
    oComando.CreateParameter "@tContacto", adVarChar, adParamInput, 4, ""
    oComando.CreateParameter "@nDescuento", adDouble, adParamInput, 10, xDescuento
    oComando.CreateParameter "@tDescuento", adVarChar, adParamInput, 3, sCodigoDescuento
    oComando.CreateParameter "@tObservacionDescuento", adVarChar, adParamInput, 250, sDescripcionDescuento
    oComando.CreateParameter "@tAutorizaDescuento", adVarChar, adParamInput, 15, IIf(sCodigoDescuento = "", "", tAutorizaDescuento)
    oComando.CreateParameter "@nTiempoDelivery", adInteger, adParamInput, 10, nTiempoDelivery
    oComando.CreateParameter "@tTienda", adVarChar, adParamInput, 3, ""
    oComando.CreateParameter "@fDiaContable", adDate, adParamInput, 10, obtieneDiaContable
    oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 20, IIf(txtFechaEntrega.Caption = "", Null, Format(txtFechaEntrega.Caption, "dd/MM/yyyy HH:mm"))
    'invitado2013
    oComando.CreateParameter "@tCodigoInvitado", adVarChar, adParamInput, 10, sCodigoInvitado
    'pariente2014
    oComando.CreateParameter "@tCodigopariente", adVarChar, adParamInput, 7, sCodigoParienteSeleccionado
    'entregara
    oComando.CreateParameter "@tEntregarA", adVarChar, adParamInput, 20, IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20))
    oComando.CreateParameter "@nTiempoAntesEnvio", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nMontoMaximo", adInteger, adParamInput, 250, 0
    oComando.CreateParameter "@tPedido", adVarChar, adParamOutput, 10, Pedido
    
    'origen de ventas
    If vOrigenVentas = Null Or vOrigenVentas = "" Then
        vOrigenVentas = "00"
    End If
    oComando.CreateParameter "@codigoOrigenVentas", adVarChar, adParamInput, 2, vOrigenVentas
    
    'Pin Totem3
    oComando.CreateParameter "@tTotemMesa", adVarChar, adParamInput, 10, ""
           
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    If Not oComando.ExecSP Then
       Set oComando = Nothing
       Exit Sub
    Else
       Pedido = oComando.GetParameterValue("@tPedido")
       
       Cn.Execute "UPDATE MPEDIDO SET FDIACONTABLE='" & Format(obtieneDiaContable, "yyyyMMdd") & "'  where tcodigopedido='" & Pedido & "' "
    End If
                                  
    'Actualiza el Numero de Pedido en el Detalle Temporal
    Cn.Execute "Update [" & sDetalle & "] Set tCodigoPedido = '" & Pedido & "'"
    
    'AGREGAR GCAA PARA OBTENER tAplicaExoneracion, tAplicaInafectacion
    
    ' Isql = "select tAplicaExoneracion, tAplicaInafectacion from vPRODUCTO where Codigo= '" + Trim(Pedido) + "'"
    ' Set RsProductox = Lib.OpenRecordset(Isql, Cn)
    
    
   ' Dim vtAplicaExoneracion As Bite
   ' vtAplicaExoneracion = RsProductox!tAplicaExoneracion
   ' Dim vtAplicaInafectacion As Bite
   ' vtAplicaInafectacion = RsProductox!tAplicaInafectacion
    
    'Inserta el Detalle
    '07102020 - GCAA
    Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
               "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
               "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,fregistro, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tsubalmacen,tCodigoEtiqueta,tAplicaExoneracion,tAplicaInafectacion ,tunidadnegocio,fDiaContable, tCajaD ) " & _
               "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
               "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
               "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,getDate(), nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubalmacen,tCodigoEtiqueta,0,0,'" & sUnidadNegocio & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "', '" & sCaja & "' " & _
               "From [" & sDetalle & "] where tEstadoItem='N'"
    
    
   ' Dim oComando As clsComando
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spUpdate_DPEDIDO_Ina", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    oComando.CreateParameter "@pedido", adVarChar, adParamInput, 10, Pedido
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If

    If Not oComando.ExecSP Then
       Set oComando = Nothing
       Exit Sub
     End If




    'Actualiza el Numero de Pedido en el Detalle Combos
    Cn.Execute "Update [" & sComboDetalle & "] Set tCodigoPedido = '" & Pedido & "'"
    
    'Inserta Combo
    Cn.Execute "Insert into CPEDIDO select * from " & sComboDetalle
    
    'Inserta las propiedades de los Combos
    Cn.Execute "Insert into TCOMBOPROPIEDAD select '" & Pedido & "', tItem, tItemCombo, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra, 1,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario from " & sComboPropiedad
    
    'Inserta las propiedades
    Cn.Execute "Insert into TPRODUCTOPROPIEDAD select '" & Pedido & "', tItem, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra, 1,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario from " & sProductoPropiedad
    sPedido = Pedido
End Sub



Public Sub GrabaProducto()
   Screen.MousePointer = vbHourglass
   
   Isql = "Update [" & sDetalle & "] Set nPrecioNeto = " & nPBase & ", " & _
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
           "nImpuesto3 = " & nImpuesto3 * nCantidad & ", " & _
           "tCortesia = '" & sCortesia & "', " & _
           "lImprime = 0 " & _
           "where tItem = '" & sItem & "'"
           
           Cn.Execute Isql
           RsDetalle.Requery
           RsDetalle.MoveFirst
           RsDetalle.Find "tItem = '" & sItem & "'"
   nMonto = Format(Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn), "#,###,##0.00")
   Screen.MousePointer = vbDefault
End Sub

Public Sub CalculaPrecio()
    Dim Acumulado As Double
    
    If nPVenta = 0 Then
       txtDPorcentaje.Caption = "0.00"
       txtRPorcentaje.Caption = "0.00"
       nRecargo = 0
       nDescuento = 0
       nImpuesto1 = 0
       nImpuesto2 = 0
       nImpuesto3 = 0
    Else
        Select Case pais 'ok
            Case "001" 'Bolivia
                    nPVenta = nOficial - nDescuento + nRecargo
                    
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = (Acumulado / 100)
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                    txtDPorcentaje.Caption = Format(nDescuento * 100 / nOficial, "###,###,###,##0.00")
                    txtRPorcentaje.Caption = Format(nRecargo * 100 / nOficial, "###,###,###,##0.00")
        
            Case Else 'Peru, Ecuador
                    nPVenta = nOficial - nDescuento + nRecargo
                    
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = 1 + (Acumulado / 100)
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                    txtDPorcentaje.Caption = Format(nDescuento * 100 / nOficial, "###,###,###,##0.00")
                    txtRPorcentaje.Caption = Format(nRecargo * 100 / nOficial, "###,###,###,##0.00")
        
        End Select
    End If
    txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
    txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
    txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
    
    txtNeto.Caption = Format(nPBase, "###,###,##0.00")
    txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
    txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
End Sub

Public Sub Impuesto()
   Label1(10).Caption = sImpuesto1 & " :"
   Label1(11).Caption = sImpuesto2 & " :"
   Label1(12).Caption = sImpuesto3 & " :"
   
   Label1(10).Visible = IIf(sImpuesto1 = "", False, True)
   Label1(11).Visible = IIf(sImpuesto2 = "", False, True)
   Label1(12).Visible = IIf(sImpuesto3 = "", False, True)
   
   txtImpuesto1.Visible = IIf(sImpuesto1 = "", False, True)
   txtImpuesto2.Visible = IIf(sImpuesto2 = "", False, True)
   txtImpuesto3.Visible = IIf(sImpuesto3 = "", False, True)
   
   cmdImpuesto(0).Caption = sImpuesto1
   cmdImpuesto(1).Caption = sImpuesto2
   cmdImpuesto(2).Caption = sImpuesto3
   
   cmdImpuesto(0).Visible = IIf(sImpuesto1 = "", False, True)
   cmdImpuesto(1).Visible = IIf(sImpuesto2 = "", False, True)
   cmdImpuesto(2).Visible = IIf(sImpuesto3 = "", False, True)
End Sub



















Private Sub grdCombo_DblClick()
    If grdCombo.col = 3 Then
        cmdCombo_Click 1
        grdCombo.col = 0
    End If
End Sub

Private Sub grdCombo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  AsignaCombo
  If fraPropiedad.Visible = True Then
     cmdOpcion_Click (6)
  End If
  
  On Error Resume Next
  If grdCombo.col = 3 Then
    cmdCombo_Click 1
    grdCombo.col = 0
  End If
End Sub












Private Sub grdDetalle_DblClick()
    If grdDetalle.col = 6 Then
        cmdDetalle_Click 0
    End If
End Sub

Private Sub grdDetalle_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   
   If fraPropiedad.Visible = True Then
      nPos = RsDetalle.AbsolutePosition
      RsDetalle.Requery
      RsDetalle.AbsolutePosition = nPos
   End If
   AsignaProducto
   
   If lPropiedad Then
      lPropiedad = False
      cmdDetalle_Click (5)
   End If

   If lComboActivo Then
      lComboActivo = False
      cmdOpcion_Click (10)
   End If
   On Error Resume Next
    If grdDetalle.col = 6 Then
        cmdDetalle_Click 0
        grdDetalle.col = 0
    End If
   
   txtBarra.SetFocus
End Sub

Public Sub AsignaCombo()
   If Not RsCombo.EOF Then
      sCombo = IIf(IsNull(RsCombo!tProductoCombo), "", RsCombo!tProductoCombo)
      xItem = IIf(IsNull(RsCombo!tItemCombo), "001", RsCombo!tItemCombo)
      lblObservacion.Text = IIf(IsNull(RsCombo!tObservacion), "", RsCombo!tObservacion)
      ListarOperadoresConFiltro (sCombo) 'OO
      AsignaComboPropiedad
   End If
End Sub

Private Sub cmdPropiedad_Click(Index As Integer)
    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMObra As Double
    
    Dim Cantidad As Double
    Dim ncantidadPropiedad As Double
    
    RsPropiedad.MoveFirst
    RsPropiedad.Find ("Descripcion = '" & cmdPropiedad(Index).Caption & "'")
      If Not (RsOperador.EOF Or RsOperador.BOF) Then
     nOperadorPropiedad = Calcular("select isnull(ncontrol,0) as codigo from voperador where codigo='" & RsOperador!codigo & "'", Cn)
     End If
  
    
    
    If cmdPropiedad(Index).FontBold = True Then
       cmdPropiedad(Index).FontBold = False
       If Not RsPropiedad.EOF Then
          If wAgregaCombo Then
             Cantidad = Calcular("select isnull(ncantidad,1) as codigo from " & sComboPropiedad & " where      titem='" & sItem & "' and titemcombo='" & xItem & "' and  tproducto='" & sCombo & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' ", Cn)
             If RsPropiedad!nPrecio <> 0 Then
                nMonto = CambiaPrecio(nPVenta - (RsPropiedad!nPrecio * Cantidad))
                txtMonto.Caption = Format(nMonto, "###,##0.00")
             End If
             Cn.Execute "delete " & sComboPropiedad & " where tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and tProducto='" & sCombo & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "'"
          Else
             Cantidad = Calcular("select isnull(ncantidad,1) as codigo from " & sProductoPropiedad & " where     titem='" & sItem & "' and tproducto='" & sProducto & "' and tcodigopropiedad='" & RsPropiedad!codigo & "'  ", Cn)
             If RsPropiedad!nPrecio <> 0 Then
                nMonto = CambiaPrecio(nPVenta - (RsPropiedad!nPrecio * Cantidad))
                txtMonto.Caption = Format(nMonto, "###,##0.00")
             End If
             Cn.Execute "delete " & sProductoPropiedad & " where tItem = '" & sItem & "' and tProducto='" & sProducto & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "'"
          End If
          
          If Cantidad <> 1 Then
              lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & cmdPropiedad(Index).Caption & ": (" & Cantidad & "), ", "")
            ' lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", ", "")
          Else
            lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & cmdPropiedad(Index).Caption & ", ", "")
          End If
          
          'lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", ", "")
       End If
    Else
    
            ncantidadPropiedad = 1
             If RsPropiedad!lsolicitacantidad = 1 Or RsPropiedad!lsolicitacantidad = True Then
                sTipo = "Prepintado"
                
                sCodigo = ncantidadPropiedad
                
                frmNumPad.Show vbModal
                If wEnter And val(sDescrip) > 0 Then
                
                            ncantidadPropiedad = sDescrip
                            
                End If
            End If
               
        
    
    
       If nOperadorPropiedad > 0 Then
          If wAgregaCombo Then
             Isql = "SELECT COUNT(" & sComboPropiedad & ".tCodigoPropiedad) AS codigo " & _
                    "FROM " & sComboPropiedad & " INNER JOIN dbo.TPROPIEDAD ON " & sComboPropiedad & ".tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND " & sComboPropiedad & ".tProducto = dbo.TPROPIEDAD.tProducto " & _
                    "where tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and " & sComboPropiedad & ".tProducto='" & sCombo & "'  and tOperador='" & RsOperador!codigo & "'"
             If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                Exit Sub
             End If
          Else
             Isql = "SELECT COUNT(" & sProductoPropiedad & ".tCodigoPropiedad) AS codigo FROM " & sProductoPropiedad & " INNER JOIN " & _
                    "dbo.TPROPIEDAD ON " & sProductoPropiedad & ".tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND " & sProductoPropiedad & ".tProducto = dbo.TPROPIEDAD.tProducto " & _
                    "where tItem = '" & sItem & "' and tOperador='" & RsOperador!codigo & "'"
             If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                Exit Sub
             End If
          End If
       End If
       
       cmdPropiedad(Index).FontBold = True
       If Not RsPropiedad.EOF Then
          nInsumo = IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo)
          nGasto = IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto)
          nMObra = IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra)
       
          If wAgregaCombo Then
             Cn.Execute "Insert into " & sComboPropiedad & " values ('" & sItem & "', '" & xItem & "', '" & RsPropiedad!codigo & "', '" & sCombo & "', '" & RsPropiedad!tEnlace & "', " & IIf(IsNull(RsPropiedad!nInsumo), 0, ncantidadPropiedad * RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, ncantidadPropiedad * RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, ncantidadPropiedad * RsPropiedad!nManoObra) & ", " & ncantidadPropiedad & ", " & IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra) & ") "
             If RsPropiedad!nPrecio <> 0 Then
                nMonto = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad))
                txtMonto.Caption = Format(nMonto, "###,##0.00")
             End If
          Else
             Cn.Execute "Insert into " & sProductoPropiedad & " values ('" & sItem & "', '" & RsPropiedad!codigo & "', '" & sProducto & "', '" & RsPropiedad!tEnlace & "', " & IIf(IsNull(RsPropiedad!nInsumo), 0, ncantidadPropiedad * RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, ncantidadPropiedad * RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, ncantidadPropiedad * RsPropiedad!nManoObra) & ", " & ncantidadPropiedad & "," & IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra) & " )"
             If RsPropiedad!nPrecio <> 0 Then
                nMonto = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad))
                txtMonto.Caption = Format(nMonto, "###,##0.00")
             End If
          End If
          'lblResumen.Text = lblResumen.Text & RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", "
          
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


'
Public Sub AsignaPropiedad()
    Dim i As Integer
    If RsOperador.RecordCount > 0 Then
       RsPropiedad.Filter = "tOperador = '" & RsOperador!codigo & "' and tProducto='" & sProducto & "'"
       nOperadorPropiedad = IIf(IsNull(RsOperador!nControl), 0, RsOperador!nControl)
    Else
       RsPropiedad.Filter = "tOperador = '  ' and tProducto='" & sProducto & "'"
       nOperadorPropiedad = 0
    End If
    AsignaComando 19, RsPropiedad, cmdPropiedad()

    ' Validaciòn de visibilidad para Propiedades -BR
    With RsPropiedad
        If .RecordCount > 0 Then
            .MoveFirst
            For i = 1 To IIf(.RecordCount > 19, 19, .RecordCount)
                If Not IsNull(!visibilidad) And !visibilidad Then
                    Select Case sTipoPedido
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
                    cmdPropiedad(i).Visible = True
                End If
                cmdPropiedad(i).Caption = !descripcion
                .MoveNext
            Next i
        End If
    End With
    
    For i = 1 To 19
        cmdPropiedad(i).FontBold = False
    Next i
    
    
    lblResumen.Text = ""
    RsProductoPropiedad.Filter = "tItem='" & sItem & "'"
    If Not RsProductoPropiedad.EOF Then
       RsProductoPropiedad.MoveFirst
       Do While Not RsProductoPropiedad.EOF
          For i = 1 To 19
              If cmdPropiedad(i).Caption = RsProductoPropiedad!descripcion And RsOperador!descripcion = RsProductoPropiedad!Operador Then
                 cmdPropiedad(i).FontBold = True
                 Exit For
              End If
          Next i
          If RsProductoPropiedad!nCantidad = 1 Then
            lblResumen.Text = lblResumen.Text & LTrim(RsProductoPropiedad!Operador) & " " & LTrim(RsProductoPropiedad!descripcion) & ", "
          Else
             lblResumen.Text = lblResumen.Text & LTrim(RsProductoPropiedad!Operador) & " " & LTrim(RsProductoPropiedad!descripcion) & ": (" & RsProductoPropiedad!nCantidad & "), "
          End If
          RsProductoPropiedad.MoveNext
       Loop
    End If

End Sub










'Private Sub Timer1_Timer()
'If lIntPatio Then
''        Me.bocina.url = App.Path & "\bmps\patio.mp3"
'        Isql = "select top 50 tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion=3 and testadopedido<>'03' and ttipopedido in( '02','01')order by fregistro desc "
'        Set RsPedido = Lib.OpenRecordset(Isql, Cn)
'        If RsPedido.RecordCount > 0 Then
'           ' Me.cmdOpcion(2).Caption = "(" + CStr(RsPedido.RecordCount) + ") Patio de Comida"
'            'If Me.cmdOpcion(10).backColor = &H8000000F Then
'                Me.TimerColor.Enabled = True
'                'Me.cmdOpcion(2).backColor = &H808000
''                MMControl1.FileName = App.Path & "\bmps\patio.wav"
''                MMControl1.Command = "open"
''                MMControl1.Command = "play"
'            'Else
'                'Me.cmdOpcion(10).backColor = &H8000000F
'            'End If
'        Else
'            Me.TimerColor.Enabled = False
'            'Me.cmdOpcion(2).backColor = &H8000000F
'           ' Me.cmdOpcion(2).Caption = "Refrescar"
'            'MMControl1.Command = "stop"
'        End If
'
''        Call comandaAuto
'
'    End If
'End Sub

'Private Sub TimerColor_Timer()
' If Me.cmdOpcion(2).backColor = &H8000000F Then
'        Me.cmdOpcion(2).backColor = &H808000
'    Else
'        Me.cmdOpcion(2).backColor = &H8000000F
'    End If
'End Sub

Private Sub txtBarra_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 And txtBarra.Text <> "" Then
        Dim xxx As String
        
        If lRotulado = True Then
              'CESAR ROTULADO
              Dim rCodigoEtiqueta As String
              Dim rCodigoProducto As String
              Dim rLenBarra As String
              Dim rCantidad As Double
              Dim X As Integer
              
              xxx = RsProducto.Filter
              
              rLenBarra = Len(Trim(txtBarra.Text))
              X = rLenBarra - 31
              
              If nLongitudBarra <> 0 Then
                rCodigoProducto = Mid(txtBarra.Text, nLongitudBarra + 1, 7)
                rCodigoEtiqueta = Mid(txtBarra.Text, 1, nLongitudBarra)
                
                If lCapturaPeso Then
                   rCantidad = val(Mid(txtBarra.Text, 31 + 1, X))
                   InsertaProductoRotulado rCodigoProducto, rCantidad, rCodigoEtiqueta
                Else
                   rCantidad = 1
                   InsertaProductoRotulado rCodigoProducto, rCantidad, rCodigoEtiqueta
                End If
              Else
                  MsgBox "Error: Longitud de barra no registrada", vbCritical, sMensaje
              End If
              txtBarra.Text = ""
              sProducto = ""
              RsProducto.Filter = IIf(xxx = "0", "", xxx)
        Else
     
            xxx = RsProducto.Filter
            RsProducto.Filter = adFilterNone
            RsProducto.MoveFirst
            RsProducto.Find "tbarra = '" & Trim(txtBarra.Text) & "'"
            
            If Not RsProducto.EOF Then
               sProducto = RsProducto!codigo
               
               If lBal And RsProducto!lBalanza Then
                  Dim nResultado As Double
                  nResultado = Pesar(nBalanzaPuerto)
                  nResultado = Format(nResultado, "#,##0.00")
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
                     Dim nCantidadBarra As Double
                       
                        If lCapturaPeso Then
                        
                         If EAN13 Then
                                nCantidadBarra = val(Mid(txtBarra.Text, nLongitudBarra + 1, 1) + "." + Mid(txtBarra.Text, nLongitudBarra + 2, 3))
                             Else
                                nCantidadBarra = val(Mid(txtBarra.Text, nLongitudBarra + 1, 2) + "." + Mid(txtBarra.Text, nLongitudBarra + 3, 4))
                             End If
                            'nCantidadBarra = Val(Mid(txtBarra.Text, nLongitudBarra + 1, 2) + "." + Mid(txtBarra.Text, nLongitudBarra + 3, 4))
                           InsertaProducto nCantidadBarra
                        Else
                           nCantidadBarra = val(Mid(txtBarra.Text, nLongitudBarra + 1, 3) + "." + Mid(txtBarra.Text, nLongitudBarra + 4, 3))
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
            txtBarra.Text = ""
            sProducto = ""
            RsProducto.Filter = IIf(xxx = "0", "", xxx)
        End If

   End If
   
End Sub
Public Sub InsertaProductoRotulado(codigoProducto As String, xCantidad As Double, codigoEtiqueta As String)
    'CESAR ROTULADO
    Dim nValor As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    'ORDEN
    Dim RsOrd As Recordset
    Dim nOrden As Integer
    'DETALLE
    Dim lProductoMultiArea As Boolean
    Dim tsubalmacen As String
    Dim tAreaProduccion As String
    'OFERTA
    Dim tOferta As String
    Dim nOferta As Double
    tOferta = ""
    nOferta = 0
    
    RsProducto.Filter = adFilterNone
    RsProducto.MoveFirst
    RsProducto.Find "Codigo = '" & Trim(codigoProducto) & "'"
    
    
    If Not RsProducto.EOF Then

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
        
        'PRECIO OFICIAL
        nOficial = IIf(RsProducto!tMONEDA = "02", nOficial * nTC, nOficial)
        
        'DESCUENTO
        If xDescuento <> 0 And RsProducto!lDescuento Then
              nPVenta = nOficial - (nOficial * xDescuento / 100)
              nDescuento = nOficial - nPVenta
        Else
              nPVenta = nOficial - nOferta
              nDescuento = nOficial - nPVenta
        End If


        Select Case pais
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
            
            sItem = Lib.Correlativo(Calcular("select max(tItem) as codigo from [" & sDetalle & "]", Cn), 3)
            'CALCULAR ITEM
            If RsDetalle.RecordCount = 0 Then
               'sitem = "001"
               nOrden = 1
            Else
               'sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from [" & sDetalle & "]", Cn), 3)
               If lOrden Then
                  Set RsOrd = Lib.OpenRecordset("select nOrden, lImprime from " & sDetalle & " Order by nOrden DESC", Cn)
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
            
            
            'TSUBALMACEN
            lProductoMultiArea = Calcular("select isnull(lmultiarea,0) as codigo from tproducto where tcodigoproducto='" & RsProducto.Fields("codigo") & "'", Cn)
    
            If lProductoMultiArea = False Then
                tsubalmacen = ""
            Else
                tAreaProduccion = Calcular("select isnull(tsubalmacen,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
                
                tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
                
                If tsubalmacen = "0" Then
                    tsubalmacen = ""
                End If
            End If
            

            'DETALLE TEMPORAL
             Isql = "insert into [" & sDetalle & "] " & _
                    "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                    "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
                    "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
                    "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem, tCodigoEtiqueta, tsubalmacen, toferta) " & _
                    "Values( '" & Pedido & "', '01', " _
                            & "'" & sItem & "', " _
                            & "'" & codigoProducto & "', " _
                            & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                            & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                            & nPBase & ", " & nRecargo & ", " _
                            & nDescuento & ", " _
                            & nOficial & ", " _
                            & nImpuesto1 & ", " & nImpuesto2 & ", " & nImpuesto3 & ", " _
                            & nPVenta & ", " & xCantidad & ", " _
                            & nPVenta * xCantidad & ", " _
                            & nImpuesto1 * xCantidad & ", " & nImpuesto2 * xCantidad & ", " & nImpuesto3 * xCantidad & ", " _
                            & "0, '" & RsProducto!tArea & "', " & IIf(RsProducto!lImprimeArea, -1, 0) & "," _
                            & IIf(RsProducto!lCombinacion, 1, 0) & ", " & RsProducto!nCombinacion & ", " _
                            & nInsumo & ", " _
                            & nGasto & ", " _
                            & nMObra & ", " _
                            & nOrden & ", " _
                            & "'N','" & codigoEtiqueta & "','" & tsubalmacen & "','" & tOferta & "') "
             Cn.Execute Isql
             RsDetalle.Requery
              nMonto = Format(Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn), "#,###,##0.00")
             RsDetalle.MoveLast
             Screen.MousePointer = vbDefault
             
    Else
       MsgBox "Producto no encontrado", vbCritical, sMensaje
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
       xItem = Lib.Correlativo(Calcular("select max(tItemCombo) as codigo from " & sComboDetalle & " where tItem = '" & sItem & "'", Cn), 3)
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
          nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto10
          lImp2 = RsProducto!lImpuesto11
          lImp3 = RsProducto!lImpuesto12
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
    
   Select Case pais 'ok
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
    
    Dim xCantCombo As Double
    xCantCombo = Calcular("Select ISNULL(nCantidad,0) as Codigo From TCOMBO where tCombo = '" & RsDetalle!tCodigoProducto & "' and tCodigoProducto='" & wProducto & "'", Cn)
    
    Isql = "insert into " & sComboDetalle & " " & _
           "(tCodigoPedido, tProducto, tItem, tItemCombo, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden ) " & _
           "Values(   '" & Pedido & "', " _
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
            txtMonto.Caption = Format(CambiaPrecio(nPVenta + RstCombo!nAumento / nCantidad), "#,###,##0.00")
        End If
    End If
    'Fin Oscar Ortega--------------------------------------------------------
    RsCombo.Requery
    RsCombo.MoveLast
    AsignaCombo
    Screen.MousePointer = vbDefault
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
   RsCombo.Filter = "tItem='" & sItem & "'"
   
   Select Case Funcion
          Case Is = "A"
               Do While Not RsTemp.EOF
                  If RsCombo.RecordCount = 0 Then
                     xItem = "001"
                  Else
                     xItem = Lib.Correlativo(Calcular("select max(tItemCombo) as codigo from " & sComboDetalle & " where tItem = '" & sItem & "'", Cn), 3)
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
               
                  Select Case pais 'ok
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
                                     
                  Isql = "insert into " & sComboDetalle & _
                       " (tCodigoPedido, tProducto, tItem, tItemCombo, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden ) " & _
                         "Values(   '" & Pedido & "', " _
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
                                 & IIf(RsProducto!lImprimeArea, -1, 0) & ", 0, " _
                                 & RsProducto!nOrden & ") "

                  Cn.Execute Isql
                  RsCombo.Requery
                  RsTemp.MoveNext
               Loop
          Case Is = "M"
               Do While Not RsTemp.EOF
                  Dim X As Double
                  X = Calcular("select nCantidad as Codigo FROM TCOMBO where tCombo='" & RsTemp!tCombo & "' and tCodigoproducto ='" & RsTemp!tCodigoProducto & "'", Cn)
                  Isql = "update " & sComboDetalle & " set nCantidad = " & X * Cantidad & " where tCodigoPedido='" & Pedido & "' and tItem='" & sItem & "' and tProductocombo='" & RsTemp!tCodigoProducto & "'"
                  Cn.Execute Isql
                  RsCombo.Requery
                  RsTemp.MoveNext
               Loop
          
          Case Is = "D"
               For i = 1 To Cantidad
                   Isql = "DELETE from " & sComboDetalle & " " & _
                          "where tCodigoPedido ='" & Pedido & "' and tProducto='" & Combo & "' and tItem='" & sItem & "'"
                   Cn.Execute Isql
                   RsCombo.Requery
               Next i
   End Select
   Set RsTemp = Nothing
End Sub

Public Sub Inicializar()
   sMozo = "0000"
   sCodigoDescuento = ""
   sClienteFrecuente = ""
   sCodigoParienteSeleccionado = ""
   sCodigoInvitado = ""
   xDescuento = 0
   nTope = 0
   ltope = False
   tAutorizaDescuento = ""
   sObser = ""
   sDescripcionDescuento = ""
   xDescuento = 0
   txtTitulo.Caption = " Caja Rápida : " & sCaja & " Mesero: Sin Mesero" '& IIf(sPedido = "", "", " - " & sPedido)
   txtObservacion.Caption = sObser
   txtEntregar.Caption = ""
   fraMozo.Visible = False
   fraDetalle.Visible = False
   fraPropiedad.Visible = False
   Pedido = ""
   wCombo = False
   wAgregaCombo = False
   nCombo = 0
   nMonto = 0
   sMesa = ""
   txtTelefono.Caption = ""
   txtCliente.Caption = ""
   VisualizaMonto
   Cn.Execute "delete from " & sComboDetalle
   RsCombo.Requery
   txtFechaEntrega.Caption = ""
   vOrigenVentas = ""

    cmdTipoDocumento(0).backColor = &H8BDD79
    cmdTipoDocumento(1).backColor = &HF2FAED
    cmdTipoDocumento(2).backColor = &HF2FAED
    cmdTipoDocumento(3).backColor = &HF2FAED
    cmdTipoDocumento(4).backColor = &HF2FAED
   
   
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
    Inicializar
    fraEliminacion.Visible = False
    tabProducto.Visible = True
    If Sw Then
       EliminaCabecera
    Else
         'KDS2
        If lKDS Then
            Dim kdsRsCabecera As Recordset
            Isql = "SELECT * From vPedidoCabecera Where Codigo = '" & Pedido & "' Order By codigo "
            Set kdsRsCabecera = Lib.OpenRecordset(Isql, Cn)
            Call KDS_EliminarProducto(kdsRsCabecera, sItem, 0)
        End If
       EliminaItem
    End If
    
    'JSJR 09/01/2025
    wAnulaDocumento = False
    
    Sw = False
    ActivaCabecera True
End Sub

Public Sub ActivaCabecera(Activa As Boolean)
   Dim i As Integer
   For i = 1 To 4
       cmdTipoDocumento(i).Enabled = Activa
       If i <> 1 And i <> 6 Then
            cmdNavegarPicture(i).Enabled = Activa
       End If
   Next i
   
   cmdDetallePicture(1).Enabled = Activa
   cmdDetallePicture(2).Enabled = Activa
   cmdDetallePicture(3).Enabled = Activa
   
   cmdNavegarPicture(5).Enabled = Activa
   cmdNavegar(6).Enabled = Activa
   'cmdDetallePicture(0).Enabled = Activa
   cmdDetallePicture(5).Enabled = Activa
   cmdDetallePicture(6).Enabled = Activa
   cmdDetallePicture(7).Enabled = Activa
   cmdDetallePicture(8).Enabled = Activa
   cmdDetallePicture(9).Enabled = Activa
   cmdDetallePicture(10).Enabled = Activa
   cmdDetallePicture(12).Enabled = Activa
   cmdDetallePicture(13).Enabled = Activa
   cmdDetallePicture(14).Enabled = Activa
   cmdOpcionPicture(0).Enabled = Activa
   cmdOpcionPicture(1).Enabled = Activa
   cmdOpcionPicture(2).Enabled = IIf(lMultiCajero, False, Activa)
   cmdOpcionPicture(5).Enabled = Activa
   cmdCabecera(1).Enabled = Activa
   cmdCabecera(2).Enabled = Activa
   FramePieForm.Visible = Activa
   cmdOpcionPicture(14).Enabled = Activa
   
   If lPrinter Then
      cmdOpcionPicture(8).Enabled = Activa
   End If

End Sub

Public Sub EliminaItem()
    Dim xMax As String
    Dim sMotivo As String
    
    If lPrinter Then
       'Impresion del Pedidos Anulados
       sMotivo = Calcular("select Descripcion as Codigo from vMotivoEliminacion where Codigo='" & sCodigo & "'", Cn)
       
       Isql = "select *, '" & sMotivo & "' as MotivoEliminacion FROM dbo.vPedido " & _
              "WHERE Codigo = '" & Pedido & "' and tItem = '" & sItem & "' and lImprime = 1 And lImprimeArea = 1 " & _
              "ORDER BY tItem,tetiqueta,combo"
                                                                                             
       Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
       Dim i As Integer
       If RsImpresion.RecordCount = 0 Then
          LimpiaRs
       Else
          RsArea.MoveFirst
          For i = 1 To RsArea.RecordCount
              RsImpresion.Filter = "tArea = '" & RsArea!tArea & "'"
              If RsArea!tIcono = "" Or sSalon = RsArea!tIcono Or (sSalon = "" And RsArea!nValor = 1) Then
                 If RsImpresion.RecordCount <> 0 Then
                    RsImpresion.MoveFirst
                    sPedido = Pedido
                    ImprimePedido RsImpresion, "A", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsProductoPropiedad, "Rapido"
                    sPedido = ""
                 End If
              End If
              RsArea.MoveNext
          Next i
          RsDetalle.Requery
       End If
    End If
    
    xMax = Calcular("select max(tItem) as Codigo from APEDIDO where tCodigoPedido='" & Pedido & "'", Cn)
    xMax = Lib.Correlativo(xMax, 3)
    Isql = "insert into APEDIDO (tCodigoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, " & _
           "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, " & _
           "tComanda, lImprime, tUsuario, fRegistro, tUsuarioAnulado, fRegistroAnulado, " & _
           "tObservacion, tObservacionAnulado, tEstadoItem, lImprimeArea, tArea, tMotivoEliminacion, tTurnoAnulado,fDiaContable) " & _
           "select '" & Pedido & "' as tCodigoPedido, '" & xMax & "' as tItem, tCodigoProducto, tCodigoGRupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, " & _
           "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tComanda, lImprime, " & _
           "'" & Mid(sUsuario, 1, 15) & "' as tUsuario, getDate() as fRegistro, " & _
           "'" & sUsuarioAutoriza & "' as tUsuarioAnulado, getDate() as fRegistroAnulado, " & _
           "tObservacion, '" & sDescrip & "' as tObservacion, tEstadoItem, lImprimeArea, tArea, '" & sCodigo & "', '" & sTurno & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
           "from " & sDetalle & _
           " where tCodigoPedido = '" & Pedido & "' and tItem = '" & sItem & "'"
    Cn.Execute Isql
    
        
 
      'INSUMOCRITICO23
    Dim rstItems As New ADODB.Recordset
    Set rstItems = New ADODB.Recordset
    'Set rstItems = Lib.OpenRecordset("select tcodigoinsumo,ncantidad from dpedido inner join tproducto on dpedido.tcodigoproducto=tproducto.tcodigoproducto where tcodigopedido='" & sPedido & "' and titem='" & sitem & "' and tproducto.lControlInsumoCritico=1 and isnull(tproducto.tcodigoinsumo,'')<>''  and isnull(dpedido.limprime,0)=1 ", Cn)
    Set rstItems = Lib.OpenRecordset("  usp_Inforest_RevertirInsumosCriticos '" & sPedido & "','" & sItem & "' ", Cn)
    If Not (rstItems.EOF Or rstItems.BOF) Then
        modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
    End If
       
    Cn.Execute "delete from DPEDIDO where tCodigoPedido = '" & Pedido & "' and tItem = '" & sItem & "'"
    Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido='" & Pedido & "' and tItem = '" & sItem & "'"
    Cn.Execute "delete CPEDIDO where tCodigoPedido='" & Pedido & "' and tItem = '" & sItem & "'"
    Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido='" & Pedido & "' and tItem = '" & sItem & "'"
    
    Cn.Execute "delete from " & sDetalle & " where tItem = '" & sItem & "'"
    Cn.Execute "delete from " & sComboDetalle & " where tItem ='" & sItem & "'"
    Cn.Execute "delete from " & sProductoPropiedad & " where tItem ='" & sItem & "'"
    Cn.Execute "delete from " & sComboPropiedad & " where tItem ='" & sItem & "'"
                          
    RsDetalle.Requery
    RsCombo.Requery
    RsPropiedad.Requery
    RsProductoPropiedad.Requery
        
    nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
    If RsDetalle.RecordCount = 0 Then
       nMonto = 0
       sProducto = ""
       wCombo = False
       nCombo = 0
       sItem = ""
    Else
       RsDetalle.MoveLast
       sItem = RsDetalle!tItem
       nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
    End If
    VisualizaMonto
End Sub

Public Sub SoloEliminaItem()
    'INSUMOCRITICO2013
    Dim rstItems As New ADODB.Recordset
    Set rstItems = New ADODB.Recordset
    Cn.Execute "DELETE FROM " & sInsumoCombo
    Cn.Execute "INSERT INTO " & sInsumoCombo & " select tcodigoinsumo,ncantidad from " & sDetalle & " inner join tproducto on " & sDetalle & ".tcodigoproducto=tproducto.tcodigoproducto where titem='" & sItem & "' and tproducto.lControlInsumoCritico=1 and isnull(tproducto.tcodigoinsumo,'')<>''  and isnull(" & sDetalle & ".limprime,0)=1"
    Cn.Execute "insert into " & sInsumoCombo & " SELECT     " & sDetalle & ".nCantidad * " & sComboDetalle & ".nCantidad AS nCantidad , dbo.TPRODUCTO.tCodigoInsumo FROM   " & sDetalle & " INNER JOIN  " & sComboDetalle & " ON " & sDetalle & ".tItem = " & sComboDetalle & ".tItem INNER JOIN dbo.TPRODUCTO ON " & sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto WHERE   " & sComboDetalle & ".tItem = '" & sItem & "' and  (dbo.TPRODUCTO.lControlInsumoCritico = 1) AND (ISNULL(dbo.TPRODUCTO.tCodigoInsumo, '') <> '') AND (ISNULL(" & sComboDetalle & ".lImprime, 0) = 1) "



    Set rstItems = Lib.OpenRecordset("select tCodigoInsumo,  SUM(ncantidad) as ncantidad from " & sInsumoCombo & "  group by tCodigoInsumo order by 2 ", Cn)
    
    If Not (rstItems.EOF Or rstItems.BOF) Then
        modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
    End If
    'INSUMOCRITICO
    

    Cn.Execute "delete from " & sDetalle & " where tItem = '" & sItem & "'"
    Cn.Execute "delete from " & sComboDetalle & " where tItem ='" & sItem & "'"
    Cn.Execute "delete from " & sComboPropiedad & " where tItem ='" & sItem & "'"
    Cn.Execute "delete from " & sProductoPropiedad & " where tItem='" & sItem & "'"
                              
    If RsDetalle.RecordCount <> 0 Then
       nMonto = nMonto - (grdDetalle.Columns(4).Text * nPVenta)
    Else
       nMonto = 0
    End If
    
    RsProductoPropiedad.Requery
    RsComboPropiedad.Requery
    RsCombo.Requery
    RsDetalle.Requery
    
    If RsDetalle.RecordCount = 0 Then
       txtMonto.Caption = "0.00"
       sProducto = ""
       wCombo = False
       nCombo = 0
       sItem = ""
    Else
       nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
       RsDetalle.MoveLast
       sItem = RsDetalle!tItem
    End If
    VisualizaMonto
End Sub

Public Sub VisualizaMonto()
    Dim vley1 As Double
    vley1 = Format(Calcular("select sum(ncantidad)*(" & sValorLey1 & ") as Codigo FROM [" & sDetalle & "] inner join vproducto on [" & sDetalle & "].tcodigoproducto=vproducto.codigo where vproducto.lLey1=1 and tEstadoItem = 'N'", Cn), "#,###,##0.00")
    'txtMonto.Caption = nMonto + vLey1
    lblLey1.Caption = sLey1 & " " & Format(vley1, "##0.00")
    
   txtMonto.Caption = Format(nMonto + vley1, "#,###,##0.00")
   txtMontoLetras.Caption = NumeroCadena(str(nMonto))
   If nPuerto > 0 Then
      If RsDetalle.RecordCount > 0 Then
         Visor "Total:" & sMonN & Right(String(10, " ") & Format(nMonto, "###,##0.00,"), 10), RsDetalle!nCantidad & " " & RsDetalle!Producto, nPuerto, "N"
      End If
   End If
   TxtDescuento.Caption = Format(Calcular("select sum(nDescuento*nCantidad) as Codigo FROM " & sDetalle, Cn), "#,###,##0.00")
    If lvisor Then
        Call InsertVisor8
    End If
End Sub

Public Sub ActualizaPedido()
      'Actualiza el Numero de Pedido en el Detalle Temporal
                  Dim oComando As clsComando
                  Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("spUpd_MPEDIDO", Cn) Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                        oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 7, sClienteFrecuente
                        oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, sTipoPedido
                        oComando.CreateParameter "@lPrioridad", adBoolean, adParamInput, 1, 1
                        oComando.CreateParameter "@tTipoAtencion", adVarChar, adParamInput, 2, "01"
                        oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, Right(sMozo, 4)
                        oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 4, sMotorizado
                        oComando.CreateParameter "@tObservacion", adVarChar, adParamInput, 250, txtObservacion.Caption
                        oComando.CreateParameter "@nTiempo", adInteger, adParamInput, 10, 0
                        oComando.CreateParameter "@tPuntoVenta", adVarChar, adParamInput, 2, ""
                        oComando.CreateParameter "@tHabitacion", adVarChar, adParamInput, 6, ""
                        oComando.CreateParameter "@tReserva", adVarChar, adParamInput, 6, ""
                        oComando.CreateParameter "@tPasajero", adVarChar, adParamInput, 50, ""
                        oComando.CreateParameter "@tCompania", adVarChar, adParamInput, 5, ""
                        oComando.CreateParameter "@tContacto", adVarChar, adParamInput, 4, ""
                        oComando.CreateParameter "@nDescuento", adDouble, adParamInput, 10, xDescuento
                        oComando.CreateParameter "@tDescuento", adVarChar, adParamInput, 3, sCodigoDescuento
                        oComando.CreateParameter "@tObservacionDescuento", adVarChar, adParamInput, 250, IIf(sCodigoDescuento = "000", sDescripcionDescuento, "")
                        oComando.CreateParameter "@tAutorizaDescuento", adVarChar, adParamInput, 15, IIf(sCodigoDescuento = "", "", tAutorizaDescuento)
                        oComando.CreateParameter "@tTienda", adVarChar, adParamInput, 3, ""
                        oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 20, IIf(txtFechaEntrega.Caption = "", Null, Format(txtFechaEntrega.Caption, "dd/MM/yyyy HH:mm"))
                        oComando.CreateParameter "@tCodigoInvitado", adVarChar, adParamInput, 10, sCodigoInvitado
                        oComando.CreateParameter "@tCodigopariente", adVarChar, adParamInput, 7, sCodigoParienteSeleccionado
                        oComando.CreateParameter "@tEntregarA", adVarChar, adParamInput, 20, IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20))
                        oComando.CreateParameter "@nTiempoAntesEnvio", adInteger, adParamInput, 10, 0
                        oComando.CreateParameter "@nMontoMaximo", adInteger, adParamInput, 250, 0
                        oComando.CreateParameter "@tPedido", adVarChar, adParamInput, 10, Pedido
                        
                        'origen de ventas
                        If vOrigenVentas = Null Or vOrigenVentas = "" Then
                            vOrigenVentas = "00"
                        End If
                        oComando.CreateParameter "@codigoOrigenVentas", adVarChar, adParamInput, 2, vOrigenVentas
                        oComando.CreateParameter "@tMesa", adVarChar, adParamInput, 3, sMesa
                        
                        'Pin Totem2
                        oComando.CreateParameter "@tTotemMesa", adVarChar, adParamInput, 10, ""
                        
                  If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  

      'Cn.Execute "Update MPEDIDO set fProgramacion = '" & IIf(Len(txtFechaEntrega.Caption) = 0, "", Left(Me.txtFechaEntrega.Caption, 16)) & "' , tclientedelivery='" & sClienteFrecuente & "',tcodigopariente='" & sCodigoParienteSeleccionado & "', tcodigoinvitado='" & sCodigoInvitado & "' , tObservacion='" & txtObservacion.Caption & "', tMozo='" & sMozo & "', tentregara='" & IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20)) & "' where tCodigoPedido='" & Pedido & "'"
      
      Cn.Execute "Update [" & sDetalle & "] Set tCodigoPedido = '" & Pedido & "'"
      'Cn.Execute "Update [" & sDetalle & "] Set lImprimeArea = 0 where lImprimeArea=1"
      
      
      'JSJR 09/01/2025
      Dim condicion As String
      
      condicion = ""
      If wAnulaDocumento = False Then
        condicion = " and lImprime=0"
      End If
      
      
      'Inserta el Detalle
      'Cn.Execute "delete DPEDIDO where tCodigoPedido = '" & Pedido & "' and lImprime=0"
      
      Cn.Execute "delete DPEDIDO where tCodigoPedido = '" & Pedido & "'" & condicion
      
      'Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                 "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                 "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,fregistro, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubalmacen,tCodigoEtiqueta,tunidadnegocio,fDiaContable, fEnvio, nEnvio, tCajaD ) " & _
                 "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                 "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                 "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,getdate(), nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tsubalmacen,tCodigoEtiqueta,'" & sUnidadNegocio & "' ,'" & Format(obtieneDiaContable, "yyyyMMdd") & "', fEnvio, nEnvio, '" & sCaja & "' " & _
                 "From [" & sDetalle & "]  where tEstadoItem='N' and lImprime=0 "
      
       Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                 "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                 "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,fregistro, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubalmacen,tCodigoEtiqueta,tunidadnegocio,fDiaContable, fEnvio, nEnvio, tCajaD ) " & _
                 "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                 "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                 "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,getdate(), nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tsubalmacen,tCodigoEtiqueta,'" & sUnidadNegocio & "' ,'" & Format(obtieneDiaContable, "yyyyMMdd") & "', fEnvio, nEnvio, '" & sCaja & "' " & _
                 "From [" & sDetalle & "]  where tEstadoItem='N'" & condicion
                 
      
      'Actualiza el Numero de Pedido en el Detalle Combos
      Cn.Execute "Update [" & sComboDetalle & "] Set tCodigoPedido = '" & Pedido & "'"
      
      'ACTUALIZA  tAplicaExoneracion, tAplicaInafectacion
      
        Set oComando = New clsComando
        If Not oComando.CreateCmdSp("spUpdate_DPEDIDO_Ina", Cn) Then
           Set oComando = Nothing
           Exit Sub
        End If
        
        oComando.CreateParameter "@pedido", adVarChar, adParamInput, 10, Pedido
        
        If Not oComando.GetParamOK Then
           Set oComando = Nothing
           Exit Sub
        End If
    
        If Not oComando.ExecSP Then
           Set oComando = Nothing
           Exit Sub
         End If

      
      
      'Inserta Combo
      Cn.Execute "delete CPEDIDO where tCodigoPedido = '" & Pedido & "'"
      Cn.Execute "Insert into CPEDIDO select * from " & sComboDetalle
      
      'Inserta las propiedades
      Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido = '" & Pedido & "'"
      Cn.Execute "Insert into TPRODUCTOPROPIEDAD select '" & Pedido & "', tItem,  tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra, 1,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario from " & sProductoPropiedad
      
      'Inserta las propiedades de los Combos
      Cn.Execute "delete tcombopropiedad where tCodigoPedido = '" & Pedido & "'"
      Cn.Execute "Insert into TCOMBOPROPIEDAD select '" & Pedido & "', tItem, tItemCombo, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra, 1,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario from " & sComboPropiedad
End Sub

Private Sub cmdPunto_Click(Index As Integer)
   Dim i As Integer
   For i = 1 To 9
       cmdPunto(i).FontBold = False
   Next i
   cmdPunto(Index).FontBold = True
   rsPuntoVenta.MoveFirst
   rsPuntoVenta.Find "Descripcion = '" & cmdPunto(Index).Caption & "'"
   'cmdCabecera(0).Caption = rsPuntoVenta!Descripcion
   sPuntoVenta = rsPuntoVenta!codigo
   tabProducto.Visible = True
   fraPuntoVenta.Visible = False
End Sub

Public Sub AsignaComboPropiedad()
    Dim i As Integer
    If RsOperador.RecordCount > 0 Then
        If RsOperador.EOF Then
            RsOperador.MoveFirst
        End If
       RsPropiedad.Filter = "tOperador = '" & RsOperador!codigo & "' and tProducto='" & sCombo & "'"
    Else
       RsPropiedad.Filter = "tOperador = '  ' and tProducto='" & sCombo & "'"
    End If
    AsignaComando 19, RsPropiedad, cmdPropiedad()
    ' Validaciòn de visibilidad para Propiedades -BR
    With RsPropiedad
        If .RecordCount > 0 Then
            .MoveFirst
            For i = 1 To IIf(.RecordCount > 19, 19, .RecordCount)
                If Not IsNull(!visibilidad) And !visibilidad Then
                    Select Case sTipoPedido
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
                    cmdPropiedad(i).Visible = True
                End If
                cmdPropiedad(i).Caption = !descripcion
                .MoveNext
            Next i
        End If
    End With
    
    For i = 1 To 19
        cmdPropiedad(i).FontBold = False
    Next i
    
    lblResumen.Text = ""
    RsComboPropiedad.Filter = "tItem='" & sItem & "' and tItemCombo='" & xItem & "'"
    If Not RsComboPropiedad.EOF Then
       RsComboPropiedad.MoveFirst
       Do While Not RsComboPropiedad.EOF
          For i = 1 To 19
              If cmdPropiedad(i).Caption = RsComboPropiedad!descripcion And RsOperador!descripcion = RsComboPropiedad!Operador Then
                 cmdPropiedad(i).FontBold = True
                 Exit For
              End If
          Next i
          
          If RsComboPropiedad!nCantidad = 1 Then
            lblResumen.Text = lblResumen.Text & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!descripcion) & ", "
          Else
            lblResumen.Text = lblResumen.Text & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!descripcion) & ": (" & RsComboPropiedad!nCantidad & "), "
          End If
          
          'lblResumen.Text = lblResumen.Text & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!Descripcion) & ", "
          RsComboPropiedad.MoveNext
       Loop
    End If
End Sub

Public Function CambiaPrecio(Valor As Double)
    nPVenta = val(Valor)
    nOficial = nPVenta
    Dim Acumulado As Double
    Select Case pais 'ok
        Case "001" 'Bolivia
                    Acumulado = 0
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = (Acumulado / 100)
                    
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta * nPorcentaje3 / 100, 0)
                    
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        Case Else 'Peru, Ecuador
                    Acumulado = 0
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = 1 + (Acumulado / 100)
                    
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                    
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        
    End Select
    
    Isql = "Update " & sDetalle & " Set nPrecioNeto = " & nPBase & ", " & _
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
           "where tItem = '" & sItem & "'"
           Cn.Execute Isql
    CambiaPrecio = Calcular("select sum(nVenta) as Codigo FROM " & sDetalle, Cn)
    txtMonto.Caption = Format(CambiaPrecio, "#,###,##0.00")
    nMonto = CambiaPrecio
End Function



Public Function CambiaPrecioCombo(Valor As Double)
    nPVenta = val(Valor)
    'nOficial = nPVenta
    Dim Acumulado As Double
    Select Case pais 'ok
        Case "001" 'Bolivia
                    Acumulado = 0
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = (Acumulado / 100)
                    
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta * nPorcentaje3 / 100, 0)
                    
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        Case Else 'Peru, Ecuador
                    Acumulado = 0
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = 1 + (Acumulado / 100)
                    
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                    
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        
    End Select
    '"nPrecioOficial = " & nOficial & ", " &
    
    Isql = "Update " & sDetalle & " Set nPrecioNeto = " & nPBase & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & " " & _
           "where tItem = '" & sItem & "'"
           Cn.Execute Isql
           
    CambiaPrecioCombo = Calcular("select sum(nVenta) as Codigo FROM " & sDetalle, Cn)
    txtMonto.Caption = Format(CambiaPrecioCombo, "#,###,##0.00")
    nMonto = CambiaPrecioCombo
End Function
Public Sub EliminaCabecera()
    'KDS2
    If (lKDS = True) Then
        Dim kdsRsCabecera As Recordset
        Isql = "SELECT * From vPedidoCabecera Where Codigo = '" & Pedido & "' Order By codigo "
        Set kdsRsCabecera = Lib.OpenRecordset(Isql, Cn)
        Call KDS_EliminarOrden(kdsRsCabecera)
    End If
    
   Dim i As Integer
   Screen.MousePointer = vbHourglass
   Dim sMotivo As String
   
   If lPrinter Then
      sMotivo = Calcular("select Descripcion as Codigo from vMotivoEliminacion where Codigo='" & sCodigo & "'", Cn)
      Isql = "select *, Descripcion as MotivoEliminacion FROM dbo.vPedido LEFT OUTER JOIN dbo.vMotivoEliminacion ON dbo.vPedido.tMotivoEliminacion = dbo.vMotivoEliminacion.Codigo " & _
             "WHERE vPedido.Codigo='" & sPedido & "' AND lImprime=1 AND lImprimeArea=1 " & _
             "ORDER BY tItem,tetiqueta,combo"
                     
       Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
       
       If Not RsImpresion.EOF Then
          RsArea.MoveFirst
          For i = 1 To RsArea.RecordCount
              RsImpresion.Filter = "tArea = '" & RsArea!tArea & "'"
              If RsArea!tIcono = "" Or sSalon = RsArea!tIcono Or (sSalon = "" And RsArea!nValor = 1) Then
                 If RsImpresion.RecordCount <> 0 Then
                    RsImpresion.MoveFirst
                    'sPedido = Pedido
                    ImprimePedido RsImpresion, "A", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsProductoPropiedad, "Rapido"
                    'sPedido = ""
                 End If
              End If
              RsArea.MoveNext
          Next i
       End If
   End If
   
   If lInfhotel Then
      CnInfhotel.Execute "update MCOMANDA set TESTADO ='04', TOBSERVACIONANULA = 'Anulado por Inforest - " & sUsuarioAutoriza & " " & Pedido & " - " & Trim(sDescrip) & "' where tComanda ='" & sComandaInfhotel & "'"
   End If

   'INSUMOCRITICO
   Cn.Execute "delete from " & sInsumoCombo
   
    Dim rstItems As New ADODB.Recordset
    Dim j As Integer
    Set rstItems = New ADODB.Recordset
    Cn.Execute "insert into " & sInsumoCombo & " select tcodigoinsumo,ncantidad from " & sDetalle & " inner join tproducto on " & sDetalle & ".tcodigoproducto=tproducto.tcodigoproducto where  tproducto.lControlInsumoCritico=1 and isnull(tproducto.tcodigoinsumo,'')<>''  and isnull(" & sDetalle & ".limprime,0)=1 "
    Cn.Execute "insert into " & sInsumoCombo & " SELECT " & sDetalle & ".nCantidad * " & sComboDetalle & ".nCantidad AS nCantidad , dbo.TPRODUCTO.tCodigoInsumo FROM " & sDetalle & " INNER JOIN " & sComboDetalle & "  on " & sDetalle & ".tItem = " & sComboDetalle & ".tItem INNER JOIN dbo.TPRODUCTO ON " & sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto WHERE     (dbo.TPRODUCTO.lControlInsumoCritico = 1) AND (ISNULL(dbo.TPRODUCTO.tCodigoInsumo, '') <> '') AND (ISNULL(" & sComboDetalle & ".lImprime, 0) = 1)"
    
    Set rstItems = Lib.OpenRecordset("select tCodigoInsumo,  SUM(ncantidad) as ncantidad from " & sInsumoCombo & "  group by tCodigoInsumo order by 2", Cn)
    If Not (rstItems.EOF Or rstItems.BOF) Then
        rstItems.MoveFirst
        For j = 0 To rstItems.RecordCount - 1
            modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
            rstItems.MoveNext
        Next j
    End If
   'INSUMOCRITIC

   Cn.Execute "delete from " & sDetalle
   Cn.Execute "delete from " & sComboDetalle
   Cn.Execute "delete from " & sProductoPropiedad
   Cn.Execute "delete from " & sComboPropiedad
   Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='" & sCodigo & "', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='" & sDescrip & "'  where tCodigoPedido ='" & sPedido & "'"
   Cn.Execute "Update DPEDIDO Set tEstadoItem = 'A' where tCodigoPedido = '" & sPedido & "'"
   Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido='" & sPedido & "'"
   Cn.Execute "delete CPEDIDO where tCodigoPedido='" & sPedido & "'"
   Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido='" & sPedido & "'"
                      
   RsDetalle.Requery
   RsCombo.Requery
   RsComboPropiedad.Requery
   RsProductoPropiedad.Requery
 
   nMonto = 0
   Pedido = ""
   sProducto = ""
   wCombo = False
   nCombo = 0
   sItem = ""
   VisualizaMonto
   Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaProductoCombo()
    Dim i As Integer
    RsProductoCombo.Filter = "tCombo = '" & sProducto & "'"
    AsignaComandoColor 52, RsProductoCombo, cmdProductoCombo()
    'AsignaComando 52, RsProductoCombo, cmdProductoCombo()
End Sub

Sub PUltimaComanda()
    Dim RsPuntoVentaU   As ADODB.Recordset
    Isql = "Select tPuntoVenta as Codigo, tDescripcion as Descripcion, nUltimoComanda, tmoneda" & _
           " From tPuntoVenta " & _
           " where tHotel='" & sHotel & "' AND lActivo=1 and lInforest=1 and tPuntoVenta='" & sPuntoVenta & "'"
    Set RsPuntoVentaU = Lib.OpenRecordset(Isql, CnInfhotel)
    UltimaComanda = IIf(IsNull(RsPuntoVentaU!nUltimoComanda), "", RsPuntoVentaU!nUltimoComanda)
    Set RsPuntoVentaU = Nothing
End Sub

Private Sub ListarOperadoresConFiltro(ByVal tProducto As String)
    Isql = "select * from vOperador where lActivo = 1 " & _
           "AND ((select Count(tCodigoPropiedad) " & _
           "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
           "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 " & _
           "And TPROPIEDAD.tOperador = vOperador.Codigo and tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) " & _
           "order by Codigo"
    Set RsOperador = Lib.OpenRecordset(Isql, Cn)
    AsignaBoton 13, RsOperador, cmdOperador()
    If RsOperador.RecordCount > 0 Then
        RsOperador.MoveFirst
        Dim i As Integer
        
        For i = 1 To RsOperador.RecordCount
            If RsOperador!nBoton <> 0 Then
                cmdOperador(RsOperador!nBoton).backColor = &HD1F2CA
            End If
            RsOperador.MoveNext
        Next i
        RsOperador.MoveFirst
        xOperador = RsOperador!codigo

        If RsOperador!nBoton <> 0 Then
            cmdOperador(RsOperador!nBoton).backColor = &H8BDD79
        End If
    End If
End Sub

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
                cmdOperador(i).backColor = &HD1F2CA
            Next i
        End If
        RsOperador.MoveFirst
        While RsOperador.EOF = False
            If RsOperador!lObligaPropiedad = True Then
                If wAgregaCombo Then
                    Isql = "Select * From " & sComboPropiedad & " Where tCodigoPropiedad In (Select tCodigoPropiedad from TPropiedad Where tOperador = '" & RsOperador!codigo & "' And tProducto = '" & tProducto & "' ) And tItem = '" & sItem & "' And tItemCombo = '" & xItem & "'"
                Else
                    Isql = "Select * From " & sProductoPropiedad & " Where tCodigoPropiedad In (Select tCodigoPropiedad from TPropiedad Where tOperador = '" & RsOperador!codigo & "' And tProducto = '" & tProducto & "' ) And tItem = '" & sItem & "' "
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
            cmdOperador(i).backColor = &HD1F2CA
        Next i
        
        'GCAA 030302021
        If RsOperador!nBoton > 0 Then
           cmdOperador(RsOperador!nBoton).backColor = &H8BDD79
        End If
     
        
        'AsignaPropiedad
    End If
End Function

'OO
Private Function ExistenPropiedadesPendientesEnPedido() As Boolean
    Dim oRsDPedidoNoImp As Recordset 'Lista de Productos no impresos
    Set oRsDPedidoNoImp = Obtener_ProductosNoImpresosPorPedido()
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
            Set oRsPropiedadesDeOperador = Obtener_PropiedadesSeleccionadasPorProducto(oRsDPedidoNoImp!tItem, oRsDPedidoNoImp!tCodigoProducto, oRsOperadoresObligados!codigo)
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

'OO
Private Function Obtener_ProductosNoImpresosPorPedido() As Recordset
    If wAgregaCombo Then
        Isql = "Select [" & sDetalle & "].*,TProducto.tDetallado from [" & sDetalle & "] Left Join TProducto On [" & sDetalle & "].tProductoCombo = TProducto.tCodigoProducto " & _
               "Where [" & sDetalle & "].lImprime = '0' And tItem = '" & sItem & "' "
    Else
        Isql = "Select [" & sDetalle & "].*,TProducto.tDetallado from [" & sDetalle & "] Left Join TProducto On [" & sDetalle & "].tCodigoProducto = TProducto.tCodigoProducto Where [" & sDetalle & "].lImprime = '0'"
    End If
    Set Obtener_ProductosNoImpresosPorPedido = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_OperadoresObligatoriosPorProducto(ByVal tProducto) As Recordset
    Isql = "select * from vOperador where lActivo = 1 AND lObligaPropiedad = 1" & _
       "AND ((select Count(tCodigoPropiedad) " & _
       "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
       "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 " & _
       "And TPROPIEDAD.tOperador = vOperador.Codigo and tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) " & _
       "order by Codigo"
    Set Obtener_OperadoresObligatoriosPorProducto = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_PropiedadesSeleccionadasPorProducto(ByVal tItem, ByVal tProducto, ByVal tOperador) As Recordset
    If wAgregaCombo Then
        Isql = "Select * From TComboPropiedad " & _
            "Where tItem = '" & tItem & "' And tCodigoPropiedad In ( " & _
            "Select tCodigoPropiedad from TPropiedad " & _
            "Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    Else
        Isql = "Select * From " & sProductoPropiedad & " " & _
            "Where tItem = '" & tItem & "' And tCodigoPropiedad In ( " & _
            "Select tCodigoPropiedad from TPropiedad " & _
            "Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    End If
    Set Obtener_PropiedadesSeleccionadasPorProducto = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function ExistenPropiedadesPendientesEnCombos() As Boolean
    Dim oRsCombosPedido As Recordset
    Dim oi, oj, ok As Integer
    Dim oMensaje As String
    oMensaje = "Combo(s) con productos con propiedades obligatorias: "
    Set oRsCombosPedido = Obtener_TodosLosCombosDelPedido()
    Dim oflag As Boolean
    oflag = True
    For oi = 1 To oRsCombosPedido.RecordCount
        Dim oRsProductoNoImpCombo As Recordset
        Set oRsProductoNoImpCombo = Obtener_LosProductosNoImpDelCombo(oRsCombosPedido!tItem)
        For oj = 1 To oRsProductoNoImpCombo.RecordCount
            Dim oRsOperadores As Recordset
            Set oRsOperadores = Obtener_OperadoresObligatoriosDeUnProducto(oRsProductoNoImpCombo!tProductoCombo)
            For ok = 1 To oRsOperadores.RecordCount
                Dim oRsPropiedadesDeOperador As Recordset
                Set oRsPropiedadesDeOperador = Obtener_PropiedadesSeleccionadasPorProductoDeCombo(oRsCombosPedido!tItem, oRsProductoNoImpCombo!tItemCombo, oRsProductoNoImpCombo!tProductoCombo, oRsOperadores!codigo)
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

'OO
Private Function Obtener_TodosLosCombosDelPedido() As Recordset
    Isql = "Select  [" & sDetalle & "].*,TProducto.tDetallado From [" & sDetalle & "] Left Join TProducto On [" & sDetalle & "].tCodigoProducto = TProducto.tCodigoProducto " & _
           "Where   [" & sDetalle & "].lCombinacion = '1'"
    Set Obtener_TodosLosCombosDelPedido = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_LosProductosNoImpDelCombo(ByVal tItem As String) As Recordset
    Isql = "Select * from [" & sComboDetalle & "] Where tItem = '" & tItem & "' And lImprime = '0'"
    Set Obtener_LosProductosNoImpDelCombo = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_PropiedadesSeleccionadasPorProductoDeCombo(ByVal tItem As String, ByVal tItemCombo As String, ByVal tProducto As String, ByVal tOperador As String) As Recordset
    Isql = "Select * From " & sComboPropiedad & " " & _
           "Where tItem = '" & tItem & "' And tItemCombo = '" & tItemCombo & "' And tCodigoPropiedad  In ( " & _
           "Select tCodigoPropiedad from TPropiedad Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    Set Obtener_PropiedadesSeleccionadasPorProductoDeCombo = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_OperadoresObligatoriosDeUnProducto(ByVal tProducto As String) As Recordset
    Isql = "select * from vOperador where lActivo = 1 AND lObligaPropiedad = 1 AND ( " & _
           "(select Count(tCodigoPropiedad) FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
           "Where TPROPIEDAD.lActivo = 1 And IsNull(tOperador.lStockMenos, 0) <> 1 And TPROPIEDAD.tOperador = vOperador.Codigo " & _
           "And tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) order by Codigo"
    Set Obtener_OperadoresObligatoriosDeUnProducto = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_ProductoDeCombo(ByVal tCombo As String, ByVal tCodigoProducto As String) As Recordset
    Isql = "Select lUnico,tEtiqueta From TCOMBO Where tCombo = '" & tCombo & "' And tCodigoProducto = '" & tCodigoProducto & "'"
    Set Obtener_ProductoDeCombo = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function ObtenerSumaCantidadesEnElCombo(ByVal tItem As String, ByVal tEtiqueta As String) As Double
    Isql = "Select ISNULL(Sum(nCantidad),0) as nCantidad from [" & sComboDetalle & "] " & _
           "Where   tItem = '" & tItem & "' And tProductoCombo IN ( " & _
           "Select tCodigoProducto From TCOMBO Where tEtiqueta = '" & tEtiqueta & "' And lUnico = '1') "
    Dim Suma As Double
    Dim oRsResultado As Recordset
    Set oRsResultado = Lib.OpenRecordset(Isql, Cn)
    Suma = oRsResultado!nCantidad
    ObtenerSumaCantidadesEnElCombo = Suma
End Function

'OO
Private Function ObtenerDetalleProducto(ByVal tItem As String) As Recordset
    'Isql = "Select D.* ,P.nPrecioVenta as 'PrecioProducto' from [" & sDetalle & "] As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tItem ='" & tItem & "'"
    'Isql = "Select D.* ,CASE (D.tTipoPedido)  WHEN '01' THEN P.nprecioventa when '02' then p.npreciodelivery when '03' then p.npreciollevar when '04' then p.npreciocanal4 when '05' then p.npreciocanal5 END as 'PrecioProducto' from [" & sDetalle & "] As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tItem ='" & tItem & "'"
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
           "            when '05' then p.npreciocanal5 End) end as 'PrecioProducto',P.tMoneda as 'tMonedaProducto' from [" & sDetalle & "] As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tItem ='" & tItem & "'"
'    Isql = " Select D.* ,d.nprecioventa   as 'PrecioProducto',P.tMoneda as 'tMonedaProducto' from [" & sDetalle & "] As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tItem ='" & tItem & "'"

    Set ObtenerDetalleProducto = Lib.OpenRecordset(Isql, Cn)
End Function
'OO
Private Function Obtener_CantidadMaximaDeUnicoEtiqueta(ByVal tItem As String, cantidadActual As Double) As Double
    
    Isql = "Select Sum(P.nCantidad) as Cantidad " & _
           "from [" & sComboDetalle & "] as P Left Join TCOMBO as C ON P.tProducto = C.tCombo AND   P.tProductoCombo = C.tCodigoProducto " & _
           "where P.tItem ='" & tItem & "' And C.lUnico = '1' " & _
           "Group By C.tEtiqueta,c.tcodigoproducto order by 1 desc"
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

Private Function ObtenerCodigoMozo(ByVal tResumido As String) As String
    Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by nBoton"
    Set RsMozo = Lib.OpenRecordset(Isql, Cn)
    RsMozo.Filter = "tResumido = '" & tResumido & "'"
    If RsMozo.RecordCount = 0 Then
       txtTitulo.Caption = " Caja Rápida : " & sCaja & Mesero & ": Sin " & Mesero & IIf(sPedido = "", "", " - " & sPedido)
       ObtenerCodigoMozo = "0000"
    Else
       txtTitulo.Caption = " Caja Rápida : " & sCaja & Mesero & ": " & tResumido & IIf(sPedido = "", "", " - " & sPedido)
       ObtenerCodigoMozo = RsMozo!codigo
    End If
End Function

Private Function CalculaDescuento() As Boolean
    Dim sCriterio As String
    Dim lAcumulable As Boolean
    Dim nOferta As Double
    Dim nSuma As Double
    
    nSuma = Calcular("SELECT sum(nPrecioOficial*nCantidad) as Codigo FROM " & sDetalle & " LEFT OUTER JOIN dbo.TPRODUCTO ON " & sDetalle & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where lDescuento=1", Cn)

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
             Isql = "Update " & sDetalle & " Set nPrecioNeto = " & nPBase & ", " & _
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
                    "where tItem = '" & RsDetalle!tItem & "' "
                    Cn.Execute Isql
          End If
       RsDetalle.MoveNext
       Loop
    End If

End Function

Public Function verificaCantidadDeItemsCombos(ByVal tItem As String, ByVal numeroDeItemCombos As Double, ByVal ncantidadNueva As Double) As Boolean
    Dim cantidadMaximoNueva As Double
    Dim cantidadMaximoPosible As Double
    Dim X As Integer
    Dim oRstRecorriendoCombo As New Recordset
    Dim oRstCantidadNUnicos As New Recordset
    
    verificaCantidadDeItemsCombos = False
    cantidadMaximoNueva = numeroDeItemCombos * ncantidadNueva
    cantidadMaximoPosible = 0
    Set oRstRecorriendoCombo = Lib.OpenRecordset("select tcombo.ncantidad ,cpedido.tproductocombo from [" & sComboDetalle & "]  AS cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where  cpedido.titem='" & tItem & "' and lfijo=1", Cn)
    If Not (oRstRecorriendoCombo.EOF Or oRstRecorriendoCombo.BOF) Then
        oRstRecorriendoCombo.MoveFirst
        For X = 0 To oRstRecorriendoCombo.RecordCount - 1
            cantidadMaximoPosible = cantidadMaximoPosible + (oRstRecorriendoCombo!nCantidad * ncantidadNueva)
            oRstRecorriendoCombo.MoveNext
        Next X
    End If
    
    Set oRstCantidadNUnicos = Lib.OpenRecordset("select isnull(sum(cpedido.ncantidad),0) from [" & sComboDetalle & "]  AS cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where   titem='" & tItem & "' and tcombo.lfijo=0 ", Cn)
    If Not (oRstCantidadNUnicos.EOF Or oRstCantidadNUnicos.BOF) Then
        cantidadMaximoPosible = cantidadMaximoPosible + oRstCantidadNUnicos.Fields(0)
    End If
    
    If cantidadMaximoNueva >= cantidadMaximoPosible Then
        verificaCantidadDeItemsCombos = True
    End If
    
End Function


'insumo critico23
Public Function obtieneProductos(tcodigoinsumo As String) As String
 obtieneProductos = ""
 Dim rstProductos As New ADODB.Recordset
 Dim k As Integer
 Set rstProductos = Lib.OpenRecordset("select isnull(tdetallado,'') as producto FROM         dbo.TPRODUCTO INNER JOIN  " & sDetalle & " ON dbo.TPRODUCTO.tCodigoProducto = " & sDetalle & ".tCodigoProducto INNER JOIN  dbo.TINSUMO ON dbo.TPRODUCTO.tcodigoInsumo = dbo.TINSUMO.tcodigo where   tproducto.tcodigoinsumo='" & tcodigoinsumo & "' and lcontrolinsumocritico=1 and isnull(limprime,0)=0 AND   (tinsumo.lactivo=1) group by isnull(tdetallado,'') ", Cn)
 If Not (rstProductos.EOF Or rstProductos.BOF) Then
        rstProductos.MoveFirst
        For k = 0 To rstProductos.RecordCount - 1
                    obtieneProductos = IIf(Len(obtieneProductos) = 0, rstProductos.Fields(0), obtieneProductos & " / " & rstProductos.Fields(0))
            rstProductos.MoveNext
        Next k
 End If
 
 
End Function
'insumo critico23
 
'luchiinsumo
Public Sub verificatitulo()
        'INSUMOCRITICO23
        Dim rsInsumo As New ADODB.Recordset
        If Calcular("select isnull(lControlInsumoCritico,0) as codigo from tproducto  INNER JOIN " & sDetalle & "  on tproducto.tcodigoproducto=" & sDetalle & ".tcodigoproducto where  titem='" & sItem & "'", Cn) = True Then
                        Set rsInsumo = Lib.OpenRecordset("select isnull(tcodigoinsumo,'') tcodigoinsumo , isnull(tinsumo.descripcion,'') ,isnull(nstock,0) , " & sDetalle & ".ncantidad from tproducto inner join tinsumo on tproducto.tcodigoinsumo =tinsumo.tcodigo inner join " & sDetalle & "  on tproducto.tcodigoproducto=" & sDetalle & ".tcodigoproducto  where     titem='" & sItem & "' and tinsumo.lactivo=1", Cn)
                        If Not (rsInsumo.EOF Or rsInsumo.BOF) Then
                                Label2.Caption = "   Insumo Crítico ->   " & rsInsumo.Fields(1) & " =  Stock: " & str(rsInsumo.Fields(2)) & "      Solicitado: " + str(rsInsumo.Fields(3))
                                Label2.Visible = True
                        End If
                Else
                        Label2.Caption = muestra
                        Label2.Visible = False
        End If
        'INSUMOCRITICO
    If lvisor Then
        Call InsertVisor8
    End If
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
                fImpresionDiaContable = obtieneDiaContable
End Function
'diaContable
Private Sub CalculaAplicaTope(nTope As Double)
    Dim sCriterio As String
    Dim lAcumulable As Boolean
    Dim nOferta As Double
    Dim nSuma As Double
    
    nSuma = Calcular("SELECT sum(nPrecioOficial*nCantidad) as Codigo FROM " & sDetalle & " LEFT OUTER JOIN dbo.TPRODUCTO ON " & sDetalle & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where lDescuento=1", Cn)

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
         Isql = "Update " & sDetalle & " Set nPrecioNeto = " & nPBase & ", " & _
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
                "where tItem = '" & RsDetalle!tItem & "'"
                Cn.Execute Isql
      End If
   RsDetalle.MoveNext
   Loop
End If

End Sub
Private Sub InsertVisor8()
On Error GoTo fin
    If lvisor And sCaja <> "" Then
        Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
        Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado) values(1,'" & sCaja & "','" & sCaja & "',1)"
        Cn.Execute "delete from Visor_Dpedido where tcajad='" & sCaja & "'"
         Cn.Execute "insert into visor_dpedido " & _
           "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
           "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
           "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem,tsubalmacen,toferta,tCajaD) " & _
            " Select '" & sCaja & "'+tItem, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
           "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
           "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem,tsubalmacen,toferta,'" & sCaja & "' " & _
           " From [" & sDetalle & "]"
    End If
fin:
    
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
               Calcular("select isnull(sum(nCantidad),0) codigo FROM " & sComboDetalle & " WHERE " & _
               " tproducto='" & Trim(.Fields(3)) & "' and titem='" & RsDetalle!tItem & "'", Cn)

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
End Function

Private Sub txtFechaEntrega_Click()
    cmdCabecera_Click 6
End Sub


