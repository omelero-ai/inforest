VERSION 5.00
Begin VB.Form frmJuntarMesas 
   BorderStyle     =   0  'None
   Caption         =   "Juntar mesas"
   ClientHeight    =   15270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15270
   ScaleWidth      =   16755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrActualiza 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12000
      Top             =   120
   End
   Begin VB.Frame fraFooter 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   72
      Top             =   8760
      Width           =   11655
      Begin VB.PictureBox picCancelar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         Picture         =   "frmJuntarMesas.frx":0000
         ScaleHeight     =   660
         ScaleWidth      =   2220
         TabIndex        =   75
         Top             =   120
         Width           =   2220
      End
      Begin VB.PictureBox picActualizar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   6840
         Picture         =   "frmJuntarMesas.frx":14BA
         ScaleHeight     =   660
         ScaleWidth      =   2220
         TabIndex        =   74
         Top             =   120
         Width           =   2220
      End
      Begin VB.PictureBox picGuardar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   9240
         Picture         =   "frmJuntarMesas.frx":2B39
         ScaleHeight     =   660
         ScaleWidth      =   2220
         TabIndex        =   73
         Top             =   120
         Width           =   2220
      End
   End
   Begin VB.Frame fraSalon 
      BackColor       =   &H00F7F6F5&
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   11655
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   1
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   68
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   2
         Left            =   1560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   67
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   3
         Left            =   3000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   66
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   4
         Left            =   4440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   65
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   5
         Left            =   5880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   64
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   6
         Left            =   7320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   63
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   7
         Left            =   8760
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   62
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   8
         Left            =   10200
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   61
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   9
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   60
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   10
         Left            =   1560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   59
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   11
         Left            =   3000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   58
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   12
         Left            =   4440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   57
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   13
         Left            =   5880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   56
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   14
         Left            =   7320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   55
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   15
         Left            =   8760
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   54
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   16
         Left            =   10200
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   53
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   17
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   52
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   18
         Left            =   1560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   51
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   19
         Left            =   3000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   50
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   20
         Left            =   4440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   49
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   21
         Left            =   5880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   48
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   22
         Left            =   7320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   47
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   23
         Left            =   8760
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   46
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   24
         Left            =   10200
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   45
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   25
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   44
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   26
         Left            =   1560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   43
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   27
         Left            =   3000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   42
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   28
         Left            =   4440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   41
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   29
         Left            =   5880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   40
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   30
         Left            =   7320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   39
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   31
         Left            =   8760
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   38
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   32
         Left            =   10200
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   37
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   33
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   36
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   34
         Left            =   1560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   35
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   35
         Left            =   3000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   34
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   36
         Left            =   4440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   33
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   37
         Left            =   5880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   32
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   38
         Left            =   7320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   31
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   39
         Left            =   8760
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   30
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   40
         Left            =   10200
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   29
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   41
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   28
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   42
         Left            =   1560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   27
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   43
         Left            =   3000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   26
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   44
         Left            =   4440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   25
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   45
         Left            =   5880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   24
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   46
         Left            =   7320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   23
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   47
         Left            =   8760
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   22
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   48
         Left            =   10200
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   21
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   49
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   20
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   50
         Left            =   1560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   19
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   51
         Left            =   3000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   18
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   52
         Left            =   4440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   17
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   53
         Left            =   5880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   16
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   54
         Left            =   7320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   15
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   55
         Left            =   8760
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   14
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   56
         Left            =   10200
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   13
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   57
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   12
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   58
         Left            =   1560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   11
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   59
         Left            =   3000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   10
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   60
         Left            =   4440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   9
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   61
         Left            =   5880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   8
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   62
         Left            =   7320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   7
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   63
         Left            =   8760
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   6
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   64
         Left            =   10200
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   5
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox Picture67 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   960
         Picture         =   "frmJuntarMesas.frx":4525
         ScaleHeight     =   270
         ScaleWidth      =   9225
         TabIndex        =   4
         Top             =   6720
         Width           =   9225
      End
      Begin VB.Label lblMesas 
         BackColor       =   &H00F7F6F5&
         Caption         =   "Salón 0 tiene 0/0 mesas libres"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   69
         Top             =   120
         Width           =   10935
      End
   End
   Begin VB.Frame fraHead 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   11655
      Begin VB.Frame fraNumMesasJuntar 
         BackColor       =   &H00D8D3CF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   9600
         TabIndex        =   78
         Top             =   600
         Width           =   975
         Begin VB.TextBox txtNumMesasJuntar 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   79
            Text            =   "0"
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame fraMesasJuntar 
         BackColor       =   &H00D8D3CF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2040
         TabIndex        =   71
         Top             =   600
         Width           =   3615
         Begin VB.TextBox txtMesasJuntar 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   76
            Top             =   0
            Width           =   3375
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Número de mesas juntas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   77
         Top             =   600
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009E9188&
         X1              =   0
         X2              =   11640
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mesas a juntar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblSalon 
         BackColor       =   &H80000004&
         Caption         =   "Salon 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   5655
      End
   End
   Begin VB.Label txtTitulo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mesas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmJuntarMesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim RsMesa As Recordset
Dim RsSalon As Recordset
Dim RsPedido As Recordset
Dim RsGrilla As Recordset
Dim RsMesaSeleccionada As Recordset
Dim RsPedidoMesaTabla As Recordset

Dim sTitulo As String
Dim sMesaActual As String
Dim nIndex As Integer
Dim xIndex As Integer
Dim xSalon As String
Dim sMesa() As String
Dim Estado As String
Dim wClick As Boolean
Dim sSinMesa As String
Dim tcolorSeleccion As String
Dim tcolorDeseleccion As String
Dim tcolorMesaOriginal As String

Dim tmesaOriginal As String
Dim tnumAdulto As String
Dim tnumNinio As String

Dim sSalonActual As String
'Dim miArrayMesas() As String

Public Sub parametrosIni(ByVal pMesaOriginal As String, ByVal pnumAdulto As String, ByVal pnumNinio As String)
'tcolorMesaOriginal = tmesaOriginal
 tmesaOriginal = pMesaOriginal
 tnumAdulto = pnumAdulto
 tnumNinio = pnumNinio
End Sub
Private Sub Form_Load()
    Call OrganizaComponentes
    Centrar Me
    
    tcolorSeleccion = "#40160d"
    tcolorDeseleccion = "#CFD3D8"
    tcolorMesaOriginal = "#76E4E6"
    
    Dim i As Integer
    wClick = False
    sCodigo = ""
    sDescrip = ""
    sTemp = ""
    xIndex = 0
    wMesa = False
    wCambioMesa = False
    
    Set RsPedidoMesaTabla = Lib.OpenRecordset("select * from TPEDIDOMESA", Cn)
    
    Set RsSalon = Lib.OpenRecordset("select tCodigo as tSalon, tDetallado as dSalon, tiCono from TTABLA where tTABLA ='Salon' and lActivo = 1 ", Cn)
    RsSalon.Filter = "tSalon = '" & sSalon & "'"
    lblSalon = RsSalon!dSalon
    'inicializarSalon 7
    inicializarMesa 64
    'SalonSeleccionado 7, picSalon(1)
    asignarMesa RsSalon!tSalon
    'asignarMesa "01"
    'asignarMesa sSalon
    
    'Set RsPedido = Lib.OpenRecordset("select tCodigoPedido as Codigo, tMesa, fFecha, nPrecuenta from MPEDIDO where tEstadoPedido = '01' order by tMesa", Cn)
    'Set RsGrilla = Lib.OpenRecordset("select tCodigoPedido,tObservacion from MPEDIDO where tEstadoPedido = '01' and tTipoPedido <> '04' and not len(rtrim(tObservacion)) = 0 and tCaja = '" & sCaja & "' order by tObservacion", Cn)
    ''Configuración de Grillas
    'Call ConfGrilla(2, grdPedidos, "Pedido", 2, "tCodigoPedido", 1200, 0, 0, "", _
    '                               "Observación", 2, "tObservacion", 2350, 0, 0, "")
    'Set grdPedidos.DataSource = RsGrilla
   '
   ' inicializarCantidadClientes
    
    
    
End Sub

Private Sub limpiarMesas(cantidadMaxMesas As Integer)
    Dim i As Integer
    'For i = 1 To RsMesa.RecordCount
    ' Si el recordset está vacío, salir del sub
    RsMesa.Requery
    If RsMesa.EOF And RsMesa.BOF Then Exit Sub
    RsMesa.MoveFirst
    
    For i = 1 To cantidadMaxMesas
        PicboxText picMesa(i), "-", ""
        PicboxBgColor picMesa(i), "#FFFFFF" '#FFFFFF
        PicboxBorder picMesa(i), "#CFD3D8" '#FFFFFF
        'PicboxTextColor picMesa(i), "#FF0000"
        PicboxTextBold picMesa(i), 0
        'PicboxTextSize picMesa(2), 12
        picMesa(i).Visible = False
        ' Mover al siguiente registro sólo si no es el último
        If Not RsMesa.EOF Then
            RsMesa.MoveNext
        End If
    Next i
End Sub

Private Sub asignarMesa(codigoSalon As String)
'colores disponibles:
    'vbHexLista "01"
    'vbHexOcupada "02"
    'vbHexPrecuenta "--"
    'vbHexSucia "04"
    'vbHexMesasjuntas "--"
    'vbHexReservada "03"
    'vbHexBloqueada "05"
    'vbHexFServicio "06"
    limpiarMesas 64
    Dim countLibre As Integer
    countLibre = 0
    
    Dim i As Integer
    'RsMesa.Filter = "tSalon = '" & RsSalon!tSalon & "'"
    RsMesa.Filter = "tSalon = '" & codigoSalon & "'"
    
    If RsMesa.RecordCount = 0 Then
      Exit Sub
    End If
    
    RsMesa.MoveFirst
    For i = 1 To RsMesa.RecordCount
            
        PicboxText picMesa(i), RsMesa!tResumido, ""
        Select Case RsMesa!tEstadoMesa
            Case Is = "01" 'Lista
                If RsMesa!tCodigoMesa = tmesaOriginal Then
                    PicboxBgColor picMesa(i), vbHexMesasjuntas
                    PicboxTextColor picMesa(i), "#232323"
                    
                    'Mesa SI seleccionada
                    PicboxBorder picMesa(i), tcolorSeleccion
                    ' Agregar a RsMesaSeleccionada (AddNew con los datos actuales)
                        RsMesaSeleccionada.AddNew
                        Dim f As Field
                        For Each f In RsMesa.Fields
                            RsMesaSeleccionada(f.name).value = f.value
                        Next
                        RsMesaSeleccionada.Update
                        
                        ActualizarTxtMesasJuntar
                Else
                    PicboxBgColor picMesa(i), vbHexLista
                    PicboxTextColor picMesa(i), "#232323"
                End If
                countLibre = countLibre + 1
            
            Case Is = "02" 'Ocupada
                'RsPedidoMesaTabla.Filter = "tMesa = '" & RsMesa!tCodigoMesa & "'"
                'If Not RsPedidoMesaTabla.EOF Then
                '    RsPedidoMesaTabla.MoveFirst
                    'PicboxBgColor picMesa(i), vbHexMesasjuntas
                    'PicboxTextColor picMesa(i), "#232323"
                'Else
                    PicboxBgColor picMesa(i), vbHexOcupada
                    PicboxTextColor picMesa(i), "#232323"
                'End If
            Case Is = "03" 'Reservada
                PicboxBgColor picMesa(i), vbHexReservada
                PicboxTextColor picMesa(i), "#232323"
            
            Case Is = "04" 'Sucia
                ''PicboxBgColor picMesa(i), vbHexSucia
                ''PicboxTextColor picMesa(i), "#232323"
                
                If RsMesa!tCodigoMesa = tmesaOriginal Then
                    PicboxBgColor picMesa(i), vbHexMesasjuntas
                    PicboxTextColor picMesa(i), "#232323"
                    
                    'Mesa SI seleccionada
                    PicboxBorder picMesa(i), tcolorSeleccion
                    ' Agregar a RsMesaSeleccionada (AddNew con los datos actuales)
                        RsMesaSeleccionada.AddNew
                        Dim f04 As Field
                        For Each f04 In RsMesa.Fields
                            RsMesaSeleccionada(f04.name).value = f04.value
                        Next
                        RsMesaSeleccionada.Update
                        
                        ActualizarTxtMesasJuntar
                Else
                    PicboxBgColor picMesa(i), vbHexSucia
                    PicboxTextColor picMesa(i), "#232323"
                End If
                countLibre = countLibre + 1
            
            Case Is = "05" 'Bloqueada
                PicboxBgColor picMesa(i), vbHexBloqueada
                PicboxTextColor picMesa(i), "#232323"
            
            Case Is = "06" 'Fuera de servicio
                PicboxBgColor picMesa(i), vbHexFServicio
                PicboxTextColor picMesa(i), "#98ACB3"
            
            'Case Is = "07" 'Original
        End Select
            
        picMesa(i).Visible = True
        RsMesa.MoveNext
    Next i
    
    
    'RsMesa.MoveFirst
    'Do While Not RsMesa.EOF
    '    If RsMesa!tEstadoMesa = "01" Then
    '        countLibre = countLibre + 1
    '    End If
    '    RsMesa.MoveNext
    'Loop

    
    lblMesas.Caption = sSalonActual & " " & countLibre & "/" & RsMesa.RecordCount & " mesas libres"
End Sub

Private Sub inicializarMesa(cantidadMesas As Integer)
    Set RsMesa = Lib.OpenRecordset("select * from TMESA where tCodigoMesa <> '000' and lActivo = 1 order by tCodigoMesa", Cn)
    
    Set RsMesaSeleccionada = New ADODB.Recordset
    With RsMesaSeleccionada
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        ' Abrimos SOLO la estructura (WHERE 1=0)
        .Open "SELECT * FROM TMESA WHERE 1=0", Cn
        ' Desconectamos de la BD
        .ActiveConnection = Nothing
    End With
    
    
    Dim i As Integer
    'For i = 1 To RsMesa.RecordCount
    
    ' Si el recordset está vacío, salir del sub
    RsMesa.Requery
    If RsMesa.EOF And RsMesa.BOF Then Exit Sub
    RsMesa.MoveFirst
    
    For i = 1 To cantidadMesas
        PicboxText picMesa(i), "-", ""
        PicboxBgColor picMesa(i), "#FFFFFF" '#FFFFFF
        PicboxBorder picMesa(i), "#CFD3D8" '#FFFFFF
        'PicboxTextColor picMesa(i), "#FF0000"
        PicboxTextBold picMesa(i), 0
        'PicboxTextSize picMesa(2), 12
        picMesa(i).Visible = False
        If Not RsMesa.EOF Then
            RsMesa.MoveNext
        End If
    Next i
    RsMesa.MoveFirst
End Sub

Private Sub OrganizaComponentes()
    On Error GoTo fin
    
        frmJuntarMesas.Width = fraSalon.Width - 50
        frmJuntarMesas.Height = txtTitulo.Height + fraHead.Height + fraSalon.Height + fraFooter.Height
        
        
        'FRAME MESASJUNTAR
        txtMesasJuntar.Height = fraMesasJuntar.Height - 40
        txtMesasJuntar.Width = fraMesasJuntar.Width - 40
        txtMesasJuntar.Top = 20
        txtMesasJuntar.Left = 20
        'FRAME MESASJUNTAR
        
        'FRAME PROPIEDADES
        txtNumMesasJuntar.Height = fraNumMesasJuntar.Height - 40
        txtNumMesasJuntar.Width = fraNumMesasJuntar.Width - 40
        txtNumMesasJuntar.Top = 20
        txtNumMesasJuntar.Left = 20
        'FRAME PROPIEDADES
        
        'FRAME MOTORIZADO
        'fraMorotizado.Height = 9015
        'fraMorotizado.Width = 6226.194
        'fraMorotizado.Top = 1440
        'fraMorotizado.Left = 9241.771
        'FRAME MOTORIZADO
        
        'FRAME MESERO
        'fraMozo.Height = 9015
        'fraMozo.Width = 6226.194
        'fraMozo.Top = 1440
        'fraMozo.Left = 9241.771
        'FRAME MESERO
        
        'FRAME ORIGEN DE VENTAS
        'fraOrigenVentas.Height = 9015
        'fraOrigenVentas.Width = 6226.194
        'fraOrigenVentas.Top = 1440
        'fraOrigenVentas.Left = 9241.771
        'FRAME ORIGEN DE VENTAS
        
        'FRAME PUNTO VENTA
        'fraPuntoVenta.Height = 9015
        'fraPuntoVenta.Width = 6226.194
        'fraPuntoVenta.Top = 1440
        'fraPuntoVenta.Left = 9241.771
        'FRAME PUNTO VENTA
        
        'FRAME PRODUCTO COMBO
        'fraProductoCombo.Height = 11325
        'fraProductoCombo.Width = 6226.193
        'fraProductoCombo.Top = 240
        'fraProductoCombo.Left = 9241.771
        'FRAME PRODUCTO COMBO
        
        'FRAME DET COMBO
        'fraCombo.Height = 9975
        'fraCombo.Width = 9256.773
        'fraCombo.Top = 1560
        'fraCombo.Left = 0
        'FRAME DET COMBO
        
        'FRAME CANAL DE VENTA
        'FrameCanalVenta.Width = 6256.2
        'FrameCanalVenta.Height = 3135
        
        'FrameCanalVenta.Top = 780
        'FrameCanalVenta.Left = 9241.772
        'FRAME CANAL DE VENTA
        
    Exit Sub
fin:
    MsgBox "Error" & Error
End Sub

Private Sub picActualizar_Click()
    inicializarMesa 64
    tmrActualiza.Enabled = True
    
End Sub

Private Sub picGuardar_Click()
    'FERM
    Dim filtroMesasAdicionales As String
    filtroMesasAdicionales = ""
    Dim xmesaOriginal As String
    Dim xmesaOriDescripcion As String
    Dim Contador As Integer
    Dim respuesta As VbMsgBoxResult
    
    'RsMesa.MoveFirst
    If RsMesaSeleccionada.RecordCount > 0 Then
        RsMesaSeleccionada.MoveFirst
        If Not RsMesaSeleccionada.EOF Then
            Do While Not RsMesaSeleccionada.EOF
            If Contador = 0 Then
                xmesaOriginal = RsMesaSeleccionada!tCodigoMesa
                xmesaOriDescripcion = RsMesaSeleccionada!tDetallado
            Else
                If filtroMesasAdicionales <> "" Then
                    filtroMesasAdicionales = filtroMesasAdicionales & " or "
                End If
                filtroMesasAdicionales = filtroMesasAdicionales & " tCodigoMesa ='" & RsMesaSeleccionada!tCodigoMesa & "'"
            End If
                
                
                RsMesaSeleccionada.MoveNext
                Contador = Contador + 1
            Loop
            If IsPedidoCajaRapida Then
                respuesta = MsgBox("¿Desea asignar una mesa?", vbQuestion + vbYesNo, "Mesa: " & xmesaOriginal & " - " & xmesaOriDescripcion)
                If respuesta = vbYes Then
                    Unload Me
                    Unload frmMesas
                    'frmCajaRapida.AsignaMesaCajaRapida xmesaOriginal, tnumAdulto, tnumNinio, filtroMesasAdicionales
                End If
            Else
                respuesta = MsgBox("¿Desea crear un pedido con esta mesa?", vbQuestion + vbYesNo, "Mesa: " & xmesaOriginal & " - " & xmesaOriDescripcion)
                If respuesta = vbYes Then
                    Unload Me
                    Unload frmMesas
                    'frmCajaRapida.CrearPedidoMesa xmesaOriginal, tnumAdulto, tnumNinio, filtroMesasAdicionales
                End If
            End If
        End If
    End If
End Sub

Private Sub picMesa_Click(Index As Integer)
    'PicboxBorder picMesa(Index), tcolorSeleccion '"#8f7a76" '"#602113" '"#40160d" '"#bd33ff"
    Dim actualColorHex As String
    Dim tCodigoMesa As String
    Dim estadoMesa As String
    
    RsMesa.MoveFirst
    RsMesa.Move Index - 1
    estadoMesa = Trim(RsMesa!tEstadoMesa)
    tCodigoMesa = RsMesa!tCodigoMesa
    actualColorHex = PicboxGetBorderHex(picMesa(Index))
    
    Select Case RsMesa!tEstadoMesa
            Case Is = "01" 'Lista
                If RsMesa!tCodigoMesa = tmesaOriginal Then
                    Exit Sub
                End If
                If actualColorHex = tcolorSeleccion Then
                    'Mesa NO seleccionada
                    PicboxBorder picMesa(Index), tcolorDeseleccion
                     ' Buscar y eliminar de RsMesaSeleccionada
                    If RsMesaSeleccionada.RecordCount > 0 Then
                        RsMesaSeleccionada.MoveFirst
                        Do While Not RsMesaSeleccionada.EOF
                            If RsMesaSeleccionada!tCodigoMesa = tCodigoMesa Then
                                RsMesaSeleccionada.Delete
                                Exit Do
                            End If
                            RsMesaSeleccionada.MoveNext
                        Loop
                    End If
                Else
                    'Mesa SI seleccionada
                    PicboxBorder picMesa(Index), tcolorSeleccion
                    ' Agregar a RsMesaSeleccionada (AddNew con los datos actuales)
                        RsMesaSeleccionada.AddNew
                        Dim f As Field
                        For Each f In RsMesa.Fields
                            RsMesaSeleccionada(f.name).value = f.value
                        Next
                        RsMesaSeleccionada.Update
                End If
                    
            ActualizarTxtMesasJuntar
            
            Case Is = "02" 'Ocupada
                'PicboxBgColor picMesa(i), vbHexOcupada
                'PicboxTextColor picMesa(i), "#232323"
            
            Case Is = "03" 'Reservada
                'PicboxBgColor picMesa(i), vbHexReservada
                'PicboxTextColor picMesa(i), "#232323"
            
            Case Is = "04" 'Sucia
                'PicboxBgColor picMesa(i), vbHexSucia
                'PicboxTextColor picMesa(i), "#232323"
                If RsMesa!tCodigoMesa = tmesaOriginal Then
                    Exit Sub
                End If
                If actualColorHex = tcolorSeleccion Then
                    'Mesa NO seleccionada
                    PicboxBorder picMesa(Index), tcolorDeseleccion
                     ' Buscar y eliminar de RsMesaSeleccionada
                    If RsMesaSeleccionada.RecordCount > 0 Then
                        RsMesaSeleccionada.MoveFirst
                        Do While Not RsMesaSeleccionada.EOF
                            If RsMesaSeleccionada!tCodigoMesa = tCodigoMesa Then
                                RsMesaSeleccionada.Delete
                                Exit Do
                            End If
                            RsMesaSeleccionada.MoveNext
                        Loop
                    End If
                Else
                    'Mesa SI seleccionada
                    PicboxBorder picMesa(Index), tcolorSeleccion
                    ' Agregar a RsMesaSeleccionada (AddNew con los datos actuales)
                        RsMesaSeleccionada.AddNew
                        Dim f04 As Field
                        For Each f04 In RsMesa.Fields
                            RsMesaSeleccionada(f04.name).value = f04.value
                        Next
                        RsMesaSeleccionada.Update
                End If
                    
            ActualizarTxtMesasJuntar
                
            
            Case Is = "05" 'Bloqueada
                'PicboxBgColor picMesa(i), vbHexBloqueada
                'PicboxTextColor picMesa(i), "#232323"
            
            Case Is = "06" 'Fuera de servicio
                'PicboxBgColor picMesa(i), vbHexFServicio
                'PicboxTextColor picMesa(i), "#98ACB3"
            
            'Case Is = "07" 'Original
        End Select
End Sub

Private Sub tmrActualiza_Timer()
' Este metodo es usado para dar un tiempo de espera y dar la sensacion que si se actualizó (es decir, es para uso visual)
    ' Detén el timer para que se ejecute solo una vez
    tmrActualiza.Enabled = False
    LimpiarMesasSeleccionadas
    asignarMesa RsSalon!tSalon
End Sub
Private Sub LimpiarMesasSeleccionadas()
    ' 1) Reinstanciar y abrir SOLO la estructura
    Set RsMesaSeleccionada = New ADODB.Recordset
    With RsMesaSeleccionada
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open "SELECT * FROM TMESA WHERE 1=0", Cn
        .ActiveConnection = Nothing
    End With

    ' 2) Limpiar los controles
    txtMesasJuntar.Text = ""
    txtNumMesasJuntar.Text = "0"
End Sub



Private Sub picCancelar_Click()
    Unload Me
End Sub
Private Sub ActualizarTxtMesasJuntar()
    Dim sTexto As String
    Dim numMesas As Integer
    sTexto = ""
    numMesas = 0

    If RsMesaSeleccionada Is Nothing Then Exit Sub
    If RsMesaSeleccionada.RecordCount = 0 Then
        txtMesasJuntar.Text = ""
        txtNumMesasJuntar.Text = "0"
        Exit Sub
    End If

    RsMesaSeleccionada.MoveFirst
    Do While Not RsMesaSeleccionada.EOF
        If Trim(RsMesaSeleccionada!tResumido) <> "" Then
            If sTexto <> "" Then sTexto = sTexto & ", "
            sTexto = sTexto & Trim(RsMesaSeleccionada!tResumido)
            numMesas = numMesas + 1
        End If
        RsMesaSeleccionada.MoveNext
    Loop

    ' Actualizar los controles con la información
    txtMesasJuntar.Text = sTexto
    txtNumMesasJuntar.Text = numMesas
End Sub


