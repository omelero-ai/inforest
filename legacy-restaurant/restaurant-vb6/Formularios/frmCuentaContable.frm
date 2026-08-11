VERSION 5.00
Begin VB.Form frmCuentaContable 
   Caption         =   "Cuentas Contables"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
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
      Left            =   1500
      Picture         =   "frmCuentaContable.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2850
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
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
      Index           =   0
      Left            =   2970
      Picture         =   "frmCuentaContable.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2850
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
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
      Left            =   30
      Picture         =   "frmCuentaContable.frx":0624
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2850
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cuentas Contables "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   4395
      Begin VB.TextBox txtClienteD 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   14
         Top             =   630
         Width           =   1755
      End
      Begin VB.TextBox txtRedondeo 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2310
         Width           =   1755
      End
      Begin VB.TextBox txtImpuesto 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1050
         Width           =   1755
      End
      Begin VB.TextBox txtServicio 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1470
         Width           =   1755
      End
      Begin VB.TextBox txtDescuento 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1890
         Width           =   1755
      End
      Begin VB.TextBox txtCliente 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   1
         Top             =   210
         Width           =   1755
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Deudor :"
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
         Left            =   330
         TabIndex        =   15
         Top             =   690
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Redondeo :"
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
         Left            =   720
         TabIndex        =   10
         Top             =   2370
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descuento :"
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
         Left            =   675
         TabIndex        =   8
         Top             =   1950
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Servicio :"
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
         Left            =   900
         TabIndex        =   7
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Impuesto :"
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
         Left            =   825
         TabIndex        =   6
         Top             =   1110
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Contado :"
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
         Left            =   255
         TabIndex        =   2
         Top             =   300
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmCuentaContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tCtaCli As String
Dim tCtaCde As String
Dim tCtaImp As String
Dim tCtaSer As String
Dim tCtaDes As String
Dim tCtaRed As String
Dim tTipo As String

Private Sub cmdGrabar_Click()
    If txtCliente.Text = "" Then
       MsgBox "Cuenta contable de cliente" & Chr(13) & "no puede quedar en blanco.", vbOKOnly, "Transferencia"
       Exit Sub
    End If
    If txtClienteD.Text = "" Then
       MsgBox "Cuenta contable de cliente" & Chr(13) & "no puede quedar en blanco.", vbOKOnly, "Transferencia"
       Exit Sub
    End If
    If txtImpuesto.Text = "" Then
       MsgBox "Cuenta contable de impuesto" & Chr(13) & "no puede quedar en blanco.", vbOKOnly, "Transferencia"
       txtCliente.SetFocus
       Exit Sub
    End If
    If txtServicio.Text = "" Then
       MsgBox "Cuenta contable de servicio" & Chr(13) & "no puede quedar en blanco.", vbOKOnly, "Transferencia"
       txtClienteD.SetFocus
       Exit Sub
    End If
    If txtDescuento.Text = "" Then
       MsgBox "Cuenta contable de descuento" & Chr(13) & "no puede quedar en blanco.", vbOKOnly, "Transferencia"
       txtDescuento.SetFocus
       Exit Sub
    End If
    If txtRedondeo.Text = "" Then
       MsgBox "Cuenta contable de redondeo" & Chr(13) & "no puede quedar en blanco.", vbOKOnly, "Transferencia"
       txtRedondeo.SetFocus
       Exit Sub
    End If
    tCtaCli = txtCliente.Text
    tCtaCde = txtClienteD.Text
    tCtaImp = txtImpuesto.Text
    tCtaSer = txtServicio.Text
    tCtaDes = txtDescuento.Text
    tCtaRed = txtRedondeo.Text
    Open "CUENTA.INI" For Output As #1
    Print #1, tCtaCli
    Print #1, tCtaCde
    Print #1, tCtaImp
    Print #1, tCtaSer
    Print #1, tCtaDes
    Print #1, tCtaRed
    Close #1
    txtCliente.Enabled = False
    txtClienteD.Enabled = False
    txtImpuesto.Enabled = False
    txtServicio.Enabled = False
    txtDescuento.Enabled = False
    txtRedondeo.Enabled = False
    cmdGrabar.Enabled = False
    cmdModificar.Enabled = True
End Sub

Private Sub cmdModificar_Click()
    txtCliente.Enabled = True
    txtClienteD.Enabled = True
    txtImpuesto.Enabled = True
    txtServicio.Enabled = True
    txtDescuento.Enabled = True
    txtRedondeo.Enabled = True
    cmdGrabar.Enabled = True
    cmdModificar.Enabled = False
End Sub

Private Sub cmdSalir_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Open "CUENTA.INI" For Input As #1  'Abre el Archivo
    Do While Not EOF(1)
       Input #1, tCtaCli, tCtaCde, tCtaImp, tCtaSer, tCtaDes, tCtaRed
    Loop
    Close #1   ' Cierra el archivo.
'                     ctaCl = "121000    "
'                     ctaIm = "401010    "
'                     ctaSe = "469110    "
'                     ctaDe = "741010    "
'                     ctaRe = "949990    "
    txtCliente.Text = tCtaCli
    txtClienteD.Text = tCtaCde
    txtImpuesto.Text = tCtaImp
    txtServicio.Text = tCtaSer
    txtDescuento.Text = tCtaDes
    txtRedondeo.Text = tCtaRed
    cmdGrabar.Enabled = False
End Sub

