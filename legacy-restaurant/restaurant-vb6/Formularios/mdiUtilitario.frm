VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiUtilitario 
   BackColor       =   &H8000000C&
   Caption         =   "Utilitarios"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8460
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5385
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3889
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            Text            =   "Now"
            TextSave        =   "27/09/2004"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            TextSave        =   "10:09"
         EndProperty
      EndProperty
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
   Begin VB.Menu mnuGeneral 
      Caption         =   "General"
   End
   Begin VB.Menu mnuInforest 
      Caption         =   "Inforest"
      Begin VB.Menu mnuReiniciar 
         Caption         =   "Reiniciar Datos"
      End
   End
   Begin VB.Menu mnuAlmacen 
      Caption         =   "Almacen && Costos"
   End
End
Attribute VB_Name = "mdiUtilitario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAlmacen_Click()

End Sub

Private Sub mnuGeneral_Click()

End Sub

Private Sub mnuReiniciar_Click()
   wEnter = False
   
   'Cambia el MsgBox
   If MsgBox("Seguro de Eliminar todos lo movimientos?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
      Exit Sub
   End If
   
   frmPassword.Show vbModal
   If wEnter Then
      If sDescrip <> sPAdmin Then
         MsgBox "Password incorrecto", vbCritical, sMensaje
         Exit Sub
      End If
   Else
      Exit Sub
   End If
   
   'Cambia el MsgBox
   If MsgBox("Si prosigue se eliminará todos los datos", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   Cn.Execute "delete from MPEDIDO"
   Cn.Execute "delete from DPEDIDO"
   Cn.Execute "delete from CPEDIDO"
   Cn.Execute "delete from APEDIDO"
   
   Cn.Execute "delete from MDOCUMENTO"
   Cn.Execute "delete from DDOCUMENTO"
   Cn.Execute "delete from DPAGODOCUMENTO"
   Cn.Execute "delete from DPREPAGO"
   Cn.Execute "delete from MNOTACREDITO"
   
   Cn.Execute "delete from MTURNO"
   Cn.Execute "delete from MCIERRE"
   Cn.Execute "delete from TTABLATEMP"
   Cn.Execute "delete from MEGRESO"
   Cn.Execute "delete from MINGRESO"
   Cn.Execute "delete from TRESERVA"

   Cn.Execute "update TMESA set tEstadoMesa='01'"
   Cn.Execute "update TPARAMETRO set nCorrelativo=0"
   Cn.Execute "update TCODIGO set Recibo= '" & Mid(Year(Date), 3, 2) & "00000000" & "', Pedido='" & Mid(Year(Date), 3, 2) & "00000000', Turno='" & Mid(Year(Date), 3, 2) & "00000000', Reserva='" & Mid(Year(Date), 3, 2) & "00000000'"
   Cn.Execute "update TCOMPANIA set nConsumo = 0"
   Screen.MousePointer = vbDefault
   
   MsgBox "El reinicio se realizó con exito", vbInformation, sMensaje
   
End Sub
