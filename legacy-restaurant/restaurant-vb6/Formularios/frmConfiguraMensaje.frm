VERSION 5.00
Begin VB.Form frmConfiguraMensaje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Mensajes"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   ForeColor       =   &H00808080&
   Icon            =   "frmConfiguraMensaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   11805
   Begin VB.CommandButton cmdMensajeImpresora 
      Caption         =   "Impresora"
      Height          =   300
      Index           =   5
      Left            =   8040
      TabIndex        =   32
      Top             =   2210
      Width           =   2055
   End
   Begin VB.CommandButton cmdMensajeImpresora 
      Caption         =   "Impresora"
      Height          =   300
      Index           =   4
      Left            =   8040
      TabIndex        =   31
      Top             =   1822
      Width           =   2055
   End
   Begin VB.CommandButton cmdMensajeImpresora 
      Caption         =   "Impresora"
      Height          =   300
      Index           =   3
      Left            =   8040
      TabIndex        =   30
      Top             =   1434
      Width           =   2055
   End
   Begin VB.CommandButton cmdMensajeImpresora 
      Caption         =   "Impresora"
      Height          =   300
      Index           =   2
      Left            =   8040
      TabIndex        =   29
      Top             =   1046
      Width           =   2055
   End
   Begin VB.CommandButton cmdMensajeImpresora 
      Caption         =   "Impresora"
      Height          =   300
      Index           =   1
      Left            =   8040
      TabIndex        =   28
      Top             =   658
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos Generales "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5115
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10365
      Begin VB.CommandButton cmdMensajeImpresora 
         Caption         =   "Impresora"
         Height          =   300
         Index           =   11
         Left            =   8040
         TabIndex        =   38
         Top             =   4545
         Width           =   2055
      End
      Begin VB.CommandButton cmdMensajeImpresora 
         Caption         =   "Impresora"
         Height          =   300
         Index           =   10
         Left            =   8040
         TabIndex        =   37
         Top             =   4150
         Width           =   2055
      End
      Begin VB.CommandButton cmdMensajeImpresora 
         Caption         =   "Impresora"
         Height          =   300
         Index           =   9
         Left            =   8040
         TabIndex        =   36
         Top             =   3762
         Width           =   2055
      End
      Begin VB.CommandButton cmdMensajeImpresora 
         Caption         =   "Impresora"
         Height          =   300
         Index           =   8
         Left            =   8040
         TabIndex        =   35
         Top             =   3374
         Width           =   2055
      End
      Begin VB.CommandButton cmdMensajeImpresora 
         Caption         =   "Impresora"
         Height          =   300
         Index           =   7
         Left            =   8040
         TabIndex        =   34
         Top             =   2986
         Width           =   2055
      End
      Begin VB.CommandButton cmdMensajeImpresora 
         Caption         =   "Impresora"
         Height          =   300
         Index           =   6
         Left            =   8040
         TabIndex        =   33
         Top             =   2598
         Width           =   2055
      End
      Begin VB.CommandButton cmdMensajeImpresora 
         Caption         =   "Impresora"
         Height          =   300
         Index           =   0
         Left            =   8040
         TabIndex        =   27
         Top             =   270
         Width           =   2055
      End
      Begin VB.TextBox txtF12 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   11
         Top             =   4545
         Width           =   5385
      End
      Begin VB.TextBox txtF11 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   10
         Top             =   4150
         Width           =   5385
      End
      Begin VB.TextBox txtF10 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   9
         Top             =   3762
         Width           =   5385
      End
      Begin VB.TextBox txtF9 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   8
         Top             =   3374
         Width           =   5385
      End
      Begin VB.TextBox txtF8 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2986
         Width           =   5385
      End
      Begin VB.TextBox txtF7 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2598
         Width           =   5385
      End
      Begin VB.TextBox txtF6 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2210
         Width           =   5385
      End
      Begin VB.TextBox txtF5 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1822
         Width           =   5385
      End
      Begin VB.TextBox txtF4 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1434
         Width           =   5385
      End
      Begin VB.TextBox txtF3 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1046
         Width           =   5385
      End
      Begin VB.TextBox txtF2 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   1
         Top             =   658
         Width           =   5385
      End
      Begin VB.TextBox txtF1 
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
         Left            =   2385
         MaxLength       =   50
         TabIndex        =   0
         Top             =   270
         Width           =   5385
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F12] :"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   26
         Top             =   4590
         Width           =   2160
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F11] :"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   25
         Top             =   4195
         Width           =   2160
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F10] :"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   24
         Top             =   3807
         Width           =   2160
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F9] :"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   23
         Top             =   3419
         Width           =   2070
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F8] :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   22
         Top             =   3031
         Width           =   2070
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F7] :"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   21
         Top             =   2643
         Width           =   2070
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F6] :"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   20
         Top             =   2255
         Width           =   2070
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F5] :"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   19
         Top             =   1867
         Width           =   2070
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F4] :"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   1479
         Width           =   2070
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F3] :"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   1091
         Width           =   2070
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F2] :"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   703
         Width           =   2070
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje Tecla Funcion [F1] :"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   315
         Width           =   2070
      End
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
      Height          =   615
      Index           =   1
      Left            =   5550
      Picture         =   "frmConfiguraMensaje.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5175
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
      Index           =   0
      Left            =   6780
      Picture         =   "frmConfiguraMensaje.frx":010E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5175
      Width           =   1170
   End
End
Attribute VB_Name = "frmConfiguraMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsMensaje As Recordset

Private Sub cmdMensajeImpresora_Click(Index As Integer)

    Select Case Index
           Case Is = 0 'f1
                teclaFunction = "tF1"
                mensajeImpresora = txtF1.Text
           Case Is = 1 'f2
                teclaFunction = "tF2"
                mensajeImpresora = txtF2.Text
           Case Is = 2 'f3
                teclaFunction = "tF3"
                mensajeImpresora = txtF3.Text
           Case Is = 3 'f4
                teclaFunction = "tF4"
                mensajeImpresora = txtF4.Text
           Case Is = 4 'f5
                teclaFunction = "tF5"
                mensajeImpresora = txtF5.Text
           Case Is = 5 'f6
                teclaFunction = "tF6"
                mensajeImpresora = txtF6.Text
           Case Is = 6 'f7
                teclaFunction = "tF7"
                mensajeImpresora = txtF7.Text
           Case Is = 7 'f8
                teclaFunction = "tF8"
                mensajeImpresora = txtF8.Text
           Case Is = 8 'f9
                teclaFunction = "tF9"
                mensajeImpresora = txtF9.Text
           Case Is = 9 'f10
                teclaFunction = "tF10"
                mensajeImpresora = txtF10.Text
           Case Is = 10 'f11
                teclaFunction = "tF11"
                mensajeImpresora = txtF11.Text
           Case Is = 11 'f12
                teclaFunction = "tF12"
                mensajeImpresora = txtF12.Text
    End Select
    
    If mensajeImpresora = "" Or teclaFunction = "" Then
       MsgBox "Ingrese el texto para la tecla de Funcion " & teclaFunction, vbExclamation, sMensaje
       Exit Sub
    End If
    
    If Calcular("select count(tImpresora) as Codigo From TIMPRESORA where tcaja='" & sCaja & "'", Cn) > 0 Then
        frmConfiguraMensajeImpresora.Show vbModal
    Else
       MsgBox "No hay impresoras configuradas para esta caja. Por favor agregue una impresora desde el Mantenimiento de Impresoras. ", vbExclamation, sMensaje
       Exit Sub
    End If
   
    
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 0 Then
      Screen.MousePointer = vbHourglass
      Isql = "Update TMENSAJE Set " & _
              "tF1 = '" & txtF1.Text & "', " & _
              "tF2 = '" & txtF2.Text & "', " & _
              "tF3 = '" & txtF3.Text & "', " & _
              "tF4 = '" & txtF4.Text & "', " & _
              "tF5 = '" & txtF5.Text & "', " & _
              "tF6 = '" & txtF6.Text & "', " & _
              "tF7 = '" & txtF7.Text & "', " & _
              "tF8 = '" & txtF8.Text & "', " & _
              "tF9 = '" & txtF9.Text & "', " & _
              "tF10 = '" & txtF10.Text & "', " & _
              "tF11 = '" & txtF11.Text & "', " & _
              "tF12 = '" & txtF12.Text & "'"
      Cn.Execute Isql
      Screen.MousePointer = vbDefault
      MsgBox "Parámetros Actualizados", vbInformation, sMensaje
      Unload Me
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   Isql = "select * from TMENSAJE"
   Set RsMensaje = Lib.OpenRecordset(Isql, Cn)
   txtF1.Text = IIf(IsNull(RsMensaje!tF1), "", RsMensaje!tF1)
   txtF2.Text = IIf(IsNull(RsMensaje!tF2), "", RsMensaje!tF2)
   txtF3.Text = IIf(IsNull(RsMensaje!tF3), "", RsMensaje!tF3)
   txtF4.Text = IIf(IsNull(RsMensaje!tF4), "", RsMensaje!tF4)
   txtF5.Text = IIf(IsNull(RsMensaje!tF5), "", RsMensaje!tF5)
   txtF6.Text = IIf(IsNull(RsMensaje!tF6), "", RsMensaje!tF6)
   txtF7.Text = IIf(IsNull(RsMensaje!tF7), "", RsMensaje!tF7)
   txtF8.Text = IIf(IsNull(RsMensaje!tF8), "", RsMensaje!tF8)
   txtF9.Text = IIf(IsNull(RsMensaje!tF9), "", RsMensaje!tF9)
   txtF10.Text = IIf(IsNull(RsMensaje!tF10), "", RsMensaje!tF10)
   txtF11.Text = IIf(IsNull(RsMensaje!tF11), "", RsMensaje!tF11)
   txtF12.Text = IIf(IsNull(RsMensaje!tF12), "", RsMensaje!tF12)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   RsMensaje.Close
   Set RsMensaje = Nothing
   Set frmConfiguraMensaje = Nothing
End Sub

