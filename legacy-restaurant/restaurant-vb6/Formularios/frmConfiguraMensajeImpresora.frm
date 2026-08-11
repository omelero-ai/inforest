VERSION 5.00
Begin VB.Form frmConfiguraMensajeImpresora 
   Caption         =   "Asignación de Impresoras"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
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
      Left            =   2310
      Picture         =   "frmConfiguraMensajeImpresora.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   1170
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
      Left            =   1080
      Picture         =   "frmConfiguraMensajeImpresora.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.VScrollBar VScroll1 
         Height          =   2370
         Left            =   4080
         TabIndex        =   4
         Top             =   1040
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   3975
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            Height          =   2235
            Left            =   0
            ScaleHeight     =   2175
            ScaleWidth      =   3915
            TabIndex        =   5
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Titulo"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Selecciona las impresoras que serán asignadas al mensaje."
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmConfiguraMensajeImpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOpcion_Click(Index As Integer)
     If Index = 0 Then
      Screen.MousePointer = vbHourglass
      Cn.Execute "delete from TMENSAJEIMPRESORA where tCaja ='" & sCaja & "' and teclaFuncion ='" & teclaFunction & "'"
      
       ' Recorrer checkboxes y guardar seleccionados
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is CheckBox Then
                If ctrl.Name Like "chk*" And ctrl.Name <> "chkTodos" Then
                    If ctrl.Value = 1 Then
                        idImpresora = ctrl.Tag  ' El ID está en Tag
                        Dim sqlInsert As String
                        sqlInsert = "INSERT INTO TMENSAJEIMPRESORA (tcaja, teclaFuncion, timpresora, tUsuarioReg, fRegistro ) " & _
                            "VALUES ('" & sCaja & "', '" & teclaFunction & "', '" & idImpresora & "', '" & sUsuario & "', getdate() )"
                        Cn.Execute sqlInsert
                    End If
                End If
            End If
        Next
      Screen.MousePointer = vbDefault
      MsgBox "Parámetros Actualizados", vbInformation, sMensaje
      Unload Me
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Load()
Centrar Me
Label1.Caption = "Tecla " & teclaFunction & " - " & mensajeImpresora
Dim RsImpresoras As Recordset
Dim RsImpresorasExists As Recordset
Dim topOffset As Integer
Dim texto As String
Dim chk As CheckBox

 ' Posicionamiento inicial dentro del Frame
 Set chkHandlers = New Collection
    topOffset = 100
   Isql = "select * from TIMPRESORA where tcaja='" & sCaja & "' order by TIMPRESORA asc"
   Set RsImpresoras = Lib.OpenRecordset(Isql, Cn)
   
   Isql = "select * from TMENSAJEIMPRESORA where tcaja='" & sCaja & "' and teclaFuncion ='" & teclaFunction & "'"
   Set RsImpresorasExists = Lib.OpenRecordset(Isql, Cn)
   
'   Set chk = Me.Controls.Add("VB.CheckBox", "chkTodos", Picture1)
'    With chk
'        .Caption = "Seleccionar todos"
'        .Left = 120
'        .Top = topOffset
'        .Width = Me.TextWidth(.Caption) + 300
'        .Visible = True
'    End With
'
'    topOffset = topOffset + 500
   
   Do Until RsImpresoras.EOF
       
        texto = RsImpresoras("tDescripcion")
        Set chk = Me.Controls.Add("VB.CheckBox", "chk" & RsImpresoras.AbsolutePosition, Picture1)
        With chk
            .Caption = texto
            .Tag = RsImpresoras("timpresora") ' Guardar ID real
            .Left = 120
            .Width = Me.TextWidth(texto) + 300
            .Top = topOffset
            .Visible = True
            
            If Not (RsImpresorasExists.EOF And RsImpresorasExists.BOF) Then
            'Cruzar con el otro Recordset
                RsImpresorasExists.MoveFirst
                Do Until RsImpresorasExists.EOF
                    If RsImpresorasExists("teclaFuncion") = teclaFunction And RsImpresorasExists("tcaja") = RsImpresoras("tcaja") And RsImpresorasExists("timpresora") = RsImpresoras("timpresora") Then
                        .Value = 1  ' Marcar como seleccionado
                        Exit Do
                    End If
                    RsImpresorasExists.MoveNext
                Loop
            End If
            
        End With

         ' Asignar manejador de eventos
        
        topOffset = topOffset + 500  ' Espacio entre CheckBoxes
        RsImpresoras.MoveNext
    Loop
     ' Ajustar tamaño del PictureBox si hay más contenido del que se ve
        Picture1.Height = topOffset + 300
        Picture1.Top = 0
        ' Configurar ScrollBar
        With VScroll1
            .Min = 0
            .Max = Picture1.Height - Frame2.Height
            .SmallChange = 100
            .LargeChange = 500
        End With
End Sub

Private Sub VScroll1_Change()
    Picture1.Top = -VScroll1.Value
End Sub
Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub
Private Sub chkTodos_Click()
    Dim ctrl As Control
    Dim marcarTodos As Boolean

    marcarTodos = Me.Controls("chkTodos").Value = 1

    For Each ctrl In Me.Controls
        If TypeOf ctrl Is CheckBox Then
            If ctrl.Name Like "chk*" And ctrl.Name <> "chkTodos" Then
                ctrl.Value = IIf(marcarTodos, 1, 0)
            End If
        End If
    Next
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Forzar ejecución del evento chkTodos.Click si cambia
    On Error Resume Next
    ' Si se hace clic en cualquier parte, revisa si chkTodos cambió
    Dim chkTodos As CheckBox
    Set chkTodos = Me.Controls("chkTodos")
    
    Static lastValue As Integer
    
    ' Ejecutar solo si el valor cambió
    If chkTodos.Value <> lastValue Then
        lastValue = chkTodos.Value
        chkTodos_Click
    End If
End Sub
