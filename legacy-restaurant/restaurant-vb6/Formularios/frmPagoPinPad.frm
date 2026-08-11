VERSION 5.00
Begin VB.Form frmPagoPinPad 
   Caption         =   "Elige PinPad de Pago Niubiz"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTerminal 
      Caption         =   "PinPad Nº"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmPagoPinPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim RsTerminal As ADODB.Recordset
    Dim Isql As String
    Dim idx As Integer         ' índice para el control array (comienza en 1)
    Dim Pos As Integer         ' contador de posición (0-based) para layout
    Dim btnWidth As Integer, btnHeight As Integer
    Dim perRow As Integer
    Dim leftStart As Integer, topStart As Integer
    
    Dim countX As Integer
    Dim countY As Integer

    On Error GoTo errHandler

    ' --- parámetros layout ---
    btnWidth = 2000: btnHeight = 800
    perRow = 10
    leftStart = 100: topStart = 100

    ' --- limpiar botones anteriores (si hay) ---
    On Error Resume Next
    Dim k As Integer
    k = 1
    Do
        Unload cmdTerminal(k)
        k = k + 1
    Loop While Err.Number = 0
    Err.Clear
    On Error GoTo errHandler

    ' --- abrir recordset ---
    Isql = "select * from TCAJATERMINAL inner join TTERMINAL on TCAJATERMINAL.tTerminal = TTERMINAL.Codigo where TTERMINAL.tIntegracion = '01' and TCAJATERMINAL.tCaja = '" & sCaja & "' and TCAJATERMINAL.lActivo = 1 order by tTerminal.Codigo"
    Set RsTerminal = Lib.OpenRecordset(Isql, Cn)

    If RsTerminal Is Nothing Then
        MsgBox "No se pudo cargar los terminales, Revisar Configuración.", vbExclamation
        Exit Sub
    End If

    If RsTerminal.State <> adStateOpen Then
        MsgBox "Se cerró la conexión", vbExclamation
        Exit Sub
    End If

    If RsTerminal.EOF Then
        ' no hay registros: dejar plantilla invisible y salir
        cmdTerminal(0).Visible = False
        RsTerminal.Close
        Set RsTerminal = Nothing
        Exit Sub
    End If

    cmdTerminal(0).Visible = False ' plantilla oculta

    idx = 1    ' primer índice que vamos a cargar (0 ya existe)
    Pos = 0    ' posición 0-based para calcular Left/Top

    RsTerminal.MoveFirst
    Do While Not RsTerminal.EOF
        Load cmdTerminal(idx)

        With cmdTerminal(idx)
            .Caption = CStr(RsTerminal!tDescripcion)
            .Left = leftStart
            .Top = topStart
            .Width = btnWidth
            .Height = btnHeight
            .Tag = CStr(RsTerminal!codigo) ' guarda el identificador en Tag
            .Visible = True
        End With
        
        countX = countX + 1
        
        leftStart = leftStart + btnWidth
        
        If countX > 2 Then
            topStart = topStart + btnHeight
            leftStart = 100
            countX = 0
        End If

        idx = idx + 1
        Pos = Pos + 1
        RsTerminal.MoveNext
    Loop

    ' limpieza
    If Not RsTerminal Is Nothing Then
        If RsTerminal.State = adStateOpen Then RsTerminal.Close
    End If
    Set RsTerminal = Nothing

    Exit Sub

errHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    On Error Resume Next
    If Not RsTerminal Is Nothing Then
        If RsTerminal.State = adStateOpen Then RsTerminal.Close
    End If
    Set RsTerminal = Nothing
End Sub

Private Sub cmdTerminal_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To cmdTerminal.UBound
        cmdTerminal(i).Enabled = False
    Next i
    
    frmPago.TerminalSeleccionado = cmdTerminal(Index).Tag
    Unload Me
End Sub
