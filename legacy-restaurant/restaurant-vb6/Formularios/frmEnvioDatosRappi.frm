VERSION 5.00
Begin VB.Form frmEnvioDatosRappi 
   Caption         =   "Confirmación de Bolsas Y Bebidas"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton btnEnviar 
         Caption         =   "Enviar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox ChkBebidas 
         Alignment       =   1  'Right Justify
         Caption         =   "Bebida fuera de la bolsa :"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtBolsas 
         Height          =   375
         Left            =   2550
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese Cantidad de Bolsas :"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmEnvioDatosRappi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StoreId As String
Public OrderIdd As String
Public IdLocall As Integer
Public CodPedido As String
Dim textval As String
Dim numval As String

Private Sub btnEnviar_Click()
    Dim nBolsas As Integer
    Dim bebidas As Integer
    
    nBolsas = CInt(txtBolsas.Text)
    bebidas = IIf(ChkBebidas.Value, 1, 0)
    
    ProcesarInformacion nBolsas, bebidas
    
End Sub

Public Sub ProcesarInformacion(nBolsas As Integer, bebidas As Integer)
    Dim sql As String
    Dim Valor As String
    Dim Rsss As ADODB.Recordset
    
    If (Calcular("Select ISNULL(JsonRespuestaEnvioDatos,'') as codigo From OrdenesRappiDespachador Where Aplicativo = 'RAPPI' and StoreId = '" & StoreId & "' and OrderId = '" & OrderIdd & "'", Cn)) <> "" Then
        Valor = Calcular("Select CodProductoConfirmado as Codigo From OrdenesRappiDespachador Where Aplicativo = 'RAPPI' and StoreId = '" & StoreId & "' and OrderId = '" & OrderIdd & "'", Cn)
    Else
        sql = "exec  SP_DepachoPedidosRappi " & 4 & ", '" & StoreId & "', '" & OrderIdd & "'," & IdLocall & ", '',''," & nBolsas & "," & bebidas & ""

        Set Rsss = Lib.OpenRecordset(sql, Cn)

        If Rsss.EOF = False Then
            If Rsss.RecordCount > 0 Then
                Rsss.MoveFirst
                Valor = Rsss!respuesta
            End If
        End If
    End If

    If Valor = "200" Then
        MsgBox "Los datos se guardaron de manera correcta.", vbExclamation, sMensaje
    Else
        MsgBox "Ocurrió un error al generar el código QR. Por favor, revise los registros de errores del sistema para obtener más detalles.", vbExclamation, sMensaje
    End If
    
    frmOrdenesConsola.RsCabecera.Requery
    
    Unload Me
    
End Sub

Private Sub txtBolsas_KeyPress(KeyAscii As Integer)
    ' Solo permitir dígitos (0–9) y la tecla Backspace (ASCII 8)
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        ' Se permite
    Else
        KeyAscii = 0 ' Bloquea cualquier otro carácter (letras, símbolos, espacio, etc.)
    End If
End Sub
