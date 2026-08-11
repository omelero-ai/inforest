VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServidorEnlace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locales"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   Icon            =   "frmServidorEnlace.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdVerificarConexion 
      Caption         =   "&Verifica Conexión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   5450
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000004&
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   6615
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5895
      Width           =   6675
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
         Height          =   525
         Index           =   3
         Left            =   5400
         Picture         =   "frmServidorEnlace.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton CmdActualiza 
         Caption         =   "Actualizar"
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
         Left            =   3000
         Picture         =   "frmServidorEnlace.frx":00FE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1125
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
         Height          =   525
         Index           =   0
         Left            =   4200
         Picture         =   "frmServidorEnlace.frx":01C6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1125
      End
   End
   Begin MSComctlLib.ListView lstServidores 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9551
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmServidorEnlace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTotal As Double
Dim sCriterio As String
Dim sTemporal As String
Dim eAlmacen As Boolean
Dim carga As Boolean
Dim conectado As Boolean
Public nPredeterm As Integer

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
        Case 0
        Screen.MousePointer = vbHourglass

            Aceptar
                Screen.MousePointer = vbDefault

        Case 3
            Unload Me
   End Select
End Sub
Public Sub cargaModo(X As Boolean)
    carga = X
End Sub

Private Sub Form_Initialize()
     carga = True
End Sub

Public Sub Form_Load()
    If CnAlmacen.State <> 0 Then
        If carga = True Then
         llenaServidores
            Centrar Me
        End If
    Else
        MsgBox "No existe enlace con Almacen para cargar los locales", vbCritical, sMensaje
        Unload Me
    End If
End Sub
Public Sub llenaServidores()
 configuraListaServidores
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo sale:
   If rsListaServidores.RecordCount > 0 Then
    If conectado = True Then
       wEnter = True
    Else
         wEnter = False
         Unload Me
     End If
        ultimoConectado = True
      Set frmServidorEnlace = Nothing
  End If
  Exit Sub
sale:
     MsgBox "No Hay Locales Creados o Se Encuentran Desactivados", vbCritical, sMensaje
End Sub
Private Sub cmdVerificarConexion_Click()
    Dim conectado As Boolean
    If (rsListaServidores.RecordCount > 0) And (Not rsListaServidores Is Nothing) Then
        If MsgBox("La verificación de la conexión con el Local seleccionado" & vbCrLf & "puede tardar algunos minutos ¿Desea continuar?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                
            If rsListaServidores.RecordCount > 0 Then
            For q = 1 To lstServidores.ListItems.Count
                If lstServidores.ListItems.Item(q).Selected = True Then
            
                    conectado = validaConexion(lstServidores.ListItems.Item(q).SubItems(3), lstServidores.ListItems.Item(q).SubItems(4))
                    If conectado = False Then
                        MsgBox "No se puede establecer conexión con: " & lstServidores.ListItems.Item(q).SubItems(2), vbCritical, sMensaje
                        
                    Else
                        MsgBox "Prueba de conexón satisfactoria con el local: " & lstServidores.ListItems.Item(q).SubItems(2), vbInformation, sMensaje
                    End If
                End If
           Next q
           End If
    
        End If
    End If
End Sub

Public Sub Aceptar()
 'Screen.MousePointer = vbHourglass
On Error Resume Next
Dim q, l As Integer
l = 0
   If rsListaServidores.RecordCount > 0 And Not (rsListaServidores Is Nothing) Then
    For q = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(q).Selected = True Then
         '   If lstServidores.ListItems.Item(q).SubItems(3) <> "Desconectado" Then
                l = l + 1
                conectado = conectaServidores(lstServidores.ListItems.Item(q).SubItems(3), lstServidores.ListItems.Item(q).SubItems(4), lstServidores.ListItems.Item(q).SubItems(5), lstServidores.ListItems.Item(q).SubItems(6))
                If conectado = False Then
                    MsgBox "Error de Conexiòn. Verificar los datos del Local Seleccionado", vbCritical, sMensaje
                    Exit Sub
                End If
                wEnter = True
                localConectado = lstServidores.ListItems.Item(q).SubItems(2)
                MsgBox "Se ha Conectado a: " & localConectado
                Select Case moduloUso
                Case "Consulta"
                    mdiConsulta.reinicia
                Case "Administracion"
                    mdiAdministracion.reinicia
                Case "PuntoVenta"
                    mdiPuntoVenta.reinicia
                End Select
        '   Else
         '       MsgBox "El local seleccionado se encuentra desconectado", vbCritical, sMensaje
          '  End If
        End If
    Next q
    If l = 0 Then
         MsgBox " Debe seleccionar un local ", vbInformation, sMensaje
         Exit Sub
     End If
      Unload Me
   End If
End Sub
Private Sub CmdActualiza_Click()
    Screen.MousePointer = vbHourglass
   configuraListaServidores
    MsgBox "Lista De Servidores Actualizada", vbInformation, sMensaje
    Screen.MousePointer = vbDefault
End Sub

Private Sub lstServidores_DblClick()
    Screen.MousePointer = vbHourglass
    Aceptar
    Screen.MousePointer = vbDefault
End Sub

Public Sub configuraListaServidores()
 Dim rsS As New ADODB.Recordset
 Set rsS = devuelveServidoresEnlaceMultilocal

'    If (rsS Is Nothing Or rsS.RecordCount = 0) Then
'        Unload Me
'    End If

    lstServidores.ColumnHeaders.Clear
    lstServidores.ListItems.Clear
    Dim Item As ListItem
    With lstServidores
         .ColumnHeaders.Add , , " Código ", 1000
         .ColumnHeaders.Add , , " Sector ", 2500
         .ColumnHeaders.Add , , " Local ", 2500
       '  .ColumnHeaders.Add , , " Estado ", 1500
         .ColumnHeaders.Add , , "Ip", 0
         .ColumnHeaders.Add , , "Bd", 0
         .ColumnHeaders.Add , , "BdAlm", 0
         .ColumnHeaders.Add , , "IpAlm", 0
         .View = lvwReport
         .GridLines = True
         
     Dim X As Integer
     If rsS.RecordCount > 0 Then
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1)
            Item.SubItems(2) = rsS.Fields(2)
          ' Item.SubItems(3) = rsS.Fields(6)
            Item.SubItems(3) = rsS.Fields(3)
            Item.SubItems(4) = rsS.Fields(4)
            Item.SubItems(5) = rsS.Fields(5)
            Item.SubItems(6) = rsS.Fields(6)
        rsS.MoveNext
        Next X
    End If
    End With
End Sub

Private Sub lstServidores_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Screen.MousePointer = vbHourglass
            Aceptar
            Screen.MousePointer = vbDefault
    End If
End Sub

Public Function devuelveServidoresEnlaceMultilocal() As ADODB.Recordset
    Dim sTemporal As String
    Dim cadena As Boolean
    Dim Estado As String
    Dim RST As New ADODB.Recordset
    sTemporal = dbTemporalEnlace(sCaja, 8, "codigo", "nvarchar(3)", "Sector", "nvarchar(30)", "Local", "nvarchar(30)", "Ip", "nvarchar(30)", "BD", "nvarchar(30)", "BDALM", "nvarchar(30)", "Estado", "nvarchar(15)", "IpAlm", "nvarchar(30)")
    
    Isql = "select * from vLocalidades WHERE  ip<>'' and bdinf<>'' and LACTIVO=1"
    Set RST = Lib.OpenRecordset(Isql, cnAlmacenDefault)
    If Not (RST.EOF Or RST.BOF) Then
            RST.MoveFirst
            For X = 0 To RST.RecordCount - 1

                  Estado = "estado"
                 Isql = " insert into " & sTemporal & " values ( '" & RST.Fields(0) & "','" & RST.Fields(1) & "','" & RST.Fields(3) & "','" & RST.Fields(4) & "','" & RST.Fields(5) & "','" & RST.Fields(7) & "','" & Estado & "','" & RST.Fields(6) & "' )"
                 cnDefault.Execute Isql
             RST.MoveNext
            Next X
            Isql = " select codigo,sector,local,ip,bd,bdalm,ipalm from " & sTemporal & " order by 2,3"
            Set RST = Lib.OpenRecordset(Isql, cnDefault)
             Set devuelveServidoresEnlaceMultilocal = RST
             Set rsListaServidores = RST
    Else
        Set devuelveServidoresEnlaceMultilocal = RST
        
        Set rsListaServidores = RST
        Set RST = Nothing
    End If
    
End Function

