VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGrupoReplica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupo de Usuarios"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   6615
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Administración Centralizada"
      TabPicture(0)   =   "frmGrupoReplica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tvCentralizada"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.TreeView tvCentralizada 
         Height          =   6135
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   10821
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         SingleSel       =   -1  'True
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9870
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6885
      Width           =   9930
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
         Height          =   615
         Index           =   3
         Left            =   8640
         Picture         =   "frmGrupoReplica.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
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
         Index           =   1
         Left            =   7440
         Picture         =   "frmGrupoReplica.frx":010E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   6825
      Left            =   0
      Picture         =   "frmGrupoReplica.frx":0640
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2280
   End
End
Attribute VB_Name = "frmGrupoReplica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrupoAcceso As ADODB.Recordset
Dim RsAcceso As ADODB.Recordset
Dim lOferta As Boolean

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
              ' ActivarBotones (False)
               Blanquear Me
          Case Is = 1 ' Grabar
                   Sw = False
                   'Cambiar el SQL
                       GrabaArbol
                       MsgBox "Registro Guardado", vbInformation, sMensaje
           '  End If
          Case Is = 3 ' Salir
               Unload Me
   End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Tablas "
   
   ' Set RsAcceso = Lib.OpenRecordset("select tCodigoAcceso, tDescripcion, tModulo,isnull(ttabla,'NO DEF') TTABLA, isnull(lenvia,0) lenvia  from TACCESO where tTipoObjeto='MN' and tmodulo='03' and isnull(ttabla,'NO DEF') <>'NO DEF' ", Cn)
   'Set RsAcceso = Lib.OpenRecordset("SELECT dbo.TACCESO.tCodigoAcceso, dbo.TACCESO.tDescripcion, dbo.TACCESO.tModulo, ISNULL(dbo.TACCESOENVIA.lEnvia, 0) AS lenvia, dbo.TACCESO.tFormulario , dbo.TACCESO.tNombreObjeto, dbo.TACCESO.tTipoObjeto FROM         dbo.TACCESO LEFT OUTER JOIN dbo.TACCESOENVIA ON dbo.TACCESO.tCodigoAcceso = dbo.TACCESOENVIA.tCodigoAcceso WHERE     (dbo.TACCESO.tTipoObjeto = 'MN') AND (dbo.TACCESO.tModulo = '03') AND (SUBSTRING(dbo.TACCESO.tCodigoAcceso, 1, 3) <> '104')", Cn)
   Set RsAcceso = Lib.OpenRecordset("select * from vTablasCentralizada order by 1", Cn)
   Asignar
    
   ssTab.Tab = 0
   Screen.MousePointer = vbDefault
End Sub

Public Sub Asignar()
    AgregaArbol
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGrupoAccesoDetalle = Nothing
End Sub

Private Sub tvAdministracion_NodeCheck(ByVal Node As MSComctlLib.Node)
   Node.Checked = True
End Sub

Public Sub AgregaArbol()
    Dim xNode As Node
    Dim yNode As Node
    tvCentralizada.Nodes.Clear
    Do While Not RsAcceso.EOF
            If Mid(RsAcceso!tCodigoAcceso, 4, 5) = "00000" Then
                 Set xNode = tvCentralizada.Nodes.Add(, , "M" & Mid(RsAcceso!tCodigoAcceso, 2, 2), RsAcceso!tDescripcion)
                 xNode.Checked = True
                 xNode.Expanded = True
            ElseIf Mid(RsAcceso!tCodigoAcceso, 7, 2) = "00" Then
                 Set xNode = tvCentralizada.Nodes.Add("M" & Mid(RsAcceso!tCodigoAcceso, 2, 2), tvwChild, "M" & Mid(RsAcceso!tCodigoAcceso, 2, 5), RsAcceso!tDescripcion)
                 xNode.Checked = True
                 xNode.Expanded = True
            Else
                 Set xNode = tvCentralizada.Nodes.Add("M" & Mid(RsAcceso!tCodigoAcceso, 2, 5), tvwChild, "M" & Mid(RsAcceso!tCodigoAcceso, 2, 7), RsAcceso!tDescripcion)
                 xNode.Checked = True
            End If
            If RsAcceso.Fields("lenvia") Then
                xNode.Checked = True
            Else
                xNode.Checked = False
            End If
       RsAcceso.MoveNext
    Loop
End Sub

Public Sub GrabaArbol()
    Cn.Execute "DELETE FROM TACCESOENVIA"
    Dim i As Integer
    For i = 1 To tvCentralizada.Nodes.Count
    
        If tvCentralizada.Nodes.Item(i).Checked = True Then
            sCodigo = Left("1" & Mid(tvCentralizada.Nodes(i).Key, 2) + "0000000", 8)
             If Right(sCodigo, 5) <> "00000" Then
            Cn.Execute "INSERT INTO TACCESOENVIA VALUES ('" & sCodigo & "',1)"
            End If
        End If
    Next i

End Sub

Private Sub tvCentralizada_Click()
If lOferta = True Then
    MsgBox "Para activar la actualización de ofertas" & vbCrLf & "debe activar antes la actualización de los productos", vbInformation, sMensaje
    tvCentralizada.Nodes.Item(30).Checked = False
    End If
End Sub

Private Sub tvCentralizada_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim ind As String
    Dim chk As Boolean
    'chk = Node.Checked
        If Len(Node.Key) = 3 Then
            ind = Node.Key
            chk = Node.Checked
            activaNodosHijos ind, chk
        End If
'        MsgBox Node.Index
'        MsgBox Node.Key
'
        'grupo de usuario y usuario
        If Node.Key = "M01750" Or Node.Key = "M01800" Then
            tvCentralizada.Nodes.Item(3).Checked = Node.Checked
            tvCentralizada.Nodes.Item(4).Checked = Node.Checked
        End If
        
        'todo menu productos
        If Node.Key = "M03010" Or Node.Key = "M03015" Or Node.Key = "M03020" Or Node.Key = "M03030" Or Node.Key = "M03040" Or Node.Key = "M03050" Or Node.Key = "M03055" Or Node.Key = "M03080" Then
            tvCentralizada.Nodes.Item(23).Checked = Node.Checked
            tvCentralizada.Nodes.Item(24).Checked = Node.Checked
            tvCentralizada.Nodes.Item(25).Checked = Node.Checked
            tvCentralizada.Nodes.Item(26).Checked = Node.Checked
            tvCentralizada.Nodes.Item(27).Checked = Node.Checked
            tvCentralizada.Nodes.Item(28).Checked = Node.Checked
            tvCentralizada.Nodes.Item(29).Checked = Node.Checked
            tvCentralizada.Nodes.Item(31).Checked = Node.Checked
            
            If tvCentralizada.Nodes.Item(27).Checked = False Then
                tvCentralizada.Nodes.Item(30).Checked = False
            End If
        End If
        
        
        'oferta... valida si esta activo los productos
        lOferta = False
        If Node.Key = "M03060" Then
            If tvCentralizada.Nodes.Item(27).Checked = False Then
                    lOferta = True
            End If
        End If
    
End Sub

Private Sub activaNodosHijos(padre As String, Activa As Boolean)
    On Error Resume Next
    Dim pareNode As String
    pareNode = ""
    Dim i As Integer
    For i = 1 To tvCentralizada.Nodes.Count
        pareNode = tvCentralizada.Nodes.Item(i).Parent.Key
        If pareNode <> "" Then
                If tvCentralizada.Nodes.Item(i).Parent.Key = padre Then
                            tvCentralizada.Nodes.Item(i).Checked = Activa
                End If
        End If
        pareNode = ""
    Next i
End Sub

