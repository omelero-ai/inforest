VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTipoProductoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2925
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9510
   Icon            =   "frmTipoProductoDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   9510
   Begin VB.Frame fraDetalle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   1800
      TabIndex        =   17
      Top             =   -30
      Width           =   7665
      Begin VB.TextBox txtDetallado 
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
         Left            =   1980
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   641
         Width           =   5490
      End
      Begin VB.TextBox txtResumido 
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
         Left            =   1980
         MaxLength       =   15
         TabIndex        =   2
         Text            =   " "
         Top             =   982
         Width           =   2595
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   300
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   1368
         Width           =   840
      End
      Begin MSDataListLib.DataCombo cboTipoProducto 
         Height          =   315
         Left            =   1980
         TabIndex        =   4
         Top             =   1650
         Visible         =   0   'False
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
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
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Producto Infhotel :"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   22
         Top             =   1710
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   686
         Width           =   1650
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   19
         Top             =   1027
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   1305
         TabIndex        =   18
         Top             =   345
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9450
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2175
      Width           =   9510
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
         Left            =   8265
         Picture         =   "frmTipoProductoDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
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
         Index           =   2
         Left            =   7095
         Picture         =   "frmTipoProductoDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
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
         Left            =   4755
         Picture         =   "frmTipoProductoDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   5925
         Picture         =   "frmTipoProductoDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4635
         TabIndex        =   10
         Top             =   60
         Width           =   4695
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTipoProductoDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTipoProductoDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTipoProductoDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4155
            Picture         =   "frmTipoProductoDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3675
            Picture         =   "frmTipoProductoDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3195
            Picture         =   "frmTipoProductoDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
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
            Left            =   1500
            TabIndex        =   21
            Top             =   180
            Width           =   1665
         End
      End
   End
   Begin VB.Image Image 
      Height          =   2070
      Left            =   15
      Picture         =   "frmTipoProductoDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1755
   End
End
Attribute VB_Name = "frmTipoProductoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTipoProducto As Recordset

Sub Asignar()
    With frmTipoProducto.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
           
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
        
        cboTipoProducto.BoundText = IIf(IsNull(!tValor), "", Trim(!tValor))
    End With
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTipoProducto.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTipoProducto.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTipoProducto.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTipoProducto.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTipoProducto.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTipoProducto.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmTipoProducto.RsCabecera.RecordCount = 0, 0, frmTipoProducto.RsCabecera.AbsolutePosition) & " de " & frmTipoProducto.RsCabecera.RecordCount
       
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.value = 1
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
                                        
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                    
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='TIPOPRODUCTO' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, tValor, lActivo) " & _
                         "values ('TIPOPRODUCTO', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                " '" & cboTipoProducto.BoundText & "', " & _
                                       chkActivo.value & ") "
                  
                  Cn.Execute Isql
                  frmTipoProducto.RsCabecera.Sort = "Codigo ASC"
                  frmTipoProducto.RsCabecera.Requery
                  frmTipoProducto.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmTipoProducto.RsCabecera.RecordCount = 0, 0, frmTipoProducto.RsCabecera.AbsolutePosition) & " de " & frmTipoProducto.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "tValor ='" & cboTipoProducto.BoundText & "', " & _
                         "lActivo =" & chkActivo.value & ", lReplica=1 " & _
                         " where tTAbla = 'TIPOPRODUCTO' and tCodigo = '" & txtCodigo & "'"
                  Cn.Execute Isql
                  nPos = frmTipoProducto.RsCabecera.Bookmark
                  frmTipoProducto.RsCabecera.Requery
                  If frmTipoProducto.RsCabecera.RecordCount = 0 Then
                     frmTipoProducto.RsCabecera.Filter = adFilterNone
                  End If
                  frmTipoProducto.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
            Cn.Execute "update TTABLA set lNuevoModificado=1 where tTAbla = 'TIPOPRODUCTO' and tCodigo = '" & txtCodigo & "'"
          Case Is = 2 ' Eliminar
               If frmTipoProducto.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Tipo de Servicio " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'TIPOPRODUCTO' and tCodigo = '" & txtCodigo & "'"
               frmTipoProducto.RsCabecera.Requery
               If frmTipoProducto.RsCabecera.RecordCount <> 0 Then
                  frmTipoProducto.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmTipoProducto.RsCabecera.RecordCount = 0, 0, frmTipoProducto.RsCabecera.AbsolutePosition) & " de " & frmTipoProducto.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    
    If lInfhotel Then
       Dim sInfhotelRuta As String
       Dim sInfhotelMDB As String
       Dim sCajaInfhotel As String
      
       sInfhotelRuta = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Conexion", "SERVIDOR", "."))
       sInfhotelMDB = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Conexion", "BASEDATO", "INFHOTEL"))
       sCajaInfhotel = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Configuracion", "CAJA", "01"))
       sPuntoVentaInfhotel = Trim(LeerIni(App.Path + "\INFHOTEL.INI", "Configuracion", "PUNTOVENTA", "01"))
       
       Set CnInfhotel = New Connection
       CnInfhotel.Provider = "SQLOLEDB"
       CnInfhotel.CursorLocation = adUseServer
       CnInfhotel.ConnectionString = "User ID=" & sUserName & _
                                     ";password=" & sUserPassword & _
                                     ";Data Source=" & sInfhotelRuta & _
                                     ";Initial Catalog=" & sInfhotelMDB
       CnInfhotel.Open
       sHotel = Calcular("select tHotel as Codigo from vCaja where tCaja='" & sCajaInfhotel & "'", CnInfhotel)
       sHotel = IIf(sHotel = "0", "01", sHotel)
       cboTipoProducto.Visible = True
       Label(3).Visible = True
    End If
    
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Tipos de Servicio "
    fraDetalle.Caption = Me.Caption
    LlenaCombo
    
    'Ingresar la Tabla
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmTipoProducto.RsCabecera.RecordCount = 0, 0, frmTipoProducto.RsCabecera.AbsolutePosition) & " de " & frmTipoProducto.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set RsTipoProducto = Nothing
    Set frmTipoProductoDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    cmdOpcion(0).Enabled = Activa
    cmdOpcion(2).Enabled = Activa
End Sub

Public Sub LlenaCombo()
    If lInfhotel Then
        With cboTipoProducto
             Isql = "select tTipoProducto as Codigo, tDescripcionResumidaTipo as Descripcion from vTipoProducto where lActivo=1 and tHotel='" & sHotel & "'"
             Set RsTipoProducto = Lib.OpenRecordset(Isql, CnInfhotel)
             Set .RowSource = RsTipoProducto
                 .DataField = "Descripcion"
                 .ListField = "Descripcion"
                 .BoundColumn = "Codigo"
        End With
    End If
End Sub

