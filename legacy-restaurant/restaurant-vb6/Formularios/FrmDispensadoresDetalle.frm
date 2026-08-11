VERSION 5.00
Begin VB.Form FrmDispensadoresDetalle 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9420
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2445
      Width           =   9480
      Begin VB.PictureBox PicNavegacion 
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   16
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "FrmDispensadoresDetalle.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "FrmDispensadoresDetalle.frx":0542
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "FrmDispensadoresDetalle.frx":0A84
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "FrmDispensadoresDetalle.frx":0FC6
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "FrmDispensadoresDetalle.frx":1508
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "FrmDispensadoresDetalle.frx":1A4A
            Style           =   1  'Graphical
            TabIndex        =   17
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
            Left            =   1440
            TabIndex        =   23
            Top             =   150
            Width           =   1665
         End
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
         Left            =   5880
         Picture         =   "FrmDispensadoresDetalle.frx":1F8C
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   4710
         Picture         =   "FrmDispensadoresDetalle.frx":24BE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   1170
      End
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
         Left            =   7080
         Picture         =   "FrmDispensadoresDetalle.frx":29F0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Frame fraDetalle 
      Height          =   2415
      Left            =   2250
      TabIndex        =   0
      Top             =   0
      Width           =   7220
      Begin VB.CommandButton cmdBuscaTarjeta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         Picture         =   "FrmDispensadoresDetalle.frx":2AE2
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   960
         Width           =   630
      End
      Begin VB.TextBox txtCodExterno 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtTopeMax 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   255
         Left            =   1000
         TabIndex        =   4
         Top             =   2040
         Width           =   885
      End
      Begin VB.TextBox txtProducto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtDetallado 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Externo :"
         Height          =   195
         Left            =   440
         TabIndex        =   11
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TopeMax :"
         Height          =   195
         Left            =   840
         TabIndex        =   9
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Producto :"
         Height          =   195
         Left            =   320
         TabIndex        =   8
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dispensador :"
         Height          =   195
         Left            =   610
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   1000
         TabIndex        =   5
         Top             =   250
         Width           =   585
      End
   End
   Begin VB.Image Image 
      Height          =   2415
      Left            =   0
      Picture         =   "FrmDispensadoresDetalle.frx":2BE4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2205
   End
End
Attribute VB_Name = "FrmDispensadoresDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscaTarjeta_Click()
On Error GoTo fin
 
   'If IIf(chkCover.value, 1, 0) = 1 Then
       Isql = "select codigo, descripcion  from vproducto where lactivo=1 AND Codigo NOT IN  (SELECT CodigoProducto  FROM TDISPENSADOR ) "
     
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 1, "Codigo", 1300, 2, 0, "", _
                                                    "Descripción", 2, "Descripcion", 4800, 0, 0, "")
       frmBusquedaRapida.grdGrilla.Caption = "Solo Productos con Precio Venta = 1 "
       frmBusquedaRapida.Show vbModal
        

       If sCodigo <> "" Then
            txtProducto.Text = sCodigo
       End If

   ' End If

Exit Sub
fin:
MsgBox (error)
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
 Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, FrmDispensadores.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, FrmDispensadores.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, FrmDispensadores.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, FrmDispensadores.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, FrmDispensadores.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, FrmDispensadores.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & FrmDispensadores.RsCabecera.AbsolutePosition & " de " & FrmDispensadores.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
 Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.Value = 1
'               txtTiempo.Text = "0.00"
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
          
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese Nombre del Dispensador", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtProducto.Text = "" Then MsgBox "Ingrese Codigo de Producto", vbExclamation, sMensaje: txtProducto.SetFocus: Exit Sub
                    
               If Sw Then
                  'Obtiene el Numero de Orden
'                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='ZONA' ", Cn)
'                  If IsNull(nCorrela) Or nCorrela = "" Then
'                      txtCodigo.Text = "001"
'                  Else
'                      txtCodigo.Text = Lib.Correlativo(nCorrela, 3)
'                  End If
                  Sw = False
                   
                   
                   
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TDISPENSADOR", "DISPENSADORES", "01", sUsuario, sPasa, "", _
                   "TDISPENSADOR", "TDISPENSADOR", "DISPENSADORES", "id", "Codigo dispensador", txtCodigo.Text, "Descripcion", "Descripcion Detallada", txtDetallado.Text, "CodigoProducto", "Descripcion Producto", txtProducto.Text, "TopeMax", "TopeMax", txtTopeMax.Text, "CodigoExterno", "Codigo Externo", txtCodExterno.Text, "Estado", "Flag Activo", IIf(chkActivo = 1, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                   
                   
                   
                   
                  'Cambiar el SQL
                  Isql = "insert into TDISPENSADOR( " & _
                         "Descripcion, fRegistro, UsuarioRegistro, FechaModificion, UsuarioModificion, Estado, CodigoProducto, TopeMax, CodigoExterno) " & _
                         "values (" & _
                                " '" & txtDetallado.Text & "', " & _
                                " getdate(), " & _
                                " '" & sUsuario & "', " & _
                                " getdate(), " & _
                                " '" & sUsuario & "', " & _
                                " '" & chkActivo.Value & "', " & _
                                " '" & txtProducto.Text & "', " & _
                                " '" & txtTopeMax.Text & "', " & _
                                " '" & txtCodExterno.Text & "' " & ") "
           
                      Cn.Execute Isql
                      FrmDispensadores.RsCabecera.Sort = "id ASC"
                      FrmDispensadores.RsCabecera.Requery
                      FrmDispensadores.RsCabecera.MoveLast
                      MsgBox "Registro Guardado", vbInformation, sMensaje
                      ActivarBotones (True)
                      cmdTexto.Caption = "Registro " & IIf(FrmDispensadores.RsCabecera.RecordCount = 0, 0, FrmDispensadores.RsCabecera.AbsolutePosition) & " de " & FrmDispensadores.RsCabecera.RecordCount
              Else
              
               sPasa = txtCodigo.Text
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TDISPENSADOR", "DISPENSADORES", "02", sUsuario, sPasa, "", _
                   "id", "Codigo dispensador", txtCodigo.Text, "Descripcion", "Descripcion Detallada", txtDetallado.Text, "CodigoProducto", "Descripcion Producto", txtProducto.Text, "TopeMax", "TopeMax", txtTopeMax.Text, "CodigoExterno", "Codigo Externo", txtCodExterno.Text, "Estado", "Flag Activo", IIf(chkActivo = 1, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
               
               
                 'Cambiar el SQL
                 Isql = "update TDISPENSADOR set " & _
                        "Descripcion ='" & txtDetallado.Text & "', " & _
                        "CodigoProducto ='" & txtProducto.Text & "', " & _
                        "TopeMax =" & txtTopeMax.Text & ", " & _
                        "CodigoExterno ='" & txtCodExterno.Text & "', " & _
                        "UsuarioModificion = '" & sUsuario & "' ," & _
                        "FechaModificion = getdate() " & _
                        ", Estado =" & chkActivo.Value & "" & _
                        " where id = '" & txtCodigo & "'"
                      
                  Cn.Execute Isql
                  nPos = FrmDispensadores.RsCabecera.Bookmark
                  FrmDispensadores.RsCabecera.Requery
                  If FrmDispensadores.RsCabecera.RecordCount = 0 Then
                     FrmDispensadores.RsCabecera.Filter = adFilterNone
                  End If
                  FrmDispensadores.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
              End If
 
          Case Is = 3 ' Salir
               Unload Me
   End Select
End Sub
Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    cmdOpcion(0).Enabled = Activa
'    cmdOpcion(2).Enabled = Activa
End Sub
Private Sub Form_Load()
  Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Dispensadores "
    fraDetalle.Caption = Me.Caption
       
    'Ingresar la Tabla
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
      ' txtTiempo.Text = "0.00"
       chkActivo.Value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & FrmDispensadores.RsCabecera.AbsolutePosition & " de " & FrmDispensadores.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub
Sub Asignar()
    With FrmDispensadores.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!id), "", !id)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtProducto = IIf(IsNull(!codigoProducto), "", !codigoProducto)
        txtTopeMax = Format(IIf(IsNull(!topemax), "0", !topemax), "##0.00")
        txtCodExterno = Format(IIf(IsNull(!codigoexterno), "0", !codigoexterno), "##0.00")
        'Check Box
        chkActivo = IIf(!Estado = True, 1, 0)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmDispensadoresDetalle = Nothing
End Sub

