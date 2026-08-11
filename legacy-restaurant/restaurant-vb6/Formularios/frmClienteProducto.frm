VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmClienteProducto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7620
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   13050
   Icon            =   "frmClienteProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   13050
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   12990
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6870
      Width           =   13050
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
         Left            =   7515
         Picture         =   "frmClienteProducto.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   0
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
         Left            =   8685
         Picture         =   "frmClienteProducto.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Left            =   9855
         Picture         =   "frmClienteProducto.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   1170
      End
   End
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
      Height          =   6840
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   12330
      Begin VB.Frame Frame1 
         Caption         =   "Seleccionar"
         Height          =   1215
         Left            =   2640
         TabIndex        =   16
         Top             =   960
         Width           =   8295
         Begin VB.CheckBox chkExcluyente 
            Alignment       =   1  'Right Justify
            Caption         =   "Permite Otros Descuentos :"
            Height          =   195
            Left            =   2280
            TabIndex        =   7
            Top             =   840
            Width           =   2250
         End
         Begin VB.TextBox txtPrecio 
            Alignment       =   1  'Right Justify
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
            Left            =   960
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   840
            Width           =   1020
         End
         Begin VB.CommandButton cmdBuscaProducto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7320
            Picture         =   "frmClienteProducto.frx":0B68
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Precio :"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   19
            Top             =   840
            Width           =   540
         End
         Begin VB.Label txtProducto 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   960
            TabIndex        =   18
            Top             =   360
            Width           =   6240
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Producto :"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   6120
         Width           =   1680
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Quitar Producto"
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
         Index           =   9
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6120
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar Producto"
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
         Index           =   6
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1320
         Width           =   1170
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   1170
      End
      Begin VB.TextBox txtNombre 
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
         Height          =   285
         Left            =   2685
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   570
         Width           =   9375
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   3765
         Left            =   2640
         TabIndex        =   12
         Top             =   2280
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   6641
         _LayoutType     =   4
         _RowHeight      =   23
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "True/False"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   2
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   2
         Splits(0).DividerColor=   32768
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
         PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
         _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(9)   =   ":id=4,.fontname=Arial"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
         _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
         _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
         _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
         _StyleDefs(43)  =   "Named:id=34:Heading"
         _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=34,.fontname=Arial"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(53)  =   ":id=36,.fontname=Arial"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(56)  =   "Named:id=38:HighlightRow"
         _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(60)  =   "Named:id=39:EvenRow"
         _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(62)  =   "Named:id=40:OddRow"
         _StyleDefs(63)  =   ":id=40,.parent=33"
         _StyleDefs(64)  =   "Named:id=41:RecordSelector"
         _StyleDefs(65)  =   ":id=41,.parent=34"
         _StyleDefs(66)  =   "Named:id=42:FilterBar"
         _StyleDefs(67)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Productos :"
         Height          =   195
         Left            =   1740
         TabIndex        =   21
         Top             =   960
         Width           =   810
      End
      Begin VB.Image Image 
         Height          =   6645
         Left            =   0
         Picture         =   "frmClienteProducto.frx":0C6A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1710
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total de Productos: "
         Height          =   195
         Index           =   0
         Left            =   2640
         TabIndex        =   20
         Top             =   6120
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   1980
         TabIndex        =   14
         Top             =   615
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         Height          =   195
         Left            =   1965
         TabIndex        =   13
         Top             =   255
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmClienteProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim RsProducto As Recordset
  
Dim sTemporal As String
 Dim nColumna As Integer
Dim nPrecio As Double
 
Dim i As Integer

Sub Asignar()
       'Grilla
       Cn.Execute "delete from " & sTemporal
       Cn.Execute "insert into " & sTemporal & " (tcodigodelivery, tCodigoProducto,nprecio,lPermiteDescuento) select tcodigodelivery, tCodigoProducto,nPrecio,lPermiteDescuentos from tclienteproducto where tcodigodelivery='" & frmClienteDelivery.RsCabecera.Fields(0) & "'"
       RsProducto.Requery
       txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
End Sub

Private Sub CmdBuscaProducto_Click()
               Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where lActivo=1 and Codigo not in (select tCodigoProducto from " & sTemporal & " ) order by Descripcion"
               frmBusca.nPredeterm = 3
               Call ConfGrilla(5, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 800, 2, 0, "", _
                                                      "Grupo", 2, "Grupo", 1300, 0, 0, "", _
                                                      "SubGrupo", 2, "SubGrupo", 1400, 0, 0, "", _
                                                      "Plato", 2, "Descripcion", 2900, 0, 0, "", _
                                                      "Precio", 2, "nPrecioVenta", 900, 1, 0, "##,##0.00")
               frmBusca.Show vbModal
               If Not wEnter Then
                    Me.txtProducto.Caption = ""
                  Exit Sub
               End If
               Me.txtProducto.Caption = sDescrip
               Me.txtPrecio.SetFocus
End Sub

Public Sub ObtieneCodigo(X As String)
    txtCodigo.Text = X
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index

          Case Is = 1 ' Grabar
 
              If RsProducto.RecordCount = 0 Then MsgBox "Debe seleccionar un producto", vbExclamation, sMensaje: cmdBuscaProducto.SetFocus: Exit Sub
 
                  Cn.Execute "delete from tclienteproducto where tCodigoDelivery='" & txtCodigo & "'"
                  Cn.Execute "insert into tclienteproducto (tCodigoDelivery, tCodigoProducto, nPRECIO,lPermiteDescuentos,tUsuario, Fregistro) " & _
                             "select '" & txtCodigo.Text & "', tCodigoProducto, nPRECIO,LpERMITEDESCUENTO, '" & sUsuario & "', getdate() from " & sTemporal
                  RsProducto.Requery
                  txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")

                  Screen.MousePointer = vbDefault


                         
          Case Is = 2 ' Eliminar
              If RsProducto.RecordCount = 0 Then
                  Exit Sub
              End If
              'Cambia el Delete
              If MsgBox("¿Confirma que desea Eliminar Todos Los Productos relacionados a este Cliente?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                Cn.Execute "delete from tclienteproducto where tcodigodelivery = '" & txtCodigo & "'"
                Unload Me
              End If
          Case Is = 3 ' Salir
               Unload Me
                                       
          Case Is = 6 ' Agregar Producto
                If sCodigo <> "" And Me.txtProducto.Caption <> "" And Val(txtPrecio.Text) > 0 Then
                        If Calcular("select count(*) as codigo from " & sTemporal & " where tcodigoproducto='" & sCodigo & "'", Cn) = 0 Then
                                Cn.Execute "insert into " & sTemporal & " (tcodigoDelivery,tCodigoProducto,nPrecio,lPermiteDescuento) values ('" & Me.txtCodigo.Text & "','" & sCodigo & "'," & Val(Me.txtPrecio.Text) & "," & IIf(Me.chkExcluyente.value, 1, 0) & ")"
                                RsProducto.Requery
                                txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
                        Else
                                MsgBox "Ya existe el Producto seleccionado", vbCritical, sMensaje
                        
                        End If
                        txtProducto.Caption = ""
                        txtPrecio.Text = ""
                        chkExcluyente.value = 0
                        cmdBuscaProducto.SetFocus
                End If
                        
          Case Is = 9 ' Elimina Producto
              If RsProducto.RecordCount = 0 Then
                  Exit Sub
              End If
              Cn.Execute "delete " & sTemporal & " where tCodigoProducto ='" & RsProducto!tCodigoProducto & "'"
              RsProducto.Requery
              txtTotal.Text = Format(RsProducto.RecordCount, "###,##0")
    End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
'
'    'Ingrese el SubTitulo
     Me.Caption = " Productos por Cliente Frecuentes "
     fraDetalle.Caption = Me.Caption
'
  
    ActivarBotones (True)
      Asignar
 
   
     Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
 
    Set RsProducto = Nothing
 
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdOpcion(2).Enabled = Activa
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    nColumna = ColIndex
    RsProducto.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
End Sub
 
Private Sub txtPrecio_GotFocus()
   Enfoque txtPrecio
End Sub

Private Sub txtPrecio_LostFocus()
   nPrecio = Val(txtPrecio.Text)
   txtPrecio.Text = Format(nPrecio, "###0.00")
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtPrecio
End Sub
