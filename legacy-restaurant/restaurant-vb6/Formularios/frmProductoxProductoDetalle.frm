VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmProductoxProductoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7335
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11175
   Icon            =   "frmProductoxProductoDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11175
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
      Height          =   1050
      Left            =   2310
      TabIndex        =   17
      Top             =   30
      Width           =   8835
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
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   255
         Width           =   1170
      End
      Begin VB.CommandButton cmdBusca 
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
         Left            =   6840
         Picture         =   "frmProductoxProductoDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   585
         Width           =   585
      End
      Begin VB.TextBox txtDetallado 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1035
         MaxLength       =   100
         TabIndex        =   0
         Text            =   " "
         Top             =   600
         Width           =   5715
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   375
         TabIndex        =   22
         Top             =   285
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Producto :"
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   630
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Equivalencias "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   2295
      TabIndex        =   20
      Top             =   1095
      Width           =   8865
      Begin VB.CommandButton cmdOpcionGrilla 
         Caption         =   "&Agregar"
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
         Left            =   7560
         Picture         =   "frmProductoxProductoDetalle.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcionGrilla 
         Caption         =   "&Eliminar"
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
         Left            =   7560
         Picture         =   "frmProductoxProductoDetalle.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1740
         Width           =   1170
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   5070
         Left            =   120
         TabIndex        =   21
         Top             =   270
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   8943
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
      Begin VB.CommandButton cmdOpcionGrilla 
         Caption         =   "&Modificar"
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
         Left            =   7560
         Picture         =   "frmProductoxProductoDetalle.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1005
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11115
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6585
      Width           =   11175
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
         Left            =   9930
         Picture         =   "frmProductoxProductoDetalle.frx":0C7A
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
         Left            =   8760
         Picture         =   "frmProductoxProductoDetalle.frx":0D6C
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
         Left            =   6420
         Picture         =   "frmProductoxProductoDetalle.frx":0E6E
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   7590
         Picture         =   "frmProductoxProductoDetalle.frx":13A0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6255
         TabIndex        =   10
         Top             =   60
         Width           =   6315
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmProductoxProductoDetalle.frx":18D2
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
            Picture         =   "frmProductoxProductoDetalle.frx":1E14
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
            Picture         =   "frmProductoxProductoDetalle.frx":2356
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5775
            Picture         =   "frmProductoxProductoDetalle.frx":2898
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5295
            Picture         =   "frmProductoxProductoDetalle.frx":2DDA
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4815
            Picture         =   "frmProductoxProductoDetalle.frx":331C
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
            Left            =   1440
            TabIndex        =   19
            Top             =   150
            Width           =   3240
         End
      End
   End
   Begin VB.Image Image 
      Height          =   6525
      Left            =   45
      Picture         =   "frmProductoxProductoDetalle.frx":385E
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2205
   End
End
Attribute VB_Name = "frmProductoxProductoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsProductoDetalle As Recordset
Public lAgrega As Boolean
Dim nColumna As Integer

Sub Asignar()
    With frmProductoxproducto.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!tCodigoProducto), "", !tCodigoProducto)
        txtDetallado = IIf(IsNull(!tDetallado), "", !tDetallado)
        RsProductoDetalle.Filter = "tCodigoProducto='" & Trim(txtCodigo) & "'"
    End With
End Sub

Private Sub cmdBusca_Click()
    Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto order by Descripcion"
    frmBusca.nPredeterm = 3
    Call ConfGrilla(5, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 800, 2, 0, "", _
                                           "Grupo", 2, "Grupo", 1300, 0, 0, "", _
                                           "SubGrupo", 2, "SubGrupo", 1400, 0, 0, "", _
                                           "Plato", 2, "Descripcion", 2900, 0, 0, "", _
                                           "Precio", 2, "nPrecioVenta", 900, 1, 0, "##,##0.00")
    frmBusca.Show vbModal
    If Not wEnter Then
       Exit Sub
    End If
    txtCodigo.Text = sCodigo
    txtDetallado.Text = sDescrip
    RsProductoDetalle.Filter = "tcodigoProducto='" & sCodigo & "'"
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmProductoxproducto.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmProductoxproducto.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmProductoxproducto.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmProductoxproducto.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmProductoxproducto.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmProductoxproducto.grdGrilla
    End Select
    Asignar
    cmdTexto.Caption = "Registro " & frmProductoxproducto.RsCabecera.AbsolutePosition & " de " & frmProductoxproducto.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
        Case Is = 0 ' Agregar
            Sw = True
            ActivarBotones (False)
            Blanquear Me
            
        
        Case Is = 1 ' Grabar
            Dim nCorrela As String
            Dim nPos     As Integer
            
            'Chequea Datos
            If txtCodigo.Text = "" Then MsgBox "Ingrese el Producto", vbExclamation, sMensaje: cmdBusca.SetFocus: Exit Sub
            If RsProductoDetalle.RecordCount > 0 Then
                If Sw Then
                    'Obtiene el Numero de Orden
                    Sw = False
                    
                    'Cambiar el SQL
                    frmProductoxproducto.RsCabecera.Sort = "tCodigoproducto ASC"
                    frmProductoxproducto.RsCabecera.Requery
                    If frmProductoxproducto.RsCabecera.RecordCount > 0 Then
                       frmProductoxproducto.RsCabecera.MoveLast
                    End If
                    MsgBox "Registro Guardado", vbInformation, sMensaje
                    ActivarBotones (True)
                    cmdTexto.Caption = "Registro " & IIf(frmProductoxproducto.RsCabecera.RecordCount = 0, 0, frmProductoxproducto.RsCabecera.AbsolutePosition) & " de " & frmProductoxproducto.RsCabecera.RecordCount
                Else
                    'Cambiar el SQL
                     nPos = frmProductoxproducto.RsCabecera.AbsolutePosition
                     frmProductoxproducto.RsCabecera.Requery
                     If frmProductoxproducto.RsCabecera.RecordCount > 0 Then
                        frmProductoxproducto.RsCabecera.AbsolutePosition = nPos
                     End If
                     MsgBox "Registro Modificado", vbInformation, sMensaje
                End If
            Else
                MsgBox "Ingrese detalle de productos ", vbExclamation, sMensaje
                'cmdOpcionGrilla(0).SetFocus
                Exit Sub
            End If
            Asignar
        
        Case Is = 2 ' Eliminar
            If frmProductoxproducto.RsCabecera.RecordCount = 0 Then
                Exit Sub
            End If
            'Cambia el MsgBox
            If MsgBox("Seguro de eliminar el producto " & txtDetallado & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                Exit Sub
            End If
            
            'Cambia el Delete
            Cn.Execute "delete from TPRODUCTOXPRODUCTO where tCodigoProducto = '" & txtCodigo & "'"
            frmProductoxproducto.RsCabecera.Requery
            If frmProductoxproducto.RsCabecera.RecordCount <> 0 Then
                frmProductoxproducto.RsCabecera.MoveLast
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmProductoxproducto.RsCabecera.RecordCount = 0, 0, frmProductoxproducto.RsCabecera.AbsolutePosition) & " de " & frmProductoxproducto.RsCabecera.RecordCount
            Else
                ActivarBotones False
                Blanquear Me
                Sw = True
                
            End If
        
        Case Is = 3 ' Salir
            Unload Me
    End Select
End Sub

Private Sub cmdOpcionGrilla_Click(Index As Integer)
'Dim RsTemporal As New ADODB.Recordset

    Select Case Index
        Case Is = 0 ' Agregar
             If txtCodigo = "" Then MsgBox "Seleccione el producto principal", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
             sVar1 = txtCodigo.Text
             lAgrega = True
             frmProductoxProductoCantidad.Show vbModal
                        
        Case Is = 1 ' Modificar
             If RsProductoDetalle.RecordCount = 0 Then
                Exit Sub
             End If
             sVar1 = txtCodigo.Text
             lAgrega = False
             frmProductoxProductoCantidad.Show vbModal
                        
        Case Is = 2 ' Eliminar
            If txtCodigo = "" Then MsgBox "Seleccione el producto principal", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
            If RsProductoDetalle.RecordCount = 0 Then
                Exit Sub
            End If
            
            'Cambia el MsgBox
            If MsgBox("Seguro de eliminar la equivalencia " & RsProductoDetalle!producto & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                Exit Sub
            End If
            
            'Cambia el Delete
            Cn.Execute "Delete from TPRODUCTOXPRODUCTO where tCodigoProducto='" & txtCodigo & "' and tsubProducto = '" & RsProductoDetalle!tSubProducto & "'"
            frmProductoxProductoDetalle.RsProductoDetalle.Requery
                
        
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    
    Me.Caption = " Mantenimiento de Equivalencias de Productos"
    fraDetalle.Caption = Me.Caption
    
    Isql = "SELECT tCodigoProducto, tSubProducto, Producto, nCantidad FROM vProductoxProducto"
    Set RsProductoDetalle = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(3, grdGrilla, "Codigo", 2, "tSubProducto", 800, 2, 0, "", _
                                  "Equivalencia", 2, "Producto", 4950, 0, 0, "", _
                                  "Cantidad", 2, "nCantidad", 1000, 1, 0, "##,##0.000")

    'Ingresar la Tabla
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       RsProductoDetalle.Filter = "tcodigoProducto=''"
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    Set grdGrilla.DataSource = RsProductoDetalle
    cmdTexto.Caption = "Registro " & frmProductoxproducto.RsCabecera.AbsolutePosition & " de " & frmProductoxproducto.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmProductoxProductoDetalle = Nothing
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

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    nColumna = ColIndex
    RsProductoDetalle.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
End Sub

Private Sub txtDetallado_Change()
  If Len(Trim(txtDetallado.Text)) = 1 Then
     Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto order by Descripcion"
     frmBusca.txtCriterio = Trim(txtDetallado.Text)
     frmBusca.txtCriterio.SelStart = 2
     Call cmdBusca_Click
  End If
End Sub

Private Sub txtDetallado_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub
