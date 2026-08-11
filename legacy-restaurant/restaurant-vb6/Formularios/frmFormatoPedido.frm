VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmFormatoPedido 
   Caption         =   "Formato de Orden de Producción (Cabecera)"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
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
      Left            =   8325
      Picture         =   "frmFormatoPedido.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7605
      Width           =   1170
   End
   Begin VB.CommandButton cmdGrabar 
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
      Left            =   7065
      Picture         =   "frmFormatoPedido.frx":00F2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7605
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Caption         =   " Cabecera de la Orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7440
      Left            =   2880
      TabIndex        =   3
      Top             =   90
      Width           =   6630
      Begin VB.CommandButton Command4 
         Caption         =   "Borrar última línea"
         Height          =   555
         Left            =   1395
         TabIndex        =   5
         Top             =   6795
         Width           =   1140
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Borrar Todo"
         Height          =   555
         Left            =   135
         TabIndex        =   4
         Top             =   6795
         Width           =   1140
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   6390
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   11271
         _LayoutType     =   4
         _RowHeight      =   18
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
   End
   Begin VB.Frame Frame1 
      Caption         =   " Opciones de Impresión "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7440
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2715
      Begin VB.CommandButton cmdPAX 
         Caption         =   "Cantidad de Personas"
         Height          =   330
         Left            =   135
         TabIndex        =   24
         Top             =   6135
         Width           =   2400
      End
      Begin VB.CommandButton cmdRC 
         Caption         =   "Razon Comercial"
         Height          =   330
         Left            =   135
         TabIndex        =   23
         Top             =   661
         Width           =   2400
      End
      Begin VB.CommandButton cmdCli 
         Caption         =   "Cliente (Delivery)"
         Height          =   330
         Left            =   135
         TabIndex        =   22
         Top             =   5744
         Width           =   2400
      End
      Begin VB.CommandButton cmdObs 
         Caption         =   "Observación de Pedido"
         Height          =   330
         Left            =   135
         TabIndex        =   21
         Top             =   5353
         Width           =   2400
      End
      Begin VB.CommandButton cmdFecha 
         Caption         =   "Fecha y Hora"
         Height          =   330
         Left            =   135
         TabIndex        =   20
         Top             =   4962
         Width           =   2400
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "Número de Pedido"
         Height          =   330
         Left            =   135
         TabIndex        =   19
         Top             =   4571
         Width           =   2400
      End
      Begin VB.CommandButton cmdCanal 
         Caption         =   "Canal de Venta"
         Height          =   330
         Left            =   135
         TabIndex        =   18
         Top             =   4180
         Width           =   2400
      End
      Begin VB.CommandButton cmdArea 
         Caption         =   "Area de Producción"
         Height          =   330
         Left            =   135
         TabIndex        =   17
         Top             =   3789
         Width           =   2400
      End
      Begin VB.CommandButton cmdMozo 
         Caption         =   "Mozo / Motorizado"
         Height          =   330
         Left            =   135
         TabIndex        =   16
         Top             =   3398
         Width           =   2400
      End
      Begin VB.CommandButton cmdMesa 
         Caption         =   "Salon / Mesa (En el Local)"
         Height          =   330
         Left            =   135
         TabIndex        =   15
         Top             =   3007
         Width           =   2400
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "Correo Electrónico"
         Height          =   330
         Left            =   135
         TabIndex        =   14
         Top             =   2616
         Width           =   2400
      End
      Begin VB.CommandButton cmdWeb 
         Caption         =   "Página Web"
         Height          =   330
         Left            =   135
         TabIndex        =   13
         Top             =   2225
         Width           =   2400
      End
      Begin VB.CommandButton cmdRUC 
         Caption         =   "Identificación Tributaria"
         Height          =   330
         Left            =   135
         TabIndex        =   12
         Top             =   1834
         Width           =   2400
      End
      Begin VB.CommandButton cmdTel 
         Caption         =   "Teléfono (s)"
         Height          =   330
         Left            =   135
         TabIndex        =   11
         Top             =   1443
         Width           =   2400
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "Dirección"
         Height          =   330
         Left            =   135
         TabIndex        =   10
         Top             =   1052
         Width           =   2400
      End
      Begin VB.CommandButton cmdRS 
         Caption         =   "Razon Social"
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   270
         Width           =   2400
      End
      Begin VB.CommandButton cmdLinea 
         Caption         =   "Líneas de Separación"
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   6930
         Width           =   2400
      End
      Begin VB.CommandButton cmdEspacio 
         Caption         =   "Espcacio en Blanco"
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   6526
         Width           =   2400
      End
   End
End
Attribute VB_Name = "frmFormatoPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sTempral As String
Dim i As Integer
Dim RsCabecera As ADODB.Recordset
Dim RsGrilla As ADODB.Recordset

Private Sub cmdArea_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'ARE','Area  : Cocina' )"
   RsGrilla.Requery
End Sub

Private Sub cmdCanal_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'TIP','T/Pedido : Salon' )"
   RsGrilla.Requery
End Sub

Private Sub cmdCli_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'CLI','Cliente: Jose Luis Ramos' )"
   RsGrilla.Requery
End Sub

Private Sub cmdDir_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'DIR','" & sDireccion & "')"
   RsGrilla.Requery
End Sub

Private Sub cmdEmail_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'EMA','" & sMail & "' )"
   RsGrilla.Requery
End Sub

Private Sub cmdEspacio_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'ESP','(espacio en blanco)' )"
   RsGrilla.Requery
End Sub

Private Sub cmdFecha_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'FEC','Fecha   :02/04/2007 10:00am' )"
   RsGrilla.Requery
End Sub

Private Sub cmdGrabar_Click()
   Cn.Execute "delete from TFORMATOPEDIDO"
   Cn.Execute "Insert into TFORMATOPEDIDO (nItem, tTexto) select NItem, tTexto from " & sTemporal
   MsgBox "Formato Guardado", vbInformation, sMensaje
End Sub

Private Sub cmdLinea_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'LIN','(linea)' )"
   RsGrilla.Requery
End Sub

Private Sub cmdMesa_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'MES','Mesa    : Salon 1 - Mesa 04' )"
   RsGrilla.Requery
End Sub

Private Sub cmdMozo_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'MOZ','Mozo    : Jose Luis' )"
   RsGrilla.Requery
End Sub

Private Sub cmdNum_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'PED','Pedido  : 0800001452' )"
   RsGrilla.Requery
End Sub

Private Sub cmdObs_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'OBS','Observ. : Cliente Importante' )"
   RsGrilla.Requery
End Sub

Private Sub cmdPAX_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'PAX','Personas: 2' )"
   RsGrilla.Requery
End Sub

Private Sub cmdRC_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'RCO','" & sRazonComercial & "' )"
   RsGrilla.Requery
End Sub

Private Sub cmdRS_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'RSO','" & sRazonSocial & "' )"
   RsGrilla.Requery
End Sub

Private Sub cmdRUC_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'RUC','" & sRUC & "' )"
   RsGrilla.Requery
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdTel_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'TEL','" & sTelefono & "' )"
   RsGrilla.Requery
End Sub

Private Sub cmdWeb_Click()
   i = i + 1
   Cn.Execute "Insert into " & sTemporal & " values (" & i & ", 'WEB','" & sWeb & "' )"
   RsGrilla.Requery
End Sub

Private Sub Command3_Click()
   Cn.Execute "delete from " & sTemporal
   RsGrilla.Requery
   i = 0
End Sub

Private Sub Command4_Click()
   Cn.Execute "delete from " & sTemporal & " where nItem=" & i
   i = i - 1
   RsGrilla.Requery
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
        
    'Ingresar la Vista de la Grilla
    sTemporal = dbTemporal(sCaja, 3, "nItem", "int", "tTexto", "nvarchar(50)", "tDescrip", "nvarchar(200)")
    Set RsCabecera = Lib.OpenRecordset("select * from TFORMATOPEDIDO order by nItem", Cn)
    Isql = "select * from " & sTemporal & " order by nItem"
    Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
    If RsCabecera.RecordCount > 0 Then
       Do While Not RsCabecera.EOF
          InsertaDetalle RsCabecera!tTexto
          RsCabecera.MoveNext
       Loop
    End If
    Call ConfGrilla(2, grdGrilla, "linea", 2, "nItem", 750, 1, 0, "", _
                                  "Campo (Impresión)", 2, "tDescrip", 5000, 2, 0, "")
    Set grdGrilla.DataSource = RsGrilla
    
    i = RsGrilla.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmFormatoPedido = Nothing
   Set RsGrilla = Nothing
   Set RsCabecera = Nothing
End Sub

Public Sub InsertaDetalle(sClave As String)
   Select Case sClave
          Case Is = "ARE"
               cmdArea_Click
          Case Is = "TIP"
               cmdCanal_Click
          Case Is = "CLI"
               cmdCli_Click
          Case Is = "DIR"
               cmdDir_Click
          Case Is = "EMA"
               cmdEmail_Click
          Case Is = "ESP"
               cmdEspacio_Click
          Case Is = "FEC"
               cmdFecha_Click
          Case Is = "MES"
               cmdMesa_Click
          Case Is = "LIN"
               cmdLinea_Click
          Case Is = "MOZ"
               cmdMozo_Click
          Case Is = "NUM"
               cmdNum_Click
          Case Is = "OBS"
               cmdObs_Click
          Case Is = "RCO"
               cmdRC_Click
          Case Is = "RSO"
               cmdRS_Click
          Case Is = "RUC"
               cmdRUC_Click
          Case Is = "TEL"
               cmdTel_Click
          Case Is = "WEB"
               cmdWeb_Click
          Case Is = "PAX"
               cmdPAX_Click
   End Select
End Sub




