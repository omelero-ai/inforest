VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmBuscaAlmacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   Icon            =   "frmBuscaAlmacen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000004&
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   7905
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5670
      Width           =   7965
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
         Left            =   6645
         Picture         =   "frmBuscaAlmacen.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1245
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
         Height          =   525
         Index           =   1
         Left            =   5400
         Picture         =   "frmBuscaAlmacen.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   525
         Index           =   3
         Left            =   1710
         Picture         =   "frmBuscaAlmacen.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   525
         Index           =   4
         Left            =   2280
         Picture         =   "frmBuscaAlmacen.frx":0B88
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   525
         Index           =   5
         Left            =   2850
         Picture         =   "frmBuscaAlmacen.frx":10CA
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   525
         Index           =   0
         Left            =   0
         Picture         =   "frmBuscaAlmacen.frx":160C
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   525
         Index           =   2
         Left            =   1140
         Picture         =   "frmBuscaAlmacen.frx":1B4E
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   525
         Index           =   1
         Left            =   570
         Picture         =   "frmBuscaAlmacen.frx":2090
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   570
      End
   End
   Begin VB.ComboBox cboCriterio 
      Height          =   315
      ItemData        =   "frmBuscaAlmacen.frx":25D2
      Left            =   30
      List            =   "frmBuscaAlmacen.frx":25D4
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   30
      Width           =   2355
   End
   Begin VB.TextBox txtCriterio 
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
      Left            =   2430
      TabIndex        =   0
      Top             =   30
      Width           =   5505
   End
   Begin TrueOleDBGrid80.TDBGrid grdGrilla 
      Height          =   4920
      Left            =   30
      TabIndex        =   1
      Top             =   390
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   8678
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
   Begin VB.Label cmdTexto 
      Caption         =   "Total : 350 de 4,500 Registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   5370
      Width           =   4305
   End
End
Attribute VB_Name = "frmBuscaAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As ADODB.Recordset
Dim nTotal As Double
Dim sCriterio As String
Public nPredeterm As Integer

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 0 Then
      Aceptar
   Else
      wEnter = False
      Unload Me
   End If
End Sub

Private Sub Form_Activate()
   LlenaBusqueda
   cboCriterio.ListIndex = nPredeterm
End Sub

Private Sub Form_Load()
   wEnter = False
   Centrar Me
   Set Rs = Lib.OpenRecordset(Isql, CnAlmacen)
   
   'Configuración de la Grilla
   Set grdGrilla.DataSource = Rs
   txtCriterio = ""
   
   nTotal = Rs.RecordCount
   cmdTexto.Caption = "Total : " & Format(nTotal, "##,##0") & " de " & Format(nTotal, "##,##0") & " Registros"
   Screen.MousePointer = vbDefault
End Sub

Sub LlenaBusqueda()
    Dim i As Integer
    With cboCriterio
        For i = 0 To grdGrilla.Columns.Count - 1
            If grdGrilla.Columns(i).ValueItems.Presentation <> dbgCheckBox Then
                .AddItem grdGrilla.Columns(i).Caption
                .ItemData(.NewIndex) = i
            End If
        Next i
    End With
End Sub

Public Sub Filtrar()
   If Len(Trim(txtCriterio.Text)) > 0 Then
      Select Case VarType(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).Value)
             Case 2 To 6
             sCriterio = Trim(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).DataField) & "= " & val(txtCriterio.Text)
             Case 7
             sCriterio = Trim(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).DataField) & " >= #" & txtCriterio.Text & "# and " & Trim(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).DataField) & " <= #" & txtCriterio.Text & " 23:59#"
             Case Else
             sCriterio = Trim(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).DataField) & " Like " & "'*" & txtCriterio.Text & "*'"
      End Select
      
      Screen.MousePointer = vbHourglass
      With Rs
           .Filter = sCriterio
           If .BOF = False Then
              .MoveFirst
           End If
      End With
      cmdTexto.Caption = "Total : " & Format(Rs.RecordCount, "##,##0") & " de " & Format(nTotal, "##,##0") & " Registros"
      Screen.MousePointer = vbDefault
   Else
      Screen.MousePointer = vbHourglass
      Rs.Filter = adFilterNone
      Rs.Requery
      Rs.MoveFirst
      cmdTexto.Caption = "Total : " & Format(nTotal, "##,##0") & " de " & Format(nTotal, "##,##0") & " Registros"
      Screen.MousePointer = vbDefault
      sCriterio = ""
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set Rs = Nothing
End Sub

Private Sub grdGrilla_DblClick()
   cmdOpcion_Click (0)
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      Aceptar
   ElseIf KeyCode = 27 Then
      wEnter = False
      Unload Me
   End If
End Sub

Private Sub txtCriterio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      Aceptar
   ElseIf KeyCode = 27 Then
      wEnter = False
      Unload Me
   End If
End Sub

Private Sub txtCriterio_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           '    numeros   letras    sp  numpad     delet  ñ
           Case 48 To 57, 65 To 90, 32, 96 To 105, 8, 46, 192
                Filtrar
    End Select
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    Rs.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
    cboCriterio.ListIndex = ColIndex
    txtCriterio.SetFocus
End Sub

Public Sub Aceptar()
   If Rs.RecordCount > 0 Then
      wEnter = True
      sCodigo = Rs!codigo
      sDescrip = Rs!Descripcion
      Unload Me
   End If
End Sub
Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero 1, grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero 5, grdGrilla
           Case Is = 2 'Previo
                MoverPuntero 3, grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero 4, grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero 6, grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero 2, grdGrilla
    End Select
End Sub
