VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCajon 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   7365
   ClientLeft      =   2535
   ClientTop       =   1725
   ClientWidth     =   11850
   Icon            =   "frmCajon.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGrilla 
      Height          =   5865
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11805
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   5610
         Left            =   90
         TabIndex        =   22
         Top             =   180
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   9895
         _LayoutType     =   4
         _RowHeight      =   25
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
         Caption         =   "Correlativo"
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
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   1425
      Left            =   0
      ScaleHeight     =   1365
      ScaleWidth      =   11790
      TabIndex        =   0
      Top             =   5940
      Width           =   11850
      Begin VB.Frame fraCriterio 
         Caption         =   " Criterio "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2850
         TabIndex        =   18
         Top             =   660
         Width           =   4005
         Begin VB.TextBox txtCriterio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   3870
         End
      End
      Begin VB.Frame fraCampo 
         Caption         =   " Campo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   60
         TabIndex        =   16
         Top             =   660
         Width           =   2745
         Begin VB.ComboBox cboCriterio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   90
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   240
            Width           =   2535
         End
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
         Left            =   6990
         Picture         =   "frmCajon.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Modificar"
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
         Left            =   8190
         Picture         =   "frmCajon.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Emite"
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
         Index           =   5
         Left            =   9390
         Picture         =   "frmCajon.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Index           =   6
         Left            =   10590
         Picture         =   "frmCajon.frx":0FA8
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "KeyBoard"
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
         Index           =   7
         Left            =   6990
         Picture         =   "frmCajon.frx":109A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Buscar"
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
         Left            =   8190
         Picture         =   "frmCajon.frx":119C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Filtrar"
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
         Left            =   9390
         Picture         =   "frmCajon.frx":129E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "No Filtrar"
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
         Index           =   4
         Left            =   10590
         Picture         =   "frmCajon.frx":13A0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   1170
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   6750
         TabIndex        =   1
         Top             =   30
         Width           =   6810
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmCajon.frx":14A2
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmCajon.frx":19E4
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmCajon.frx":1F26
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   6195
            Picture         =   "frmCajon.frx":2468
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5655
            Picture         =   "frmCajon.frx":29AA
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   5115
            Picture         =   "frmCajon.frx":2EEC
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro"
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
            Left            =   1740
            TabIndex        =   21
            Top             =   120
            Width           =   3315
         End
      End
   End
End
Attribute VB_Name = "frmCajon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RsCabecera As Recordset
Dim RsReporte As Recordset
Dim CriterioB As String
Dim CriterioF As String
Dim nColumna As Integer
Dim Reporte As New dsrReporte

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

Private Sub cboCriterio_Click()
   txtCriterio.Text = ""
   txtCriterio.SetFocus
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, grdGrilla
    End Select

End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0  'Nuevo
                Sw = True
                'Cambiar el Nombre del Formulario Detalle
                frmCajonDetalle.Show vbModal
                      
           Case Is = 1  'Modifica
                If RsCabecera.RecordCount > 0 Then
                    Sw = False
                    'Cambiar el Nombre del Formulario Detalle
                    frmCajonDetalle.Show vbModal
                Else
                    MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
           
           Case Is = 2  'Buscar
                If Len(cboCriterio) > 0 And Len(Trim(txtCriterio)) > 0 Then
                   Select Case VarType(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).Value)
                          Case 2 To 6
                          CriterioB = (Trim(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).DataField)) & " = " & Val(txtCriterio.Text)
                          Case 7
                          CriterioB = Trim(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).DataField) & "= #" & txtCriterio.Text & "#"
                          Case Else
                          CriterioB = Trim(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).DataField) & " Like " & "'*" & txtCriterio.Text & "*'"
                   End Select
                   Screen.MousePointer = vbHourglass
                   With RsCabecera
                        .Requery
                        .MoveFirst
                        .Find CriterioB
                        If .EOF = True Then
                           MsgBox "Criterio No Encontrado", vbExclamation, sMensaje
                           .MoveLast
                        End If
                   End With
                Else
                    MsgBox "Datos Incompletos", vbExclamation, sMensaje
                End If
                Screen.MousePointer = vbDefault
           
           Case Is = 3  'Filtrar
                If Len(cboCriterio) > 0 And Len(Trim(txtCriterio.Text)) > 0 Then
                   Select Case VarType(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).Value)
                          Case 2 To 6
                          CriterioF = Trim(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).DataField) & "= " & Val(txtCriterio.Text)
                          Case 7
                          CriterioF = Trim(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).DataField) & " >= #" & txtCriterio.Text & "# and " & Trim(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).DataField) & " <= #" & txtCriterio.Text & " 23:59:59#"
                          Case Else
                          CriterioF = Trim(grdGrilla.Columns(cboCriterio.ItemData(cboCriterio.ListIndex)).DataField) & " Like " & "'*" & txtCriterio.Text & "*'"
                   End Select
                   
                   Screen.MousePointer = vbHourglass
                   With RsCabecera
                        .Filter = CriterioF
                        If .BOF = True Then
                           MsgBox "Criterio No Encontrado", vbExclamation, sMensaje
                           .Filter = adFilterNone
                           CriterioF = 0
                        Else
                           .MoveFirst
                        End If
                   End With
                   
                Else
                   MsgBox "Datos Incompletos", vbExclamation, sMensaje
                End If
                cmdTexto.Caption = "Registro " & RsCabecera.AbsolutePosition & " de " & RsCabecera.RecordCount
                Screen.MousePointer = vbDefault
           
           Case Is = 4  'No Filtrar
                Screen.MousePointer = vbHourglass
                RsCabecera.Filter = adFilterNone
                RsCabecera.Requery
                If Not RsCabecera.EOF Then
                   RsCabecera.MoveLast
                End If
                Screen.MousePointer = vbDefault
                CriterioF = ""
           
           Case Is = 5  'Emite
                If RsCabecera.RecordCount = 0 Then MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema": Exit Sub
                Screen.MousePointer = vbHourglass
                Set RsReporte = RsCabecera.Clone
                RsReporte.Filter = CriterioF
                RsReporte.Sort = grdGrilla.Columns(nColumna).DataField & " ASC"
                
                Reporte.DiscardSavedData
                Reporte.ReportTitle = Me.Caption
                Reporte.Text3.SetText "Resumido"
                Reporte.Text7.SetText "Estado"
                Reporte.Database.SetDataSource RsReporte
                Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                frmEmite.CRViewer.ReportSource = Reporte
                frmEmite.CRViewer.DisplayGroupTree = False
                frmEmite.CRViewer.ViewReport
                frmEmite.Show vbModal
           
           Case Is = 6  'Salir
                Unload Me
           
           Case Is = 7  'Keyboard
                frmKeyBoard.Show vbModal
                txtCriterio.Text = IIf(wEnter, sDescrip, "")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
 
    ' Ingrese el Titulo
    Me.Caption = " Mantenimiento de Cajon de Dinero "
    grdGrilla.Caption = Me.Caption
    Centrar Me
    nColumna = 2
    
    ' Ingrese la Vista
    Isql = "select * from vCajon where tCaja='" & sCaja & "'"
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    'Configuración de la Grilla
    Call ConfGrilla(4, grdGrilla, "Codigo", 2, "Codigo", 1200, 2, 0, "", _
                                  "Descripción Detallada", 2, "Descripcion", 5500, 0, 0, "", _
                                  "Descripción Resumida", 2, "tResumido", 3200, 0, 0, "", _
                                  "Activo", 2, "lactivo", 1100, 2, 4, "")
    
    Set grdGrilla.DataSource = RsCabecera
    LlenaBusqueda
    txtCriterio = ""
    
    cmdTexto.Caption = "Registro " & RsCabecera.AbsolutePosition & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
   fraGrilla.Height = IIf(Me.Height - 2000 > 0, Me.Height - 2000, 0)
   grdGrilla.Height = IIf(fraGrilla.Height - 500 > 0, fraGrilla.Height - 500, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCajon = Nothing
End Sub

Private Sub grdGrilla_DblClick()
    cmdOpcion_Click (1)
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    nColumna = ColIndex
    RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
End Sub

