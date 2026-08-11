VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmClienteDeliveryFrecuente 
   Caption         =   "Clientes Frecuentes"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
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
      Height          =   6855
      Left            =   8280
      TabIndex        =   2
      Top             =   0
      Width           =   6195
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
         Left            =   2040
         Picture         =   "frmClienteDeliveryFrecuente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3840
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
         Left            =   3360
         Picture         =   "frmClienteDeliveryFrecuente.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3840
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
         Left            =   4680
         Picture         =   "frmClienteDeliveryFrecuente.frx":0A64
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3840
         Width           =   1290
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
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   2205
         MaxLength       =   50
         TabIndex        =   7
         Text            =   " "
         Top             =   2310
         Width           =   3780
      End
      Begin VB.TextBox txtTelefono 
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
         Left            =   2205
         MaxLength       =   15
         TabIndex        =   6
         Text            =   " "
         Top             =   1050
         Width           =   2610
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   2205
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1905
         Width           =   3780
      End
      Begin VB.TextBox txtApellido 
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
         Left            =   2205
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1515
         Width           =   3780
      End
      Begin VB.TextBox txtIdentidad 
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
         Left            =   2205
         MaxLength       =   15
         TabIndex        =   3
         Text            =   " "
         Top             =   3150
         Width           =   2610
      End
      Begin MSDataListLib.DataCombo cboTipoCliente 
         Height          =   315
         Left            =   2205
         TabIndex        =   9
         Top             =   630
         Width           =   2610
         _ExtentX        =   4604
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
      Begin MSDataListLib.DataCombo cboTipoIdentidad 
         Height          =   315
         Left            =   2205
         TabIndex        =   10
         Top             =   2760
         Width           =   2610
         _ExtentX        =   4604
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
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Teléfono :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   27
         Left            =   1410
         TabIndex        =   18
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   1545
         TabIndex        =   17
         Top             =   285
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail :"
         Height          =   195
         Index           =   6
         Left            =   1605
         TabIndex        =   16
         Top             =   2355
         Width           =   525
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Razón Social / Nombres :"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   15
         Top             =   1950
         Width           =   1830
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Razón Comercial / Apellido :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   2010
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cliente Frecuente :"
         Height          =   195
         Index           =   15
         Left            =   435
         TabIndex        =   13
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Número Identidad :"
         Height          =   195
         Index           =   25
         Left            =   780
         TabIndex        =   12
         Top             =   3225
         Width           =   1350
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Identidad :"
         Height          =   195
         Index           =   16
         Left            =   795
         TabIndex        =   11
         Top             =   2820
         Width           =   1335
      End
   End
   Begin VB.Frame fraGrilla 
      Height          =   6840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   6555
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   11562
         _LayoutType     =   4
         _RowHeight      =   21
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).FooterText=   "Footer Text"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "True/False"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).AnchorRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   3
         Splits(0).AllowColMove=   -1  'True
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   3
         Splits(0).DividerColor=   32768
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos(0).NeedTotalPageCount=   -1  'True
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Correlativo"
         TabAction       =   2
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         MultiSelect     =   0
         DataView        =   2
         GroupByCaption  =   "Arrastre aqui la columna a agrupar"
         DeadAreaBackColor=   -2147483637
         ScrollTrack     =   -1  'True
         RowDividerColor =   12632256
         RowSubDividerColor=   -2147483648
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         CollapseColor   =   16711935
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(10)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=3,.fontname=Calibri"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=63,.parent=1,.namedParent=33"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=64,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=65,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=66,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=80,.parent=6,.bgcolor=&HE7FAB6&,.fgcolor=&H80000009&"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=79,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=90,.parent=63"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=64,.alignment=0"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=65"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
         _StyleDefs(36)  =   "Named:id=33:Normal"
         _StyleDefs(37)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(38)  =   ":id=33,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(39)  =   ":id=33,.fontname=Calibri"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H80000003&"
         _StyleDefs(42)  =   ":id=34,.fgcolor=&H800000&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(43)  =   ":id=34,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(44)  =   ":id=34,.fontname=Arial Black"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(49)  =   ":id=36,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(50)  =   ":id=36,.fontname=Calibri"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&HFFFF&"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H800000&,.locked=0,.bold=-1"
         _StyleDefs(55)  =   ":id=38,.fontsize=2025,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(56)  =   ":id=38,.fontname=Calibri"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(59)  =   ":id=39,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(60)  =   ":id=39,.fontname=Calibri"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33,.bgcolor=&H80000005&,.borderColor=&H80000002&,.bold=0"
         _StyleDefs(63)  =   ":id=40,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(64)  =   ":id=40,.fontname=Calibri"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34,.bgcolor=&H80000003&,.fgcolor=&H80000009&,.bold=-1"
         _StyleDefs(67)  =   ":id=41,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(68)  =   ":id=41,.fontname=Calibri"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33,.bgcolor=&H80000018&,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(71)  =   ":id=42,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(72)  =   ":id=42,.fontname=Calibri"
         _StyleDefs(73)  =   "Named:id=13:CabeceraTitulo"
         _StyleDefs(74)  =   ":id=13,.parent=34,.bold=-1,.fontsize=1125,.italic=0,.underline=0"
         _StyleDefs(75)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(76)  =   ":id=13,.fontname=Arial Black"
      End
   End
End
Attribute VB_Name = "frmClienteDeliveryFrecuente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim nColumna As Integer
Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim lAsc            As Boolean

Dim RsTipoCliente As Recordset
Dim RsTipoIdentidad As Recordset

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
    
        Case Is = 0 ' Limpiar
            txtCodigo.Text = ""
            txtTelefono.Text = ""
            txtApellido.Text = ""
            txtNombre.Text = ""
            txtEmail.Text = ""
            txtIdentidad.Text = ""
    
        Case Is = 1 ' Grabar
            Dim nCorrela As String
            Dim tValorDelivery As String
                                                         
            If pais = "002" Then ' ECUADOR
                    If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & cboTipoIdentidad.BoundText & "'", Cn) Then
                        If Len(Trim(txtIdentidad.Text)) = 13 Or Len(Trim(txtIdentidad.Text)) = 10 Then
            
                        Else
                           MsgBox "La longitud del Identificador debe ser 10(Cédula) ó 13(RUC)", vbCritical, sMensaje
                           Exit Sub
                        End If
                    End If
            Else
                    'PERU - BOLIVIA
                    If Me.cboTipoIdentidad.BoundText = "02" Then
                            If lLongitud And Len(Trim(txtIdentidad.Text)) <> nLongitud Then
                               MsgBox "La longitud del Identificador debe ser " & nLongitud, vbCritical, sMensaje
                               Exit Sub
                            ElseIf Not lLongitud And Len(Trim(txtIdentidad.Text)) < nLongitud Then
                               MsgBox "La longitud del Identificador debe ser mayor igual a " & nLongitud, vbCritical, sMensaje
                               Exit Sub
                            End If
                    End If
            End If
            
            If txtApellido.Text = "" Then MsgBox "Ingrese el Apellido", vbExclamation, sMensaje: txtApellido.SetFocus: Exit Sub
            If txtNombre.Text = "" Then MsgBox "Ingrese el Nombre", vbExclamation, sMensaje: txtNombre.SetFocus: Exit Sub
                                         
               'Chequea Datos
               If txtCodigo.Text = "" Then
                    If tValorDelivery = "1" Then
                          If Calcular("select count(tTelefono) as Codigo from TDELIVERY where tTelefono='" & Trim(txtTelefono.Text) & "'", Cn) > 0 Then
                             MsgBox "Teléfono Existente", vbExclamation, sMensaje
                             Exit Sub
                          End If
                    End If
                    
                    nCorrela = Calcular("select max(tCodigodelivery) as Codigo from TDELIVERY", Cn)
                    If Mid(nCorrela, 1, 1) = "0" Or Mid(nCorrela, 1, 1) = "1" Or Mid(nCorrela, 1, 1) = "2" Or Mid(nCorrela, 1, 1) = "3" Or Mid(nCorrela, 1, 1) = "4" Or Mid(nCorrela, 1, 1) = "5" Or Mid(nCorrela, 1, 1) = "6" Or Mid(nCorrela, 1, 1) = "7" Or Mid(nCorrela, 1, 1) = "8" Or Mid(nCorrela, 1, 1) = "9" Then
                              If IsNull(nCorrela) Or nCorrela = "" Then
                                  txtCodigo.Text = "0000001"
                              Else
                                  txtCodigo.Text = Lib.Correlativo(nCorrela, 7)
                              End If
                     Else
                              If IsNull(nCorrela) Or nCorrela = "" Then
                                  txtCodigo.Text = "0000001"
                              Else
                                  If Mid(nCorrela, 2, 6) = "999999" Then
                                  txtCodigo.Text = Chr(Asc(Mid(nCorrela, 1, 1)) + 1) + "000001"
                                  Else
                                  txtCodigo.Text = Mid(nCorrela, 1, 1) + Lib.Correlativo(Mid(nCorrela, 2, 6), 6)
                                  End If
                     
                              End If
                     End If
                    'Sw = False
                     
                    'Cambiar el SQL
                    Isql = "insert into TDELIVERY( " & _
                            "tCodigodelivery, tApellido, tNombre, tTelefono, tEMail, lActivo, tTipoCliente, tUsuario, fRegistro, tTipoIdentidad, tIdentidad) " & _
                           "values ('" & txtCodigo.Text & "', " & _
                                  " '" & txtApellido.Text & "', " & _
                                  " '" & txtNombre.Text & "', " & _
                                  " '" & txtTelefono.Text & "', " & _
                                  " '" & txtEmail.Text & "', " & _
                                         1 & ", " & _
                                  " '" & cboTipoCliente.BoundText & "', " & _
                                    " '" & sUsuario & "', " & _
                                    " getdate() ,'" & cboTipoIdentidad.BoundText & "','" & txtIdentidad.Text & "' )"
                                    
                    Cn.Execute Isql
                                    
                    frmClienteDeliveryFrecuente.RsCabecera.Sort = "Codigo ASC"
                    frmClienteDeliveryFrecuente.RsCabecera.Requery
                    frmClienteDeliveryFrecuente.RsCabecera.MoveLast
                                                        
                    MsgBox "Registro Guardado", vbInformation, sMensaje

               Else
                    'Cambiar el SQL
                    If tValorDelivery = "1" Then
                          If Calcular("select count(tTelefono) as Codigo from TDELIVERY where tCodigoDelivery<>'" & txtCodigo.Text & "' and tTelefono='" & Trim(txtTelefono.Text) & "'", Cn) > 0 Then
                             MsgBox "Teléfono Existente", vbExclamation, sMensaje
                             Exit Sub
                          End If
                    End If
                    
                    Isql = "update TDELIVERY set " & _
                           "tApellido ='" & txtApellido.Text & "', " & _
                           "tNombre ='" & txtNombre.Text & "', " & _
                           "tTelefono ='" & txtTelefono.Text & "', " & _
                           "tEmail ='" & txtEmail.Text & "', " & _
                           "tTipoCliente='" & cboTipoCliente.BoundText & "', " & _
                           "fModificacion = getdate(), tUsuario='" & sUsuario & "', tTipoIdentidad= '" & cboTipoIdentidad.BoundText & "', tIdentidad='" & txtIdentidad.Text & "' " & _
                           " where tCodigoDelivery = '" & txtCodigo & "'"
                     Cn.Execute Isql
                                 
                     nPos = frmClienteDeliveryFrecuente.RsCabecera.Bookmark
                     frmClienteDeliveryFrecuente.RsCabecera.Requery
                     If frmClienteDeliveryFrecuente.RsCabecera.RecordCount = 0 Then
                        frmClienteDeliveryFrecuente.RsCabecera.Filter = adFilterNone
                     End If
                     frmClienteDeliveryFrecuente.RsCabecera.Bookmark = nPos
                     Screen.MousePointer = vbDefault
                     MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
                           
        Case Is = 3 ' Salir
            Unload Me
        
    End Select
End Sub

Private Sub Form_Resize()
'   fraGrilla.Height = IIf(Me.Height - 1300 > 0, Me.Height - 1300, 0)
'   fraGrilla.Width = IIf(Me.Width - 300 > 0, Me.Width - 300, 0)
'
'   grdGrilla.Height = IIf(fraGrilla.Height - 300 > 0, fraGrilla.Height - 300, 0)
'   grdGrilla.Width = IIf(fraGrilla.Width - 300 > 0, fraGrilla.Width - 300, 0)
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    nColumna = ColIndex
    lAsc = Not lAsc
    If lAsc Then
       RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
    Else
       RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " DESC"
    End If
End Sub

Private Sub grdGrilla_FilterChange()
    On Error GoTo errHandler
    Set cols = grdGrilla.Columns
    Dim c As Integer
    c = grdGrilla.col
    grdGrilla.HoldFields
    RsCabecera.Filter = getFilter(cols)
    grdGrilla.col = c
    grdGrilla.EditActive = True
    Exit Sub

errHandler:
    MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & err.Description, vbCritical + vbOKOnly, sMensaje
    
    Call cmdClearFilter_Click(cols)
    RsCabecera.Filter = adFilterNone
End Sub

Sub LlenaCombo()
    With cboTipoCliente
        Isql = "Select * from vTipoClienteFrecuente  where codigo <> '00' and lActivo = 1 order by Codigo"
        Set RsTipoCliente = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsTipoCliente
        .DataField = "Descripcion"
        .ListField = "Descripcion"
        .BoundColumn = "Codigo"
    End With
    
    With cboTipoIdentidad
         Isql = "Select * from vTipoIdentidad where lActivo=1 order by Descripcion"
         Set RsTipoIdentidad = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoIdentidad
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
End Sub
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    wEnter = False
    ' Ingrese el Titulo
    Me.Caption = " Clientes Frecuentes "
    grdGrilla.Caption = Me.Caption
    Centrar Me
    nColumna = 2
    
    LlenaCombo
    
    ' Ingrese la Vista
    Isql = "select * from vDelivery"
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    'Configuración de la Grilla
    Call ConfGrilla(4, grdGrilla, "Cod", 2, "Codigo", 1000, 2, 0, "", _
                                  "Tipo", 2, "TipoCliente", 1800, 0, 0, "", _
                                  "Cliente", 2, "Cliente", 3400, 0, 0, "", _
                                  "Teléfono", 2, "tTelefono", 1000, 0, 0, "")
                
    Set grdGrilla.DataSource = RsCabecera
        
    Screen.MousePointer = vbDefault
End Sub

'Private Sub grdGrilla_Click()
'    If RsCabecera.RecordCount > 0 Then
'        Asignar
'    Else
'        MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
'    End If
'End Sub

Private Sub grdGrilla_DblClick()
    If RsCabecera.RecordCount > 0 Then
        sCodigo = txtCodigo.Text
        sDescrip = txtApellido.Text + " " + txtNombre.Text
        sTelefonoReserva = txtTelefono.Text
        wEnter = True
        Unload Me
    Else
        MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
    End If
End Sub

Sub Asignar()
    With RsCabecera
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtApellido = IIf(IsNull(!tApellido), "", !tApellido)
        txtNombre = IIf(IsNull(!tNombre), "", !tNombre)
        txtTelefono = IIf(IsNull(!tTelefono), "", !tTelefono)
        txtEmail = IIf(IsNull(!tEmail), "", !tEmail)
        cboTipoCliente.BoundText = IIf(IsNull(!tTipoCliente), "00", !tTipoCliente)
        txtIdentidad.Text = IIf(IsNull(!tIdentidad), "", !tIdentidad)
        cboTipoIdentidad.BoundText = IIf(IsNull(!tTipoIdentidad), "", !tTipoIdentidad)
    End With

End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If RsCabecera.RecordCount > 0 Then
        Asignar
    Else
        MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
    End If
End Sub
