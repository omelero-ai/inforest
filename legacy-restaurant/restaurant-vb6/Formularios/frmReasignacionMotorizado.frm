VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmReasignacionMotorizado 
   Caption         =   "Reasignación Motorizado"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Desasigna Motorizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Asigna Motorizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   5160
      Picture         =   "frmReasignacionMotorizado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   1455
   End
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
      Left            =   5160
      Picture         =   "frmReasignacionMotorizado.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " Opciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6645
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
         Left            =   1560
         TabIndex        =   10
         Top             =   1320
         Width           =   3255
      End
      Begin VB.CommandButton cmdBuscar 
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
         Left            =   5040
         Picture         =   "frmReasignacionMotorizado.frx":0624
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   825
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   102432769
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   102432769
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3525
         TabIndex        =   4
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   102432771
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3525
         TabIndex        =   5
         Top             =   825
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   102432771
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   735
         TabIndex        =   8
         Top             =   1350
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   450
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   330
         TabIndex        =   6
         Top             =   915
         Width           =   1125
      End
   End
   Begin TrueOleDBGrid80.TDBGrid grdDetalle2 
      Height          =   3690
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   6509
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
      Caption         =   "Pedidos"
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
Attribute VB_Name = "frmReasignacionMotorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim RsMotorizado As Recordset

Dim nColumna As Integer


Private Sub cmdBuscar_Click()

   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
   
   Screen.MousePointer = vbHourglass

   Isql = "Select m.tCodigoPedido, mz.Descripcion, m.nTarifaMotorizadoN from MPEDIDO m " & _
          "LEFT JOIN vMOTORIZADO mz ON m.tMotorizadoN = mz.Codigo where m.tTipoPedido ='02' AND  m.tEstadoPedido ='02' AND  m.tMotorizado <>'0000' AND " & _
          "(m.tCodigoPedido like '%" & txtCriterio.Text & "') AND m.fRegistro>= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' AND m.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "'"

   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
   Set grdDetalle2.DataSource = RsCabecera
   
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub cmdOpcion_Click(Index As Integer)

 Select Case Index
 
    Case Is = 1  'Asignar
    
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
            
            Dim nMaximoAsignacionMotorizado As Double
            Dim nCantidadPedidosYaAsignados As Integer
            Dim nCantidadPedidosYaAsignados1 As Integer
            Dim nDiaActual As Integer
            Dim nTarifaDiaria As Double
            Dim nMontoAsignado As Double
            Dim nMaximo As Boolean
            nMaximoAsignacionMotorizado = Calcular("select isnull(nAsignacionMotorizado,0) as Codigo FROM TPARAMETRO", Cn)
                
               sTemp = ""
               Isql = "select * from vMotorizado where lActivo = 1 and Codigo <> '0000' order by Descripcion"
               Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 2, "Codigo", 1800, 2, 0, "", _
                                                               "Motorizado", 2, "Descripcion", 6400, 0, 0, "")
                
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
            
               If wEnter Then
               
                        If nMaximoAsignacionMotorizado > 0 Then
                            
                            nCantidadPedidosYaAsignados = Calcular("SELECT COUNT(ISNULL(nTarifaMotorizado,0)) AS Codigo FROM dbo.vMotorizado INNER JOIN dbo.MPEDIDO ON dbo.vMotorizado.Codigo = dbo.MPEDIDO.tMotorizado WHERE ISNULL(dbo.MPEDIDO.nTarifaExtra,0)=0  AND (CONVERT(nvarchar, dbo.MPEDIDO.fasignacion, 112) = '" & Format(obtieneFechaServidor, "yyyyMMdd") & "') GROUP BY dbo.MPEDIDO.tMotorizado HAVING (dbo.MPEDIDO.tMotorizado = '" & sCodigo & "')  ", Cn)
                            nCantidadPedidosYaAsignados1 = Calcular("SELECT COUNT(ISNULL(nTarifaMotorizado,0)) AS Codigo FROM dbo.vMotorizado INNER JOIN dbo.MPEDIDO ON dbo.vMotorizado.Codigo = dbo.MPEDIDO.tMotorizadoN WHERE  ISNULL(dbo.MPEDIDO.nTarifaExtraN,0)=0 AND (CONVERT(nvarchar, dbo.MPEDIDO.fasignacion, 112) = '" & Format(obtieneFechaServidor, "yyyyMMdd") & "') GROUP BY dbo.MPEDIDO.tMotorizadoN HAVING (dbo.MPEDIDO.tMotorizadoN = '" & sCodigo & "')  ", Cn)
                            
                            Select Case tTarifaActualMotorizado
                                Case "Tarifa Lun - Sab"
                                            nTarifaDiaria = Calcular("select isnull(ntarifalv,0) as codigo  from tmotorizadodatos where tcodigo='" & sCodigo & "'", Cn)
            
                                Case "Tarifa Dom"
                                            nTarifaDiaria = Calcular("select isnull(ntarifasd,0) as codigo  from tmotorizadodatos where tcodigo='" & sCodigo & "'", Cn)
            
                                Case "Tarifa Especial"
                                
                                            nTarifaDiaria = Calcular("select isnull(ntarifaes,0) as codigo  from tmotorizadodatos where tcodigo='" & sCodigo & "'", Cn)
                                
                                Case Else
                                            nTarifaDiaria = 0
                            End Select
                      
                            
                            
                            nMontoAsignado = Round((nTarifaDiaria * (nCantidadPedidosYaAsignados + nCantidadPedidosYaAsignados1)) + nTarifaDiaria, 2)
                            
                                    ' Validar Tarifa Extra
                                    If nMaximoAsignacionMotorizado < nMontoAsignado Then
                                    
                                        If MsgBox("Con la Asignacion de este Pedido, sobre pasa el Monto Máximo de Asignación Por Motorizado, Desea Continuar?", vbQuestion + vbYesNo) = vbYes Then
            
                                               nMaximo = True
                                            
                                        Else
                                            Exit Sub
                
                                        End If
                                                                                        
                                    End If
                                    
                        End If
               
                        'Tarifa Oficial
                        Cn.Execute "update MPEDIDO set nTarifaMotorizadoN=" & IIf(IsNull(nTarifaDiaria), 0, nTarifaDiaria) & ", tMotorizadoN='" & sCodigo & "', nTarifaExtraN=0  where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                        
                        'Tarifa Extra
                        If nMaximo = True Then
                        Cn.Execute "update MPEDIDO set nTarifaExtraN=1  where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                        End If
                  
                  End If
                  
                  RsCabecera.Requery
          
          
    Case Is = 2  'Desasignar
    
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                  
                      Cn.Execute "Update MPEDIDO set  nTarifaMotorizadoN = null, nTarifaExtraN = null , tMotorizadoN = null where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                      
                RsCabecera.Requery
        
  End Select
          
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()

   Centrar Me
      
   dtpFecIni.Value = Date
   dtpFecFin.Value = Date
   
   nColumna = 2
   
   Isql = "Select m.tCodigoPedido, mz.Descripcion, m.nTarifaMotorizadoN From MPEDIDO m LEFT JOIN vMOTORIZADO mz ON m.tMotorizadoN = mz.Codigo Where m.tTipoPedido ='02'  AND m.tEstadoPedido ='02' " & _
          "AND m.tMotorizado <>'0000' AND m.fRegistro > '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " 00:00' AND m.fRegistro < '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " 23:59'"
   
   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
   
       'Configuración de la Grilla
       Call ConfGrilla(3, grdDetalle2, "Pedido", 2, "tCodigoPedido", 1100, 2, 0, "", _
                                       "Motorizado", 2, "Descripcion", 2300, 0, 0, "", _
                                       "Tarifa", 2, "nTarifaMotorizadoN", 800, 0, 0, "")
                                     
       Set grdDetalle2.DataSource = RsCabecera

End Sub
