VERSION 5.00
Object = "{B1C46850-3E6A-11D2-8FEB-00104B9E07A7}#3.0#0"; "SSDW3AO.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Begin VB.Form frmBuscarPropiedad 
   Caption         =   "Busca Propiedades"
   ClientHeight    =   5580
   ClientLeft      =   2310
   ClientTop       =   1395
   ClientWidth     =   10845
   Icon            =   "frmBuscarPropiedad.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   10845
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdFiltrar 
      Height          =   555
      Left            =   1410
      TabIndex        =   2
      Top             =   4995
      Width           =   1320
      _Version        =   196612
      _ExtentX        =   2328
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Filtrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscarPropiedad.frx":0442
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdUltimo 
      Height          =   555
      Left            =   4770
      TabIndex        =   1
      Top             =   4410
      Width           =   1320
      _Version        =   196612
      _ExtentX        =   2328
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Ultimo Registro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscarPropiedad.frx":075C
      Font3D          =   3
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdPrimero 
      Height          =   555
      Left            =   4770
      TabIndex        =   0
      Top             =   2700
      Width           =   1320
      _Version        =   196612
      _ExtentX        =   2328
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Primer Registro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscarPropiedad.frx":0CAE
      Font3D          =   3
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdAceptar 
      Height          =   555
      Left            =   9495
      TabIndex        =   3
      Top             =   4995
      Width           =   1320
      _Version        =   196612
      _ExtentX        =   2328
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Aceptar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscarPropiedad.frx":1200
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdCancelar 
      Height          =   555
      Left            =   8100
      TabIndex        =   4
      Top             =   4995
      Width           =   1320
      _Version        =   196612
      _ExtentX        =   2328
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Cancela"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscarPropiedad.frx":151A
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdPrevio 
      Height          =   555
      Left            =   4770
      TabIndex        =   5
      Top             =   3270
      Width           =   1320
      _Version        =   196612
      _ExtentX        =   2328
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Previo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscarPropiedad.frx":1DF4
      Font3D          =   3
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdSiguiente 
      Height          =   555
      Left            =   4770
      TabIndex        =   6
      Top             =   3840
      Width           =   1320
      _Version        =   196612
      _ExtentX        =   2328
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Siguiente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscarPropiedad.frx":2346
      Font3D          =   3
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdTodos 
      Height          =   555
      Index           =   1
      Left            =   45
      TabIndex        =   7
      Top             =   4995
      Width           =   1320
      _Version        =   196612
      _ExtentX        =   2328
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Todos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscarPropiedad.frx":2898
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdPaginaUp 
      Height          =   555
      Left            =   4770
      TabIndex        =   8
      Top             =   945
      Width           =   1320
      _Version        =   196612
      _ExtentX        =   2328
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "No Pasar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscarPropiedad.frx":3172
      Font3D          =   3
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdPasar 
      Height          =   555
      Left            =   4770
      TabIndex        =   9
      Top             =   360
      Width           =   1320
      _Version        =   196612
      _ExtentX        =   2328
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Pasar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscarPropiedad.frx":36C4
      Font3D          =   3
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdBuscar 
      Height          =   555
      Left            =   2775
      TabIndex        =   10
      Top             =   4995
      Width           =   1320
      _Version        =   196612
      _ExtentX        =   2328
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscarPropiedad.frx":3C16
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdTemporal1 
      Height          =   4605
      Left            =   45
      TabIndex        =   11
      Top             =   360
      Width           =   4695
      ScrollBars      =   2
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets.count =   1
      stylesets(0).Name=   "Cabecera"
      stylesets(0).ForeColor=   16711680
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmBuscarPropiedad.frx":4068
      stylesets(0).AlignmentText=   2
      HeadFont3D      =   3
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      RowNavigation   =   1
      HeadStyleSet    =   "Cabecera"
      ForeColorEven   =   0
      BackColorEven   =   -2147483624
      BackColorOdd    =   -2147483624
      RowHeight       =   503
      ExtraHeight     =   132
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   8281
      _ExtentY        =   8123
      _StockProps     =   79
      Caption         =   "Listado de Propiedades"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdTemporal2 
      Height          =   4605
      Left            =   6120
      TabIndex        =   13
      Top             =   360
      Width           =   4695
      ScrollBars      =   2
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets.count =   1
      stylesets(0).Name=   "Cabecera"
      stylesets(0).ForeColor=   16711680
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmBuscarPropiedad.frx":4084
      stylesets(0).AlignmentText=   2
      HeadFont3D      =   3
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      RowNavigation   =   1
      HeadStyleSet    =   "Cabecera"
      ForeColorEven   =   0
      BackColorEven   =   -2147483624
      BackColorOdd    =   -2147483624
      RowHeight       =   423
      ExtraHeight     =   26
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   8281
      _ExtentY        =   8123
      _StockProps     =   79
      Caption         =   "Listado de Propiedades Aceptadas"
      ForeColor       =   0
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label txtResultado 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   45
      TabIndex        =   12
      Top             =   0
      Width           =   10725
   End
End
Attribute VB_Name = "frmBuscarPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nColumna As Integer
Dim RsTemporal1 As Recordset
Dim sTemporal1 As String
Dim RsTemporal2 As Recordset
Dim sTemporal2 As String
Public sCodigo1 As String
Public sCodigo2 As String
Public sCodigo3 As String

Private Sub cmdAceptar_AfterClick()
   wEnter = True
   If Not RsTemporal2.EOF Then
      RsTemporal2.MoveFirst
   End If
   For i = 1 To 3
       If RsTemporal2.EOF Then
          Select Case i
                 Case Is = 1
                 sCodigo1 = ""
                 Case Is = 2
                 sCodigo2 = ""
                 Case Is = 3
                 sCodigo3 = ""
          End Select
          frmVenta.txtPropiedad(i).Caption = ""
       Else
          Select Case i
                 Case Is = 1
                 sCodigo1 = RsTemporal2!Codigo
                 Case Is = 2
                 sCodigo2 = RsTemporal2!Codigo
                 Case Is = 3
                 sCodigo3 = RsTemporal2!Codigo
          End Select
          frmVenta.txtPropiedad(i).Caption = RsTemporal2!Propiedad
          RsTemporal2.MoveNext
       End If
       
   Next i
   Unload Me
End Sub

Private Sub cmdBuscar_AfterClick()
   frmKeyboard.Show vbModal
   If wEnter And Len(Trim(sDescrip)) > 0 Then
      If nColumna = 0 Then
         Criterio = "Codigo like '*" & Trim(sDescrip) & "*'"
      Else
         Criterio = "Propiedad like '*" & Trim(sDescrip) & "*'"
      End If
      Screen.MousePointer = vbHourglass
      RsTemporal1.MoveFirst
      RsTemporal1.Find Criterio
      
      If Not RsTemporal1.EOF Then
         txtResultado.Caption = RsTemporal1!Propiedad
      Else
         MsgBox "Criterio No Encontrado", vbExclamation, sMensaje
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_AfterClick()
   wEnter = False
   Unload Me
End Sub

Private Sub cmdFiltrar_AfterClick()
   frmKeyboard.Show vbModal
   If wEnter And Len(Trim(sDescrip)) > 0 Then
      If nColumna = 0 Then
         Criterio = "Codigo Like '*" & LTrim(RTrim(sDescrip)) & "*'"
      Else
         Criterio = "Propiedad Like '*" & LTrim(RTrim(sDescrip)) & "*'"
      End If
        
      Screen.MousePointer = vbHourglass
      'RsTemporal1.Requery
      RsTemporal1.Filter = Criterio
      If RsTemporal1.RecordCount <= 0 Then
         RsTemporal1.Filter = adFilterNone
         MsgBox "Criterio No Encontrado", vbExclamation, sMensaje
      Else
        RsTemporal1.MoveFirst
        txtResultado.Caption = RsTemporal1!Propiedad
      End If
      'RsTemporal1.Requery
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPaginaUp_AfterClick()
      Screen.MousePointer = vbHourglass
      If RsTemporal2.RecordCount <> 0 Then
         Isql = "insert into " & sTemporal1 & _
                "( Codigo, Propiedad, lGeneral, subGrupo, Producto ) " & _
                " values ( " & "'" & RsTemporal2!Codigo & "', " _
                             & "'" & RsTemporal2!Propiedad & "', " _
                             & Format(RsTemporal2!lGeneral, 0) & ", " _
                             & Format(RsTemporal2!subGrupo, 0) & ", " _
                             & Format(RsTemporal2!Producto, 0) & ")"
         Cn.Execute Isql
         Cn.Execute "delete from " & sTemporal2 & " where Codigo = '" & RsTemporal2!Codigo & "'"
         RsTemporal2.Requery
         RsTemporal1.Requery
      End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPasar_AfterClick()
   If RsTemporal2.RecordCount < 3 Then
      Screen.MousePointer = vbHourglass
      If RsTemporal1.RecordCount <> 0 Then
         Isql = "insert into " & sTemporal2 & _
                "( Codigo, Propiedad, lGeneral, subGrupo, Producto ) " & _
                " values ( " & "'" & RsTemporal1!Codigo & "', " _
                             & "'" & RsTemporal1!Propiedad & "', " _
                             & Format(RsTemporal1!lGeneral, 0) & ", " _
                             & Format(RsTemporal1!subGrupo, 0) & ", " _
                             & Format(RsTemporal1!Producto, 0) & ")"
         Cn.Execute Isql
         Cn.Execute "delete from " & sTemporal1 & " where Codigo = '" & RsTemporal1!Codigo & "'"
         RsTemporal2.Requery
         RsTemporal1.Requery
      End If
   Else
      MsgBox "Maximo 3 propiedades por plato", vbCritical, sMensaje
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrevio_AfterClick()
   MoverPuntero Previo, grdTemporal1
   txtResultado.Caption = RsTemporal1!Propiedad
End Sub

Private Sub cmdPrimero_AfterClick()
   MoverPuntero Primero, grdTemporal1
   txtResultado.Caption = RsTemporal1!Propiedad
End Sub
Private Sub cmdSiguiente_AfterClick()
   MoverPuntero siguiente, grdTemporal1
   txtResultado.Caption = RsTemporal1!Propiedad
End Sub

Private Sub cmdTodos_AfterClick(Index As Integer)
   Screen.MousePointer = vbHourglass
   RsTemporal1.Filter = adFilterNone
   txtResultado.Caption = ""
   Set grdTemporal1.DataSource = RsTemporal1
   Screen.MousePointer = vbDefault
   txtResultado.Caption = IIf(RsTemporal1.EOF = True, "", RsTemporal1!Propiedad)
End Sub

Private Sub cmdUltimo_AfterClick()
   MoverPuntero Ultimo, grdTemporal1
   txtResultado.Caption = RsTemporal1!Propiedad
End Sub

Private Sub Form_Load()
   Centrar Me
   nColumna = 1
   wEnter = False
   sTemporal1 = dbTemporal(sCaja, 5, "D", "Codigo", "TEXT(4)", _
                                          "Propiedad", "TEXT(50)", _
                                          "lGeneral", "LOGICAL", _
                                          "SubGrupo", "LOGICAL", _
                                          "Producto", "LOGICAL")
   Isql = "insert into " & sTemporal1 & " (Codigo, Propiedad, lGeneral, SubGrupo, Producto) select Codigo, tDetallado, lGeneral, SubGrupo, Producto from vPropiedad where lActivo = True and (lGeneral = True or tSubGrupo = '" & frmVenta.sSubGrupo & "' or tProducto = '" & frmVenta.wProducto & "')"
   Cn.Execute Isql

   Set RsTemporal1 = Lib.OpenRecordset("select * from " & sTemporal1, Cn)
   
   Call Lib.ConfGrid(5, grdTemporal1, "Cód", 2, "Codigo", 550, 2, 0, "", _
                                      "Propiedad", 2, "Propiedad", 2000, 0, 0, "", _
                                      "Grl", 2, "lGeneral", 550, 2, 2, "", _
                                      "Sgp", 2, "SubGrupo", 550, 2, 2, "", _
                                      "Prd", 2, "Producto", 550, 2, 2, "")
   
   Set grdTemporal1.DataSource = RsTemporal1
   Sleep (1000)
   sTemporal2 = dbTemporal(sCaja, 5, "D", "Codigo", "TEXT(4)", _
                                         "Propiedad", "TEXT(50)", _
                                         "lGeneral", "LOGICAL", _
                                         "SubGrupo", "LOGICAL", _
                                         "Producto", "LOGICAL")
      
   Set RsTemporal2 = Lib.OpenRecordset("select * from " & sTemporal2, Cn)
       Call Lib.ConfGrid(5, grdTemporal2, "Cód", 2, "Codigo", 550, 2, 0, "", _
                                         "Propiedad", 2, "Propiedad", 2000, 0, 0, "", _
                                         "Grl", 2, "lGeneral", 550, 2, 2, "", _
                                         "Sgp", 2, "SubGrupo", 550, 2, 2, "", _
                                         "Prd", 2, "Producto", 550, 2, 2, "")
   Set grdTemporal2.DataSource = RsTemporal2
            
   'Isql debe ser defino antes de ser llamado este formulario
   txtResultado.Caption = IIf(RsTemporal1.EOF = True, "", RsTemporal1!Propiedad)
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cn.Execute "drop table " & sTemporal1
   Cn.Execute "drop table " & sTemporal2
   Cn.Execute "delete from TABLATEMP where tabla = '" & sTemporal1 & "'"
   Cn.Execute "delete from TABLATEMP where tabla = '" & sTemporal2 & "'"
End Sub

Private Sub grdTemporal1_DblClick()
   wEnter = True
   sCodigo = IIf(RsTemporal1.EOF = True, "", RsTemporal1!Codigo)
   sDescrip = IIf(RsTemporal1.EOF = True, "", RsTemporal1!Propiedad)
   Unload Me
End Sub

Private Sub grdTemporal1_HeadClick(ByVal ColIndex As Integer)
    Screen.MousePointer = vbHourglass
    RsTemporal1.Sort = grdTemporal1.Columns(ColIndex).DataField & " ASC"
    nColumna = ColIndex
    RsTemporal1.Requery
    RsTemporal1.MoveFirst
    Screen.MousePointer = vbDefault
End Sub

Private Sub grdTemporal1_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
   txtResultado.Caption = IIf(RsTemporal1.RecordCount = 0, "", IIf(IsNull(RsTemporal1!Propiedad), "", RsTemporal1!Propiedad))
End Sub
