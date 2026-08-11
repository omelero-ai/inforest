VERSION 5.00
Begin VB.Form frmFlash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmFlash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      ForeColor       =   &H00800000&
      Height          =   4050
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5445
         Picture         =   "frmFlash.frx":000C
         ScaleHeight     =   300
         ScaleWidth      =   1500
         TabIndex        =   7
         Top             =   2760
         Width           =   1500
      End
      Begin VB.Label lblRev 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rv. 2565"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   6150
         TabIndex        =   9
         Top             =   3405
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "rv. 2356"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7170
         TabIndex        =   8
         Top             =   3360
         Width           =   60
      End
      Begin VB.Image Image 
         Height          =   1335
         Left            =   1770
         Picture         =   "frmFlash.frx":0537
         Stretch         =   -1  'True
         Top             =   210
         Width           =   5220
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Módulo de Administración"
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
         Height          =   240
         Left            =   4245
         TabIndex        =   6
         Top             =   1710
         Width           =   2700
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Software para Restaurantes"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4965
         TabIndex        =   5
         Top             =   2055
         Width           =   1980
      End
      Begin VB.Image Image1 
         Height          =   3225
         Left            =   120
         Picture         =   "frmFlash.frx":1B53
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         Caption         =   "Copyright : Infhotel Servicios Informáticos S.A.C."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   3450
         TabIndex        =   2
         Top             =   3135
         Width           =   3510
      End
      Begin VB.Label lblWarning 
         Caption         =   "Producto protegido por las leyes de derecho de autor como se describe en Acerca de."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   3720
         Width           =   6375
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 1.1.1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4740
         TabIndex        =   3
         Top             =   3390
         Width           =   1350
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Producto de la Familia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   3075
         TabIndex        =   4
         Top             =   2295
         Width           =   3870
      End
   End
   Begin VB.Timer Tiempo 
      Interval        =   1500
      Left            =   30
      Top             =   3810
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   lblVersion.Caption = "Versión " & App.Major & "." & App.Minor
   lblRev.Caption = "Rv." & App.Revision
   Screen.MousePointer = vbHourglass
   If lAlmacen Then
      Tiempo.Interval = 1
   Else
      Tiempo.Interval = 1500
   End If
   'TIPO CAMBIO
   If pais = "002" Or pais = "005" Then
      RTipoCambio
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmFlash = Nothing
End Sub

Private Sub Tiempo_Timer()
   If lAlmacen Then
      Actualiza
   End If
   If sModulo = "INTEGRACION" Then
      Integra
   End If
   Unload Me
End Sub
Public Sub Actualiza()
   Dim RsTemp As Recordset
   Screen.MousePointer = vbHourglass
   CnAlmacen.Execute "sp_ActualizaReceta"
   
   'Actualiza INFOREST los Costos de las Recetas de Venta
'   Cn.Execute "update TPRODUCTO set nInsumo=0, nGasto=0, nManoObra=0, " & _
'                                   "nInsumo2=0, nGasto2=0, nManoObra2=0, " & _
'                                   "nInsumo3=0, nGasto3=0, nManoObra3=0, " & _
'                                   "nInsumo4=0, nGasto4=0, nManoObra4=0, " & _
'                                   "nInsumo5=0, nGasto5=0, nManoObra5=0 "
   'Actualiza INFOREST los Costos de las Recetas de Propiedades
'   Cn.Execute "update TPROPIEDAD set nInsumo=0, nGasto=0, nManoObra=0 "

   Cn.Execute "usp_Inforest_InicializaCostos"
'Dim oComandox As clsComando
'Set oComandox = New clsComando
'If Not oComandox.CreateCmdSp("usp_Inforest_InicializaCostos", Cn) Then
'   Set oComandox = Nothing
'   Exit Sub
'End If
'If Not oComandox.ExecSP Then
'    Set oComandox = Nothing
'    Exit Sub
'End If

Set oComandox = New clsComando
If Not oComandox.CreateCmdSp("usp_Inforest_ActualizaCostos", Cn) Then
   Set oComandox = Nothing
   Exit Sub
End If
oComandox.CreateParameter "@tNombreInforest", adVarChar, adParamInput, 50, sMDB

If BDLink = "1" Then
    oComandox.CreateParameter "@tNombreAlmacen", adVarChar, adParamInput, 50, "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" 'CASO CORP.FERNANDEZ
Else
    oComandox.CreateParameter "@tNombreAlmacen", adVarChar, adParamInput, 50, sAlmacenMDB   'CASO CORP.FERNANDEZ
End If

oComandox.CreateParameter "@tLocal", adVarChar, adParamInput, 5, sLocal
If Not oComandox.GetParamOK Then
   Set oComandox = Nothing
   Exit Sub
End If
If Not oComandox.ExecSP Then
Set oComandox = Nothing
Exit Sub
End If

'Actualiza los precios de Venta de Transferencia a almacen

If Not oComandox.CreateCmdSp("Usp_ActualizarPreciosTransferenciaAlmacen", Cn) Then
   Set oComandox = Nothing
   Exit Sub
End If
oComandox.CreateParameter "@SubGrupo", adVarChar, adParamInput, 50, ""
If BDLink = "1" Then
    oComandox.CreateParameter "@BaseDatoAlmacen", adVarChar, adParamInput, 50, "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" 'CASO CORP.FERNANDEZ
Else
    oComandox.CreateParameter "@BaseDatoAlmacen", adVarChar, adParamInput, 50, sAlmacenMDB  'CASO CORP.FERNANDEZ
End If
oComandox.CreateParameter "@tipooper", adInteger, adParamInput, 5, 2
If Not oComandox.GetParamOK Then
   Set oComandox = Nothing
   Exit Sub
End If
If Not oComandox.ExecSP Then
    Set oComandox = Nothing
    Exit Sub
End If


'
'   Isql = "Update " & sMDB & ".dbo.TPRODUCTO set " & _
'          "nInsumo = T2.nInsumo, nGasto = T2.nGastoGeneral, nManoObra= T2.nManoObra,  " & _
'          "nInsumo2 = T2.nInsumo2, nGasto2 = T2.nGastoGeneral2, nManoObra2= T2.nManoObra2," & _
'          "nInsumo3 = T2.nInsumo3, nGasto3 = T2.nGastoGeneral3, nManoObra3= T2.nManoObra3, " & _
'          "nInsumo4 = T2.nInsumo4, nGasto4 = T2.nGastoGeneral4, nManoObra4= T2.nManoObra4, " & _
'          "nInsumo5 = T2.nInsumo5, nGasto5 = T2.nGastoGeneral5, nManoObra5= T2.nManoObra5 " & _
'          "FROM " & sMDB & ".dbo.TPRODUCTO as T1 LEFT OUTER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA as T2 ON T1.tEnlace = t2.tRecetaVenta  AND tLocal='" & sLocal & "'"
'   CnAlmacen.Execute Isql
'
'
'
'   Isql = "Update " & sMDB & ".dbo.TPROPIEDAD set " & _
'          "nInsumo = T2.nInsumo, nGasto = T2.nGastoGeneral, nManoObra= T2.nManoObra  " & _
'          "FROM " & sMDB & ".dbo.TPROPIEDAD as T1 LEFT OUTER JOIN " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD as T2 ON T1.tEnlace = t2.tRecetaPropiedad  AND tLocal='" & sLocal & "'"
'   CnAlmacen.Execute Isql
'
'   'Actualiza los DD Costos en los productos de venta
'   Isql = "Update " & sMDB & ".dbo.TPRODUCTO " & _
'          "Set nInsumo= T2.nPrecioPromedio, nInsumo2= T2.nPrecioPromedio, nInsumo3= T2.nPrecioPromedio, nInsumo4= T2.nPrecioPromedio, nInsumo5= T2.nPrecioPromedio " & _
'          "FROM " & sMDB & ".dbo.TPRODUCTO as T1 LEFT OUTER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO as T2 ON T2.tCodigoProducto = t1.tEnlace " & _
'          "where lDDirecto=1"
'   CnAlmacen.Execute Isql
'
'   'Actualiza los DD Costos en los productos de propiedades
'   Isql = "Update " & sMDB & ".dbo.TPROPIEDAD " & _
'          "Set nInsumo= T2.nPrecioPromedio " & _
'          "FROM " & sMDB & ".dbo.TPROPIEDAD as T1 LEFT OUTER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO as T2 ON T2.tCodigoProducto = t1.tEnlace " & _
'          "where lDDirecto=1"
'   CnAlmacen.Execute Isql
'
'
   

   
   Screen.MousePointer = vbDefault
   Exit Sub
End Sub
Public Sub Integra()
    frmServidores.cargaModo False
    frmServidores.llenaGrid
End Sub

Public Sub RTipoCambio()
   'TIPO DE CAMBIO
   Dim RsTipoCambio As Recordset
   
   Isql = "select * from TTIPOCAMBIO WHERE CONVERT(NVARCHAR,fFecha,103)= '" & FechaServidorTipoCambio() & "' "
   Set RsTipoCambio = Lib.OpenRecordset(Isql, Cn)
     
   If RsTipoCambio.RecordCount = 0 Then
                     
                  Dim oComando As clsComando
                  Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("spIns_TipoCambio", Cn) Then
                     Set oComando = Nothing
                     Exit Sub
                  End If

                  oComando.CreateParameter "@nTc", adDouble, adParamInput, 0, 1
                  oComando.CreateParameter "@tUSUARIO", adVarChar, adParamInput, 15, ""
                  oComando.CreateParameter "@nTco", adDouble, adParamInput, 0, 0
                  oComando.CreateParameter "@nTc2", adDouble, adParamInput, 0, 1
                  oComando.CreateParameter "@nTc3", adDouble, adParamInput, 0, 1
                  If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
   End If
End Sub
