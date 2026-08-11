VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmImportacion 
   Caption         =   "Importación de Data de El Polo"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   Icon            =   "frmImportacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Height          =   1230
      Left            =   45
      TabIndex        =   5
      Top             =   810
      Width           =   3750
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   330
         Left            =   1440
         TabIndex        =   8
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51707905
         CurrentDate     =   38604
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   330
         Left            =   1440
         TabIndex        =   9
         Top             =   675
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51707905
         CurrentDate     =   38604
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final :"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   743
         Width           =   915
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial :"
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   6
         Top             =   338
         Width           =   990
      End
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
      Height          =   555
      Left            =   2520
      Picture         =   "frmImportacion.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2115
      Width           =   1275
   End
   Begin VB.CommandButton cmdAceptar 
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
      Height          =   555
      Left            =   1215
      Picture         =   "frmImportacion.frx":03FC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2115
      Width           =   1275
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   2745
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
      EndProperty
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
   Begin VB.CheckBox chkMovimiento 
      Caption         =   "Actualizar Movimientos de El Polo"
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
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Value           =   1  'Checked
      Width           =   3705
   End
   Begin VB.CheckBox chkTablas 
      Caption         =   "Actualizar Tablas de El Polo"
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
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Value           =   1  'Checked
      Width           =   3705
   End
End
Attribute VB_Name = "frmImportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkMovimiento_Click()
   If chkMovimiento.Value Then
      dtpInicial.Enabled = True
      dtpFinal.Enabled = True
   Else
      dtpInicial.Enabled = False
      dtpFinal.Enabled = False
   End If
End Sub

Private Sub cmdAceptar_Click()
   If chkTablas.Value Then
      Screen.MousePointer = vbHourglass
      StatusBar.Panels(1).Text = "Actualizando Tablas"
      CnDBF.Execute "delete from TTABLA"
      CnDBF.Execute "insert into TTABLA SELECT * FROM OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User ID=sa;Password=sistemas')." & sMDB3 & ".dbo.TTABLA"
       
      StatusBar.Panels(1).Text = "Actualizando Clientes"
      CnDBF.Execute "delete from TCLIENTE"
      CnDBF.Execute "insert into TCLIENTE SELECT * FROM OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User ID=sa;Password=sistemas')." & sMDB3 & ".dbo.TCLIENTE"
        
      StatusBar.Panels(1).Text = "Actualizando Compañías"
      CnDBF.Execute "delete from TCOMPANIA"
      CnDBF.Execute "insert into TCOMPANIA SELECT * FROM OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User ID=sa;Password=sistemas')." & sMDB3 & ".dbo.TCOMPANIA"
        
      StatusBar.Panels(1).Text = "Actualizando Grupos"
      CnDBF.Execute "delete from TGRUPO"
      CnDBF.Execute "insert into TGRUPO SELECT * FROM OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User ID=sa;Password=sistemas')." & sMDB3 & ".dbo.TGRUPO"
        
      StatusBar.Panels(1).Text = "Actualizando SubGrupos"
      CnDBF.Execute "delete from TSUBGRUPO"
      CnDBF.Execute "insert into TSUBGRUPO SELECT * FROM OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User ID=sa;Password=sistemas')." & sMDB3 & ".dbo.TSUBGRUPO"
        
      StatusBar.Panels(1).Text = "Actualizando Productos"
      CnDBF.Execute "delete from TPRODUCTO"
      CnDBF.Execute "insert into TPRODUCTO SELECT * FROM OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User ID=sa;Password=sistemas')." & sMDB3 & ".dbo.TPRODUCTO"
        
      'StatusBar.Panels(1).Text = "Actualizando Ofertas"
      'CnDBF.Execute "delete from TOFERTA"
      'CnDBF.Execute "insert into TOFERTA SELECT * FROM OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User ID=sa;Password=sistemas')." & sMDB3 & ".dbo.TOFERTA"
        
      StatusBar.Panels(1).Text = "Actualizando Combos"
      CnDBF.Execute "delete from TCOMBO"
      CnDBF.Execute "insert into TCOMBO SELECT * FROM OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User ID=sa;Password=sistemas')." & sMDB3 & ".dbo.TCOMBO"
        
      StatusBar.Panels(1).Text = "Actualizando Propiedades"
      CnDBF.Execute "delete from TPROPIEDAD"
      CnDBF.Execute "insert into TPROPIEDAD SELECT * FROM OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User ID=sa;Password=sistemas')." & sMDB3 & ".dbo.TPROPIEDAD"
        
      StatusBar.Panels(1).Text = "Actualizando Mesas"
      CnDBF.Execute "delete from TMESA"
      CnDBF.Execute "insert into TMESA SELECT * FROM OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User ID=sa;Password=sistemas')." & sMDB3 & ".dbo.TMESA"
        
      StatusBar.Panels(1).Text = "Proceso Concluido"
      Screen.MousePointer = vbDefault
   End If
   
   If chkMovimiento.Value Then
      Screen.MousePointer = vbHourglass
      StatusBar.Panels(1).Text = "Actualizando Pedidos, Anulaciones"
      CnDBF.Execute "DELETE APEDIDO where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      CnDBF.Execute "INSERT into APEDIDO select * from OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.APEDIDO where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
   
      StatusBar.Panels(1).Text = "Actualizando Pedidos, Combos"
      CnDBF.Execute "DELETE CPEDIDO FROM dbo.CPEDIDO LEFT OUTER JOIN dbo.MPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      CnDBF.Execute "INSERT into CPEDIDO select T1.* from " & _
                    "OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.CPEDIDO T1 LEFT OUTER JOIN " & _
                    "OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.MPEDIDO T2 ON " & _
                    "T1.tCodigoPedido = T2.tCodigoPedido where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
                 
      StatusBar.Panels(1).Text = "Actualizando Pedidos, Cabecera"
      CnDBF.Execute "DELETE MPEDIDO  where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      CnDBF.Execute "INSERT into MPEDIDO select * from OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.MPEDIDO where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
                 
      StatusBar.Panels(1).Text = "Actualizando Pedidos, Detalle"
      CnDBF.Execute "DELETE DPEDIDO  where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      CnDBF.Execute "INSERT into DPEDIDO select * from OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.DPEDIDO where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
       
      StatusBar.Panels(1).Text = "Actualizando Documentos, Cabecera"
      CnDBF.Execute "DELETE MDOCUMENTO  where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      CnDBF.Execute "INSERT into MDOCUMENTO select * from OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.MDOCUMENTO where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      
      StatusBar.Panels(1).Text = "Actualizando Documentos, Detalle"
      CnDBF.Execute "DELETE DDOCUMENTO FROM dbo.DDOCUMENTO LEFT OUTER JOIN dbo.MDOCUMENTO ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      CnDBF.Execute "INSERT into DDOCUMENTO select T1.* from " & _
                    "OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.DDOCUMENTO T1 LEFT OUTER JOIN " & _
                    "OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.MDOCUMENTO T2 ON " & _
                    "T1.tDocumento = T2.tDocumento where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
                  
      StatusBar.Panels(1).Text = "Actualizando Documentos, Pagos"
      CnDBF.Execute "DELETE DPAGODOCUMENTO  where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      CnDBF.Execute "INSERT into DPAGODOCUMENTO select * from OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.DPAGODOCUMENTO where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"

      StatusBar.Panels(1).Text = "Actualizando Recibos de Egreso"
      CnDBF.Execute "DELETE MEGRESO where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      CnDBF.Execute "INSERT into MEGRESO select * from OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.MEGRESO where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      
      StatusBar.Panels(1).Text = "Actualizando Recibos de Ingreso"
      CnDBF.Execute "DELETE MINGRESO  where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      CnDBF.Execute "INSERT into MINGRESO select * from OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.MINGRESO where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      
      StatusBar.Panels(1).Text = "Actualizando Nota de Creditos"
      CnDBF.Execute "DELETE MNOTACREDITO  where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      CnDBF.Execute "INSERT into MNOTACREDITO select * from OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.MNOTACREDITO where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
   
      StatusBar.Panels(1).Text = "Actualizando Prepagos"
      CnDBF.Execute "DELETE DPREPAGO  where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      CnDBF.Execute "INSERT into DPREPAGO select * from OPENDATASOURCE('SQLOLEDB','Data Source=" & xRuta & ";User Id=sa;Password=sistemas')." & sMDB3 & ".dbo.DPREPAGO where fRegistro >= '" & Format(dtpInicial, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(dtpFinal, "yyyy/MM/dd") & " 23:59'"
      
      StatusBar.Panels(1).Text = "Proceso Concluido"
      Screen.MousePointer = vbDefault
   End If
   
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub dtpFinal_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
   If dtpInicial.Value > dtpFinal.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFinal.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   dtpInicial.Value = Date
   dtpFinal.Value = Date
End Sub

