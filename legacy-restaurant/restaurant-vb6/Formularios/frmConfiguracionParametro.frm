VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfiguracionParametro 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Otros 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Configuracion KDS"
      TabPicture(0)   =   "frmConfiguracionParametro.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "btnAceptar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "btnCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ChklKDSInforest"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkKDS"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmConfiguracionParametro.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdOpcion(0)"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   -67440
         Picture         =   "frmConfiguracionParametro.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3480
         Width           =   1170
      End
      Begin VB.Frame Frame3 
         Caption         =   "Descargo de Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74760
         TabIndex        =   30
         Top             =   600
         Width           =   2895
         Begin VB.CheckBox chkDescargoAutomatico 
            Alignment       =   1  'Right Justify
            Caption         =   "Descargo Automatico"
            Height          =   375
            Left            =   210
            TabIndex        =   31
            Top             =   360
            Width           =   2310
         End
         Begin MSComCtl2.DTPicker dtpHoraDescargo 
            Height          =   315
            Left            =   1200
            TabIndex        =   32
            Top             =   795
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
            CustomFormat    =   "HH:mm 'HRS'"
            Format          =   54067203
            UpDown          =   -1  'True
            CurrentDate     =   38587.2083333333
         End
         Begin VB.Label Label7 
            Caption         =   "Ejecucion :"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkKDS 
         Caption         =   "KDS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox ChklKDSInforest 
         Caption         =   "Activar KDS Inforest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   28
         Top             =   720
         Width           =   2895
      End
      Begin VB.Frame Frame1 
         Caption         =   "KDS Externos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   3960
         TabIndex        =   6
         Top             =   1320
         Width           =   4650
         Begin VB.CommandButton cmdBump 
            Caption         =   "..."
            Height          =   345
            Left            =   3870
            TabIndex        =   18
            Top             =   2115
            Width           =   615
         End
         Begin VB.TextBox txtOrderStatus 
            Height          =   330
            Left            =   135
            TabIndex        =   17
            Top             =   1320
            Width           =   3645
         End
         Begin VB.CommandButton cmdBump2 
            Caption         =   "..."
            Height          =   345
            Left            =   8880
            TabIndex        =   16
            Top             =   3435
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdOrderStatus2 
            Caption         =   "..."
            Height          =   345
            Left            =   8880
            TabIndex        =   15
            Top             =   2640
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtBumpx 
            Height          =   330
            Left            =   4800
            TabIndex        =   14
            Top             =   3435
            Visible         =   0   'False
            Width           =   4005
         End
         Begin VB.TextBox txtOrderStatusx 
            Height          =   330
            Left            =   4800
            TabIndex        =   13
            Top             =   2640
            Visible         =   0   'False
            Width           =   4005
         End
         Begin VB.CommandButton cmdOrderInfo2 
            Caption         =   "..."
            Height          =   345
            Left            =   8880
            TabIndex        =   12
            Top             =   1845
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtOrderInfox 
            Height          =   330
            Left            =   4800
            TabIndex        =   11
            Top             =   1845
            Visible         =   0   'False
            Width           =   4005
         End
         Begin VB.CommandButton cmdOrderStatus 
            Caption         =   "..."
            Height          =   345
            Left            =   3870
            TabIndex        =   10
            Top             =   1320
            Width           =   615
         End
         Begin VB.CommandButton cmdOrderInfo 
            Caption         =   "..."
            Height          =   345
            Left            =   3870
            TabIndex        =   9
            Top             =   525
            Width           =   615
         End
         Begin VB.TextBox txtBump 
            Height          =   330
            Left            =   135
            TabIndex        =   8
            Top             =   2115
            Width           =   3645
         End
         Begin VB.TextBox txtOrderInfo 
            Height          =   330
            Left            =   135
            TabIndex        =   7
            Top             =   525
            Width           =   3645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Solo se debera configurar el KDS 2, cuando se tenga 2 modelos de KDS en un mismo local"
            Height          =   195
            Left            =   1560
            TabIndex        =   27
            Top             =   3885
            Visible         =   0   'False
            Width           =   6450
         End
         Begin VB.Label Label5 
            Caption         =   "Fuente de Notificacion del Bump Bar (BumpNotification)"
            Height          =   210
            Left            =   4800
            TabIndex        =   26
            Top             =   2280
            Width           =   4170
         End
         Begin VB.Label Label4 
            Caption         =   "Destino de archivo de informacion LS (orderstatus)"
            Height          =   255
            Left            =   4800
            TabIndex        =   25
            Top             =   1515
            Width           =   3690
         End
         Begin VB.Label Label3 
            Caption         =   "Destino de archivos de las ordenes (orderinfo)"
            Height          =   225
            Left            =   4800
            TabIndex        =   24
            Top             =   1635
            Visible         =   0   'False
            Width           =   3345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Modelo de KDS - 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4800
            TabIndex        =   23
            Top             =   1200
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Modelo de KDS - 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3960
            TabIndex        =   22
            Top             =   2760
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Line Line1 
            X1              =   4680
            X2              =   4680
            Y1              =   360
            Y2              =   2880
         End
         Begin VB.Label lblBump 
            Caption         =   "Fuente de Notificacion del Bump Bar (BumpNotification)"
            Height          =   210
            Left            =   120
            TabIndex        =   21
            Top             =   1800
            Width           =   4170
         End
         Begin VB.Label lblOrderStatus 
            Caption         =   "Destino de archivo de informacion LS (orderstatus)"
            Height          =   255
            Left            =   135
            TabIndex        =   20
            Top             =   1035
            Width           =   3690
         End
         Begin VB.Label lblOrderInfo 
            Caption         =   "Destino de archivos de las ordenes (orderinfo)"
            Height          =   225
            Left            =   135
            TabIndex        =   19
            Top             =   315
            Width           =   3345
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "KDS - Infomatica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   3135
         Begin VB.CheckBox ChklKDSxArea 
            Caption         =   "Vista por Areas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   2415
         End
         Begin VB.CheckBox ChklKDSTodos 
            Caption         =   "Vista Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   960
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar"
         Height          =   600
         Left            =   2205
         TabIndex        =   2
         Top             =   3000
         Width           =   1400
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar"
         Height          =   600
         Left            =   480
         TabIndex        =   1
         Top             =   3015
         Width           =   1400
      End
   End
End
Attribute VB_Name = "frmConfiguracionParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDescAlmacen  As clsAlmacen

Private Sub btnAceptar_Click()
    Call KDS_GrabarPath(Me.txtOrderInfo.Text, Me.txtOrderStatus.Text, Me.txtBump.Text, Me.txtOrderInfox.Text, Me.txtOrderStatusx.Text, Me.txtBumpx.Text, ChklKDSInforest.Value, ChklKDSxArea.Value, ChklKDSTodos.Value, chkKDS.Value)
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub chkKDS_Click()
    ChklKDSInforest.Enabled = chkKDS.Value
    ChklKDSxArea.Enabled = chkKDS.Value
    ChklKDSTodos.Enabled = chkKDS.Value
    cmdOrderInfo.Enabled = chkKDS.Value
    cmdOrderStatus.Enabled = chkKDS.Value
    cmdBump.Enabled = chkKDS.Value
End Sub

Private Sub ChklKDSInforest_Click()
    If ChklKDSInforest.Value Then
        Frame1.Enabled = False
    Else
        Frame1.Enabled = True
    End If
End Sub

Private Sub ChklKDSTodos_Click()
  If ChklKDSTodos.Value Then
        ChklKDSxArea.Value = False
    End If
End Sub

Private Sub ChklKDSxArea_Click()
  If ChklKDSxArea.Value Then
        ChklKDSTodos.Value = False
    End If
End Sub

Private Sub cmdBump_Click()
     Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")
    If (ret <> "") Then
        txtBump.Text = ret
    End If
End Sub

Private Sub cmdBump2_Click()
 Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")
    If (ret <> "") Then
        txtBumpx.Text = ret
    End If
End Sub

Private Sub CargarConfigDescargoVentas()
    Dim RsDescargo As New Recordset
    
    Set RsDescargo = Lib.OpenRecordset("select lDescargoAutomatico,tHoraDescargo from TPARAMETRO", Cn)
    If Not IsNull(RsDescargo!lDescargoAutomatico) And Not IsNull(RsDescargo!tHoraDescargo) Then
        chkDescargoAutomatico.Value = IIf(RsDescargo!lDescargoAutomatico, 1, 0)
        dtpHoraDescargo.Value = RsDescargo!tHoraDescargo
    End If
    
End Sub

Private Sub SaveDescargoVentas()
    Dim DescargoAutomatico As Boolean
    DescargoAutomatico = False
    Dim RsDescargo As New Recordset
    
    Set clsDescAlmacen = New clsAlmacen
    
    Set RsDescargo = Lib.OpenRecordset("select lDescargoAutomatico,tHoraDescargo from TPARAMETRO", Cn)
    
    If chkDescargoAutomatico.Value Then
    
        If Not IsNull(RsDescargo!tHoraDescargo) Then
            If (Format(RsDescargo!tHoraDescargo, "HH:mm") <> Format(dtpHoraDescargo.Value, "HH:mm")) Or (RsDescargo!lDescargoAutomatico <> chkDescargoAutomatico.Value) Then
            
                Isql = "UPDATE tparametro set lDescargoAutomatico=1, tHoraDescargo='" & Format(dtpHoraDescargo.Value, "HH:mm") & "'"
                Cn.Execute Isql
                
                DescargoAutomatico = clsDescAlmacen.EjecutarDescargoAutomatico("ServicioDescargoAutomatico", False)
                DescargoAutomatico = clsDescAlmacen.EjecutarDescargoAutomatico("ServicioDescargoAutomatico", True)
                
            Else
             DescargoAutomatico = True
            End If
        End If
        
        If DescargoAutomatico Then
            Isql = "UPDATE tparametro set lDescargoAutomatico=1, tHoraDescargo='" & Format(dtpHoraDescargo.Value, "HH:mm") & "'"
            
        Else
            MsgBox "No fue posible iniciar o activar el servicio automático." & vbCrLf & _
            "Por favor, inténtelo de nuevo de forma manual o contacte a soporte técnico para verificar la disponibilidad de la función.", vbCritical, "Descargo Automatico"
               
            chkDescargoAutomatico.Value = 0
            Isql = "UPDATE tparametro set lDescargoAutomatico=0, tHoraDescargo='" & Format(dtpHoraDescargo.Value, "HH:mm") & "'"
        End If
        
    Else
        DescargoAutomatico = clsDescAlmacen.EjecutarDescargoAutomatico("ServicioDescargoAutomatico", False)
        If Not DescargoAutomatico Then
            MsgBox "No fue posible detener o desactivar el servicio automático." & vbCrLf & _
            "Por favor, inténtelo de nuevo de forma manual o contacte a soporte técnico para verificar la disponibilidad de la función.", vbCritical, "Descargo Automatico"
               
            chkDescargoAutomatico.Value = 0
        End If
        
        Isql = "UPDATE tparametro set lDescargoAutomatico=0, tHoraDescargo='" & Format(dtpHoraDescargo.Value, "HH:mm") & "'"
    End If

    Cn.Execute Isql
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
  
  SaveDescargoVentas
  MsgBox "Parámetros Actualizados", vbInformation, sMensaje

End Sub

Private Sub cmdOrderInfo_Click()
    Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")
    If (ret <> "") Then
        txtOrderInfo.Text = ret
    End If
End Sub

Function Buscar_Carpeta(Optional Titulo As String, Optional Path_Inicial As Variant) As String
On Local Error GoTo errFunction
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
       
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder(0, Titulo, 0, Path_Inicial)
       
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
       
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path
Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString
End Function

Private Sub cmdOrderInfo2_Click()
 Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")
    If (ret <> "") Then
        txtOrderInfox.Text = ret
    End If
End Sub

Private Sub cmdOrderStatus_Click()
    Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")
    If (ret <> "") Then
        txtOrderStatus.Text = ret
    End If
End Sub

Sub KDS_GrabarPath(ByVal OrderInfo As String, ByVal OrderStatus As String, ByVal Bump As String, ByVal OrderInfox As String, ByVal OrderStatusx As String, ByVal Bumpx As String, ByVal info As Boolean, ByVal Area As Boolean, ByVal Todos As Boolean, ByVal lKDS As Boolean)
    Cn.Execute "USP_KDS_GrabarPath '" & OrderInfo & "', '" & OrderStatus & "', '" & Bump & "','" & OrderInfox & "', '" & OrderStatusx & "', '" & Bumpx & "', " & IIf(info, 1, 0) & " , " & IIf(Area, 1, 0) & " , " & IIf(Todos, 1, 0) & " , " & IIf(lKDS, 1, 0) & ""
End Sub

Private Sub cmdOrderStatus2_Click()
    Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")
    If (ret <> "") Then
        txtOrderStatusx.Text = ret
    End If
End Sub

Private Sub Form_Load()
    Dim RsPath As New Recordset
    Set RsPath = Lib.OpenRecordset("USP_KDS_ObtenerPath", Cn)
    Me.txtOrderInfo.Text = RsPath!tOrderInfo
    Me.txtOrderStatus.Text = RsPath!tOrderStatus
    Me.txtBump.Text = RsPath!tBump
    
    Me.txtOrderInfox.Text = RsPath!tOrderInfox
    Me.txtOrderStatusx.Text = RsPath!tOrderStatusx
    Me.txtBumpx.Text = RsPath!tBumpx
    
    ChklKDSInforest.Value = IIf(IsNull(RsPath!lKDSInforest), 0, IIf(RsPath!lKDSInforest = True, 1, 0))
    ChklKDSxArea.Value = IIf(IsNull(RsPath!lKDSxArea), 0, IIf(RsPath!lKDSxArea = True, 1, 0))
    ChklKDSTodos.Value = IIf(IsNull(RsPath!lKDSTodos), 0, IIf(RsPath!lKDSTodos = True, 1, 0))
    chkKDS.Value = IIf(IsNull(RsPath!lKDS), 0, IIf(RsPath!lKDS = True, 1, 0))
    
    
    chkKDS_Click
    CargarConfigDescargoVentas
    
End Sub
