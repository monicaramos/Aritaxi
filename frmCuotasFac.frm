VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCuotasFac 
   Caption         =   "Facturación de Cuotas"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   2130
      TabIndex        =   4
      Top             =   2790
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   2910
      TabIndex        =   32
      Top             =   2790
      Width           =   3825
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   5250
      Visible         =   0   'False
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   360
      Top             =   5520
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   2130
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   2910
      TabIndex        =   28
      Top             =   2400
      Width           =   3825
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4290
      TabIndex        =   11
      Top             =   5640
      Width           =   1135
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5610
      TabIndex        =   12
      Top             =   5640
      Width           =   1135
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmCuotasFac.frx":0000
      Left            =   2100
      List            =   "frmCuotasFac.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2100
      TabIndex        =   2
      Text            =   "99/99/9999"
      Top             =   1920
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   2100
      TabIndex        =   0
      Top             =   1080
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   2100
      TabIndex        =   1
      Top             =   1470
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   3300
      TabIndex        =   14
      Top             =   1080
      Width           =   3435
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   3300
      TabIndex        =   13
      Top             =   1470
      Width           =   3435
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   3630
      Width           =   6495
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3390
         TabIndex        =   8
         Text            =   "99/99/9999"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3390
         TabIndex        =   7
         Text            =   "99/99/9999"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3390
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3030
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3030
         ToolTipText     =   "Buscar fecha"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2430
         TabIndex        =   24
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2430
         TabIndex        =   23
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha de servicios:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label Label4 
         Caption         =   "Mes de cuota:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Index           =   1
      Left            =   240
      TabIndex        =   25
      Top             =   3630
      Width           =   6495
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   6
         Left            =   2400
         MaxLength       =   80
         TabIndex        =   10
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   1485
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   2040
         Tag             =   "-1"
         ToolTipText     =   "Ver concepto"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Forma Pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   33
      Top             =   2820
      Width           =   1485
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   1800
      Tag             =   "-1"
      ToolTipText     =   "Buscar forma de pago"
      Top             =   2820
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Facturación Cuotas a Socios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   210
      TabIndex        =   31
      Top             =   300
      Width           =   5355
   End
   Begin VB.Label Label11 
      Caption         =   "Cuenta Cobro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      TabIndex        =   29
      Top             =   2400
      Width           =   1515
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1800
      Tag             =   "-1"
      ToolTipText     =   "Buscar cuenta"
      Top             =   2400
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Cuota:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3270
      Width           =   1605
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1920
      Width           =   1515
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   2
      Left            =   1800
      ToolTipText     =   "Buscar fecha"
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label Label5 
      Caption         =   "Socios"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   840
      Width           =   765
   End
   Begin VB.Label Label7 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1110
      TabIndex        =   16
      Top             =   1110
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1110
      TabIndex        =   15
      Top             =   1500
      Width           =   615
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   3
      Left            =   1800
      Tag             =   "-1"
      ToolTipText     =   "Buscar Socio"
      Top             =   1470
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   1800
      Tag             =   "-1"
      ToolTipText     =   "Buscar Socio"
      Top             =   1110
      Width           =   240
   End
End
Attribute VB_Name = "frmCuotasFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmSoc As frmGesSocios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCal As frmCal
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios
Attribute frmMtoBancosPro.VB_VarHelpID = -1
Private WithEvents frmFPag As frmFacFormasPago
Attribute frmFPag.VB_VarHelpID = -1

Dim Fecha As Date
Dim Cad As String
'variables para el report
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Integer

Dim Servicios As Long

Dim almac As String
Dim Prove As String
Dim NomArtic As String
Dim NomArtic2 As String
Dim NomArtic3 As String
Dim NomArtic4 As String
Dim CodTraba As String
Dim codtipom As String
Dim iva As String
Dim porIva As Currency
Dim LetraSer  As String
Dim indCodigo As String

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim CADENA As String
Dim b As Boolean

    Screen.MousePointer = vbHourglass
    
    If Not DatosOk Then
        Exit Sub
        Screen.MousePointer = vbDefault
    End If
    
    '[Monica]04/12/2012: si no hay registros a seleccionar damos un aviso.
    If VerSocios(CADENA) Then
        If TotalRegistrosConsulta(CADENA) = 0 Then
            MsgBox "No existen socios a facturar. Revise.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    'CUOTA NORMAL
    If Combo1.ListIndex = 1 Then '2 Then
        If datosok1 Then
            DesBloqueoManual ("CUOTAFAC") 'facturas de cuotas
            If Not BloqueoManual("CUOTAFAC", "1") Then
                MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            b = GenerarFactura(True)
            HacerImpresionFacturas
            DesBloqueoManual ("CUOTAFAC")
            TerminaBloquear
        Else
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
    'CUOTA EXTRAORINARIA
    ElseIf Combo1.ListIndex = 0 Then '1 Then
        If datosok2 Then
            DesBloqueoManual ("CUOTAFAC") 'facturas de cuotas
            If Not BloqueoManual("CUOTAFAC", "1") Then
                MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            b = GenerarFactura(False)
            HacerImpresionFacturas
            DesBloqueoManual ("CUOTAFAC")
            TerminaBloquear
        Else
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        MsgBox "Es necesario seleccionar un tipo de cuota a facturar.", vbExclamation
    End If
    Screen.MousePointer = vbDefault
    
    If b Then
        cmdCancelar_Click
    End If
    
End Sub

Private Sub HacerImpresionFacturas()
    cadFormula = "(" & cadFormula & " and {scafac.fecfactu}= Date(" & Year(Text1(1).Text) & "," & Month(Text1(1).Text) & "," & Day(Text1(1).Text) & "))"
    LlamarImprimir False
End Sub

Private Sub LlamarImprimir(duplicado As Boolean)
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        cadParam = cadParam & "pDuplicado= " & Abs(duplicado) & " |"
        numParam = 2
        
        With frmImprimir
        .Titulo = "Impresión de Facturas de Cuotas"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "49", "N")
    '------ > Listado 49 = rFactuCuotas.rpt
        .Opcion = 101
        .ConSubInforme = True
        .Show vbModal
    End With

End Sub

Private Function GenerarFactura(normal As Boolean) As Boolean
Dim vC As CTiposMov
Dim fac As CFactura
Dim Cad As String
Dim Sql As String
Dim totfactu As Currency
Dim BaseImp As Currency
Dim base0 As Currency
Dim base1 As Currency
Dim base2 As Currency
Dim base4 As Currency
Dim ImpIVA As Currency
Dim cli As CCliente
Dim b As Boolean
Dim CADENA As String
Dim LetraSer As String
Dim ForPago As Integer
Dim FecFactu As Date
Dim NumFactu As Long
Dim codtipom As String
Dim Cantidad As Currency
Dim total As Currency
Dim i As Currency
Dim J As Integer
Dim SqlArt As String
Dim RsArt As ADODB.Recordset
Dim SQL2 As String

Dim cad1 As String


    On Error GoTo EGenerarFacturas
    
    ' vamos a protegerlo con transacciones
    conn.BeginTrans
    ConnConta.BeginTrans
    
    
    'guardo el contador inicial por si falla para volver a guardarlo
    Set miRsAux = New ADODB.Recordset
    If normal Then
        codtipom = "FCN"
    Else
        codtipom = "FCE"
    End If
    
    'valores grales para todos los socios
    porIva = CCur(DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T"))
    LetraSer = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", codtipom, "T")
    ForPago = Text1(8).Text
    CodTraba = DevuelveDesdeBD(conAri, "codtraba", "straba", "login", vUsu.Login, "T")
    If CodTraba = "" Then CodTraba = DevuelveValor("select min(codtraba) from straba")
    'busco el minimo almacen y el minimo proveedor
    Sql = "select min(codalmac) from salmpr"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        almac = miRsAux.Fields(0)
    End If
        
    miRsAux.Close
        
    Sql = "select min(codprove) from sprove"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        Prove = miRsAux.Fields(0)
    End If
    
    'generaremos una factura por cada socio que hallamos seleccionado
    If Not VerSocios(Sql) Then
        conn.RollbackTrans
        ConnConta.RollbackTrans
        Exit Function
    End If
    
    Set miRsAux = Nothing
    
    PB1.visible = True
    
    Cantidad = 0
    Data1.RecordSource = Sql
    Data1.Refresh
    total = Data1.Recordset.RecordCount
    
    b = True
    
    'inicializamos cadenas
    Cad = ""
    While Not Data1.Recordset.EOF
        Cantidad = Cantidad + 1
        PB1.Value = Cantidad * 100 / total
        Set vC = New CTiposMov
        Set cli = New CCliente
        Set fac = New CFactura

        If vC.TipoMovimiento <> codtipom Then
            If Not vC.Leer(codtipom) Then
                Data1.Recordset.Close
                If NumRegElim > 0 Then MsgBox "Se han generado " & NumRegElim & " factura(s) antes del error", vbExclamation
                Exit Function
            End If
        End If
        vC.IncrementarContador (vC.TipoMovimiento)
        
        If normal Then 'calculo los servicios

            SqlArt = "select sclien_cuotas.*, sartic.preciove, sartic.nomartic from sclien_cuotas inner join sartic on sclien_cuotas.codartic = sartic.codartic where codsocio = " & DBSet(Data1.Recordset!CodClien, "N")
            
            BaseImp = 0
            
            Servicios = 0
            CalcularServicios BaseImp, base2, NomArtic3, Servicios

            Set RsArt = New ADODB.Recordset
            RsArt.Open SqlArt, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RsArt.EOF
                '[Monica]05/12/2013: modificado ahora cogemos el importe que pueden modificar, antes cogiamos el precio del articulo
                BaseImp = BaseImp + DBLet(RsArt!Importes, "N") 'DBLet(RsArt!preciove, "N")
                RsArt.MoveNext
            Wend
            Set RsArt = Nothing

        Else
            BaseImp = ImporteFormateado(Text1(4).Text)
        End If
        DoEvents
        fac.BaseImp = BaseImp
        fac.BrutoFac = BaseImp
        ImpIVA = (BaseImp * porIva) / 100
        totfactu = BaseImp + ImpIVA
        fac.TotalFac = totfactu
        fac.codtipom = codtipom
        FecFactu = Text1(1).Text
        fac.FecFactu = FecFactu
        fac.LetraSerie = LetraSer
        NumFactu = vC.Contador
        fac.NumFactu = NumFactu
'        fac.CuentaPrev = Text1(7).Text
        fac.ForPago = ForPago
        
        fac.BancoPr = Text1(7).Text
        fac.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", fac.BancoPr, "N")
        
        fac.Agente = vParamAplic.PorDefecto_Agente
        
        'datos del cliente
        fac.Cliente = Data1.Recordset!CodClien
        
'        If fac.Cliente = 10357 Then Stop
        
        cli.Nombre = Data1.Recordset!nomclien
        fac.NombreClien = Data1.Recordset!nomclien
        cli.Domicilio = Data1.Recordset!domclien
        fac.DomicilioClien = Data1.Recordset!domclien
        cli.CPostal = Data1.Recordset!codpobla
        fac.CPostal = Data1.Recordset!codpobla
        cli.Poblacion = Data1.Recordset!pobclien
        fac.Poblacion = Data1.Recordset!pobclien
        cli.Provincia = Data1.Recordset!proclien
        fac.Provincia = Data1.Recordset!proclien
        cli.NIF = Data1.Recordset!nifClien
        fac.NIF = Data1.Recordset!nifClien
        
        '[Monica]22/11/2013:iban
        fac.Iban = Data1.Recordset!Iban
        fac.Banco = DBLet(Data1.Recordset!codbanco, "N")
        fac.Sucursal = DBLet(Data1.Recordset!codsucur, "N")
        fac.DigControl = DBLet(Data1.Recordset!digcontr, "T")
        fac.CuentaBan = DBLet(Data1.Recordset!cuentaba, "T")
    
        'scafac
        Cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "'," & fac.Cliente & ","
        Cad = Cad & DBSet(cli.Nombre, "T") & "," & DBSet(cli.Domicilio, "T") & "," & DBSet(cli.CPostal, "T") & ","
        Cad = Cad & DBSet(cli.Poblacion, "T") & "," & DBSet(cli.Provincia, "T") & "," & DBSet(cli.NIF, "T") & "," & vParamAplic.PorDefecto_Agente
        Cad = Cad & "," & fac.ForPago & ",0,0," & TransformaComasPuntos(CStr(BaseImp)) & ",0,0," & TransformaComasPuntos(CStr(BaseImp)) & "," & iva
        Cad = Cad & "," & TransformaComasPuntos(CStr(porIva)) & "," & TransformaComasPuntos(CStr(ImpIVA)) & "," & TransformaComasPuntos(CStr(totfactu)) & ",0,NULL,"
        Cad = Cad & DBSet(Data1.Recordset!codbanco, "N", "S") & "," & DBSet(Data1.Recordset!codsucur, "N", "S") & "," & DBSet(Data1.Recordset!digcontr, "T", "S") & "," & DBSet(Data1.Recordset!cuentaba, "T", "S") & "," & DBSet(Data1.Recordset!Iban, "T") & ")"
        Sql = "INSERT INTO scafac (codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,"
        Sql = Sql & "nifclien,codagent,codforpa,dtoppago,dtognral,brutofac,impdtopp,impdtogr,baseimp1,codigiv1,porciva1,"
        Sql = Sql & "imporiv1,totalfac,intconta,coddirec, codbanco, codsucur, digcontr, cuentaba, iban) VALUES ("
        Sql = Sql & Cad
        If Not ejecutar(Sql, False) Then
            vC.DevolverContador vC.TipoMovimiento, vC.Contador
            Exit Function
        Else
            'scafac1
            If cadFormula = "" Then
                cadFormula = "{scafac.numfactu}=" & NumFactu
            Else
                cadFormula = cadFormula & " or {scafac.numfactu}=" & NumFactu
            End If
            Cad = ""
            Cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0,'"
            Cad = Cad & Format(FecFactu, FormatoFecha) & "'," & vParamAplic.PorDefecto_Envio & "," & CodTraba
            Cad = Cad & "," & CodTraba & ",NULL,NULL,NULL,NULL,NULL,NULL"
    
            Sql = "INSERT INTO scafac1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,"
            Sql = Sql & "codenvio,codtraba,codtrab2,observa1,observa2,observa3,observa4,observa5,codtrab1) VALUES ("
            Sql = Sql & Cad & ")"
            conn.Execute Sql
            'slifac
            Cad = ""
            If normal Then
'[Monica]20/04/2011
'                i = 1
'
'                '1º linea sera la cuota del mes
'                cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0," & DBSet(i, "N") & "," & almac & ","
'                cad = cad & DBSet(vParamAplic.ArtCuotaSinChofer, "T") & "," & DBSet(NomArtic, "T") & ",1," & TransformaComasPuntos(CStr(base0)) & ","
'                cad = cad & TransformaComasPuntos(CStr(base0)) & "," & TransformaComasPuntos(CStr(base0)) & "," & TransformaComasPuntos(CStr(base0)) & ","
'                cad = cad & TransformaComasPuntos(CStr(base0)) & ",0,0,'M'," & Prove & "," & TransformaComasPuntos(CStr(base0)) & ",1)"
'                SQL = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
'                SQL = SQL & "numbultos,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel,cantidad) VALUES ("
'                SQL = SQL & cad
'                conn.Execute SQL
'
'                '2º linea es la cuota si hay chofer
'                If base1 <> 0 Then
'                    i = 2
'
'                    cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0," & DBSet(i, "N") & "," & almac & ","
'                    cad = cad & DBSet(vParamAplic.ArtCuotaConChofer, "T") & "," & DBSet(NomArtic2, "T") & ",1," & TransformaComasPuntos(CStr(base1)) & ","
'                    cad = cad & TransformaComasPuntos(CStr(base1)) & "," & TransformaComasPuntos(CStr(base1)) & "," & TransformaComasPuntos(CStr(base1)) & ","
'                    cad = cad & TransformaComasPuntos(CStr(base1)) & ",0,0,'M'," & Prove & "," & TransformaComasPuntos(CStr(base1)) & ",1)"
'                    SQL = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
'                    SQL = SQL & "numbultos,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel,cantidad) VALUES ("
'                    SQL = SQL & cad
'                    conn.Execute SQL
'                End If
'
'                If base2 <> 0 Then ' si hay servicios
'                    i = 3
'                    '3º linea seran los servicios
'                    cad = ""
'                    cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0," & DBSet(i, "N") & "," & almac & ","
'                    cad = cad & DBSet(vParamAplic.ArtServCuotas, "T") & "," & DBSet(NomArtic3, "T") & ",1," & TransformaComasPuntos(CStr(vParamAplic.PrecioPorServicio)) & ","
'                    cad = cad & TransformaComasPuntos(CStr(base2)) & "," & TransformaComasPuntos(CStr(base2)) & "," & TransformaComasPuntos(CStr(base2)) & ","
'                    cad = cad & TransformaComasPuntos(CStr(base2)) & ",0,0,'M'," & Prove & "," & TransformaComasPuntos(CStr(base2)) & "," & Servicios & ")"
'                    SQL = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
'                    SQL = SQL & "numbultos,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel,cantidad) VALUES ("
'                    SQL = SQL & cad
'                    conn.Execute SQL
'                End If
'
'                '4º linea el alquiler de equipos si no es socio
'                If Data1.Recordset!essocio = 0 Then
'                    i = 4
'
'                    cad = ""
'                    cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0," & DBSet(i, "N") & "," & almac & ","
'                    cad = cad & DBSet(vParamAplic.ArtAlquiler, "T") & "," & DBSet(NomArtic4, "T") & ",1," & TransformaComasPuntos(CStr(vParamAplic.PrecioPorAlquiler)) & ","
'                    cad = cad & TransformaComasPuntos(CStr(vParamAplic.PrecioPorAlquiler)) & "," & TransformaComasPuntos(CStr(vParamAplic.PrecioPorAlquiler)) & "," & TransformaComasPuntos(CStr(base2)) & ","
'                    cad = cad & TransformaComasPuntos(CStr(vParamAplic.PrecioPorAlquiler)) & ",0,0,'M'," & Prove & "," & TransformaComasPuntos(CStr(vParamAplic.PrecioPorAlquiler)) & ",1)"
'                    SQL = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
'                    SQL = SQL & "numbultos,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel,cantidad) VALUES ("
'                    SQL = SQL & cad
'                    conn.Execute SQL
'                End If
'[Monica]20/04/2011
                i = 1
                If base2 <> 0 Then ' si hay servicios
                    '3º linea seran los servicios
                    Cad = ""
                    Cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0," & DBSet(i, "N") & "," & almac & ","
                    Cad = Cad & DBSet(vParamAplic.ArtServCuotas, "T") & "," & DBSet(NomArtic3, "T") & ",1," & TransformaComasPuntos(CStr(vParamAplic.PrecioPorServicio)) & ","
                    Cad = Cad & TransformaComasPuntos(CStr(base2)) & "," & TransformaComasPuntos(CStr(base2)) & "," & TransformaComasPuntos(CStr(base2)) & ","
                    Cad = Cad & TransformaComasPuntos(CStr(base2)) & ",0,0,'M'," & Prove & "," & TransformaComasPuntos(CStr(base2)) & "," & Servicios & ")"
                    Sql = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
                    Sql = Sql & "numbultos,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel,cantidad) VALUES ("
                    Sql = Sql & Cad
                    conn.Execute Sql
                End If
                
                Set RsArt = New ADODB.Recordset
                RsArt.Open SqlArt, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not RsArt.EOF
                    i = i + 1
                    '[Monica]27/04/2015: para taxivip
                    If vParamAplic.Cooperativa = 1 Then
                        If UCase(vParamAplic.ArtServCuotas) = UCase(DBLet(RsArt!codArtic, "T")) Then
                            NomArtic = Mid(RsArt!NomArtic, 1, 18)
                        Else
                            NomArtic = Mid(RsArt!NomArtic, 1, 18) & " " & Mid(UCase(Combo2.Text), 1, 10) & "-" & Year(CDate(Text1(1).Text))
                        End If
                    Else
                            NomArtic = Mid(RsArt!NomArtic, 1, 18) & " " & Mid(UCase(Combo2.Text), 1, 10) & "-" & Year(CDate(Text1(1).Text))
                    End If
                    
                    Cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0," & DBSet(i, "N") & "," & almac & ","
                    '[Monica]05/12/2013: antes DBSet(RsArt!preciove, "N") ahora importes
                    Cad = Cad & DBSet(RsArt!codArtic, "T") & "," & DBSet(NomArtic, "T") & ",1," & DBSet(RsArt!Importes, "N") & ","
                    Cad = Cad & DBSet(RsArt!Importes, "N") & "," & DBSet(RsArt!Importes, "N") & "," & DBSet(RsArt!Importes, "N") & ","
                    Cad = Cad & DBSet(RsArt!Importes, "N") & ",0,0,'M'," & Prove & "," & DBSet(RsArt!Importes, "N") & ",1)"
                    
                    Sql = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
                    Sql = Sql & "numbultos,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel,cantidad) VALUES ("
                    Sql = Sql & Cad
                    
                    conn.Execute Sql
                                
                    RsArt.MoveNext
                Wend
                Set RsArt = Nothing

            Else

                Cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0,1," & almac & ","
                Cad = Cad & DBSet(vParamAplic.ArtCuotaExtraor, "T") & "," & DBSet(NomArtic, "T") & ",1," & TransformaComasPuntos(CStr(BaseImp)) & ","
                Cad = Cad & TransformaComasPuntos(CStr(BaseImp)) & "," & TransformaComasPuntos(CStr(BaseImp)) & "," & TransformaComasPuntos(CStr(BaseImp)) & ","
                Cad = Cad & TransformaComasPuntos(CStr(BaseImp)) & ",0,0,'M'," & Prove & "," & TransformaComasPuntos(CStr(BaseImp)) & "," & DBSet(Text1(6).Text, "T") & ",1)"
                Sql = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
                Sql = Sql & "numbultos,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel,ampliaci,cantidad) VALUES ("
                Sql = Sql & Cad
                conn.Execute Sql
            End If
            
'[Monica]20/04/2011 antes
'            'insertar en tesoreria
'            fac.Text33csb = NomArtic '"CUOTA MES DE " && v_nommes clipped && " DE " && year(f_fecexp) using "&&&&"
'            fac.Text41csb = "FACTURA: " & Format(DBLet(NumFactu, "N"), "#######") & " DE FECHA " & Format(DBLet(FecFactu, "N"), "dd/mm/yyyy")
'
'            fac.Text42csb = ""
'            If normal Then
'                fac.Text42csb = "REALIZADOS " & Format(Servicios, "###0") & " SERVICIOS EN EL "
'            Else
'                fac.Text42csb = Trim(Text1(6).Text)
'            End If
'            fac.Text43csb = "EMITIDA POR: C.V.L. TELE-TAXI"
'                                '1234567890          1234567890           123456789012                                     34567890
'
'            fac.Text51csb = ""
'            If normal Then
'                fac.Text51csb = "PERIODO DEL " & Format(Text1(0).Text, "dd/mm/yy") & " AL " & Format(Text1(2).Text, "dd/mm/yy")
'                                '123456789012345                                     67890123      4567                                     8901234567
'            End If
'            fac.Text52csb = "N.I.F. Nro " & vParam.CifEmpresa
'
'            fac.Text53csb = ""
'            If normal Then
'                fac.Text53csb = "CUOTA:" & Format(base0, "##0.00") & " CHOF:" & Format(base1, "##0.00") & " SERV:" & Format(base2, "##,##0.00")
'                                '123456                   7890123      45678                   901234      567890                   12345678901234
'            Else
'                fac.Text53csb = "CUOTA:" & Format(BaseImp, "##0.00")
'            End If
'            fac.Text61csb = "CLIENTE: " & cli.Nombre
'            fac.Text62csb = ""
'
''            If Data1.Recordset!essocio = 0 And normal Then
'            If base4 Then
'                fac.Text62csb = "EQUIPOS:" & Format(base4, "##0.00")
'            End If
'
'            fac.Text63csb = ""
'            fac.Text63csb = "N.I.F.: " & DBLet(cli.NIF, "N")
'
'            fac.Text71csb = "B.IMPONIBLE:" & Format(BaseImp, "##0.00") & " I.V.A.    :" & Format(ImpIVA, "##0.00")
'            fac.Text72csb = "IMPORTE TOTAL FACTURA: " & Format(fac.TotalFac, "##0.00")
'            fac.Text73csb = ""
'
'            fac.Text81csb = ""
'            fac.Text82csb = "INSCRITA EN EL REG.DE COOP. Nro.CV-240"
'            fac.Text83csb = "SERVICIOS Y CUOTAS AL 18% I.V.A."
        
'Ahora
'[Monica]22/11/2013: iban
'insertar en tesoreria
'            fac.Text41csb = "FACTURA: " & Format(DBLet(NumFactu, "N"), "#######") & " DE FECHA " & Format(DBLet(FecFactu, "N"), "dd/mm/yyyy")
'            fac.Text43csb = "EMITIDA POR: C.V.L. TELE-TAXI"
'            fac.Text52csb = "INSCRITA EN EL REG.DE COOP. Nro.CV-240"
'            fac.Text61csb = "N.I.F. Nro " & vParam.CifEmpresa
'            fac.Text63csb = "CLIENTE: " & cli.Nombre
'            fac.Text72csb = "N.I.F.: " & DBLet(cli.NIF, "N")
'            fac.Text81csb = "B.IMPONIBLE:" & Format(BaseImp, "##0.00") & " I.V.A. " & Format(porIva, "#0.00") & "%:" & Format(ImpIVA, "##0.00")
'            fac.Text83csb = "IMPORTE TOTAL FACTURA: " & Format(fac.TotalFac, "##0.00")
'
'            fac.Text33csb = ""
'            If normal Then
'                fac.Text33csb = Format(Servicios, "###0") & " SERVICIOS DE " & Format(Text1(0).Text, "dd/mm/yy") & "-" & Format(Text1(2).Text, "dd/mm/yy")
'            Else
'                fac.Text33csb = Trim(Text1(6).Text)
'            End If
'
'            fac.Text42csb = ""
'            fac.Text51csb = ""
'            fac.Text53csb = ""
'            fac.Text62csb = ""
'            fac.Text71csb = ""
'            fac.Text73csb = ""
'            fac.Text82csb = ""
'
'            If normal Then
'                J = 2
'                SQL2 = "select codartic, nomartic, importel from slifac where codtipom = 'FCN' and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
'
'                Set RsArt = New ADODB.Recordset
'                RsArt.Open SQL2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                While Not RsArt.EOF
'                    If RsArt!codArtic <> vParamAplic.ArtServCuotas Then
'                        If fac.Text42csb = "" Then
'                            fac.Text42csb = Mid(DBLet(RsArt!NomArtic, "T"), 1, 33) & ":" & Format(RsArt!ImporteL, "##0.00")
'                        ElseIf fac.Text51csb = "" Then
'                            fac.Text51csb = Mid(DBLet(RsArt!NomArtic, "T"), 1, 33) & ":" & Format(RsArt!ImporteL, "##0.00")
'                        ElseIf fac.Text53csb = "" Then
'                            fac.Text53csb = Mid(DBLet(RsArt!NomArtic, "T"), 1, 33) & ":" & Format(RsArt!ImporteL, "##0.00")
'                        ElseIf fac.Text62csb = "" Then
'                            fac.Text62csb = Mid(DBLet(RsArt!NomArtic, "T"), 1, 33) & ":" & Format(RsArt!ImporteL, "##0.00")
'                        ElseIf fac.Text71csb = "" Then
'                            fac.Text71csb = Mid(DBLet(RsArt!NomArtic, "T"), 1, 33) & ":" & Format(RsArt!ImporteL, "##0.00")
'                        ElseIf fac.Text73csb = "" Then
'                            fac.Text73csb = Mid(DBLet(RsArt!NomArtic, "T"), 1, 33) & ":" & Format(RsArt!ImporteL, "##0.00")
'                        ElseIf fac.Text82csb = "" Then
'                            fac.Text82csb = Mid(DBLet(RsArt!NomArtic, "T"), 1, 33) & ":" & Format(RsArt!ImporteL, "##0.00")
'                        Else
'                            ' no se graba, ya no cabe
'                        End If
'                    Else
'                        fac.Text33csb = fac.Text33csb & ":" & Format(RsArt!ImporteL, "##0.00")
'                    End If
'                    J = J + 1
'                    RsArt.MoveNext
'                Wend
'
'                Set RsArt = Nothing
'          End If

'[Monica]22/11/2013: iban
'            fac.Text33csb = ""
'            fac.Text41csb = ""
'
'            cad1 = "FRA.CUOTA " & Mid(UCase(Combo2.Text), 1, 3) & "-" & Year(CDate(Text1(1).Text)) & " NRO." & Format(DBLet(NumFactu, "N"), "0000000") & " DE " & _
'                        Format(DBLet(FecFactu, "N"), "dd/mm/yy") & " de " & Format(fac.TotalFac, "##0.00") & "€ "
'            If normal Then
'                cad1 = cad1 & Format(Servicios, "###0") & " SERV." & Format(Text1(0).Text, "dd/mm/yy") & "-" & Format(Text1(2).Text, "dd/mm/yy")
'            Else
'                cad1 = Trim(Text1(6).Text)
'            End If
'
'            fac.Text33csb = cad1
'
'            If normal Then
'                J = 2
'                SQL2 = "select codartic, nomartic, importel from slifac where codtipom = 'FCN' and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
'
'                Set RsArt = New ADODB.Recordset
'                RsArt.Open SQL2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                cad1 = ""
'
'                While Not RsArt.EOF
'                    If RsArt!codArtic <> vParamAplic.ArtServCuotas Then
'                        cad1 = cad1 & Mid(DBLet(RsArt!NomArtic, "T"), 1, 6) & ":" & Format(RsArt!ImporteL, "##0.00") & " "
'                    End If
'                    J = J + 1
'                    RsArt.MoveNext
'                Wend
'
'                Set RsArt = Nothing
'                fac.Text41csb = Mid(cad1, 1, 40)
'          End If
            
'[Monica]08/01/2014: otra vez
            fac.Text33csb = ""
            fac.Text41csb = ""
                   '1234567890123                                  456                                  7890123      4
            cad1 = "FRA.CUOTA " & Mid(UCase(Combo2.Text), 1, 3) & "-N." & Format(DBLet(NumFactu, "N"), "0000000") & " " & _
                        Format(DBLet(FecFactu, "N"), "dd/mm/yy") & " "
                                                    ' 56789012
            If normal Then                      '3456      7890
                cad1 = cad1 & Format(Servicios, "###0") & " SER" '& Format(Text1(0).Text, "dd/mm/yy") & "-" & Format(Text1(2).Text, "dd/mm/yy")
'            Else
'                cad1 = Trim(Text1(6).Text)
            End If

            fac.Text41csb = cad1

            If normal Then
                J = 2
                SQL2 = "select codartic, nomartic, importel from slifac where codtipom = 'FCN' and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")

                Set RsArt = New ADODB.Recordset
                RsArt.Open SQL2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

                cad1 = ""

                While Not RsArt.EOF
                    If RsArt!codArtic <> vParamAplic.ArtServCuotas Then
                        cad1 = cad1 & Mid(DBLet(RsArt!NomArtic, "T"), 1, 10) & ":" & Format(RsArt!ImporteL, "##0.00") & " "
                    End If
                    J = J + 1
                    RsArt.MoveNext
                Wend

                Set RsArt = Nothing
                fac.Text33csb = cad1
            Else
                fac.Text33csb = Text1(6)
            End If

            
            
            b = fac.InsertarEnTesoreriaCuotasSoc("Error al pasar a tesoreria")
        
        End If
        Set vC = Nothing
        Data1.Recordset.MoveNext
    Wend
    Data1.Recordset.Close
    PB1.visible = False

EGenerarFacturas:
    If Err.Number <> 0 Or Not b Then
        GenerarFactura = False
        conn.RollbackTrans
        ConnConta.RollbackTrans
        MsgBox "Error al generar facturas: " & Err.Description
    Else
        MsgBox "Proceso realizado correctamente.", vbExclamation
        GenerarFactura = True
        conn.CommitTrans
        ConnConta.CommitTrans
    End If
End Function

Private Sub CalcularServicios(ByRef BaseImp As Currency, ByRef base2 As Currency, ByRef NomArtic2 As String, ByRef Servicios As Long)
Dim RS As ADODB.Recordset
Dim Sql As String

    Set RS = New ADODB.Recordset
    Sql = "select count(*) from shilla where codsocio=" & Data1.Recordset!CodClien
    Sql = Sql & " and fecha >='" & Format(Text1(0).Text, FormatoFecha) & "' and fecha <='"
    Sql = Sql & Format(Text1(2).Text, FormatoFecha) & "'"
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        NomArtic2 = RS.Fields(0) & " servicios del " & Text1(0).Text & "-"
        NomArtic2 = NomArtic2 & Text1(2).Text
        base2 = RS.Fields(0) * vParamAplic.PrecioPorServicio
        BaseImp = BaseImp + base2
        Servicios = RS.Fields(0)
    Else
        NomArtic2 = "0 servicios del " & Text1(0).Text & " al "
        NomArtic2 = NomArtic2 & Text1(2).Text & " a " & vParamAplic.PrecioPorServicio
        base2 = 0
        Servicios = 0
    End If
    RS.Close
    Set RS = Nothing
    
End Sub

Private Function VerSocios(ByRef Sql As String) As Boolean

    VerSocios = False
'    If Text1(3).Text = "" Then
'        If Text1(5).Text = "" Then
'            'SQL = "select * from sclien"
'            SQL = "select * from sclien inner join ssitua on ssitua.codsitua=sclien.codsitua and ssitua.generafactu=1"
'        Else
'            If MsgBox("Si no coloca un codigo de socio desde se facturará a todos. Desea continuar?", vbYesNo = vbYes) Then
'                'SQL = "select * from sclien"
'                SQL = "select * from sclien inner join ssitua on ssitua.codsitua=sclien.codsitua and ssitua.generafactu=1"
'            Else
'                Exit Function
'            End If
'        End If
'    ElseIf Text1(5).Text = "" Then
'            If MsgBox("Si no coloca un codigo de socio hasta se facturará a todos. Desea continuar?", vbYesNo = vbYes) Then
'                'SQL = "select * from sclien"
'                SQL = "select * from sclien inner join ssitua on ssitua.codsitua=sclien.codsitua and ssitua.generafactu=1"
'            Else
'                Exit Function
'            End If
'    Else
'        'SQL = "select * from sclien where codclien >= " & Text1(3).Text & " and codclien <= " & Text1(5).Text
'        SQL = "select * from sclien inner join ssitua on codclien >= " & Text1(3).Text & " and codclien <= " & Text1(5).Text & " and ssitua.codsitua=sclien.codsitua and ssitua.generafactu=1"
'    End If

'[Monica] sustituido por esto seleccionamos los socios que tengan una determinada situacion y tengan asignada una v
    Sql = "select sclien.* from sclien inner join ssitua on ssitua.codsitua=sclien.codsitua and ssitua.generafactu=1 where sclien.numeruve is not null and sclien.numeruve <> 0 "
    If Text1(3).Text <> "" Then Sql = Sql & " and sclien.codclien >= " & DBSet(Text1(3).Text, "N")
    If Text1(5).Text <> "" Then Sql = Sql & " and sclien.codclien <= " & DBSet(Text1(5).Text, "N")
    
'    ' solo los clientes que tengan algun tipo de cuota
'    Sql = Sql & " and sclien.codclien in (select codsocio from sclien_cuotas) "

    VerSocios = True
         
End Function

Private Function datosok1() As Boolean
Dim Sql As String

    datosok1 = False
    If Text1(0).Text = "" Then
        MsgBox "Es necesario introducir una fecha desde para la facturación de los servicios.", vbExclamation
        PonerFoco Text1(0)
        Exit Function
    End If
    If Text1(2).Text = "" Then
        MsgBox "Es necesario introducir una fecha hasta para la facturación de los servicios.", vbExclamation
        PonerFoco Text1(2)
        Exit Function
    End If
    
    If Text1(7).Text <> "" Then
        Sql = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", Text1(7).Text, "N")
        If Sql = "" Then
            MsgBox "La cta. prevista de cobro debe tener valor.", vbExclamation
            Exit Function
        End If
    End If
    ' forma de pago
    If Text1(8).Text <> "" Then
        Sql = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", Text1(8).Text, "N")
        If Sql = "" Then
            MsgBox "La forma de pago debe tener valor.", vbExclamation
            Exit Function
        End If
    End If
    
    
    datosok1 = True
End Function
Private Function datosok2() As Boolean
Dim Sql As String

    datosok2 = False
    If Text1(4).Text <> "" Then
        If IsNumeric(Text1(4).Text) Then
            If Text1(4).Text = 0 Then
                MsgBox "Es necesario introducir un importe para facturar.", vbExclamation
                PonerFoco Text1(4)
                Exit Function
            End If
        Else
            MsgBox "El importe debe ser númerico", vbExclamation
            PonerFoco Text1(4)
            Exit Function
        End If
    Else
        MsgBox "Es necesario introducir un importe para facturar.", vbExclamation
        PonerFoco Text1(4)
        Exit Function
    End If
    
    If Text1(7).Text <> "" Then
        Sql = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", Text1(7).Text, "N")
        If Sql = "" Then
            MsgBox "La cta. prevista de cobro debe tener valor.", vbExclamation
            Exit Function
        End If
    End If
    ' forma de pago
    If Text1(8).Text <> "" Then
        Sql = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", Text1(8).Text, "N")
        If Sql = "" Then
            MsgBox "La forma de pago debe tener valor.", vbExclamation
            Exit Function
        End If
    End If
    
    datosok2 = True
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Combo1_Click()
Dim mes As Byte
Dim Anyo As Integer

    If Combo1.ListIndex = 1 Then '2 Then
    
        If Text1(1).Text = "" Then Exit Sub
    
        Frame1(0).visible = True
        Frame1(1).visible = False
        CargarComboMes
        mes = Month(Text1(1).Text)
        If mes = 1 Then
            Combo2.Text = Combo2.List(11)
        Else
            Combo2.Text = Combo2.List(mes - 1)
        End If
        Anyo = Year(Text1(1).Text)
        Select Case Combo2.ListIndex
        Case 0
            Text1(0).Text = "01/12/" & Anyo - 1
            Text1(2).Text = "31/12/" & Anyo - 1
            PonerFoco Text1(0)
        Case 1
            Text1(0).Text = "01/01/" & Anyo
            Text1(2).Text = "31/01/" & Anyo
        Case 2
            If EsBiciesto(Anyo) Then
                Text1(0).Text = "01/02/" & Anyo
                Text1(2).Text = "29/02/" & Anyo
            Else
                Text1(0).Text = "01/02/" & Anyo
                Text1(2).Text = "28/02/" & Anyo
            End If
        Case 3
            Text1(0).Text = "01/03/" & Anyo
            Text1(2).Text = "31/03/" & Anyo
        Case 4
            Text1(0).Text = "01/04/" & Anyo
            Text1(2).Text = "30/04/" & Anyo
        Case 5
            Text1(0).Text = "01/05/" & Anyo
            Text1(2).Text = "31/05/" & Anyo
        Case 6
            Text1(0).Text = "01/06/" & Anyo
            Text1(2).Text = "30/06/" & Anyo
        Case 7
            Text1(0).Text = "01/07/" & Anyo
            Text1(2).Text = "31/07/" & Anyo
        Case 8
            Text1(0).Text = "01/08/" & Anyo
            Text1(2).Text = "31/08/" & Anyo
        Case 9
            Text1(0).Text = "01/09/" & Anyo
            Text1(2).Text = "30/09/" & Anyo
        Case 10
            Text1(0).Text = "01/10/" & Anyo
            Text1(2).Text = "31/10/" & Anyo
        Case 11
            Text1(0).Text = "01/11/" & Anyo
            Text1(2).Text = "30/11/" & Anyo
        End Select

'        Combo2.SetFocus
        
    ElseIf Combo1.ListIndex = 0 Then '1 Then
        Frame1(1).visible = True
        Frame1(0).visible = False
    Else
        Frame1(0).visible = False
        Frame1(1).visible = False
    End If
End Sub

Private Sub Combo2_click()
Dim Anyo As Integer

    Anyo = Year(Text1(1).Text)

    Select Case Combo2.ListIndex
        Case 0
            Text1(0).Text = "01/12/" & Anyo - 1
            Text1(2).Text = "31/12/" & Anyo - 1
            PonerFoco Text1(0)
        Case 1
            Text1(0).Text = "01/01/" & Anyo
            Text1(2).Text = "31/01/" & Anyo
        Case 2
            If EsBiciesto(Anyo) Then
                Text1(0).Text = "01/02/" & Anyo
                Text1(2).Text = "29/02/" & Anyo
            Else
                Text1(0).Text = "01/02/" & Anyo
                Text1(2).Text = "28/02/" & Anyo
            End If
        Case 3
            Text1(0).Text = "01/03/" & Anyo
            Text1(2).Text = "31/03/" & Anyo
        Case 4
            Text1(0).Text = "01/04/" & Anyo
            Text1(2).Text = "30/04/" & Anyo
        Case 5
            Text1(0).Text = "01/05/" & Anyo
            Text1(2).Text = "31/05/" & Anyo
        Case 6
            Text1(0).Text = "01/06/" & Anyo
            Text1(2).Text = "30/06/" & Anyo
        Case 7
            Text1(0).Text = "01/07/" & Anyo
            Text1(2).Text = "31/07/" & Anyo
        Case 8
            Text1(0).Text = "01/08/" & Anyo
            Text1(2).Text = "31/08/" & Anyo
        Case 9
            Text1(0).Text = "01/09/" & Anyo
            Text1(2).Text = "30/09/" & Anyo
        Case 10
            Text1(0).Text = "01/10/" & Anyo
            Text1(2).Text = "31/10/" & Anyo
        Case 11
            Text1(0).Text = "01/11/" & Anyo
            Text1(2).Text = "30/11/" & Anyo
    End Select
End Sub

Private Function EsBiciesto(Anyo As Integer) As Boolean

EsBiciesto = False
If ((Anyo Mod 4 = 0) And (Anyo Mod 100 <> 0) Or (Anyo Mod 400 = 0)) Then
    EsBiciesto = True
End If

End Function

Private Sub Form_Activate()
    cadFormula = ""
    cadParam = ""
    numParam = 0
End Sub

Private Sub Form_Load()
Dim i As Integer

    'Icono del form
    Me.Icon = frmPpal.Icon

    'ocultamos los frames
    Frame1(0).visible = False
    Frame1(1).visible = False
    Text1(1).Text = Date
    'cargamos los tipos de cuotas
    CargarComboCuota
    Combo1.ListIndex = 1
    
'    Frame1(1).visible = True
'    Frame1(0).visible = False
    
    For i = 0 To Me.imgBuscar.Count - 1
        imgBuscar(i).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next
    For i = 0 To Me.imgFecha.Count - 1
        imgFecha(i).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next
    
    
    
    
    
    cadFormula = ""
    cadParam = ""
    numParam = 0
    Data1.ConnectionString = conn
End Sub

Private Sub CargarComboCuota()
    Combo1.Clear
'    Combo1.AddItem ""
'    Combo1.ItemData(Combo1.NewIndex) = 0
    
    Combo1.AddItem "Extraordinaria"
    Combo1.ItemData(Combo1.NewIndex) = 1
    Combo1.AddItem "Normal"
    Combo1.ItemData(Combo1.NewIndex) = 2
End Sub
Private Sub CargarComboMes()
    Combo2.Clear
    Combo2.AddItem "Enero"
    Combo2.ItemData(Combo2.NewIndex) = 0
    
    Combo2.AddItem "Febrero"
    Combo2.ItemData(Combo2.NewIndex) = 1
    
    Combo2.AddItem "Marzo"
    Combo2.ItemData(Combo2.NewIndex) = 2

    Combo2.AddItem "Abril"
    Combo2.ItemData(Combo2.NewIndex) = 3

    Combo2.AddItem "Mayo"
    Combo2.ItemData(Combo2.NewIndex) = 4

    Combo2.AddItem "Junio"
    Combo2.ItemData(Combo2.NewIndex) = 5

    Combo2.AddItem "Julio"
    Combo2.ItemData(Combo2.NewIndex) = 6

    Combo2.AddItem "Agosto"
    Combo2.ItemData(Combo2.NewIndex) = 7

    Combo2.AddItem "Septiembre"
    Combo2.ItemData(Combo2.NewIndex) = 8

    Combo2.AddItem "Octubre"
    Combo2.ItemData(Combo2.NewIndex) = 9

    Combo2.AddItem "Noviembre"
    Combo2.ItemData(Combo2.NewIndex) = 10
    
    Combo2.AddItem "Diciembre"
    Combo2.ItemData(Combo2.NewIndex) = 11

End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Cad = CadenaDevuelta
End Sub

Private Sub frmCal_Selec(vFecha As Date)
    Fecha = vFecha
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub frmFPag_DatoSeleccionado(CadenaSeleccion As String)
    Text1(8).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(8).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    Select Case Index
        Case 4, 3
            If Index = 4 Then
                indCodigo = 3
            Else
                indCodigo = 5
            End If
            Set frmSoc = New frmGesSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            If CadenaDesdeOtroForm <> "" Then
                Text1(indice).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                Text2(indice).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            End If
        Case 1
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing

        Case 0
            CadenaDesdeOtroForm = Text1(6).Text
            frmFacClienteObser.Modificar = True
            frmFacClienteObser.Text1 = CadenaDesdeOtroForm
            frmFacClienteObser.Show vbModal
            If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(6).Text = Mid(CadenaDesdeOtroForm, 3)
            CadenaDesdeOtroForm = ""
            PonerFoco Text1(6)
            
        Case 2 ' forma de pago
            Set frmFPag = New frmFacFormasPago
            frmFPag.DatosADevolverBusqueda = "0|1|"
            frmFPag.Show vbModal
            Set frmFPag = Nothing
        
    End Select
End Sub
Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    Text1(7).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub
Private Sub imgFecha_Click(Index As Integer)
Dim indice As String

    Screen.MousePointer = vbHourglass
   
    Set frmCal = New frmCal
    frmCal.Fecha = Now
    Select Case Index
        Case 0
            indice = 0
        Case 1
            indice = 2
        Case 2
            indice = 1
    End Select
    PonerFormatoFecha Text1(indice)
    If Text1(indice).Text <> "" Then frmCal.Fecha = CDate(Text1(indice).Text)
    Screen.MousePointer = vbDefault
    frmCal.Show vbModal
    Set frmCal = Nothing
    If Fecha <> "" Then Text1(indice).Text = Fecha
    PonerFoco Text1(indice)
End Sub



Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim encontrado As String

    Select Case Index
        Case 0, 2, 1
            If Text1(Index).Text <> "" Then
                PonerFormatoFecha Text1(Index)
            End If
        Case 5, 3 'socio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(Index).Text, "T")
            End If
        
        Case 4 'importe
            If Text1(Index).Text <> "" Then
                If Not IsNumeric(Text1(Index).Text) Then
                    MsgBox "El importe a facturar debe ser numérico.", vbExclamation
                    PonerFoco Text1(Index)
                    Exit Sub
                End If
                PonerFormatoDecimal Text1(Index), 1
            End If
            
        Case 7 ' cta de banco
            If Text1(Index).Text <> "" Then
                encontrado = DevuelveDesdeBD(conAri, "nombanpr", "sbanpr", "codbanpr", Text1(Index).Text, "T")
                If encontrado = "" Then
                    MsgBox "El banco introducido no existe", vbExclamation
                    PonerFoco Text1(Index)
                Else
                    Text2(Index).Text = encontrado
                End If
            End If
        
            
        Case 8 ' forma de pago
            PonerFormatoEntero Text1(8)
            Text2(8).Text = PonerNombreDeCod(Text1(8), conAri, "sforpa", "nomforpa", "codforpa", Text1(8).Text, "N")
    End Select

End Sub


Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim miRsAux As ADODB.Recordset
Dim Sql As String

    b = True

    If Text1(1).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una fecha de factura.", vbExclamation
        b = False
    End If

    Select Case Me.Combo1.ListIndex
        Case 0 ' extraordinarias
            codtipom = "FCE"
            
            If b Then
                If vParamAplic.ArtCuotaExtraor = "" Then
                    MsgBox "No está configurado el artículo de cuotas extraordinarias en parámetros. Revise", vbExclamation
                    b = False
                Else
                    'busco el iva del articulo
                    iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArtCuotaExtraor, "T")
                    If iva = "" Then
                        MsgBox "El artículo de cuota extraordinaria no tiene asignado el iva. Revise.", vbExclamation
                        b = False
                    Else
                        NomArtic = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", vParamAplic.ArtCuotaExtraor, "T")
                    End If
                End If
            End If
        Case 1 ' cuota normal
            codtipom = "FCN"
            
            If b Then
                If vParamAplic.ArtCuotaSinChofer = "" Then
                    MsgBox "No está configurado el artículo de cuotas sin chofer en parámetros. Revise", vbExclamation
                    b = False
                Else
                    'busco el iva del articulo
                    iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArtCuotaSinChofer, "T")
                    If iva = "" Then
                        MsgBox "El artículo de cuota sin chofer no tiene asignado el iva. Revise.", vbExclamation
                        b = False
                    Else
                        'busco el nombre del articulo
                        NomArtic = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtCuotaSinChofer, "T")
                        NomArtic = Trim(NomArtic) & " " & UCase(Combo2.Text) & " de " & Year(CDate(Text1(1).Text))
                    End If
                End If
            End If
            If b Then
                If vParamAplic.ArtCuotaConChofer = "" Then
                    MsgBox "No está configurado el artículo de cuotas con chofer en parámetros. Revise", vbExclamation
                    b = False
                Else
                    'busco el nombre del articulo
                    NomArtic2 = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtCuotaConChofer, "T")
                End If
            End If
            If b Then
                If vParamAplic.ArtServCuotas = "" Then
                    MsgBox "No está configurado el artículo de servicios de cuotas en parámetros. Revise", vbExclamation
                    b = False
                Else
                    'busco el nombre del articulo
                    NomArtic3 = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtServCuotas, "T")
                End If
            End If
            If b Then
                If vParamAplic.ArtAlquiler = "" Then
                    MsgBox "No está configurado el artículo de alquiler en parámetros. Revise", vbExclamation
                    b = False
                Else
                    'busco el nombre del articulo
                    NomArtic4 = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtAlquiler, "T")
                End If
            End If
    End Select

    If b Then ' comprobamos que desde / hasta socio son correctos
        If Text1(3).Text <> "" And Text1(5).Text <> "" Then
            If CLng(Text1(3).Text) > CLng(Text1(5).Text) Then
                MsgBox "Desde no puede ser superior a hasta.", vbExclamation
                b = False
                PonerFoco Text1(3)
            End If
        End If
    End If
    
    If b Then
        'valores grales para todos los socios
        Sql = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T")
        If Sql = "" Then
            MsgBox "No existe el tipo de iva en contabilidad. Revise.", vbExclamation
            b = False
        End If
    End If
    If b Then
        LetraSer = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", codtipom, "T")
        If LetraSer = "" Then
            MsgBox "No existe la letra de serie. Revise.", vbExclamation
            b = False
        End If
    End If
    
    If b Then 'forma de pago
        If Text1(8).Text <> "" Then
            Sql = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", Text1(8).Text, "N")
            If Sql = "" Then
                MsgBox "No existe la forma de pago. Revise.", vbExclamation
                b = False
            End If
        Else
            MsgBox "Debe introducir una forma de pago.", vbExclamation
            b = False
            PonerFoco Text1(8)
        End If
    End If
    
    If b Then
        Set miRsAux = New ADODB.Recordset
    
        'busco el minimo almacen y el minimo proveedor
        Sql = "select min(codalmac) from salmpr"
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        If Not miRsAux.EOF Then
            Sql = miRsAux.Fields(0)
        End If
        If Sql = "" Then
            MsgBox "No existe el almacén. Revise"
            b = False
        End If
            
        miRsAux.Close
            
        If b Then
            Sql = "select min(codprove) from sprove"
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Sql = ""
            If Not miRsAux.EOF Then
                Sql = miRsAux.Fields(0)
            End If
            If Sql = "" Then
                MsgBox "No existe el proveedor. Revise"
                b = False
            End If
        End If
        
        Set miRsAux = Nothing
        
    End If

    If b Then
        If Text1(7).Text = "" Then
            MsgBox "Debe introducir un banco de cobro.", vbExclamation
            b = False
            PonerFoco Text1(7)
        Else
            Sql = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", Text1(7).Text, "N")
            If Sql = "" Then
                MsgBox "El banco no tiene cuenta asociada. Revise", vbExclamation
                b = False
            End If
        End If
    End If

    DatosOk = b
End Function


