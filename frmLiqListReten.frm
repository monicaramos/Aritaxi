VERSION 5.00
Begin VB.Form frmLiqListReten 
   Caption         =   "Informes"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRecibosReten 
      Height          =   5835
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   6795
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   6
         Left            =   2010
         TabIndex        =   11
         Top             =   4290
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2790
         TabIndex        =   45
         Top             =   4290
         Width           =   3675
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   96
         Left            =   2010
         MaxLength       =   10
         TabIndex        =   8
         Top             =   2760
         Width           =   1185
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   2010
         MaxLength       =   10
         TabIndex        =   9
         Top             =   3270
         Width           =   1185
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   82
         Left            =   2010
         TabIndex        =   4
         Tag             =   "Num vehiculo|N|N|||shilla|codclien|000000|S|"
         Top             =   1230
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   82
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1230
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   14
         Left            =   5460
         TabIndex        =   14
         Top             =   4980
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepRecibos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4410
         TabIndex        =   13
         Top             =   4980
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   83
         Left            =   2010
         TabIndex        =   5
         Tag             =   "Num vehiculo|N|N|||shilla|codclien|000000|S|"
         Top             =   1590
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   83
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1590
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   102
         Left            =   2010
         TabIndex        =   6
         Top             =   2190
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   103
         Left            =   4140
         TabIndex        =   7
         Top             =   2190
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   2010
         TabIndex        =   10
         Top             =   3780
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   61
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   3780
         Width           =   3675
      End
      Begin VB.Label Label3 
         Caption         =   "Cargando tabla temporal..."
         Height          =   195
         Index           =   1
         Left            =   570
         TabIndex        =   47
         Top             =   4920
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.Label Label9 
         Caption         =   "Cuenta Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   570
         TabIndex        =   46
         Top             =   4290
         Width           =   1125
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   5
         Left            =   1710
         Picture         =   "frmLiqListReten.frx":0000
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta"
         Top             =   4290
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   1710
         Picture         =   "frmLiqListReten.frx":0102
         Top             =   3270
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recibo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   48
         Left            =   570
         TabIndex        =   23
         Top             =   3300
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   47
         Left            =   570
         TabIndex        =   22
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   1140
         TabIndex        =   21
         Top             =   1590
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   17
         Left            =   570
         TabIndex        =   20
         Top             =   990
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   38
         Left            =   1725
         Picture         =   "frmLiqListReten.frx":018D
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Recibos de Retenciones"
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
         Height          =   345
         Left            =   570
         TabIndex        =   19
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   88
         Left            =   1140
         TabIndex        =   18
         Top             =   1230
         Width           =   465
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   39
         Left            =   1725
         Picture         =   "frmLiqListReten.frx":028F
         ToolTipText     =   "Buscar socio"
         Top             =   1590
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   1710
         Picture         =   "frmLiqListReten.frx":0391
         Top             =   2190
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   3840
         Picture         =   "frmLiqListReten.frx":041C
         Top             =   2190
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   87
         Left            =   570
         TabIndex        =   17
         Top             =   1950
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   86
         Left            =   3270
         TabIndex        =   16
         Top             =   2190
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   85
         Left            =   1140
         TabIndex        =   15
         Top             =   2190
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   50
         Left            =   570
         TabIndex        =   12
         Top             =   3780
         Width           =   1005
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   8
         Left            =   1695
         Picture         =   "frmLiqListReten.frx":04A7
         ToolTipText     =   "Buscar f.pago"
         Top             =   3780
         Width           =   240
      End
   End
   Begin VB.Frame FrameListado 
      Height          =   5865
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   6825
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   5670
         TabIndex        =   42
         Top             =   4110
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   5670
         TabIndex        =   41
         Top             =   3660
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5400
         TabIndex        =   33
         Top             =   4860
         Width           =   1035
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   32
         Top             =   4860
         Width           =   1035
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   86
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   30
         Top             =   2895
         Width           =   1035
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   85
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   29
         Top             =   2520
         Width           =   1035
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   28
         Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
         Top             =   1710
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1710
         Width           =   3765
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   27
         Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
         Top             =   1365
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1365
         Width           =   3765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Resumen"
         Height          =   225
         Left            =   870
         TabIndex        =   25
         Top             =   3750
         Width           =   2265
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   5370
         Picture         =   "frmLiqListReten.frx":05A9
         ToolTipText     =   "Buscar fecha"
         Top             =   3690
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Importe Base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   4500
         TabIndex        =   44
         Top             =   4110
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Informe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   4500
         TabIndex        =   43
         Top             =   3360
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label10 
         Caption         =   "Retenciones Servicios a Crédito"
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
         Height          =   345
         Index           =   0
         Left            =   510
         TabIndex        =   40
         Top             =   390
         Width           =   5655
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   1410
         Picture         =   "frmLiqListReten.frx":0634
         Top             =   2910
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   900
         TabIndex        =   39
         Top             =   2940
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   900
         TabIndex        =   38
         Top             =   2550
         Width           =   450
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   7
         Left            =   510
         TabIndex        =   37
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   1410
         Picture         =   "frmLiqListReten.frx":06BF
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   900
         TabIndex        =   36
         Top             =   1710
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   1
         Left            =   1410
         Picture         =   "frmLiqListReten.frx":074A
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   900
         TabIndex        =   35
         Top             =   1365
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   510
         TabIndex        =   34
         Top             =   1020
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   0
         Left            =   1410
         Picture         =   "frmLiqListReten.frx":084C
         Top             =   1365
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmLiqListReten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OpcionListado As Integer


Dim Tabla As String
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim codtipom As String
Dim cadSelect As String
Dim indCodigo As Long
Dim cadNombreRPT As String
Dim cadTitulo As String
Dim ConSubInforme As Boolean
Dim conSubRPT As Boolean

Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmMtoV As frmGesSocios ' socios
Attribute frmMtoV.VB_VarHelpID = -1

Public WithEvents frmFP As frmFacFormasPago ' formas de pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios ' bancos propios
Attribute frmMtoBancosPro.VB_VarHelpID = -1

' Importes para la Grabacion de Cabecera de Facturas de Socio
Dim TotalFac As Currency
Dim TotalLiq As Currency
Dim BaseImpo As Currency
Dim BaseReten As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim vPorcIva As String
Dim PorceIVA As Currency

Dim tipoMov As String
Dim CodSocio As String

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Sub cmdAceptar_Click()
Dim Codigo As String
Dim FecFac As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


    If Not DatosOK Then Exit Sub
    
    InicializarVbles
    
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 1
   
    'Desde/Hasta numero de V
    '---------------------------------------------
    If txtCodigo(0).Text <> "" Or txtCodigo(1).Text <> "" Then
        Codigo = "{" & Tabla & ".codsocio}"
        If Not PonerDesdeHasta(Codigo, "N", 0, 1, "pDHSocio=""") Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(85).Text <> "" Or txtCodigo(86).Text <> "" Then
        Codigo = "{" & Tabla & ".fecfactu}"
        If Not PonerDesdeHasta(Codigo, "F", 85, 86, "pDHFecha=""") Then Exit Sub
    End If
    
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub

    If Check1.Value Then
        cadNombreRPT = "rListRetencionesRes.rpt"
    Else
        cadNombreRPT = "rListRetenciones.rpt"
    End If
    
    cadTitulo = "Retenciones Servicios de Crédito"

    cadParam = cadParam & "pFecFac= """ & txtCodigo(3).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pTitulo= ""Retenciones Servicios de Crédito""|"
    numParam = numParam + 1
    cadParam = cadParam & "pBase=" & TransformaComasPuntos(ImporteSinFormato(txtCodigo(2).Text)) & "|"
    numParam = numParam + 1
    
    ConSubInforme = False
    
    LlamarImprimir False
        
    cmdCancelar_Click

End Sub

Private Sub LlamarImprimir(duplicado As Boolean)
        
    With frmImprimir
        .Titulo = cadTitulo
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        'El nombre es el del documento
        .NombreRPT = cadNombreRPT
        .Opcion = 101
        .ConSubInforme = ConSubInforme
        .Show vbModal
    End With
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdAcepRecibos_Click()
Dim Codigo As String
Dim FecFac As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


    If Not DatosOK Then Exit Sub
    
    InicializarVbles
    
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 1
   
    'Desde/Hasta numero de V
    '---------------------------------------------
    If txtCodigo(82).Text <> "" Or txtCodigo(83).Text <> "" Then
        Codigo = "{" & Tabla & ".codsocio}"
        If Not PonerDesdeHasta(Codigo, "N", 82, 83, "pDHSocio=""") Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(102).Text <> "" Or txtCodigo(103).Text <> "" Then
        Codigo = "{" & Tabla & ".fecfactu}"
        If Not PonerDesdeHasta(Codigo, "F", 102, 103, "pDHFecha=""") Then Exit Sub
    End If
    
'    If Not AnyadirAFormula(cadFormula, "{sreten.tiporeten} = 0") Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, "{sreten.tiporeten} = 0") Then Exit Sub
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub

    If CargarTablaTemporal(Tabla, cadSelect) Then
        If Not HayRegParaInforme("tmpinformes", "codusu= " & vUsu.Codigo) Then Exit Sub

        cadNombreRPT = "rRecRetenciones.rpt"
        cadTitulo = "Recibos Retenciones Servicios de Crédito"
    
        cadParam = cadParam & "pFecFac= """ & txtCodigo(3).Text & """|"
        numParam = numParam + 1
        cadParam = cadParam & "pTitulo= ""Retenciones Servicios de Crédito""|"
        numParam = numParam + 1
        cadParam = cadParam & "pBase=" & TransformaComasPuntos(ImporteSinFormato(txtCodigo(2).Text)) & "|"
        numParam = numParam + 1
        
        ConSubInforme = False
        
        ' llamamos a la impresion de recibo
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        LlamarImprimir False
        
        If MsgBox("¿Impresion correcta para actualizar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            If ActualizarRegistros Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancelar_Click
            End If
        End If
    End If

End Sub

Private Function ActualizarRegistros() As Boolean
Dim Sql As String
Dim SQL2 As String
Dim RS As ADODB.Recordset
Dim Sql2Values As String
Dim fac As CFacturaCom
Dim b As Boolean
Dim Socio As Long
Dim FormatSocio As String
Dim cuenta As String
Dim vSocio As CSocio
Dim MenError As String
Dim Mens As String


    On Error GoTo eActualizarRegistros
    
    ActualizarRegistros = False
        
    Label3(1).visible = True
    Label3(1).Caption = "Insertando registros..."
    DoEvents
        
    Screen.MousePointer = vbHourglass

    conn.BeginTrans
    ConnConta.BeginTrans
    
    Sql = "select * from tmpinformes where codusu = " & vUsu.Codigo
    Sql = Sql & " order by codigo1, importe1"
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    SQL2 = "insert into sreten (codsocio, numeruve, fecfactu, numfactu, impreten, tiporeten, desdefec, hastafec) values "
    b = True
    While Not RS.EOF And b
        Sql2Values = "(" & DBSet(RS!Codigo1, "N") & "," & DBSet(RS!importe1, "N") & "," & DBSet(txtCodigo(4).Text, "F") & ","
        Sql2Values = Sql2Values & "0," & DBSet(RS!importe2 * (-1), "N") & ",1," & DBSet(txtCodigo(102).Text, "F") & ","
        Sql2Values = Sql2Values & DBSet(txtCodigo(103).Text, "F") & ")"
        
        conn.Execute SQL2 & Sql2Values

'desde aqui
        Set fac = New CFacturaCom
    
        fac.TotalFac = DBLet(RS!importe2, "N")
        fac.FecFactu = txtCodigo(4).Text
        fac.NumFactu = "R-" & Format(RS!Codigo1, "0000") & Format(RS!importe1, "0000")
        
        fac.Proveedor = DBLet(RS!Codigo1, "N")
        fac.NombreProv = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", RS!Codigo1, "T")
        fac.DomicilioProv = DevuelveDesdeBD(conAri, "domclien", "sclien", "codclien", RS!Codigo1, "T")
        fac.CPostalProv = DevuelveDesdeBD(conAri, "codpobla", "sclien", "codclien", RS!Codigo1, "T")
        fac.PoblacionProv = DevuelveDesdeBD(conAri, "pobclien", "sclien", "codclien", RS!Codigo1, "T")
        fac.ProvinciaProv = DevuelveDesdeBD(conAri, "proclien", "sclien", "codclien", RS!Codigo1, "T")
        fac.NIFProv = DevuelveDesdeBD(conAri, "nifclien", "sclien", "codclien", RS!Codigo1, "T")
        fac.ForPago = txtCodigo(61).Text
        
        'Cuenta Prevista de Cobro de las Facturas
        fac.BancoPr = txtCodigo(6).Text
        fac.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", fac.BancoPr, "N")
        'cuenta contable de proveedor
        'comprobamos q la cuenta contable exista en contabilidad
        Socio = DBLet(RS!Codigo1, "N")
        FormatSocio = String((vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior), "0")
        cuenta = Trim(vParamAplic.Raiz_Cta_Reten_Soc & Format(Socio, FormatSocio))
        Sql = ""
        Sql = DevuelveDesdeBD(conConta, "nommacta", "cuentas", "codmacta", cuenta, "T")
        If Sql = "" Then
            MsgBox "La cuenta contable del socio: " & Format(Socio, "000000") & " no existe.", vbExclamation
            conn.RollbackTrans
            ConnConta.RollbackTrans
            Screen.MousePointer = vbDefault
            Me.Label3(1).visible = False
            DoEvents
            Exit Function
        End If
        fac.CtaProve = cuenta
        
        '[Monica]añadido no se cargaba la ccc del socio en tesoreria
        Set vSocio = New CSocio
        If vSocio.LeerDatos(CStr(Socio)) Then
            '[Monica]22/11/2013: iban
            fac.CCC_Iban = vSocio.Iban
            fac.CCC_Entidad = vSocio.Banco
            fac.CCC_Oficina = vSocio.Sucursal
            fac.CCC_CC = vSocio.DigControl
            fac.CCC_CTa = vSocio.CuentaBan
        End If
        Set vSocio = Nothing
      
        MenError = "Error al pasar a tesoreria"
        '[Monica]26/01/2012: cambiamos el parametro opcional para que imprima en texto de csb otra cosa
        fac.Proveedor = Year(CDate(txtCodigo(103).Text))
        b = fac.InsertarEnTesoreria(MenError, True) ' true = indicamos que venimos de pago de retenciones
        
        Set fac = Nothing
'hasta aqui
    
        RS.MoveNext
    Wend
    Set RS = Nothing

    conn.CommitTrans
    ConnConta.CommitTrans
    
    ActualizarRegistros = True
    
    Screen.MousePointer = vbDefault
    Me.Label3(1).visible = False
    DoEvents
    
    Exit Function

eActualizarRegistros:
    If Err.Number <> 0 Or Not b Then
        Mens = ""
        If Not b Then Mens = Mens & MenError
        MuestraError Err.Number, "Actualizar registros", Mens & vbCrLf & Err.Description
        conn.RollbackTrans
        ConnConta.RollbackTrans
        Me.Label3(1).visible = False
        DoEvents
        Screen.MousePointer = vbDefault
    End If
End Function


Private Function CargarTablaTemporal(Tabla As String, cadSelect As String) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim SqlValues As String
Dim Importe As Currency
    
    On Error GoTo eCargarTablaTemporal

    CargarTablaTemporal = False
    
    Me.Label3(1).visible = True
    Label3(1).Caption = "Cargando tabla temporal..."
    DoEvents
    
    Screen.MousePointer = vbHourglass
    
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "select codsocio, numeruve, sum(if(impreten is null,0,impreten)) as Importe from sreten "
    If cadSelect <> "" Then Sql = Sql & " where " & cadSelect
    Sql = Sql & " group by 1 "
    Sql = Sql & " order by 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SqlValues = ""
    While Not RS.EOF
        Importe = DBLet(RS!Importe, "N") - ComprobarCero(txtCodigo(96).Text)
        
        If Importe > 0 Then
            SqlValues = SqlValues & "(" & vUsu.Codigo & "," & DBLet(RS!CodSocio, "N") & "," & DBLet(RS!NumerUve, "N") & "," & DBSet(Importe, "N") & "," & DBSet(txtCodigo(103).Text, "F") & "),"
        End If
    
        RS.MoveNext
    Wend
    
    If SqlValues <> "" Then
        Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, fecha1) values "
        Sql = Sql & Mid(SqlValues, 1, Len(SqlValues) - 1)
        
        conn.Execute Sql
    End If
    
    Set RS = Nothing

    CargarTablaTemporal = True
    Screen.MousePointer = vbDefault
    Me.Label3(1).visible = False
    DoEvents
    Exit Function
    
eCargarTablaTemporal:
    Screen.MousePointer = vbDefault
    Me.Label3(1).visible = False
    DoEvents
    MuestraError Err.Number, "Cargando Tabla Temporal", Err.Description
End Function


Private Sub Form_Activate()
    cadFormula = ""
    numParam = 0
    cadParam = ""


    PonerFoco txtCodigo(0)

End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer

    'Icono del form
    Me.Icon = frmPpal.Icon
    
    Me.FrameListado.visible = False
    Me.FrameRecibosReten.visible = False
    
    Select Case OpcionListado
        Case 0
            PonerFrameListadoVisible True, H, W

            Tabla = "sreten"
        
            txtCodigo(3).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(2).Text = "0,00"
            
        Case 1
            PonerFrameRecibosRetenVisible True, H, W

            Tabla = "sreten"
        
    End Select

End Sub

Private Sub PonerFrameListadoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de los mantenimientos de tabla: smarca, stipar,...

    H = 6405
    W = 7095
    PonerFrameVisible Me.FrameListado, visible, H, W

End Sub

Private Sub PonerFrameRecibosRetenVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de los mantenimientos de tabla: smarca, stipar,...

    H = 6405
    W = 7095
    PonerFrameVisible Me.FrameRecibosReten, visible, H, W

End Sub



Private Function DatosOK() As Boolean
Dim encontrado As String
Dim Codigo As String

    DatosOK = True
    
    Select Case OpcionListado
        Case 0 ' listado de retenciones
            If txtCodigo(2).Text = "" Then
                MsgBox "Debe introducir obligatoriamente un Importe Base.", vbExclamation
                DatosOK = False
                Exit Function
            End If
            If txtCodigo(3).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la fecha de listado.", vbExclamation
                DatosOK = False
                Exit Function
            End If
        Case 1 ' Impresion de recibos de retenciones
            'fecha de recibo
            If txtCodigo(4).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la Fecha de Recibo.", vbExclamation
                PonerFoco txtCodigo(4)
                DatosOK = False
                Exit Function
            End If
            'forma de pago
            If txtCodigo(61).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la Forma de Pago.", vbExclamation
                PonerFoco txtCodigo(61)
                DatosOK = False
                Exit Function
            End If
            'cuenta de pago
            If txtCodigo(6).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la Cuenta de Pago.", vbExclamation
                PonerFoco txtCodigo(6)
                DatosOK = False
                Exit Function
            End If
            
            '[Monica]01/10/2012: obligamos a meter el desde/hasta fecha
            If txtCodigo(102).Text = "" Or txtCodigo(103).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la Fecha Desde / Hasta.", vbExclamation
                PonerFoco txtCodigo(102)
                DatosOK = False
                Exit Function
            End If
            
    End Select
        
End Function



Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoV_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 0, 1 ' Vsocio
            indCodigo = Index
            
            Set frmMtoV = New frmGesSocios
            frmMtoV.DatosADevolverBusqueda = "0|1|"
            frmMtoV.Show vbModal
            Set frmMtoV = Nothing
        
        Case 38, 39 ' socio
            indCodigo = Index + 44
            
            Set frmMtoV = New frmGesSocios
            frmMtoV.DatosADevolverBusqueda = "0|1|"
            frmMtoV.Show vbModal
            Set frmMtoV = Nothing
        
        Case 8 ' forma de pago
            indCodigo = Index + 53
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0|1|"
            frmFP.Show vbModal
            Set frmFP = Nothing
        
        Case 5 ' cuenta de pago
            indCodigo = Index + 1
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
        
        
        
    End Select
    PonerFoco txtCodigo(indCodigo)

End Sub

Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 23, 24 'fechas de factura
            indCodigo = Index + 62
        Case 0
            indCodigo = 3
   End Select
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
End Sub

Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    If txtCodigo(indD).Text <> "" Then
        cad = cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        cad = cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then cad = cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
Private Sub KEYpress(KeyAscii As Integer)
'Dim cerrar As Boolean
'
'    KEYpressGnral KeyAscii, 2, cerrar
'    If cerrar Then Unload Me
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If

End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Tabla As String
Dim codCampo As String, NomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean
Dim encontrado As String


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    Select Case Index
        Case 85, 86, 102, 103 'FECHA Desde Hasta
            PonerFormatoFecha txtCodigo(Index)
            
        Case 0, 1 'V Socio
            PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sclien", "nomclien", "codclien", "N")
            
        Case 82, 83 'Socio
            PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sclien", "nomclien", "codclien", "N")
            
            
        Case 2 ' importe base
            PonerFormatoDecimal txtCodigo(Index), 3
            
        Case 3, 4 ' fecha de listado
            PonerFormatoFecha txtCodigo(Index)
            
        Case 96 ' importe
            PonerFormatoDecimal txtCodigo(Index), 3
            
        Case 6 ' cta de banco
            If txtCodigo(Index).Text <> "" Then
                encontrado = DevuelveDesdeBD(conAri, "nombanpr", "sbanpr", "codbanpr", txtCodigo(Index).Text, "T")
                If encontrado = "" Then
                    MsgBox "El banco introducido no existe", vbExclamation
                    PonerFoco txtCodigo(Index)
                Else
                    txtNombre(Index).Text = encontrado
                End If
            End If
        
        Case 61 ' forma de pago
            If txtCodigo(Index).Text <> "" Then
                If Not IsNumeric(txtCodigo(Index).Text) Then
                    MsgBox "La forma de pago debe ser numérica.", vbExclamation
                    PonerFoco txtCodigo(Index)
                    Exit Sub
                End If
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
                encontrado = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", txtCodigo(Index).Text, "T")
                If encontrado = "" Then
                    MsgBox "La forma de pago introducida no existe.", vbExclamation
                    PonerFoco txtCodigo(Index)
                Else
                    txtNombre(Index).Text = encontrado
                End If
            End If
        
        
        
    End Select
End Sub



