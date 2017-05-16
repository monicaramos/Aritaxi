VERSION 5.00
Begin VB.Form frmCuotasReImp 
   Caption         =   "Reimpresión de facturas"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Ordenar por "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   390
      TabIndex        =   31
      Top             =   4920
      Width           =   2085
      Begin VB.OptionButton Option3 
         Caption         =   "Código Socio"
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
         Left            =   210
         TabIndex        =   33
         Top             =   240
         Value           =   -1  'True
         Width           =   1785
      End
      Begin VB.OptionButton Option4 
         Caption         =   "V Socio"
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
         Left            =   210
         TabIndex        =   32
         Top             =   540
         Width           =   1455
      End
   End
   Begin VB.TextBox txtNombre 
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
      Index           =   1
      Left            =   2790
      TabIndex        =   30
      Top             =   2910
      Width           =   3585
   End
   Begin VB.TextBox txtNombre 
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
      Index           =   0
      Left            =   2790
      TabIndex        =   29
      Top             =   2550
      Width           =   3585
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   1680
      TabIndex        =   6
      Tag             =   "Cod. Cliente|N|N|0|999999|scafac|codclien|0000|N|"
      Top             =   2910
      Width           =   1065
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   1680
      TabIndex        =   5
      Tag             =   "Cod. Cliente|N|N|0|999999|scafac|codclien|0000|N|"
      Top             =   2550
      Width           =   1065
   End
   Begin VB.TextBox txtNombre 
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
      Left            =   2790
      TabIndex        =   22
      Top             =   1920
      Width           =   3585
   End
   Begin VB.TextBox txtNombre 
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
      Left            =   2790
      TabIndex        =   21
      Top             =   1560
      Width           =   3585
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   1680
      TabIndex        =   4
      Tag             =   "Cod. Cliente|N|N|0|999999|scafac|codclien|000000|N|"
      Top             =   1920
      Width           =   1065
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   1680
      TabIndex        =   3
      Tag             =   "Cod. Cliente|N|N|0|999999|scafac|codclien|000000|N|"
      Top             =   1560
      Width           =   1065
   End
   Begin VB.CheckBox chk_duplicado 
      Caption         =   "Duplicado"
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
      Left            =   4080
      TabIndex        =   20
      Top             =   4950
      Width           =   1575
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   86
      Left            =   4185
      MaxLength       =   10
      TabIndex        =   10
      Top             =   4365
      Width           =   1245
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   85
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   9
      Top             =   4350
      Width           =   1245
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   5580
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   12
      Top             =   5580
      Width           =   1215
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   36
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "codtalle|N|N|||scaord|codtalle|00|S|"
      Top             =   3630
      Width           =   1245
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
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
      Index           =   37
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "codtalle|N|N|||scaord|codtalle|00|S|"
      Top             =   3630
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   660
      Width           =   6015
      Begin VB.OptionButton Option2 
         Caption         =   "Extraordinarias"
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
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   2085
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normales"
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
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Image imgAyuda 
      Height          =   240
      Index           =   0
      Left            =   2610
      MousePointer    =   4  'Icon
      Tag             =   "-1"
      ToolTipText     =   "Ayuda"
      Top             =   5070
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1410
      Tag             =   "-1"
      ToolTipText     =   "Buscar V Socio"
      Top             =   2910
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1410
      Tag             =   "-1"
      ToolTipText     =   "Buscar V Socio"
      Top             =   2550
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "V Socio"
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
      Left            =   390
      TabIndex        =   28
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label Label2 
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
      Left            =   690
      TabIndex        =   27
      Top             =   2550
      Width           =   615
   End
   Begin VB.Label Label1 
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
      Left            =   690
      TabIndex        =   26
      Top             =   2910
      Width           =   615
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   1410
      Tag             =   "-1"
      ToolTipText     =   "Buscar Socio"
      Top             =   1950
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   3
      Left            =   1410
      Tag             =   "-1"
      ToolTipText     =   "Buscar Socio"
      Top             =   1590
      Width           =   240
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
      Left            =   690
      TabIndex        =   25
      Top             =   1950
      Width           =   615
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
      Left            =   690
      TabIndex        =   24
      Top             =   1590
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Socio"
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
      Left            =   390
      TabIndex        =   23
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   24
      Left            =   3900
      Top             =   4380
      Width           =   240
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   8
      Left            =   3300
      TabIndex        =   19
      Top             =   4380
      Width           =   570
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   7
      Left            =   720
      TabIndex        =   18
      Top             =   4380
      Width           =   600
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Fact."
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
      Height          =   240
      Index           =   7
      Left            =   390
      TabIndex        =   17
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   23
      Left            =   1410
      Top             =   4350
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   690
      TabIndex        =   16
      Top             =   3660
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   3300
      TabIndex        =   15
      Top             =   3690
      Width           =   570
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Nº Factura"
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
      Height          =   240
      Index           =   4
      Left            =   390
      TabIndex        =   14
      Top             =   3390
      Width           =   1140
   End
   Begin VB.Label Label10 
      Caption         =   "Reimpresión Facturas Cuotas"
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
      Left            =   360
      TabIndex        =   0
      Top             =   210
      Width           =   5295
   End
End
Attribute VB_Name = "frmCuotasReImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tabla As String
Dim cadFormula As String
Dim cadSelect As String
Dim cadParam As String
Dim numParam As Integer
Dim codtipom As String
Dim indCodigo As Integer
Dim PrimeraVez As Boolean


Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmSoc As frmGesSocios  'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmMtoV As frmGesVSocio ' V socios
Attribute frmMtoV.VB_VarHelpID = -1

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Sub cmdAceptar_Click()
Dim Codigo As String

    Tabla = "scafac"
    If Option1.Value Then
        codtipom = "FCN"
    End If
    If Option2.Value Then
        codtipom = "FCE"
    End If
    
    InicializarVbles
    
    'Desde/Hasta codigo de socio
    '---------------------------------------------
    If txtcodigo(3).Text <> "" Or txtcodigo(5).Text <> "" Then
        Codigo = "{" & Tabla & ".codclien}"
        If Not PonerDesdeHasta(Codigo, "N", 3, 5, "") Then Exit Sub
    End If
    
    
    'Desde/Hasta V de socio
    '---------------------------------------------
    If txtcodigo(0).Text <> "" Or txtcodigo(1).Text <> "" Then
        Codigo = "{sclien.numeruve}"
        If Not PonerDesdeHasta(Codigo, "N", 0, 1, "") Then Exit Sub
    End If
    
    'Desde/Hasta numero de FACTURA
    '---------------------------------------------
    If txtcodigo(36).Text <> "" Or txtcodigo(37).Text <> "" Then
        Codigo = "{" & Tabla & ".numfactu}"
        If Not PonerDesdeHasta(Codigo, "N", 36, 37, "") Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtcodigo(85).Text <> "" Or txtcodigo(86).Text <> "" Then
        Codigo = "{" & Tabla & ".fecfactu}"
        If Not PonerDesdeHasta(Codigo, "F", 85, 86, "") Then Exit Sub
    End If
    
    If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".codtipom} = """ & codtipom & """") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".codtipom} = """ & codtipom & """") Then Exit Sub
    
'    If CBool(Me.chk_duplicado.Value) Then
'        cadParam = "pDuplicado=1|"
'    Else
'        cadParam = "pDuplicado=0|"
'    End If
'
    
    Tabla = Tabla & " INNER JOIN sclien ON scafac.codclien = sclien.codclien "
    
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 1
    
    If Option3.Value Then
        cadParam = cadParam & "|pGroup={scafac.codclien}|"
    Else
        cadParam = cadParam & "|pGroup={sclien.numeruve}|"
    End If
    numParam = numParam + 1
    
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    LlamarImprimir True
End Sub

Private Sub LlamarImprimir(duplicado As Boolean)
        cadParam = cadParam & "pDuplicado= " & Abs(duplicado) & " |"
        numParam = numParam + 1
        
        With frmImprimir
        .Titulo = "Impresión de Facturas de Cuotas"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        'El nombre es el del documento
        .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "49", "N")
        '------ > Listado 49 = rFactuCuotas.rpt
        .Opcion = 101
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
    cadFormula = ""
    numParam = 0
    cadParam = ""
    
    If PrimeraVez Then
        PonerFoco txtcodigo(3)
        PrimeraVez = False
    End If
        
End Sub

Private Sub Form_Load()
Dim i As Integer

    'Icono del form
    Me.Icon = frmPpal.Icon
    
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i
    
    For i = 0 To 1
        imgBuscar(i).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next i
    For i = 3 To 4
        imgBuscar(i).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next i
    
    For i = 23 To 24
        imgFecha(i).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next i
    
    
    PrimeraVez = True
End Sub

Private Function DatosOk() As Boolean
Dim encontrado As String
Dim Codigo As String

    DatosOk = True
End Function


Private Function ejecutaselect(CADENA As String) As String
Dim RS As Recordset
Dim C As String

ejecutaselect = ""
Set RS = New ADODB.Recordset
RS.Open CADENA, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RS.EOF Then
    If Not IsNull((RS.Fields(0))) Then
        C = RS.Fields(0)
    Else
        C = 0
    End If
End If
RS.Close
Set RS = Nothing
ejecutaselect = C


End Function

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub frmMtoV_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub


Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "El valor de la V socio es que tiene en el momento actual, puede que" & vbCrLf & _
                      "no coincida con la V que tenia cuando se hizo la factura." & vbCrLf & vbCrLf

                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    Select Case Index
        Case 4, 3
            If Index = 3 Then
                indCodigo = 3
            Else
                indCodigo = 5
            End If
            Set frmSoc = New frmGesSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            If CadenaDesdeOtroForm <> "" Then
                txtcodigo(indice).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                txtnombre(indice).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            End If
            
        Case 0, 1 ' Vsocio
            indCodigo = Index
            
            Set frmMtoV = New frmGesVSocio
            frmMtoV.DeConsulta = True
            frmMtoV.DatosADevolverBusqueda = "0|1|2|"
            frmMtoV.Show vbModal
            Set frmMtoV = Nothing
        
        
    End Select
    
    PonerFoco txtcodigo(indCodigo)
    
End Sub

Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 23, 24 'fechas de factura
            indCodigo = Index + 62
   End Select
   
   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtcodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Tabla As String
Dim codCampo As String, NomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean


    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    EsNomCod = False
    TipCampo = "N" 'Casi todos son numericos
    
    Select Case Index
        Case 5, 3 'socio
            txtnombre(Index).Text = ""
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtnombre(Index).Text = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtcodigo(Index).Text, "N")
            End If
            
        Case 0, 1 'v socio
            txtnombre(Index).Text = ""
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtnombre(Index).Text = DevuelveDesdeBD(conAri, "nomclien", "sclien", "numeruve", txtcodigo(Index).Text, "N")
            End If
            
    
        Case 85, 86  'FECHA Desde Hasta
            If txtcodigo(Index).Text = "" Then Exit Sub
            PonerFormatoFecha txtcodigo(Index)
            
        Case 36, 37 'Nº de FACTURA
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000000")
            End If
    End Select
    
End Sub

Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
'            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
'            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtcodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


