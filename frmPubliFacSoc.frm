VERSION 5.00
Begin VB.Form frmPubliFacSoc 
   Caption         =   "Informes"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5190
      TabIndex        =   11
      Top             =   5250
      Width           =   1135
   End
   Begin VB.TextBox Text1 
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
      Index           =   8
      Left            =   4110
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2580
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Index           =   7
      Left            =   1530
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2580
      Width           =   1215
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
      Index           =   6
      Left            =   2490
      TabIndex        =   25
      Top             =   4770
      Width           =   3825
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
      Index           =   6
      Left            =   1530
      TabIndex        =   9
      Top             =   4770
      Width           =   915
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
      Left            =   2490
      TabIndex        =   22
      Top             =   2010
      Width           =   3825
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
      Left            =   2490
      TabIndex        =   21
      Top             =   1650
      Width           =   3825
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
      Left            =   1530
      TabIndex        =   3
      Top             =   2010
      Width           =   915
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
      Left            =   1530
      TabIndex        =   2
      Top             =   1650
      Width           =   915
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
      Index           =   2
      Left            =   2490
      TabIndex        =   17
      Top             =   3480
      Width           =   3825
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   5190
      TabIndex        =   23
      Top             =   5250
      Visible         =   0   'False
      Width           =   1135
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
      Left            =   3900
      TabIndex        =   10
      Top             =   5250
      Width           =   1135
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
      Left            =   1530
      TabIndex        =   8
      Top             =   3990
      Width           =   4785
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
      Index           =   2
      Left            =   1530
      TabIndex        =   7
      Top             =   3480
      Width           =   915
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
      Left            =   1530
      TabIndex        =   6
      Text            =   "99/99/9999"
      Top             =   3000
      Width           =   1305
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
      Index           =   0
      Left            =   2490
      TabIndex        =   16
      Top             =   960
      Width           =   3855
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
      Left            =   1500
      TabIndex        =   1
      Top             =   960
      Width           =   945
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   13
      Left            =   3810
      Top             =   2595
      Width           =   240
   End
   Begin VB.Label Label4 
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
      Index           =   46
      Left            =   570
      TabIndex        =   28
      Top             =   2640
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
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
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Top             =   2340
      Width           =   630
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   12
      Left            =   1260
      Top             =   2595
      Width           =   240
   End
   Begin VB.Label Label4 
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
      Index           =   37
      Left            =   3150
      TabIndex        =   26
      Top             =   2640
      Width           =   570
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   5
      Left            =   1260
      Tag             =   "-1"
      ToolTipText     =   "Buscar cuenta"
      Top             =   4800
      Width           =   240
   End
   Begin VB.Label Label9 
      Caption         =   "Cuenta Pago"
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
      TabIndex        =   24
      Top             =   4470
      Width           =   1815
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   1260
      Tag             =   "-1"
      ToolTipText     =   "Buscar socio"
      Top             =   1680
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   3
      Left            =   1260
      Tag             =   "-1"
      ToolTipText     =   "Buscar socio"
      Top             =   2040
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
      Left            =   570
      TabIndex        =   20
      Top             =   2040
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
      Left            =   570
      TabIndex        =   19
      Top             =   1680
      Width           =   615
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
      TabIndex        =   18
      Top             =   1380
      Width           =   825
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1260
      Tag             =   "-1"
      ToolTipText     =   "Buscar forma de pago"
      Top             =   3510
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1260
      Tag             =   "-1"
      ToolTipText     =   "Ver observaciones"
      Top             =   3990
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   1260
      Tag             =   "-1"
      ToolTipText     =   "Buscar cliente"
      Top             =   960
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   2
      Left            =   1260
      ToolTipText     =   "Buscar fecha"
      Top             =   3030
      Width           =   240
   End
   Begin VB.Label Label6 
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
      TabIndex        =   15
      Top             =   3990
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "F.Pago"
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
      TabIndex        =   14
      Top             =   3510
      Width           =   885
   End
   Begin VB.Label Label3 
      Caption         =   "F.Factura"
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
      TabIndex        =   13
      Top             =   3030
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
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
      TabIndex        =   12
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Facturaci�n Publicidad a Socios"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5355
   End
End
Attribute VB_Name = "frmPubliFacSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 604


Public WithEvents frmCal As frmCal
Attribute frmCal.VB_VarHelpID = -1
Public WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Public WithEvents frmFP As frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Public WithEvents frmSoc As frmGesSocios
Attribute frmSoc.VB_VarHelpID = -1
Public WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios
Attribute frmMtoBancosPro.VB_VarHelpID = -1


Dim PrimeraVez As Boolean
Dim cadFormula As String
Dim cadParam As String
Dim NumFactu As Long
Dim FecFactu As Date
Dim numParam As Integer
Dim Modo As Byte
Dim Cad As String
Dim indCodigo As Integer

Dim kCampo As Integer


Private Sub cmdAceptar_Click()
Dim Sql As String
Dim SQL2 As String
    
    Set miRsAux = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    
    If DatosOk Then
        If HaySocios Then
            Sql = "select * from sclien_publicidad where codclien=" & Text1(0).Text & " and situacio=0 and codsocio >=" & Text1(3).Text & " and codsocio <=" & Text1(5).Text & " and situacio=0 and sclien_publicidad.desdefec >= " & DBSet(Text1(7).Text, "F") & " and sclien_publicidad.hastafec <=" & DBSet(Text1(8).Text, "F")
            SQL2 = "select nomclien,sclien_publicidad.importes,sclien_publicidad.desdefec,sclien_publicidad.hastafec from sclien_publicidad inner join sclien on "
            SQL2 = SQL2 & "sclien_publicidad.codClien=" & Text1(0).Text & " and sclien_publicidad.situacio=0 and sclien_publicidad.codsocio >=" & Text1(3).Text & " and sclien_publicidad.codsocio <=" & Text1(5).Text & " and sclien_publicidad.codsocio=sclien.codclien"
            SQL2 = SQL2 & " and sclien_publicidad.desdefec >= " & DBSet(Text1(7).Text, "F") & " and sclien_publicidad.hastafec <=" & DBSet(Text1(8).Text, "F")
        Else
            Sql = "select * from sclien_publicidad where codclien=" & Text1(0).Text & " and situacio=0 and sclien_publicidad.desdefec >= " & DBSet(Text1(7).Text, "F") & " and sclien_publicidad.hastafec <=" & DBSet(Text1(8).Text, "F")
            SQL2 = "select nomclien,sclien_publicidad.importes,sclien_publicidad.desdefec,sclien_publicidad.hastafec from sclien_publicidad inner join sclien on "
            SQL2 = SQL2 & "sclien_publicidad.codClien=" & Text1(0).Text & " and sclien_publicidad.situacio=0 and sclien_publicidad.codsocio=sclien.codclien"
            SQL2 = SQL2 & " and sclien_publicidad.desdefec >= " & DBSet(Text1(7).Text, "F") & " and sclien_publicidad.hastafec <=" & DBSet(Text1(8).Text, "F")
        End If
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If TotalRegistrosConsulta(Sql) = 0 Then
            MsgBox "No hay registros para facturar", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Not MostrarFacturas(SQL2) Then Exit Sub
        DesBloqueoManual ("PUBLIFAC") 'facturas de publicidad
        If Not BloqueoManual("PUBLIFAC", "1") Then
            MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        If GenerarFacturas2 Then
            HacerImpresionFacturas
        End If
        
        DesBloqueoManual ("PUBLIFAC")
        TerminaBloquear
        
        Set miRsAux = Nothing
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Function MostrarFacturas(CADENA As String) As Boolean

MostrarFacturas = False

frmFacturas.Socio = True
frmFacturas.Sql = CADENA
frmFacturas.Caption = "Facturas de Publicidad a Socios"
frmFacturas.Show vbModal

If CadenaDesdeOtroForm <> "" Then
    MostrarFacturas = True
End If

End Function

Private Function HaySocios() As Boolean

HaySocios = False
If Text1(3).Text <> "" And Text1(5).Text <> "" Then
    HaySocios = True
End If

End Function

Private Function DatosOk() As Boolean
Dim Sql As String

    DatosOk = False
    
    'cliente
    If Text1(0).Text = "" Then
        MsgBox "Es necesario introducir un cliente para facturar.", vbExclamation
        PonerFoco Text1(0)
        Exit Function
    End If
    'fecha factu
    If Text1(1).Text = "" Then
        MsgBox "Es necesario introducir fecha de factura.", vbExclamation
        PonerFoco Text1(1)
        Exit Function
    End If
    'forma de pago
    If Text1(2).Text = "" Then
        MsgBox "Es necesario introducir la forma de pago para facturar.", vbExclamation
        PonerFoco Text1(2)
        Exit Function
    End If
    'concepto
    If Text1(4).Text = "" Then
        MsgBox "Es necesario introducir el concepto de la factura.", vbExclamation
        PonerFoco Text1(4)
        Exit Function
    End If
    If Text1(7).Text = "" Or Text1(8).Text = "" Then
        MsgBox "Debe introducir el rango de fechas a facturar.", vbExclamation
        PonerFoco Text1(7)
        Exit Function
    Else
        ' comprobamos que desde es menor que hasta
        If CDate(Text1(7).Text) > CDate(Text1(8).Text) Then
            MsgBox "Desde es mayor que hasta. Revise.", vbExclamation
            PonerFoco Text1(7)
            Exit Function
        End If
    End If
    
    If Text1(6).Text = "" Then
        MsgBox "Debe introducir un banco de cobro.", vbExclamation
        PonerFoco Text1(7)
        Exit Function
    Else
        Sql = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", Text1(6).Text, "N")
        If Sql = "" Then
            MsgBox "La cuenta prevista de cobro debe tener un valor. Revise", vbExclamation
            Exit Function
        End If
    End If
    
    DatosOk = True

End Function

Private Sub HacerImpresionFacturas()
    cadFormula = "{sfactusoc.fecfactu}= Date(" & Year(FecFactu) & "," & Month(FecFactu) & "," & Day(FecFactu) & ")"
    
    '[Monica]29/06/2018: falta que si me ha puesto socios d/h que lo seleccione
    If Text1(3).Text <> "" Then cadFormula = cadFormula & " and {sfactusoc.codsocio} >= " & Text1(3).Text
    If Text1(5).Text <> "" Then cadFormula = cadFormula & " and {sfactusoc.codsocio} <= " & Text1(5).Text
    
    cadFormula = cadFormula & " and {sfactusoc.codtipom}='FPS'"
    LlamarImprimir False
End Sub

Private Function GenerarFacturas() As Boolean
Dim vC As CTiposMov
Dim fac As CFacturaCom
Dim Cad As String
Dim Sql As String
Dim iva As Integer
Dim porIva As Currency
Dim totfactu As Currency
Dim BaseImp As Currency
Dim ImpIVA As Currency
Dim b As Boolean
Dim Socio As Long
Dim FormatSocio As String
Dim cuenta As String
Dim vDevuelve As String

Dim vSocio As CSocio

    On Error GoTo EGenFactu
    
    GenerarFacturas = False
    Cad = "FPS"

    conn.BeginTrans
    ConnConta.BeginTrans

    Set vC = New CTiposMov
'    Set fac = New CFacturaCom
    
    '[Monica]: modificado, el iva lo saco del articulo de publicidad
    iva = DevuelveValor("select codigiva from sartic where codartic = " & DBSet(vParamAplic.CodarticTfnia, "T"))
'    iva = vParamAplic.IVA_REA
    
    vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "N")
    porIva = 0
    If vDevuelve <> "" Then porIva = CCur(vDevuelve)
    
    b = True
    
    While Not miRsAux.EOF And b
    
        Set fac = New CFacturaCom
    
        If vC.TipoMovimiento <> Cad Then
            If Not vC.Leer(Cad) Then
                miRsAux.Close
                If NumRegElim > 0 Then MsgBox "Se han generado " & NumRegElim & " factura(s) antes del error", vbExclamation
                Exit Function
            End If
        End If
        'busco contador de cada socio y lo incremento
        NumFactu = ContadorSocio(miRsAux!codSocio, Cad, True)
        If NumFactu = 0 Then
            DesBloqueoManual ("PUBLIFAC")
            TerminaBloquear
            Exit Function
        End If
        'vC.IncrementarContador (vC.TipoMovimiento)
        fac.BaseImp = miRsAux!Importes
        fac.BrutoFac = miRsAux!Importes
        ImpIVA = Round2((fac.BaseImp * porIva) / 100, 2)
        totfactu = fac.BaseImp + ImpIVA
        fac.TotalFac = totfactu
        FecFactu = Text1(1).Text
        fac.FecFactu = FecFactu
        fac.NumFactu = NumFactu
        
        fac.Proveedor = miRsAux!codSocio
        fac.NombreProv = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", miRsAux!codSocio, "T")
        fac.DomicilioProv = DevuelveDesdeBD(conAri, "domclien", "sclien", "codclien", miRsAux!codSocio, "T")
        fac.CPostalProv = DevuelveDesdeBD(conAri, "codpobla", "sclien", "codclien", miRsAux!codSocio, "T")
        fac.PoblacionProv = DevuelveDesdeBD(conAri, "pobclien", "sclien", "codclien", miRsAux!codSocio, "T")
        fac.ProvinciaProv = DevuelveDesdeBD(conAri, "proclien", "sclien", "codclien", miRsAux!codSocio, "T")
        fac.NIFProv = DevuelveDesdeBD(conAri, "nifclien", "sclien", "codclien", miRsAux!codSocio, "T")
        fac.ForPago = Text1(2).Text
        
        'Cuenta Prevista de Cobro de las Facturas
        fac.BancoPr = Text1(6).Text
        fac.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", fac.BancoPr, "N")
        'cuenta contable de proveedor
        'comprobamos q la cuenta contable exista en contabilidad
        Socio = miRsAux!codSocio
        FormatSocio = String((vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior), "0")
        cuenta = Trim(vParamAplic.Raiz_Cta_Soc_publi & Format(Socio, FormatSocio))
        Sql = ""
        Sql = DevuelveDesdeBD(conConta, "nommacta", "cuentas", "codmacta", cuenta, "T")
        If Sql = "" Then
            MsgBox "La cuenta contable del socio: " & Text1(0).Text & " no existe.", vbExclamation
            DesBloqueoManual ("PUBLIFAC")
            TerminaBloquear
            ContadorSocio miRsAux!codSocio, Cad, False
            conn.RollbackTrans
            ConnConta.RollbackTrans
            Exit Function
        End If
        fac.CtaProve = cuenta
        
        '[Monica]a�adido no se cargaba la ccc del socio en tesoreria
        Set vSocio = New CSocio
        If vSocio.LeerDatos(CStr(Socio)) Then
            '[Monica]22/11/2013
            fac.CCC_Iban = vSocio.Iban
            fac.CCC_Entidad = vSocio.Banco
            fac.CCC_Oficina = vSocio.Sucursal
            fac.CCC_CC = vSocio.DigControl
            fac.CCC_CTa = vSocio.CuentaBan
        End If
        Set vSocio = Nothing
        
        
        'sfactusoc
        Sql = "INSERT INTO sfactusoc (codtipom,codsocio,numfactu,fecfactu,concepto,importel,baseiva1,impoiva1,"
        Sql = Sql & "codiiva1,porciva1,totalfac,impreten,intconta,codforpa) VALUES (" & DBSet(Cad, "T") & "," & miRsAux!codSocio & ","
        Sql = Sql & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "'," & DBSet(Text1(4).Text, "T") & ","
        Sql = Sql & TransformaComasPuntos(CStr(miRsAux!Importes)) & "," & TransformaComasPuntos(CStr(miRsAux!Importes)) & ","
        Sql = Sql & TransformaComasPuntos(CStr(ImpIVA)) & "," & TransformaComasPuntos(CStr(iva)) & "," & TransformaComasPuntos(CStr(porIva)) & ","
        Sql = Sql & TransformaComasPuntos(CStr(totfactu)) & ",NULL,0," & Text1(2).Text & ")"
    
        If Not ejecutar(Sql, False) Then
            DesBloqueoManual ("PUBLIFAC")
            TerminaBloquear
            ContadorSocio miRsAux!codSocio, Cad, False
            Exit Function
        End If
      
        b = fac.InsertarEnTesoreria("Error al pasar a tesoreria")
        
        Set fac = Nothing
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set vC = Nothing
    
    conn.CommitTrans
    ConnConta.CommitTrans
    
    GenerarFacturas = True
    
EGenFactu:
If Err.Number <> 0 Or Not b Then
    MsgBox "ERROR AL GENERAR FACTURAS:" & Err.Description
    DesBloqueoManual ("PUBLIFAC")
    TerminaBloquear
    ContadorSocio miRsAux!codSocio, Cad, False
    conn.RollbackTrans
    ConnConta.RollbackTrans
End If
End Function


Private Function ContadorSocio(Socio As Integer, codtipom As String, Accion As Boolean) As Integer
Dim Codigo As String
Dim Contador As String
Dim Sql As String

On Error GoTo EContador

Codigo = "codtipom='" & codtipom & "' and codsocio"
Contador = DevuelveDesdeBD(conAri, "contador", "sclien_contadores", Codigo, CStr(Socio), "T")
If Contador = "" Then
    MsgBox "Error grave, no existe contador para el movimiento:" & codtipom & " del socio:" & Socio, vbExclamation
    ContadorSocio = 0
    Exit Function
End If

If Accion Then
    ContadorSocio = CInt(Contador) + 1
    Sql = "UPDATE sclien_contadores SET contador=" & ContadorSocio & " where codsocio=" & Socio
    Sql = Sql & " and codtipom='" & codtipom & "'"
    conn.Execute Sql
Else
    Sql = "UPDATE sclien_contadores SET contador=" & CInt(Contador) - 1 & " where codsocio=" & Socio
    Sql = Sql & " and codtipom='" & codtipom & "'"
    conn.Execute Sql
End If

EContador:
If Err.Number <> 0 Then
    MsgBox "Error al modificar contador: " & Err.Description
    ContadorSocio = 0
End If
End Function


Private Sub LlamarImprimir(duplicado As Boolean)
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        cadParam = cadParam & "pDuplicado= " & Abs(duplicado) & " |"
        numParam = 2
        
        With frmImprimir
        .Titulo = "Impresi�n de Facturas de publicidad"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "48", "N")
    '------ > Listado 48 = rFacPublisoc.rpt
        .Opcion = 101
        .ConSubInforme = True
        .Show vbModal
    End With

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdRegresar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Icono del form
    Me.Icon = frmppal.Icon
    
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    
    Me.imgFecha(2).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    For kCampo = 12 To 13
        Me.imgFecha(kCampo).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    
    
    
    Text1(1).Text = Date
    Text1(4).Text = vParamAplic.ConFactuPubli
    Modo = 0
    numParam = 0
    cadFormula = ""
    cadParam = ""
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Cad = CadenaDevuelta
End Sub

Private Sub frmCal_Selec(vFecha As Date)
    Text1(indCodigo).Text = vFecha
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte
    Select Case Index
        Case 0
            CadenaDesdeOtroForm = Text1(4).Text
            frmFacClienteObser.Modificar = True
            frmFacClienteObser.Text1 = CadenaDesdeOtroForm
            frmFacClienteObser.Show vbModal
            If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(4).Text = Mid(CadenaDesdeOtroForm, 3)
            CadenaDesdeOtroForm = ""
            PonerFoco Text1(59)
        Case 1
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0|1|"
            frmFP.Show vbModal
            Set frmFP = Nothing
        Case 2
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0|1|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            PonerFoco Text1(0)
        Case 4, 3
            If Index = 4 Then
                indice = 3
            Else
                indice = 5
            End If
            Set frmSoc = New frmGesSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            If CadenaDesdeOtroForm <> "" Then
                Text1(indice).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                Text2(indice).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            End If
        Case 5
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
    End Select
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    Text1(6).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
   
    Set frmCal = New frmCal
    frmCal.Fecha = Now
    Select Case Index
        Case 2
            indCodigo = 1
            PonerFormatoFecha Text1(1)
            If Text1(1).Text <> "" Then frmCal.Fecha = CDate(Text1(1).Text)
            Screen.MousePointer = vbDefault
            frmCal.Show vbModal
            Set frmCal = Nothing
            PonerFoco Text1(1)
        Case 12
            indCodigo = 7
            PonerFormatoFecha Text1(7)
            If Text1(7).Text <> "" Then frmCal.Fecha = CDate(Text1(7).Text)
            Screen.MousePointer = vbDefault
            frmCal.Show vbModal
            Set frmCal = Nothing
            PonerFoco Text1(7)
        Case 13
            indCodigo = 8
            PonerFormatoFecha Text1(8)
            If Text1(8).Text <> "" Then frmCal.Fecha = CDate(Text1(8).Text)
            Screen.MousePointer = vbDefault
            frmCal.Show vbModal
            Set frmCal = Nothing
            PonerFoco Text1(8)
    End Select
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim encontrado As String

    Select Case Index
        Case 0 'cliente
            If Text1(Index).Text <> "" Then
                If Not IsNumeric(Text1(Index).Text) Then
                    MsgBox "El c�digo de cliente debe ser num�rico.", vbExclamation
                    PonerFoco Text1(Index)
                    Exit Sub
                End If
                Text1(Index).Text = Format(Text1(Index).Text, "000000")
                encontrado = DevuelveDesdeBD(conAri, "nomclien", "scliente", "codclien", Text1(Index).Text, "T")
                If encontrado = "" Then
                    MsgBox "El c�digo de cliente introducido no existe.", vbExclamation
                    PonerFoco Text1(Index)
                Else
                    Text2(Index).Text = encontrado
                End If
            End If
        Case 2 'forma de pago
            If Text1(Index).Text <> "" Then
                If Not IsNumeric(Text1(Index).Text) Then
                    MsgBox "La forma de pago debe ser num�rica.", vbExclamation
                    PonerFoco Text1(Index)
                    Exit Sub
                End If
                Text1(Index).Text = Format(Text1(Index).Text, "000")
                encontrado = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", Text1(Index).Text, "T")
                If encontrado = "" Then
                    MsgBox "La forma de pago introducida no existe.", vbExclamation
                    PonerFoco Text1(Index)
                Else
                    Text2(Index).Text = encontrado
                End If
            End If
        Case 1 'fecha
            If Text1(Index).Text <> "" Then
                PonerFormatoFecha Text1(Index)
            End If
        Case 3, 5 'socios
            If Text1(Index).Text <> "" Then
                If Not IsNumeric(Text1(Index).Text) Then
                    MsgBox "El c�digo de socio debe ser num�rico.", vbExclamation
                    PonerFoco Text1(Index)
                    Exit Sub
                End If
                Text1(Index).Text = Format(Text1(Index).Text, "000000")
                encontrado = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(Index).Text, "T")
                If encontrado = "" Then
                    MsgBox "El c�digo de socio introducido no existe.", vbExclamation
                    PonerFoco Text1(Index)
                Else
                    Text2(Index).Text = encontrado
                End If
            End If
        Case 6 'cuenta bancaria
            If Text1(Index).Text <> "" Then
                encontrado = DevuelveDesdeBD(conAri, "nombanpr", "sbanpr", "codbanpr", Text1(Index).Text, "T")
                If encontrado = "" Then
                    MsgBox "El banco introducido no existe", vbExclamation
                    PonerFoco Text1(Index)

                Else
                    Text2(Index).Text = encontrado
                End If
            End If
            
        Case 7, 8 'fechas
            PonerFormatoFecha Text1(Index)
        
    End Select
    
End Sub


Private Function GenerarFacturas2() As Boolean
Dim vC As CTiposMov
Dim vFactu As CFacturaSoc
Dim vSocio As CSocio

Dim codtipom As String
Dim Sql As String
Dim iva As Integer
Dim porIva As Currency
Dim totfactu As Currency
Dim BaseImp As Currency
Dim ImpIVA As Currency
Dim b As Boolean
Dim Socio As Long
Dim FormatSocio As String
Dim cuenta As String
Dim vDevuelve As String
Dim MenError As String

Dim devuelve As Long
Dim Existe As Boolean


    On Error GoTo EGenFactu
    
    GenerarFacturas2 = False
    
    codtipom = "FPS"

    conn.BeginTrans
    ConnConta.BeginTrans

    iva = DevuelveValor("select codigiva from sartic where codartic = " & DBSet(vParamAplic.CodarticTfnia, "T"))
    
    vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "N")
    porIva = 0
    If vDevuelve <> "" Then porIva = CCur(vDevuelve)
    
    b = True
    
    While Not miRsAux.EOF And b
        Set vFactu = New CFacturaSoc
        Set vSocio = New CSocio
        
        If vSocio.LeerDatos(miRsAux!codSocio) Then
            '[Monica]22/02/2018: no incrementaba el contador
            Do
                NumFactu = vSocio.ConseguirContador(codtipom)
                Sql = "select numfactu from sfactusoc where codtipom = " & DBSet(codtipom, "T") & " and numfactu = " & DBSet(NumFactu, "N") & " and codsocio = " & DBSet(vSocio.Codigo, "N")
                devuelve = DevuelveValor(Sql) 'DevuelveDesdeBDNew(cAgro, "rfacttra", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> 0 Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vSocio.IncrementarContador (codtipom)
                    NumFactu = vSocio.ConseguirContador(codtipom)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            vFactu.tipoMov = codtipom
            vFactu.BaseIVA1 = miRsAux!Importes
            vFactu.BrutoFac = miRsAux!Importes
            vFactu.TipoIVA1 = iva
            vFactu.PorceIVA1 = porIva
            
            vFactu.ImpIVA1 = Round2((vFactu.BaseIVA1 * porIva) / 100, 2)
            
            vFactu.TotalFac = vFactu.BaseIVA1 + vFactu.ImpIVA1
            
            FecFactu = Text1(1).Text
            
            vFactu.FecFactu = FecFactu
            vFactu.NumFactu = NumFactu
            vFactu.Socio = vSocio.Codigo
            vFactu.NombreSocio = vSocio.Nombre
            vFactu.DomicilioSocio = vSocio.Domicilio
            vFactu.CPostalSocio = vSocio.CPostal
            vFactu.PoblacionSocio = vSocio.Poblacion
            vFactu.ProvinciaSocio = vSocio.Provincia
            vFactu.nifSocio = vSocio.NIF
            vFactu.ForPago = Text1(2).Text
            vFactu.Concepto = Text1(4).Text
            
            'Cuenta Prevista de Cobro de las Facturas
            vFactu.BancoPr = Text1(6).Text
            vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", vFactu.BancoPr, "N")
            'cuenta contable de proveedor
            'comprobamos q la cuenta contable exista en contabilidad
            
            vFactu.CtaSocio = vSocio.CtaSocioPub
            '[Monica]22/11/2013: iban
            vFactu.CCC_Iban = vSocio.Iban
            vFactu.CCC_Entidad = vSocio.Banco
            vFactu.CCC_Oficina = vSocio.Sucursal
            vFactu.CCC_CC = vSocio.DigControl
            vFactu.CCC_CTa = vSocio.CuentaBan
            
            MenError = ""
            b = vFactu.InsertarFacturaPublicidad(MenError)
            If b Then
                 b = vFactu.InsertarEnTesoreria(MenError)
            End If
            
            If b Then b = vSocio.IncrementarContador(codtipom)

        Else
            b = False
        End If
        
        Set vFactu = Nothing
        Set vSocio = Nothing
        
        miRsAux.MoveNext
    Wend
    Set miRsAux = Nothing
    
    
EGenFactu:
If Err.Number <> 0 Or Not b Then
    MuestraError Err.Number, "Error al generar Facturas:" & Err.Description & " " & MenError
    conn.RollbackTrans
    ConnConta.RollbackTrans
    
    GenerarFacturas2 = False
    
    DesBloqueoManual ("PUBLIFAC")
    TerminaBloquear
Else
    conn.CommitTrans
    ConnConta.CommitTrans
    
    GenerarFacturas2 = True

End If
End Function


