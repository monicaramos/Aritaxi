VERSION 5.00
Begin VB.Form frmFCliFactuVar 
   Caption         =   "Informes"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
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
      Index           =   6
      Left            =   5235
      MaxLength       =   14
      TabIndex        =   5
      Top             =   3510
      Width           =   1140
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
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
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2220
      Width           =   3825
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
      Index           =   3
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "Forma de Pago|N|N|||shilla|codforpa|000||"
      Top             =   2220
      Width           =   855
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
      Left            =   1620
      MaxLength       =   14
      TabIndex        =   4
      Top             =   3495
      Width           =   1455
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
      Index           =   1
      Left            =   1620
      MaxLength       =   16
      TabIndex        =   3
      Tag             =   "Articulo|T|N|||shilla|codartic|||"
      Top             =   2865
      Width           =   1425
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
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
      Left            =   3060
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2865
      Width           =   3255
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
      Index           =   0
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Cliente|N|N|||shilla|codclien|000000|S|"
      Top             =   1020
      Width           =   885
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1020
      Width           =   3795
   End
   Begin VB.TextBox txtnombre 
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
      Left            =   2400
      TabIndex        =   16
      Top             =   5040
      Width           =   3915
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
      Left            =   1620
      MaxLength       =   4
      TabIndex        =   7
      Tag             =   "Código de Banco Propio|N|N|0|9999|sbanpr|codbanpr|0000|S|"
      Top             =   5040
      Width           =   765
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
      Left            =   5160
      TabIndex        =   9
      Top             =   5700
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
      TabIndex        =   8
      Top             =   5700
      Width           =   1135
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
      Index           =   4
      Left            =   1620
      MaxLength       =   255
      TabIndex        =   6
      Top             =   4230
      Width           =   4695
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
      Index           =   2
      Left            =   1620
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "99/99/9999"
      Top             =   1620
      Width           =   1215
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
      Left            =   5250
      TabIndex        =   14
      Top             =   5700
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "% Retención"
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
      Index           =   0
      Left            =   3735
      TabIndex        =   23
      Top             =   3555
      Width           =   1395
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1380
      Top             =   2250
      Width           =   240
   End
   Begin VB.Label Label5 
      Caption         =   "Artículo"
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
      Left            =   210
      TabIndex        =   22
      Top             =   2880
      Width           =   1035
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
      Left            =   180
      TabIndex        =   20
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   7
      Left            =   210
      TabIndex        =   19
      Top             =   3510
      Width           =   1695
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   5
      Left            =   1380
      Top             =   2895
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   1380
      Top             =   1020
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   3
      Left            =   1380
      Tag             =   "-1"
      ToolTipText     =   "Buscar cuenta"
      Top             =   5070
      Width           =   240
   End
   Begin VB.Label Label7 
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
      Height          =   375
      Left            =   210
      TabIndex        =   15
      Top             =   4650
      Width           =   1665
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1380
      Tag             =   "-1"
      ToolTipText     =   "Ver observaciones"
      Top             =   4230
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   2
      Left            =   1380
      ToolTipText     =   "Buscar fecha"
      Top             =   1650
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "Ampliación"
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
      Left            =   210
      TabIndex        =   13
      Top             =   3930
      Width           =   1905
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
      Left            =   210
      TabIndex        =   12
      Top             =   1380
      Width           =   1995
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
      Left            =   210
      TabIndex        =   11
      Top             =   1020
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Facturas Varias de Clientes"
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
      TabIndex        =   10
      Top             =   240
      Width           =   5325
   End
End
Attribute VB_Name = "frmFCliFactuVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 320


Public WithEvents frmCal As frmCal
Attribute frmCal.VB_VarHelpID = -1
Public WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Public WithEvents frmFP As frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Public WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios
Attribute frmMtoBancosPro.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmArt As frmAlmArticulos   'Form Articulos
Attribute frmArt.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim NumFactu As Long
Dim FecFactu As Date
Dim Modo As Byte
Dim Cad As String
Dim Codigo As String
Dim CadServicios As String
Dim Salir As Boolean

Dim Tabla As String
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim codtipom As String
Dim cadSelect As String
Dim indCodigo As Long
Dim FacturasaImprimir As String

Dim kCampo As Integer


Private Sub cmdAceptar_Click()
Dim Sql As String
Dim b As Boolean

        Set miRsAux = New ADODB.Recordset
        Screen.MousePointer = vbHourglass
        
        If Not DatosOk Then Exit Sub
         
        InicializarVbles
         
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = 1
        
             
        DesBloqueoManual ("FACCLI") 'facturas de publicidad
        If Not BloqueoManual("FACCLI", "1") Then
            MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
            
        If GenerarFacturas(cadSelect, Tabla) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            HacerImpresionFacturas
        End If
        
        DesBloqueoManual ("FACCLI")
        TerminaBloquear
                
        Screen.MousePointer = vbDefault
        
        cmdCancelar_Click
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub



Private Function DatosOk() As Boolean
Dim Sql As String

    DatosOk = False
    'fecha factu
    If txtcodigo(2).Text = "" Then
        MsgBox "Es necesario introducir fecha de factura.", vbExclamation
        PonerFoco txtcodigo(2)
        Exit Function
    End If
    'concepto
    If txtcodigo(4).Text = "" Then
        MsgBox "Es necesario introducir el concepto de la factura.", vbExclamation
        PonerFoco txtcodigo(4)
        Exit Function
    End If
    
    'banco
    If txtcodigo(5).Text = "" Then
        MsgBox "Es necesario introducir el banco de cobro.", vbExclamation
        PonerFoco txtcodigo(5)
        Exit Function
    Else
        Sql = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", txtcodigo(5).Text, "N")
        If Sql = "" Then
            MsgBox "La Cta.Contable prevista de cobro del banco debe tener valor.", vbExclamation
            Exit Function
        End If
    End If
    If txtcodigo(1).Text = "" Then
        MsgBox "Es necesario introducir el artículo a facturar.", vbExclamation
        PonerFoco txtcodigo(1)
        Exit Function
    Else
        txtnombre(1).Text = PonerNombreDeCod(txtcodigo(1), conAri, "sartic", "nomartic", "codartic", "Artículo", "T")
        If txtnombre(1).Text = "" Then
            MsgBox "El artículo debe de existir. Revise.", vbExclamation
            PonerFoco txtcodigo(1)
            Exit Function
        End If
    End If
    If ComprobarCero(txtcodigo(85).Text) = 0 Then
        MsgBox "El importe de factura debe ser distinto de 0. Revise.", vbExclamation
        PonerFoco txtcodigo(85)
        Exit Function
    End If
    DatosOk = True

End Function

Private Sub HacerImpresionFacturas()
    cadFormula = "({scafaccli.codtipom}= ""FVC""" & " and {scafaccli.numfactu} in [" & Mid(FacturasaImprimir, 1, Len(FacturasaImprimir) - 1) & "]"
    cadFormula = cadFormula & " and {scafaccli.fecfactu}= Date(" & Year(FecFactu) & "," & Month(FecFactu) & "," & Day(FecFactu) & "))"
    LlamarImprimir False
End Sub

Private Function GenerarFacturas(cWhere As String, cTabla As String) As Boolean
Dim vTipoMov As CTiposMov
Dim fac As CFactura
Dim TipoMovimiento As String
Dim Sql As String
Dim iva As String
Dim porIva As Currency

Dim porIvaServ As Currency
Dim porIvaGtos As Currency

Dim totfactu As Currency
Dim BaseImp As Currency
Dim ImpIVA As Currency
Dim cli As CCliente
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim o4 As String
Dim o5 As String
Dim tamanyo As Double
Dim almac As String
Dim Prove As String
Dim NomArtic As String
Dim CodTraba As String
Dim b As Boolean
Dim vDevuelve As String

Dim BaseivaGtos As Currency
Dim ImpivaGtos As Currency
Dim BaseivaServ As Currency
Dim ImpivaServ As Currency

Dim RsServ As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim devuelve As String
Dim Existe As Boolean
Dim SQL2 As String
Dim SQL3 As String

Dim linea As Long

Dim cadWHERE As String
Dim Mens As String

Dim Suplidos As Currency
Dim DtoGnral As Currency


    On Error GoTo EGenFactu

    GenerarFacturas = False
    '[Monica]03/10/2012: Nuevo tipo de movimiento
    TipoMovimiento = "FVC" '"FAC"

    conn.BeginTrans
    ConnConta.BeginTrans

    FecFactu = txtcodigo(2).Text
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    
    Set miRsAux = New ADODB.Recordset
    
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
    
    Set miRsAux = Nothing
    
    
    b = True
    
    FacturasaImprimir = ""
    
    Set cli = New CCliente
    Set fac = New CFactura
    
    If cli.LeerDatos(txtcodigo(0).Text, False) Then
        
        Set vTipoMov = New CTiposMov
        If vTipoMov.Leer(TipoMovimiento) Then
            NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
            ' si existe la factura incrementamos el contador
            Do
                devuelve = DevuelveDesdeBDNew(conAri, "scafaccli", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", CStr(FecFactu), "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (TipoMovimiento)
                    NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
        Else
            Exit Function
        End If
    
        ' calculo de base iva de GASTOS ADMON
        iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", txtcodigo(1).Text, "T")
        vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T")
        porIvaGtos = 0
        If vDevuelve <> "" Then porIvaGtos = CCur(vDevuelve)
        
        BaseivaGtos = ImporteSinFormato(txtcodigo(85).Text)
        ImpivaGtos = Round2(BaseivaGtos * porIvaGtos / 100, 2)
        
        ' Asignamos los importes a la factura
        fac.TipoIVA1 = iva
        fac.BaseIVA1 = BaseivaGtos
        fac.PorceIVA1 = porIvaGtos
        fac.ImpIVA1 = ImpivaGtos
        
        fac.BaseImp = BaseivaGtos
        fac.ImpGnral = DtoGnral
        fac.DtoGnral = cli.DtoGnral
        fac.BrutoFac = fac.BaseImp
        fac.Suplidos = 0
        
        '[Monica]28/06/2018: porcentaje de retencion
        fac.PorRet = ComprobarCero(txtcodigo(6).Text)
        If fac.PorRet <> 0 Then
            fac.ImpReten = Round2(fac.BrutoFac * fac.PorRet / 100, 2)
        Else
            fac.ImpReten = 0
        End If
        
        fac.TotalFac = BaseivaGtos + ImpivaGtos - fac.ImpReten
        
        fac.codtipom = TipoMovimiento
        
        fac.FecFactu = FecFactu
        fac.LetraSerie = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", TipoMovimiento, "T")
'            NumFactu = vTipoMov.Contador
        fac.NumFactu = NumFactu
        FacturasaImprimir = FacturasaImprimir & NumFactu & ","
        
        'Cuenta Prevista de Cobro de las Facturas
        fac.BancoPr = txtcodigo(5).Text
        fac.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", fac.BancoPr, "N")
    
        fac.Cliente = cli.Codigo
        fac.NombreClien = cli.Nombre
        fac.DomicilioClien = cli.Domicilio
        fac.CPostal = cli.CPostal
        fac.Poblacion = cli.Poblacion
        fac.Provincia = cli.Provincia
        fac.NIF = cli.NIF
        '[Monica]10/10/2012: la forma de pago la cogemos del frame
        fac.ForPago = txtcodigo(3).Text 'cli.ForPago
    
        '[Monica]22/11/2013: iban
        fac.Iban = cli.Iban
        fac.Banco = cli.Banco
        fac.Sucursal = cli.Sucursal
        fac.DigControl = cli.DigControl
        fac.CuentaBan = cli.CuentaBan
        
        
        
        Mens = "Insertando en cabecera factura"
        'scafaccli
        Sql = "INSERT INTO scafaccli (codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,"
        Sql = Sql & "nifclien,codagent,codforpa,dtoppago,dtognral,brutofac,impdtopp,impdtogr,baseimp1,codigiv1,porciva1,"
        Sql = Sql & "imporiv1,baseimp2,codigiv2,porciva2,imporiv2,totalfac,intconta,coddirec,codbanco,codsucur,digcontr,cuentaba, numservi, suplidos, iban,"
        
        '[Monica]28/06/2018: insertamos el porcentaje de retencion y el importe de retencion
        Sql = Sql & "porcret, impreten) VALUES ("
        
        Sql = Sql & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "'," & DBSet(fac.Cliente, "N") & ","
        Sql = Sql & DBSet(cli.Nombre, "T") & "," & DBSet(cli.Domicilio, "T") & "," & DBSet(cli.CPostal, "T") & ","
        Sql = Sql & DBSet(cli.Poblacion, "T") & "," & DBSet(cli.Provincia, "T") & "," & DBSet(cli.NIF, "T") & "," & vParamAplic.PorDefecto_Agente
        'cli.ForPago
        Sql = Sql & "," & DBSet(txtcodigo(3).Text, "N") & ",0," & DBSet(fac.DtoGnral, "N") & "," & DBSet(fac.BrutoFac, "N") & ",0," & DBSet(fac.ImpGnral, "N") & ","
        Sql = Sql & DBSet(fac.BaseIVA1, "N") & "," & DBSet(fac.TipoIVA1, "N")
        Sql = Sql & "," & DBSet(fac.PorceIVA1, "N") & "," & DBSet(fac.ImpIVA1, "N") & ","
        Sql = Sql & DBSet(fac.BaseIVA2, "N", "S") & "," & DBSet(fac.TipoIVA2, "N", "S") & "," & DBSet(fac.PorceIVA2, "N", "S") & ","
        Sql = Sql & DBSet(fac.ImpIVA2, "N", "S") & "," & DBSet(fac.TotalFac, "N") & ",0,NULL,"
        Sql = Sql & DBSet(fac.Banco, "N") & "," & DBSet(fac.Sucursal, "N") & "," & DBSet(fac.DigControl, "T") & "," & DBSet(fac.CuentaBan, "T") & ","
        Sql = Sql & DBSet(1, "N") & "," & DBSet(Suplidos, "N") & "," & DBSet(fac.Iban, "T") & ","
        Sql = Sql & DBSet(fac.PorRet, "N", "S") & "," & DBSet(fac.ImpReten, "N", "S") & ")"
    
    
        conn.Execute Sql
    
        o1 = DevuelveDesdeBD(conAri, "observa1", "scliente", "codclien", cli.Codigo, "N") '("select observa1 from scliente where codclien = " & DBSet(cli.Codigo, "N"))
        
        
        CodTraba = DevuelveDesdeBD(conAri, "codtraba", "straba", "login", vUsu.Login, "T")
        If CodTraba = "" Then CodTraba = DevuelveValor("select min(codtraba) from straba")
        Mens = "Insertando Albaran"
        
        Sql = "INSERT INTO scafaccli1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,"
        Sql = Sql & "codenvio,codtraba,codtrab2,observa1,observa2,observa3,observa4,observa5,codtrab1) VALUES ("
        Sql = Sql & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0,'"
        Sql = Sql & Format(FecFactu, FormatoFecha) & "'," & vParamAplic.PorDefecto_Envio & "," & CodTraba
        Sql = Sql & "," & CodTraba & "," & DBSet(o1, "T") & "," & DBSet(o2, "T") & "," & DBSet(o3, "T") & ","
        Sql = Sql & DBSet(o4, "T") & "," & DBSet(o5, "T") & ",NULL)"
        
        conn.Execute Sql
        'slifac
        
        
        Mens = "Insertando linea de articulo"
        
        NomArtic = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtcodigo(1).Text, "T")
        
        Sql = "INSERT INTO slifacCli (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
        Sql = Sql & "numbultos,cantidad,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel,ampliaci ) VALUES ("
        Sql = Sql & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0,1," & almac & ","
        Sql = Sql & DBSet(txtcodigo(1).Text, "T") & "," & DBSet(NomArtic, "T") & ",1,1," & DBSet(BaseivaGtos, "N") & ","
        Sql = Sql & DBSet(BaseivaGtos, "N") & "," & DBSet(BaseivaGtos, "N") & "," & DBSet(BaseivaGtos, "N") & ","
        Sql = Sql & DBSet(BaseivaGtos, "N") & ",0,0,'M'," & Prove & "," & DBSet(BaseivaGtos, "N") & "," & DBSet(txtcodigo(4).Text, "T") & ")"
        
        conn.Execute Sql
        
        
        'insertar en tesoreria
        fac.Agente = vParamAplic.PorDefecto_Agente
        
        b = fac.InsertarEnTesoreriaFACcli("", "Error al pasar a Tesoreria")
        'b = fac.InsertarEnTesoreriaFACcli("", "Error al pasar a tesoreria")
    
        If b Then vTipoMov.IncrementarContador (TipoMovimiento)
        
        
        Set vTipoMov = Nothing
        Set cli = Nothing
        Set fac = Nothing
    
    Else
        MsgBox "No existe el cliente " & cli.Codigo & " " & cli.Nombre
        b = False
    End If

    
    GenerarFacturas = True

EGenFactu:
    If Err.Number <> 0 Then
        Mens = "Insertando Movimiento." & vbCrLf & "----------------------------" & vbCrLf & Mens
        MuestraError Err.Number, Mens, Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        GenerarFacturas = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        GenerarFacturas = False
    End If
    
End Function

Private Sub LlamarImprimir(duplicado As Boolean)
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        cadParam = cadParam & "pDuplicado= " & Abs(duplicado) & " |"
        numParam = 2
        
        With frmImprimir
        .Titulo = "Impresión de Facturas de Cliente"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "52", "N")
    '------ > Listado 47 = rFacPubli.rpt
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

    txtcodigo(2).Text = Date
'    Text1(4).Text = vParamAplic.ConFactuPubli
    Modo = 0
    '[Monica]19/02/2018: Entra Cordoba
        '[Monica]19/11/2018: Entra Sevilla
    If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3 Then
        Tabla = "sfactclitr"
    Else
        Tabla = "shilla"
    End If
    For kCampo = 0 To 1
        Me.imgBuscar(kCampo).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 3 To 5
        Me.imgBuscar(kCampo).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    Me.imgFecha(2).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    
    
    '[Monica]28/06/2018: la retencion solo la tiene cordoba (pq  son facturas de alquiler)
        '[Monica]19/11/2018: Entra Sevilla
    Label17(0).visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3)
    Label17(0).Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3)
    txtcodigo(6).visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3)
    txtcodigo(6).Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3)
    
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Cad = CadenaDevuelta
End Sub

Private Sub frmCal_Selec(vFecha As Date)
    txtcodigo(1).Text = vFecha
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    
    If CadenaSeleccion = "Salir" Then
        Salir = True
        Exit Sub
    End If
    
    If CadenaSeleccion <> "" Then
        CadServicios = "(fecha,hora,numeruve) in (" & CadenaSeleccion & ")"
    Else
        CadServicios = "shilla.numeruve = -1"
    End If
    
    If Not AnyadirAFormula(cadSelect, CadServicios) Then Exit Sub
    
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    Select Case Index
        Case 0
            CadenaDesdeOtroForm = txtcodigo(4).Text
            frmFacClienteObser.Modificar = True
            frmFacClienteObser.Text1 = CadenaDesdeOtroForm
            frmFacClienteObser.Show vbModal
            If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then txtcodigo(4).Text = Mid(CadenaDesdeOtroForm, 3)
            CadenaDesdeOtroForm = ""
            PonerFoco txtcodigo(4)
        Case 1
            indCodigo = Index + 2
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0|1|"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 4
            indCodigo = Index - 4
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0|1|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            PonerFoco txtcodigo(indCodigo)
            
        Case 5 ' articulo a facturar
            indCodigo = 1
            Set frmArt = New frmAlmArticulos
            frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en Modo busqueda
            frmArt.DeConsulta = True
            frmArt.Show vbModal
            Set frmArt = Nothing
            
        Case 3
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
    End Select
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    txtcodigo(5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtnombre(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
   
    Set frmCal = New frmCal
    frmCal.Fecha = Now
    PonerFormatoFecha txtcodigo(1)
    If txtcodigo(1).Text <> "" Then frmCal.Fecha = CDate(txtcodigo(1).Text)
    Screen.MousePointer = vbDefault
    frmCal.Show vbModal
    Set frmCal = Nothing
    PonerFoco txtcodigo(1)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
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
    
    
    Select Case Index
            
        Case 0 'cliente
            PonerFormatoEntero txtcodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "scliente", "nomclien", "codclien", "Cliente", "N")
            If txtnombre(Index).Text <> "" Then
                txtcodigo(3).Text = DevuelveDesdeBDNew(conAri, "scliente", "codforpa", "codclien", txtcodigo(0).Text, "N")
                If txtcodigo(3).Text <> "" Then
                    txtnombre(3).Text = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", txtcodigo(3).Text, "N")
                End If
            End If
            
        Case 1 'articulo
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sartic", "nomartic", "codartic", "Artículo", "T")
            
        Case 3 ' forma de pago
            PonerFormatoEntero txtcodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sforpa", "nomforpa", "codforpa", "Forma de Pago", "N")
            
        Case 2 ' fecha de factura
            PonerFormatoFecha txtcodigo(Index)
             
        Case 5 'banco propio
            If PonerFormatoEntero(txtcodigo(5)) Then
                txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sbanpr", "nombanpr", "codbanpr", , "N")
            End If
           
        Case 85  'Importe
            PonerFormatoDecimal txtcodigo(Index), 3
           
           
    End Select
End Sub

Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
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


