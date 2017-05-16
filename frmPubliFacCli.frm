VERSION 5.00
Begin VB.Form frmPubliFacCli 
   Caption         =   "Informes"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   18
      Top             =   3480
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
      Index           =   5
      Left            =   2310
      MaxLength       =   4
      TabIndex        =   6
      Tag             =   "Código de Banco Propio|N|N|0|9999|sbanpr|codbanpr|0000|S|"
      Top             =   3480
      Width           =   945
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
      Left            =   3270
      TabIndex        =   16
      Top             =   1920
      Width           =   3885
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
      Left            =   6120
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   6120
      TabIndex        =   8
      Top             =   4080
      Width           =   1035
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
      Left            =   4980
      TabIndex        =   7
      Top             =   4080
      Width           =   1035
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
      Left            =   2310
      TabIndex        =   5
      Top             =   2850
      Width           =   4845
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
      Left            =   2310
      TabIndex        =   4
      Top             =   2400
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
      Index           =   2
      Left            =   2310
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "Codigo cliente|N|S|||shilla|codclien|000||"
      Top             =   1920
      Width           =   945
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
      Left            =   2310
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "99/99/9999"
      Top             =   1440
      Width           =   1245
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
      Left            =   3300
      TabIndex        =   14
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
      Left            =   2310
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "Codigo cliente|N|S|||shilla|codclien|000000||"
      Top             =   960
      Width           =   945
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   3
      Left            =   2040
      Tag             =   "-1"
      ToolTipText     =   "Buscar cuenta"
      Top             =   3510
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
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   2040
      Tag             =   "-1"
      ToolTipText     =   "Buscar forma de pago"
      Top             =   1920
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   2040
      Tag             =   "-1"
      ToolTipText     =   "Ver observaciones"
      Top             =   2880
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   2040
      Tag             =   "-1"
      ToolTipText     =   "Buscar cliente"
      Top             =   960
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   2
      Left            =   2040
      ToolTipText     =   "Buscar fecha"
      Top             =   1470
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
      TabIndex        =   13
      Top             =   2880
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "Importe a facturar"
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
      Top             =   2400
      Width           =   2535
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
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
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
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
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
      TabIndex        =   9
      Top             =   960
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Facturación Publicidad Clientes"
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
      Width           =   5325
   End
End
Attribute VB_Name = "frmPubliFacCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

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

Dim PrimeraVez As Boolean
Dim cadFormula As String
Dim cadParam As String
Dim NumFactu As Long
Dim FecFactu As Date
Dim numParam As Integer
Dim Modo As Byte
Dim cad As String


Dim kCampo As Integer

Private Sub cmdAceptar_Click()
Dim Sql As String
    
    Set miRsAux = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    
    If DatosOk Then
        Sql = "select count(*) from sclien_publicidad where codclien=" & Text1(0).Text
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            MsgBox "No hay vehiculos con esa publicidad.", vbExclamation
        Else
            Sql = miRsAux.Fields(0) & " vehiculos con publicidad del cliente " & Text1(0).Text
            MsgBox Sql, vbOKOnly
        End If
        
        DesBloqueoManual ("PUBLIFAC") 'facturas de publicidad
        If Not BloqueoManual("PUBLIFAC", "1") Then
            MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
            
        If GenerarFacturas Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            HacerImpresionFacturas
        End If
            
        DesBloqueoManual ("PUBLIFAC")
        TerminaBloquear
                
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Function DatosOk() As Boolean

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
    'importe facturar
    If Text1(3).Text = "" Then
        MsgBox "Es necesario introducir el importe a facturar.", vbExclamation
        PonerFoco Text1(3)
        Exit Function
    End If
    'concepto
    If Text1(4).Text = "" Then
        MsgBox "Es necesario introducir el concepto de la factura.", vbExclamation
        PonerFoco Text1(4)
        Exit Function
    End If
    
    DatosOk = True

End Function

Private Sub HacerImpresionFacturas()
    cadFormula = "({scafaccli.codclien}= " & Text1(0).Text & " and {scafaccli.numfactu}= "
    cadFormula = cadFormula & NumFactu & " and {scafaccli.fecfactu}= Date(" & Year(FecFactu) & "," & Month(FecFactu) & "," & Day(FecFactu) & "))"
    LlamarImprimir False
End Sub

Private Function GenerarFacturas() As Boolean
Dim vC As CTiposMov
Dim fac As CFactura
Dim cad As String
Dim Sql As String
Dim iva As String
Dim porIva As Currency
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


    On Error GoTo EGenFactu

    GenerarFacturas = False
    cad = "FPC"

    Set vC = New CTiposMov
    Set cli = New CCliente
    Set fac = New CFactura
    

    If vC.TipoMovimiento <> cad Then
        If Not vC.Leer(cad) Then
            miRsAux.Close
            If NumRegElim > 0 Then MsgBox "Se han generado " & NumRegElim & " factura(s) antes del error", vbExclamation
            Exit Function
        End If
    End If
    vC.IncrementarContador (vC.TipoMovimiento)
    
    iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.CodarticTfnia, "T")
    vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T")
    porIva = 0
    If vDevuelve <> "" Then porIva = CCur(vDevuelve)
    
    BaseImp = ImporteFormateado(Text1(3).Text)
    fac.BaseImp = BaseImp
    fac.BrutoFac = BaseImp
    ImpIVA = (BaseImp * porIva) / 100
    totfactu = BaseImp + ImpIVA
    fac.TotalFac = totfactu
    fac.codtipom = cad
    FecFactu = Text1(1).Text
    fac.FecFactu = FecFactu
    fac.LetraSerie = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", cad, "T")
    NumFactu = vC.Contador
    fac.NumFactu = NumFactu
    'Cuenta Prevista de Cobro de las Facturas
    fac.BancoPr = Text1(5).Text
    fac.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", fac.BancoPr, "N")
    'comprobar que la cuenta prevista de cobro tiene valor
    b = (fac.CuentaPrev <> "")
    If Not b Then
        Set fac = Nothing
        'Desbloqueamos ya no estamos facturando
        DesBloqueoManual ("PUBLIFAC")
        TerminaBloquear
        MsgBox "La cta. prevista de cobro debe tener valor.", vbExclamation
        Exit Function
    End If
    
    'datos del cliente
    Set miRsAux = New ADODB.Recordset
    
    Sql = "select * from scliente where codclien= " & Text1(0).Text
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    fac.Cliente = Text1(0).Text
    
    If Not miRsAux.EOF Then
        cli.Nombre = miRsAux!nomclien
        fac.NombreClien = miRsAux!nomclien
        cli.Domicilio = miRsAux!domclien
        fac.DomicilioClien = miRsAux!domclien
        cli.CPostal = miRsAux!codpobla
        fac.CPostal = miRsAux!codpobla
        cli.Poblacion = miRsAux!pobclien
        fac.Poblacion = miRsAux!pobclien
        cli.Provincia = miRsAux!proclien
        fac.Provincia = miRsAux!proclien
        cli.NIF = miRsAux!nifClien
        fac.NIF = miRsAux!nifClien
        
        '[Monica]04/02/2015: insertamos los datos de los bancos
        cli.Banco = Format(miRsAux!codbanco, "0000")
        fac.Banco = Format(miRsAux!codbanco, "0000")
        cli.Sucursal = Format(miRsAux!codsucur, "0000")
        fac.Sucursal = Format(miRsAux!codsucur, "0000")
        cli.DigControl = DBLet(miRsAux!digcontr, "T")
        fac.DigControl = DBLet(miRsAux!digcontr, "T")
        cli.CuentaBan = DBLet(miRsAux!cuentaba, "T")
        fac.CuentaBan = DBLet(miRsAux!cuentaba, "T")
        cli.Iban = DBLet(miRsAux!Iban, "T")
        fac.Iban = DBLet(miRsAux!Iban, "T")
        
        
    End If
    miRsAux.Close
    
    'scafac
    Sql = "INSERT INTO scafaccli (codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,"
    Sql = Sql & "nifclien,codagent,codforpa,dtoppago,dtognral,brutofac,impdtopp,impdtogr,baseimp1,codigiv1,porciva1,"
    Sql = Sql & "imporiv1,totalfac,intconta,coddirec, iban, codbanco, codsucur, digcontr, cuentaba ) VALUES ("
    Sql = Sql & DBSet(cad, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "'," & Text1(0).Text & ","
    Sql = Sql & DBSet(cli.Nombre, "T") & "," & DBSet(cli.Domicilio, "T") & "," & DBSet(cli.CPostal, "T") & ","
    Sql = Sql & DBSet(cli.Poblacion, "T") & "," & DBSet(cli.Provincia, "T") & "," & DBSet(cli.NIF, "T") & "," & vParamAplic.PorDefecto_Agente
    Sql = Sql & "," & Text1(2).Text & ",0,0," & TransformaComasPuntos(CStr(BaseImp)) & ",0,0," & TransformaComasPuntos(CStr(BaseImp)) & "," & iva
    Sql = Sql & "," & TransformaComasPuntos(CStr(porIva)) & "," & TransformaComasPuntos(CStr(ImpIVA)) & "," & TransformaComasPuntos(CStr(totfactu)) & ",0,NULL,"
    Sql = Sql & DBSet(fac.Iban, "T") & "," & DBSet(fac.Banco, "N", "S") & "," & DBSet(fac.Sucursal, "N", "S") & "," & DBSet(fac.DigControl, "T", "S") & "," & DBSet(fac.CuentaBan, "T", "S") & ")"
    
    fac.ForPago = Text1(2).Text
    
    If Not ejecutar(Sql, False) Then
        vC.DevolverContador vC.TipoMovimiento, vC.Contador
        Exit Function
    Else
        'scafac1
        'acoplamos el concepto a las observaciones de la scafac1
        tamanyo = Len(Text1(4).Text)
        tamanyo = tamanyo / 80
        Select Case tamanyo
            Case Is <= 1
                o1 = Text1(4).Text
            Case Is <= 2
                o1 = Mid(Text1(4).Text, 1, 80)
                o2 = Mid(Text1(4).Text, 81)
            Case Is <= 3
                o1 = Mid(Text1(4).Text, 1, 80)
                o2 = Mid(Text1(4).Text, 81, 160)
                o3 = Mid(Text1(4).Text, 161)
            Case Is <= 4
                o1 = Mid(Text1(4).Text, 1, 80)
                o2 = Mid(Text1(4).Text, 81, 160)
                o3 = Mid(Text1(4).Text, 161, 240)
                o4 = Mid(Text1(4).Text, 241)
            Case Else
                o1 = Mid(Text1(4).Text, 1, 80)
                o2 = Mid(Text1(4).Text, 81, 160)
                o3 = Mid(Text1(4).Text, 161, 240)
                o4 = Mid(Text1(4).Text, 241, 320)
                o5 = Mid(Text1(4).Text, 321, 400)
        End Select
        CodTraba = DevuelveDesdeBD(conAri, "codtraba", "straba", "login", vUsu.Login, "T")
        
        '[Monica]04/02/2015: si entrabamos como root no hacia nada
        If CodTraba = "" And vUsu.Login = "root" Then CodTraba = DevuelveValor("select min(codtraba) from straba")
        
        Sql = "INSERT INTO scafaccli1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,"
        Sql = Sql & "codenvio,codtraba,codtrab2,observa1,observa2,observa3,observa4,observa5,codtrab1) VALUES ("
        Sql = Sql & DBSet(cad, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0,'"
        Sql = Sql & Format(FecFactu, FormatoFecha) & "'," & vParamAplic.PorDefecto_Envio & "," & CodTraba
        Sql = Sql & "," & CodTraba & "," & DBSet(o1, "T") & "," & DBSet(o2, "T") & "," & DBSet(o3, "T") & ","
        Sql = Sql & DBSet(o4, "T") & "," & DBSet(o5, "T") & ",NULL)"
        
        conn.Execute Sql
        'slifac
        
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
        'busco el nombre del articulo
        NomArtic = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.CodarticTfnia, "T")
        
        Sql = "INSERT INTO slifacCli (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
        Sql = Sql & "numbultos,cantidad,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel) VALUES ("
        Sql = Sql & DBSet(cad, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0,0," & almac & ","
        Sql = Sql & DBSet(vParamAplic.CodarticTfnia, "T") & "," & DBSet(NomArtic, "T") & ",1,1," & TransformaComasPuntos(CStr(BaseImp)) & ","
        Sql = Sql & TransformaComasPuntos(CStr(BaseImp)) & "," & TransformaComasPuntos(CStr(BaseImp)) & "," & TransformaComasPuntos(CStr(BaseImp)) & ","
        Sql = Sql & TransformaComasPuntos(CStr(BaseImp)) & ",0,0,'M'," & Prove & "," & TransformaComasPuntos(CStr(BaseImp)) & ")"
        
        conn.Execute Sql
        
        'insertar en tesoreria
        fac.Agente = vParamAplic.PorDefecto_Agente
        b = fac.InsertarEnTesoreriaFACcli("", "Error al pasar a tesoreria")
        
    End If
    Set vC = Nothing
    
    GenerarFacturas = True
EGenFactu:
If Err <> 0 Then
    MsgBox "ERROR AL GENERAR FACTURAS:" & Err.Description
    DesBloqueoManual ("PUBLIFAC")
    TerminaBloquear
End If
End Function

Private Sub LlamarImprimir(duplicado As Boolean)
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        cadParam = cadParam & "pDuplicado= " & Abs(duplicado) & " |"
        numParam = 2
        
        With frmImprimir
        .Titulo = "Impresión de Facturas de publicidad"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "47", "N")
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
    Me.Icon = frmPpal.Icon

    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    
    Me.imgFecha(2).Picture = frmPpal.imgIcoForms.ListImages(2).Picture



    Text1(1).Text = Date
    Text1(4).Text = vParamAplic.ConFactuPubli
    Modo = 0
    numParam = 0
    cadFormula = ""
    cadParam = ""
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    cad = CadenaDevuelta
End Sub

Private Sub frmCal_Selec(vFecha As Date)
    Text1(1).Text = vFecha
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Text1(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)

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
        Case 3
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
    End Select
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    Text1(5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
   
    Set frmCal = New frmCal
    frmCal.Fecha = Now
    PonerFormatoFecha Text1(1)
    If Text1(1).Text <> "" Then frmCal.Fecha = CDate(Text1(1).Text)
    Screen.MousePointer = vbDefault
    frmCal.Show vbModal
    Set frmCal = Nothing
    PonerFoco Text1(1)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "scliente", "nomclien", "codclien", , "N")
            End If
            
        Case 2 'forma de pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa", "codforpa", , "N")
            End If
            
        Case 1 'fecha
            PonerFormatoFecha Text1(Index)
            
        Case 3 'importe
            PonerFormatoDecimal Text1(Index), 1
            
        Case 5 'banco propio
            If PonerFormatoEntero(Text1(5)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sbanpr", "nombanpr", "codbanpr", , "N")
            End If
        
    End Select
    
End Sub
