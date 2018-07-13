VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'VENTAS
'FACTURA: se corresponde con la tabla aritaxi.scafac (cabecera de facturas)
'-------------------------------------------------------------------------

'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodTipoM As String 'Tipo de Movimiento="FAV", "FAR" (Factura venta o reparacion)
Private mNumFactu As Long 'Nº Factura
Private mFecFactu As String 'FEcha Factura
Private mCodClien As String 'Cod. cliente

Private mNomClien As String 'nombre del cliente
Private mDomClien As String 'domicilio
Private mCodpobla As String 'cpostal
Private mPobClien As String 'poblacion
Private mProClien As String 'provincia
Private mNIFClien As String 'NIF proveedor
Private mTelclien As String 'telefono

Private mCoddirec As String 'codigo direc./dpto
Private mNomdirec As String  'nombre direc./dpto
Private mCodAgent As Integer 'codigo de agente

'[Monica]22/11/2013: iban
Private mIban As String 'iban

Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria

Private mCodForpa As String 'Cod. Forma de pago
Private mTipForpa As Integer 'tipo de forma de pago: efectivo,...

Private mDtoPpago As Currency 'descuento pronto pago
Private mDtoGnral As Currency 'descuento general
Private mBrutofac As Currency 'importe bruto de la factura
Private mImpPPago As Currency 'Importe al aplicar al bruto el dto pronto pago
Private mImpGnral As Currency  'Importe al aplicar al bruto el dto general
Private mBaseImp As Currency  'Base imponible de la factura (bruto - dtopp - dtogn)

Private mBaseiva1 As Currency  'base imponible tipo IVA 1
Private mBaseiva2 As Currency  'base imponible tipo IVA 2
Private mBaseiva3 As Currency  'base imponible tipo IVA 3
Private mTipoiva1 As Byte   'Cod. tipo de iva 1
Private mTipoiva2 As Byte   'Cod. tipo de iva 2
Private mTipoiva3 As Byte   'Cod. tipo de iva 3
Private mPorciva1 As Currency    '% de iva 1
Private mPorciva2 As Currency    '% de iva 2
Private mPorciva3 As Currency    '% de iva 3

Private mSuplidos As Currency  ' suplidos


Private mImpoiva1 As Currency   'Importe de iva 1
Private mImpoiva2 As Currency   'Importe de iva 2
Private mImpoiva3 As Currency   'Importe de iva 3


'DAVID.   ENERO 2008
Private mPorciva1RE As Currency    '% de iva 1
Private mPorciva2RE As Currency    '% de iva 2
Private mPorciva3RE As Currency    '% de iva 3

Private mImpoiva1RE As Currency   'Importe de iva 1
Private mImpoiva2RE As Currency   'Importe de iva 2
Private mImpoiva3RE As Currency   'Importe de iva 3





Private mTotalfac As String 'total factura

Private mLetraSer As String 'Letra de serie para los cobros
Private mBanPr As String 'Banco propio que ponemos como cuenta prevista de cobro
Private mCtaPrev As String 'Cuenta prevista de cobro

Private mNumtermi As Integer 'terminal de venta del TPV
Private mNumventa As Long 'Nº de venta del TPV


'David.
'Para las facturas de mantenimeinto, hay un campo que es "concepto factura" varchar(60)
'Si tiene valor entonces ese valor es el que grabaremos en la linea y como obervacion de la factura
'meteremos lo de antes ("texto" de: mes de año.  Ejemplo: manteimiento de Junio de 2007
'Si no tiene valor seguira como estaba
Private mObservacion As String



' VBLES para los Mantenimientos -------------------------------------------
Dim TipCoMan As String 'tipo de contrato para Facturas de Manteniminetos
Dim OpeFactu As String 'Operador conectado o el que pasamos como parametro
Dim AlmFactu As String 'Cod Almacen del trabajador conectado, para insertar lines mantenim, desplaz, bonificaciones,....
Dim MesFactu As String
Dim nummante As String
Dim ArticFac As String 'cod articulo que se va a utilizar para facturar mantenimientos
Dim manCCost As String 'Centro de coste  para las lineas
'--------------------------------------------------------------------------

' VBLES para los Desplazamientos ----------------------------------
Dim TotalKm As Integer
Dim PrecioKm As Currency
Dim ImporteL As Currency
Dim Ampliacion As String  'ampliacion de la linea: "Dia,Dia,Dia..." de los albaranes con desplazamiento
Dim codArtic As String
Dim despIVA As Integer 'tipo de IVA del articulo para desplazamientos
'------------------------------------------------------------------

' VBLES para las Bonificaciones -----------------------------------
Dim TotBonif1 As Currency 'Total de bonificacion a factura con articulo de tipo de IVA1
Dim IVABonif1 As Byte 'tipo de IVA
Dim TotBonif2 As Currency
Dim IVABonif2 As Byte
Dim TotBonif3 As Currency
Dim IVABonif3 As Byte
Dim SQLBonif As String
'------------------------------------------------------------------

' VBLES para los Cheques Regalo -----------------------------------
Dim ImpCheque As Currency
'------------------------------------------------------------------




'vbles para las facturas que vienen de una venta de ticket ---------
Private NumTicket As String
Private NumAlbTicket As String '01/09/06 laura
'-------------------------------------------------------------------

Private mAportacion As Currency   'Viene de Radival, pero se podria utilizar en mas
                                 ' Es la aportacion que hace Telefonica, en los moviles
                                 

Private mPorRet As Currency
Private mImpReten As Currency


'# Laura 14/11/2006 recuperar facturas Alzira
Private EstaRecuFac As Boolean

'***TEXTCSB
Dim mvarText33csb As String

Dim mvarText41csb As String
Dim mvarText42csb As String
Dim mvarText43csb As String

Dim mvarText51csb As String
Dim mvarText52csb As String
Dim mvarText53csb As String

Dim mvarText61csb As String
Dim mvarText62csb As String
Dim mvarText63csb As String
                    
Dim mvarText71csb As String
Dim mvarText72csb As String
Dim mvarText73csb As String

Dim mvarText81csb As String
Dim mvarText82csb As String
Dim mvarText83csb As String

'------------------------------------------------
'Propiedades del modulo CFactura
'------------------------------------------------

'**** Tipo de Movimiento de la Factura

Public Property Let codtipom(ByVal vData As String)
     mCodTipoM = vData
End Property

Public Property Get codtipom() As String
     codtipom = mCodTipoM
End Property


'**** Nº de la Factura

Public Property Let NumFactu(ByVal vData As Long)
     mNumFactu = vData
End Property

Public Property Get NumFactu() As Long
     NumFactu = mNumFactu
End Property


'**** Fecha de la Factura

Public Property Let FecFactu(ByVal vData As String)
     mFecFactu = vData
End Property

Public Property Get FecFactu() As String
     FecFactu = mFecFactu
End Property


'**** Cliente de la Factura

Public Property Let Cliente(ByVal vData As String)
     mCodClien = vData
End Property

Public Property Get Cliente() As String
     Cliente = mCodClien
End Property


'**** Nombre del Cliente de la Factura

Public Property Let NombreClien(ByVal vData As String)
     mNomClien = vData
End Property

Public Property Get NombreClien() As String
     NombreClien = mNomClien
End Property


'**** Domicilio del Cliente de la Factura

Public Property Let DomicilioClien(ByVal vData As String)
     mDomClien = vData
End Property

Public Property Get DomicilioClien() As String
     DomicilioClien = mDomClien
End Property


'**** CPostal del Cliente de la Factura

Public Property Let CPostal(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodpobla
End Property


'**** Poblacion del Cliente de la Factura

Public Property Let Poblacion(ByVal vData As String)
     mPobClien = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobClien
End Property


'**** Provincia del Cliente de la Factura

Public Property Let Provincia(ByVal vData As String)
     mProClien = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProClien
End Property


'**** NIF del Cliente de la Factura

Public Property Let NIF(ByVal vData As String)
     mNIFClien = vData
End Property

Public Property Get NIF() As String
     NIF = mNIFClien
End Property


'**** Telefono del Cliente de la Factura

Public Property Let Telefono(ByVal vData As String)
     mTelclien = vData
End Property

Public Property Get Telefono() As String
     Telefono = mTelclien
End Property


'**** Direc./Dpto de Cliente de la Factura

Public Property Let DirDpto(ByVal vData As String)
     mCoddirec = vData
End Property

Public Property Get DirDpto() As String
     DirDpto = mCoddirec
End Property


'**** Nombre Direc./Dpto de Cliente de la Factura

Public Property Let NombreDirDpto(ByVal vData As String)
     mNomdirec = vData
End Property

Public Property Get NombreDirDpto() As String
     NombreDirDpto = mNomdirec
End Property


'**** Agente de Cliente de la Factura

Public Property Let Agente(ByVal vData As Integer)
     mCodAgent = vData
End Property

Public Property Get Agente() As Integer
     Agente = mCodAgent
End Property


'**** Forma Pago de la Factura

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property


'**** Tipo de Forma Pago de la Factura
'(para tickets saber si insertar en tesoreria o no: si es efectivo no se inserta en scobro)

Public Property Let TipForPago(ByVal vData As Integer)
     mTipForpa = vData
End Property

Public Property Get TipForPago() As Integer
     TipForPago = mTipForpa
End Property


'**** Cuenta Prevista de cobro de la Factura

Public Property Let LetraSerie(ByVal vData As String)
     mLetraSer = vData
End Property

Public Property Get LetraSerie() As String
     LetraSerie = mLetraSer
End Property


'**** Banco propio de pago de la Factura

Public Property Let BancoPr(ByVal vData As String)
     mBanPr = vData
End Property

Public Property Get BancoPr() As String
     BancoPr = mBanPr
End Property


'**** Cuenta Prevista de cobro de la Factura

Public Property Let CuentaPrev(ByVal vData As String)
     mCtaPrev = vData
End Property

Public Property Get CuentaPrev() As String
     CuentaPrev = mCtaPrev
End Property

'**** IBAN de la Factura

Public Property Let Iban(ByVal vData As String)
     mIban = vData
End Property

Public Property Get Iban() As String
     Iban = mIban
End Property


'**** Banco de la Factura

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal de la Factura

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property


'**** Digito Control de la Factura

Public Property Let DigControl(ByVal vData As String)
     mDigContr = vData
End Property

Public Property Get DigControl() As String
     DigControl = mDigContr
End Property


'**** Cuenta Bancaria de la Factura

Public Property Let CuentaBan(ByVal vData As String)
     mCuentaba = vData
End Property

Public Property Get CuentaBan() As String
     CuentaBan = mCuentaba
End Property


'**** Descuento pronto pago de la Factura

Public Property Let DtoPPago(ByVal vData As Currency)
     mDtoPpago = vData
End Property

Public Property Get DtoPPago() As Currency
     DtoPPago = mDtoPpago
End Property


'**** Descuento general de la Factura

Public Property Let DtoGnral(ByVal vData As Currency)
     mDtoGnral = vData
End Property

Public Property Get DtoGnral() As Currency
     DtoGnral = mDtoGnral
End Property


'**** Importe Bruto Factura

Public Property Let BrutoFac(ByVal vData As Currency)
     mBrutofac = vData
End Property

Public Property Get BrutoFac() As Currency
     BrutoFac = mBrutofac
End Property


'**** Importe pronto pago de la Factura

Public Property Let ImpPPago(ByVal vData As Currency)
     vData = vData
     mImpPPago = vData
End Property

Public Property Get ImpPPago() As Currency
     ImpPPago = mImpPPago
End Property


'**** Importe general de la Factura

Public Property Let ImpGnral(ByVal vData As Currency)
     mImpGnral = vData
End Property

Public Property Get ImpGnral() As Currency
     ImpGnral = mImpGnral
End Property


'**** Importe Base Imponible Factura

Public Property Let BaseImp(ByVal vData As Currency)
     mBaseImp = vData
End Property

Public Property Get BaseImp() As Currency
     BaseImp = mBaseImp
End Property


'**** Base IVA1 de la Factura

Public Property Let BaseIVA1(ByVal vData As Currency)
     mBaseiva1 = vData
End Property

Public Property Get BaseIVA1() As Currency
     BaseIVA1 = mBaseiva1
End Property


'**** Base IVA2 de la Factura

Public Property Let BaseIVA2(ByVal vData As Currency)
     mBaseiva2 = vData
End Property

Public Property Get BaseIVA2() As Currency
     BaseIVA2 = mBaseiva2
End Property


'**** Base IVA3 de la Factura

Public Property Let BaseIVA3(ByVal vData As Currency)
     mBaseiva3 = vData
End Property

Public Property Get BaseIVA3() As Currency
     BaseIVA3 = mBaseiva3
End Property


'**** Tipo iva 1 de la Factura

Public Property Let TipoIVA1(ByVal vData As Byte)
     mTipoiva1 = vData
End Property

Public Property Get TipoIVA1() As Byte
     TipoIVA1 = mTipoiva1
End Property


'**** Tipo iva 2 de la Factura

Public Property Let TipoIVA2(ByVal vData As Byte)
     mTipoiva2 = vData
End Property

Public Property Get TipoIVA2() As Byte
     TipoIVA2 = mTipoiva2
End Property


'**** Tipo iva 3 de la Factura

Public Property Let TipoIVA3(ByVal vData As Byte)
     mTipoiva3 = vData
End Property

Public Property Get TipoIVA3() As Byte
     TipoIVA3 = mTipoiva3
End Property


'**** % de iva 1 de la Factura

Public Property Let PorceIVA1(ByVal vData As Currency)
     mPorciva1 = vData
End Property

Public Property Get PorceIVA1() As Currency
     PorceIVA1 = mPorciva1
End Property


'**** % de iva 2 de la Factura

Public Property Let PorceIVA2(ByVal vData As Currency)
     mPorciva2 = vData
End Property

Public Property Get PorceIVA2() As Currency
     PorceIVA2 = mPorciva2
End Property


'**** % de iva 3 de la Factura

Public Property Let PorceIVA3(ByVal vData As Currency)
     mPorciva3 = vData
End Property

Public Property Get PorceIVA3() As Currency
     PorceIVA3 = mPorciva3
End Property


'**** Importe de IVA 1 de la Factura

Public Property Let ImpIVA1(ByVal vData As Currency)
     mImpoiva1 = vData
End Property

Public Property Get ImpIVA1() As Currency
     ImpIVA1 = mImpoiva1
End Property


'**** Importe de IVA 2 de la Factura

Public Property Let ImpIVA2(ByVal vData As Currency)
     mImpoiva2 = vData
End Property

Public Property Get ImpIVA2() As Currency
     ImpIVA2 = mImpoiva2
End Property


'**** Importe de IVA 3 de la Factura

Public Property Let ImpIVA3(ByVal vData As Currency)
     mImpoiva3 = vData
End Property

Public Property Get ImpIVA3() As Currency
     ImpIVA3 = mImpoiva3
End Property


'**** Suplidos

Public Property Let Suplidos(ByVal vData As Currency)
     mSuplidos = vData
End Property

Public Property Get Suplidos() As Currency
     Suplidos = mSuplidos
End Property


'*****************************************************++
'   Recargo de equivalencia

'**** % de iva 1 de la Factura

Public Property Let PorceIVA1RE(ByVal vData As Currency)
     mPorciva1RE = vData
End Property

Public Property Get PorceIVA1RE() As Currency
     PorceIVA1RE = mPorciva1RE
End Property


'**** % de iva 2 de la Factura

Public Property Let PorceIVA2RE(ByVal vData As Currency)
     mPorciva2RE = vData
End Property

Public Property Get PorceIVA2RE() As Currency
     PorceIVA2RE = mPorciva2RE
End Property


'**** % de iva 3 de la Factura

Public Property Let PorceIVA3RE(ByVal vData As Currency)
     mPorciva3RE = vData
End Property

Public Property Get PorceIVA3RE() As Currency
     PorceIVA3RE = mPorciva3RE
End Property


'**** Importe de IVA 1 de la Factura

Public Property Let ImpIVA1RE(ByVal vData As Currency)
     mImpoiva1RE = vData
End Property

Public Property Get ImpIVA1RE() As Currency
     ImpIVA1RE = mImpoiva1RE
End Property


'**** Importe de IVA 2 de la Factura

Public Property Let ImpIVA2RE(ByVal vData As Currency)
     mImpoiva2RE = vData
End Property

Public Property Get ImpIVA2RE() As Currency
     ImpIVA2RE = mImpoiva2RE
End Property


'**** Importe de IVA 3 de la Factura

Public Property Let ImpIVA3RE(ByVal vData As Currency)
     mImpoiva3RE = vData
End Property

Public Property Get ImpIVA3RE() As Currency
     ImpIVA3RE = mImpoiva3RE
End Property

'**** Total Factura

Public Property Let TotalFac(ByVal vData As String)
     mTotalfac = vData
End Property

Public Property Get TotalFac() As String
     TotalFac = mTotalfac
End Property


'**** Nº terminal de la venta del TPV

Public Property Let NumTerminal(ByVal vData As Integer)
     mNumtermi = vData
End Property

Public Property Get NumTerminal() As Integer
     NumTerminal = mNumtermi
End Property



'**** Nº venta del TPV

Public Property Let NumVenta(ByVal vData As Long)
     mNumventa = vData
End Property

Public Property Get NumVenta() As Long
     NumVenta = mNumventa
End Property


'**** Para los MANTENIMIENTOS.  sera 1 la linea. Si esta vacio lo cargaremos con la amplicaicon normal

Public Property Let Observacion(ByVal vData As String)
     mObservacion = vData
End Property

Public Property Get Observacion() As String
     Observacion = mObservacion
End Property


'-----   Aportacion
Public Property Let Aportacion(ByVal vData As Currency)
     mAportacion = vData
End Property

Public Property Get Aportacion() As Currency
     Aportacion = mAportacion
End Property

'**** % de retencion

Public Property Let PorRet(ByVal vData As Currency)
     mPorRet = vData
End Property

Public Property Get PorRet() As Currency
     PorRet = mPorRet
End Property


'**** Importe de retencion

Public Property Let ImpReten(ByVal vData As Currency)
     mImpReten = vData
End Property

Public Property Get ImpReten() As Currency
     ImpReten = mImpReten
End Property







'**** TEXTOS CSB DE CARTERA DE COBROS

Public Property Let Text33csb(ByVal vData As String)
     mvarText33csb = vData
End Property

Public Property Get Text33csb() As String
     Text33csb = mvarText33csb
End Property

Public Property Let Text41csb(ByVal vData As String)
     mvarText41csb = vData
End Property

Public Property Get Text41csb() As String
     Text41csb = mvarText41csb
End Property

Public Property Let Text42csb(ByVal vData As String)
     mvarText42csb = vData
End Property

Public Property Get Text42csb() As String
     Text42csb = mvarText42csb
End Property

Public Property Let Text43csb(ByVal vData As String)
     mvarText43csb = vData
End Property

Public Property Get Text43csb() As String
     Text43csb = mvarText43csb
End Property

Public Property Let Text51csb(ByVal vData As String)
     mvarText51csb = vData
End Property

Public Property Get Text51csb() As String
     Text51csb = mvarText51csb
End Property

Public Property Let Text52csb(ByVal vData As String)
     mvarText52csb = vData
End Property

Public Property Get Text52csb() As String
     Text52csb = mvarText52csb
End Property

Public Property Let Text53csb(ByVal vData As String)
     mvarText53csb = vData
End Property

Public Property Get Text53csb() As String
     Text53csb = mvarText53csb
End Property


Public Property Let Text61csb(ByVal vData As String)
     mvarText61csb = vData
End Property

Public Property Get Text61csb() As String
     Text61csb = mvarText61csb
End Property

Public Property Let Text62csb(ByVal vData As String)
     mvarText62csb = vData
End Property

Public Property Get Text62csb() As String
     Text62csb = mvarText62csb
End Property

Public Property Let Text63csb(ByVal vData As String)
     mvarText63csb = vData
End Property

Public Property Get Text63csb() As String
     Text63csb = mvarText63csb
End Property

Public Property Let Text71csb(ByVal vData As String)
     mvarText71csb = vData
End Property

Public Property Get Text71csb() As String
     Text71csb = mvarText71csb
End Property

Public Property Let Text72csb(ByVal vData As String)
     mvarText72csb = vData
End Property

Public Property Get Text72csb() As String
     Text72csb = mvarText72csb
End Property

Public Property Let Text73csb(ByVal vData As String)
     mvarText73csb = vData
End Property

Public Property Get Text73csb() As String
     Text73csb = mvarText73csb
End Property

Public Property Let Text81csb(ByVal vData As String)
     mvarText81csb = vData
End Property

Public Property Get Text81csb() As String
     Text81csb = mvarText81csb
End Property

Public Property Let Text82csb(ByVal vData As String)
     mvarText82csb = vData
End Property

Public Property Get Text82csb() As String
     Text82csb = mvarText82csb
End Property

Public Property Let Text83csb(ByVal vData As String)
     mvarText83csb = vData
End Property

Public Property Get Text83csb() As String
     Text83csb = mvarText83csb
End Property



'------------------------------------------------
'Procedimientos del modulo CFactura
'------------------------------------------------

Public Function LeerDatos(vCodtipom As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla scafac
'Lee de la BD: Aritaxi, Tabla: scafac
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Mens As String
Dim Sql As String

    On Error GoTo ELeer
    
    
    LeerDatos = False
    Sql = "SELECT codtipom,numfactu,fecfactu,codclien,codforpa,codbanco,codsucur,digcontr,cuentaba,brutofac, totalfac, iban FROM scafac "
    Sql = Sql & " WHERE codtipom='" & vCodtipom & "'"
    Sql = Sql & " AND numfactu=" & vNumfactu
    Sql = Sql & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If Rs.EOF Then
        LeerDatos = False
    Else
        Mens = "Datos Factura"
        mCodTipoM = Rs!codtipom
        mNumFactu = CStr(Rs!NumFactu)
        mFecFactu = CStr(Rs!FecFactu)
        mCodClien = CStr(Rs!CodClien)
        ForPago = CStr(Rs!codforpa)
        TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N")
        mCodbanco = CStr(DBLet(Rs!codbanco, "N"))
        mCodsucur = CStr(DBLet(Rs!codsucur, "N"))
        mDigContr = DBLet(Rs!digcontr, "T")
        mCuentaba = DBLet(Rs!cuentaba, "T")
        mBrutofac = CStr(Rs!BrutoFac)
        mTotalfac = CStr(Rs!TotalFac)
        
        '[Monica]07/11/2014: Faltaba cargar la variable del IBAN
        mIban = DBLet(Rs!Iban, "T")

        mLetraSer = ObtenerLetraSerie(mCodTipoM)
        
        'si venimos de proceso de insertar tenemos la cuenta prevista de cobro que
        'se introdujo en el formulario de pasar albaranes a factura
        If CuentaPrev = "" Then
            'leer la cuenta prevista de cobros de la tabla de contabilidad: conta.scobro
            'cuando vamos a borrar una factura porque modificarmos
            'ya que este dato se pidio al facturar (SI EXISTE LA FACTURA EN TESORERIA)
            If vParamAplic.ContabilidadNueva Then
                Sql = "SELECT COUNT(*) FROM cobros WHERE numserie='" & LetraSerie & "' and numfactu=" & NumFactu
                Sql = Sql & " AND fecfactu=" & DBSet(FecFactu, "F")
                If RegistrosAListar(Sql, conConta) > 0 Then
                    Sql = DevuelveDesdeBDNew(conConta, "cobros", "ctabanc1", "numserie", mLetraSer, "T", , "numfactu", vNumfactu, "N", "fecfactu", vFecfactu, "F")
                    If Sql <> "" Then
                        mCtaPrev = Sql
                        LeerDatos = True
                    Else
                        LeerDatos = False
                        Mens = "La cuenta prevista de cobro para la factura no puede ser nula."
                    End If
                Else
                    LeerDatos = True
                End If
            
            Else
                Sql = "SELECT COUNT(*) FROM scobro WHERE numserie='" & LetraSerie & "' and codfaccl=" & NumFactu
                Sql = Sql & " AND fecfaccl=" & DBSet(FecFactu, "F")
                If RegistrosAListar(Sql, conConta) > 0 Then
                    Sql = DevuelveDesdeBDNew(conConta, "scobro", "ctabanc1", "numserie", mLetraSer, "T", , "codfaccl", vNumfactu, "N", "fecfaccl", vFecfactu, "F")
                    If Sql <> "" Then
                        mCtaPrev = Sql
                        LeerDatos = True
                    Else
                        LeerDatos = False
                        Mens = "La cuenta prevista de cobro para la factura no puede ser nula."
                    End If
                Else
                    LeerDatos = True
                End If
             End If
        Else
            LeerDatos = True
        End If
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
    LeerDatos = False
End Function


Public Function LeerDatosFACcli(vCodtipom As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla scafac
'Lee de la BD: Aritaxi, Tabla: scafac
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Mens As String
Dim Sql As String

    On Error GoTo ELeer
    
    
    LeerDatosFACcli = False
    Sql = "SELECT codtipom,numfactu,fecfactu,codclien,codforpa,codbanco,codsucur,digcontr,cuentaba,brutofac, totalfac, iban, porcret, impreten FROM scafaccli "
    Sql = Sql & " WHERE codtipom='" & vCodtipom & "'"
    Sql = Sql & " AND numfactu=" & vNumfactu
    Sql = Sql & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If Rs.EOF Then
        LeerDatosFACcli = False
    Else
        Mens = "Datos Factura"
        mCodTipoM = Rs!codtipom
        mNumFactu = CStr(Rs!NumFactu)
        mFecFactu = CStr(Rs!FecFactu)
        mCodClien = CStr(Rs!CodClien)
        ForPago = CStr(Rs!codforpa)
        TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N")
        mCodbanco = CStr(DBLet(Rs!codbanco, "N"))
        mCodsucur = CStr(DBLet(Rs!codsucur, "N"))
        mDigContr = DBLet(Rs!digcontr, "T")
        mCuentaba = DBLet(Rs!cuentaba, "T")
        mBrutofac = CStr(Rs!BrutoFac)
        mIban = CStr(DBLet(Rs!Iban, "T"))
        mTotalfac = CStr(Rs!TotalFac)
        
        mPorRet = DBLet(Rs!PorcRet, "N")
        mImpReten = DBLet(Rs!ImpReten, "N")

        mLetraSer = ObtenerLetraSerie(mCodTipoM)
        
        'si venimos de proceso de insertar tenemos la cuenta prevista de cobro que
        'se introdujo en el formulario de pasar albaranes a factura
        If CuentaPrev = "" Then
            'leer la cuenta prevista de cobros de la tabla de contabilidad: conta.scobro
            'cuando vamos a borrar una factura porque modificarmos
            'ya que este dato se pidio al facturar (SI EXISTE LA FACTURA EN TESORERIA)
            If vParamAplic.ContabilidadNueva Then
                Sql = "SELECT COUNT(*) FROM cobros WHERE numserie='" & LetraSerie & "' and numfactu=" & NumFactu
                Sql = Sql & " AND fecfactu=" & DBSet(FecFactu, "F")
                If RegistrosAListar(Sql, conConta) > 0 Then
                    Sql = DevuelveDesdeBDNew(conConta, "cobros", "ctabanc1", "numserie", mLetraSer, "T", , "numfactu", vNumfactu, "N", "fecfactu", vFecfactu, "F")
                    If Sql <> "" Then
                        mCtaPrev = Sql
                        LeerDatosFACcli = True
                    Else
                        LeerDatosFACcli = False
                        Mens = "La cuenta prevista de cobro para la factura no puede ser nula."
                    End If
                Else
                    LeerDatosFACcli = True
                End If
            
            Else
                Sql = "SELECT COUNT(*) FROM scobro WHERE numserie='" & LetraSerie & "' and codfaccl=" & NumFactu
                Sql = Sql & " AND fecfaccl=" & DBSet(FecFactu, "F")
                If RegistrosAListar(Sql, conConta) > 0 Then
                    Sql = DevuelveDesdeBDNew(conConta, "scobro", "ctabanc1", "numserie", mLetraSer, "T", , "codfaccl", vNumfactu, "N", "fecfaccl", vFecfactu, "F")
                    If Sql <> "" Then
                        mCtaPrev = Sql
                        LeerDatosFACcli = True
                    Else
                        LeerDatosFACcli = False
                        Mens = "La cuenta prevista de cobro para la factura no puede ser nula."
                    End If
                Else
                    LeerDatosFACcli = True
                End If
            End If
        Else
            LeerDatosFACcli = True
        End If
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
    LeerDatosFACcli = False
End Function






'======================================================================
'GRABAR EN TESORERIA
'======================================================================

Public Function InsertarEnTesoreria(vTextosCSB As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: aritaxi.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String, codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAuxConta As String 'para insertar en conta.scobro
Dim CadValues3 As String
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single 'importe para insertar en la svenci
Dim ImpVenci2 As Single 'importe para insertar en conta.scobro
Dim I As Byte
Dim TotalFactura2 As Currency   'Por si acaso lleva aportacion al terminal
'1 Julio 2009. Los graba en scobro
Dim CadenaDatosFiscales As String

Dim CADENA As String
Dim vvIban As String

    On Error GoTo EInsertarTesoreria

    b = False
    InsertarEnTesoreria = False


    If vParamAplic.ContabilidadNueva Then
        If vTextosCSB = "" Then
            vTextosCSB = "NULL"
        Else
            
            CadValues = RecuperaValor(vTextosCSB, 1)
            CadValues2 = RecuperaValor(vTextosCSB, 2)
            CadValues3 = RecuperaValor(vTextosCSB, 3)
            vTextosCSB = DBSet(CadValues & CadValues2 & CadValues3, "T", "S")
        End If
    
    Else
        If vTextosCSB = "" Then
            vTextosCSB = "NULL,NULL,NULL"
        Else
            
            CadValues = DBSet(RecuperaValor(vTextosCSB, 1), "T", "S")
            CadValues2 = DBSet(RecuperaValor(vTextosCSB, 2), "T", "S")
            CadValues3 = DBSet(RecuperaValor(vTextosCSB, 3), "T", "S")
            vTextosCSB = CadValues & "," & CadValues2 & "," & CadValues3
        End If
    End If
    CadValues3 = ""
    CadValues = ""
    CadValues2 = ""

    'campo para insertar en conta.scobro de Tesoreria
    textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd mmm yyyy") & "'"


    'Datos fiscales en scobro     Julio 2009
    'nomclien,domclien,pobclien, cpclien,proclien
    CadenaDatosFiscales = DBSet(mNomClien, "T") & "," & DBSet(mDomClien, "T") & "," & DBSet(mPobClien, "T")
    CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(mCodpobla, "T") & "," & DBSet(mProClien, "T")
    
    If vParamAplic.ContabilidadNueva Then
        CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(mNIFClien, "T") & ",'ES'"
    End If
    
    'Obtener el Nº de Vencimientos de la forma de pago
    Sql = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 And CCur(TotalFac) <> 0 Then
        
            'Comporbamos si el importe es <>0
        
            'Obtener los dias de pago del cliente : el socio no tiene dias de pago
            Sql = " SELECT  0 diapago1, 0 diapago2, 0 diapago3, 0 mesnogir, 0 diavtoat, '' codmacta, impacuenta as anticipo "
            Sql = Sql & " FROM sclien "
            Sql = Sql & " WHERE codclien=" & Cliente
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
'            codmacta = DBSet(RS!codmacta, "T")

            CADENA = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
            
            
            '[Monica] 31/05/2011: solo en el caso de una factura de venta de equipos se descontará el importe de anticipo
            If codtipom = "FAV" Then
                TotalFac = TotalFac - DBLet(Rs!anticipo, "N")
            End If
            
            '[Monica] 30/05/2011: Añadida la condicion del tipo de movimiento
            If codtipom = "FCN" Or codtipom = "FCE" Or codtipom = "FRC" Then
                codmacta = vParamAplic.Raiz_CtaClien_Soc & Format(Cliente, CADENA)
            Else
                codmacta = vParamAplic.Raiz_Cta_Soc_Equip & Format(Cliente, CADENA)
            End If

'            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
            
            If Not Rs.EOF Then
                cadValuesAux = "('" & codtipom & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                CadValuesAuxConta = "('" & LetraSerie & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                '                    Añadire a la cadena fija esta los valores de textcsb41,txcs
                CadValuesAuxConta = CadValuesAuxConta & vTextosCSB & ","
                '-------- Primer Vencimiento
                I = 1
                'FECHA VTO
                FecVenci = CDate(FecFactu)
                '=== Laura 23/01/2007
                'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                '===
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                If TipForPago <> 0 Then
                    FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                Else
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                End If
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(Rs!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAuxConta & I & ", "
                CadValues2 = CadValues2 & codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                TotalFactura2 = TotalFac - Aportacion
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFactura2
                    ImpVenci2 = TotalFactura2 - ImpCheque
                Else
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura2 - ImpCheque) / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFactura2 Then
                        ImpVenci = Round(ImpVenci + (TotalFactura2 - ImpVenci * rsVenci!numerove), 2)
                    End If
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If (ImpVenci2 * rsVenci!numerove) + ImpCheque <> TotalFactura2 Then
                        ImpVenci2 = Round(ImpVenci2 + (TotalFactura2 - ImpCheque - (ImpVenci2 * rsVenci!numerove)), 2)
                    End If
                End If
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                If vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', " & textcsb33 & ", " & DBSet(Agente, "N")
                Else
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                End If
                'departamento y transfer
                CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                '1 Julio 2009
                ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                 CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                 
                If vParamAplic.ContabilidadNueva Then
                    vvIban = MiFormat(Iban, "") & MiFormat(Banco, "0000") & MiFormat(Sucursal, "0000") & MiFormat(DigControl, "00") & MiFormat(CuentaBan, "0000000000")
                
                    CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                Else
                    '[Monica]22/11/2013: tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                       CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                    Else
                       CadValues2 = CadValues2 & ")"
                    End If
                End If
                
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For I = 2 To rsVenci!numerove
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(Rs!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    End If
                        
                    CadValues = CadValues & ", " & cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & ", " & codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura2 - ImpCheque) / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    End If
                    CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                    '1 Julio 2009
                    ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                    CadValues2 = CadValues2 & "," & CadenaDatosFiscales ' & ")"
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(Iban, "") & MiFormat(Banco, "0000") & MiFormat(Sucursal, "0000") & MiFormat(DigControl, "00") & MiFormat(CuentaBan, "0000000000")
                    
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                    
                    Else
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                        Else
                           CadValues2 = CadValues2 & ")"
                        End If
                    End If
                Next I
                
                '--- Cheque regalo: laura 1/12/2006   y/o    Aportacion terminal
                'si hay cheque regalo insertar una linea más para la forma de pago correspondiente y el importe del cheque
                If ImpCheque > 0 Then
                
                    
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & "," & DBSet(codmacta, "T") & ", " & vParamAplic.ForPagoChequeRegalo & ", "
                    
                    'FECHA VTO
                    CadValuesAuxConta = "primerve"
                    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", vParamAplic.ForPagoChequeRegalo, "N", CadValuesAuxConta)
                    FecVenci = CDate(FecFactu)
                    FecVenci = FecVenci + CByte(CadValuesAuxConta)
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                        MsgBox "FALTA cheque regalo con forma de pago no en EFECTIVO", vbInformation
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    
                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(ImpCheque, "N") & ", '" & CuentaPrev & "', " & textcsb33 & ", " & DBSet(Agente, "N")
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpCheque, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                    End If
                    'departamento
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL" ')"
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(Iban, "") & MiFormat(Banco, "0000") & MiFormat(Sucursal, "0000") & MiFormat(DigControl, "00") & MiFormat(CuentaBan, "0000000000")
                    
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                    Else
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                        Else
                           CadValues2 = CadValues2 & ")"
                        End If
                    End If
                End If
                
                
                'Aportacion al terminal
                'Si tiene cuenta aportacion entonces añadiremos un cobro
                If Me.Aportacion > 0 Then
                    If vParamAplic.ctaAportacion = "" Then
                        MsgBox "Error cta aportacion NULL", vbExclamation
                        Exit Function
                    End If
                        
                   'Montamos EL SQL para el cobro de la aportacion al termina
                    'ForPago
                    
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & ",'" & vParamAplic.ctaAportacion & "', " & ForPago & ", "
                    
                    
                    CadValuesAuxConta = "primerve"
                    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N", CadValuesAuxConta)
                    FecVenci = CDate(FecFactu)
                    FecVenci = FecVenci + CByte(CadValuesAuxConta)
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                        MsgBox "FALTA config con forma de pago no en EFECTIVO", vbInformation
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    
                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(Aportacion, "N") & ", '" & CuentaPrev & "', " & textcsb33 & ", " & DBSet(Agente, "N")
                    Else
                        CadValues2 = CadValues2 & DBSet(Aportacion, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                    End If
                    'departamento
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL)"
                
                End If
                
            End If
            Rs.Close
            
        Else
            'totalfac =0 and numerovtos >=1
            b = True
        End If
        
        Set Rs = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        Sql = "INSERT INTO svenci (codtipom, numfactu, fecfactu, ordefect, fecefect, impefect)"
        Sql = Sql & " VALUES " & CadValues
        conn.Execute Sql
    End If
    
    
    'Grabar tabla scobro de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        '01/09/06
'        If (NumTicket = "") Or (NumTicket <> "" And TipForPago <> 0) Then
            If CuentaPrev <> "" Then
                'antes de grabar en la scobro comprobar que existe en conta.sforpa la
                'forma de pago de la factura. Sino existe insertarla
                'vemos si existe en la conta
                b = InsertarFormaPagoEnConta(ForPago, MenError)
                
                'si hay cheque regalo comprobar q existe su forma de pago en conta
                'sino insertarla
                If b And ImpCheque > 0 Then b = InsertarFormaPagoEnConta(vParamAplic.ForPagoChequeRegalo, MenError)
                
                
                
                If b Then
                    If vParamAplic.ContabilidadNueva Then
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu,text41csb , numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, "
                        Sql = Sql & "text33csb,agente,departamento,transfer "
                        'JULIO 2009
                        'nomclien,domclien,pobclien, cpclien,proclien
                        Sql = Sql & ",nomclien,domclien,pobclien, cpclien,proclien,nifclien,codpais,iban)" ')"
                        
                    Else
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl,text41csb ,text42csb, text43csb, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur,"
                        Sql = Sql & "digcontr, cuentaba,text33csb,agente,departamento,transfer "
                        'JULIO 2009
                        'nomclien,domclien,pobclien, cpclien,proclien
                        Sql = Sql & ",nomclien,domclien,pobclien, cpclien,proclien" ')"
                        
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           Sql = Sql & ",iban)"
                        Else
                           Sql = Sql & ")"
                        End If
                    End If
                    Sql = Sql & " VALUES " & CadValues2
                    ConnConta.Execute Sql
                End If
                
            Else
                'DAVID ####
                'ENERO 2008
                'Si no inserto en tesoreria por que ctaprvsta="" entonces dejo continuar
                b = True
        
        End If





    End If
    
'    If b Then b = True
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Insertar en Tesoreria: " & vbCrLf & Err.Description
    End If
    InsertarEnTesoreria = b
End Function


Public Function InsertarEnTesoreriaFACcli(vTextosCSB As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: aritaxi.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String, codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAuxConta As String 'para insertar en conta.scobro
Dim CadValues3 As String
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single 'importe para insertar en la svenci
Dim ImpVenci2 As Single 'importe para insertar en conta.scobro
Dim I As Byte
Dim TotalFactura2 As Currency   'Por si acaso lleva aportacion al terminal
'1 Julio 2009. Los graba en scobro
Dim CadenaDatosFiscales As String
Dim vvIban As String


    On Error GoTo EInsertarTesoreria

    If vParamAplic.ContabilidadNueva Then
        If vTextosCSB = "" Then
            vTextosCSB = "NULL"
        Else
            CadValues = RecuperaValor(vTextosCSB, 1)
            CadValues2 = RecuperaValor(vTextosCSB, 2)
            CadValues3 = RecuperaValor(vTextosCSB, 3)
            vTextosCSB = DBSet(CadValues & CadValues2 & CadValues3, "T", "S")
        End If
    Else
        If vTextosCSB = "" Then
            vTextosCSB = "NULL,NULL,NULL"
        Else
            
            CadValues = DBSet(RecuperaValor(vTextosCSB, 1), "T", "S")
            CadValues2 = DBSet(RecuperaValor(vTextosCSB, 2), "T", "S")
            CadValues3 = DBSet(RecuperaValor(vTextosCSB, 3), "T", "S")
            vTextosCSB = CadValues & "," & CadValues2 & "," & CadValues3
        End If
    End If
    b = False
    InsertarEnTesoreriaFACcli = False
    CadValues3 = ""
    CadValues = ""
    CadValues2 = ""

    'campo para insertar en conta.scobro de Tesoreria
    textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd mmm yyyy") & "'"


    'Datos fiscales en scobro     Julio 2009
    'nomclien,domclien,pobclien, cpclien,proclien
    CadenaDatosFiscales = DBSet(mNomClien, "T") & "," & DBSet(mDomClien, "T") & "," & DBSet(mPobClien, "T")
    CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(mCodpobla, "T") & "," & DBSet(mProClien, "T")
    
    If vParamAplic.ContabilidadNueva Then
        CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(mNIFClien, "T") & ",'ES'"
    End If
    
    
    'Obtener el Nº de Vencimientos de la forma de pago
    Sql = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 And CCur(TotalFac) <> 0 Then
        
            'Comporbamos si el importe es <>0
        
        
        
            'Obtener los dias de pago del cliente
            Sql = " SELECT  diapago1, diapago2, diapago3, mesnogir, diavtoat, codmacta "
            Sql = Sql & " FROM scliente "
            Sql = Sql & " WHERE codclien=" & Cliente
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            codmacta = DBSet(Rs!codmacta, "T")
'            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
            
            If Not Rs.EOF Then
                cadValuesAux = "('" & codtipom & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                CadValuesAuxConta = "('" & LetraSerie & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                '                    Añadire a la cadena fija esta los valores de textcsb41,txcs
                CadValuesAuxConta = CadValuesAuxConta & vTextosCSB & ","
                '-------- Primer Vencimiento
                I = 1
                'FECHA VTO
                FecVenci = CDate(FecFactu)
                '=== Laura 23/01/2007
                'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                '===
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                If TipForPago <> 0 Then
                    FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                Else
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                End If
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(Rs!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAuxConta & I & ", "
                CadValues2 = CadValues2 & codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                TotalFactura2 = TotalFac - Aportacion
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFactura2
                    ImpVenci2 = TotalFactura2 - ImpCheque
                Else
                    
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura2 - ImpCheque) / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFactura2 Then
                        ImpVenci = Round(ImpVenci + (TotalFactura2 - ImpVenci * rsVenci!numerove), 2)
                    End If
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If (ImpVenci2 * rsVenci!numerove) + ImpCheque <> TotalFactura2 Then
                        ImpVenci2 = Round(ImpVenci2 + (TotalFactura2 - ImpCheque - (ImpVenci2 * rsVenci!numerove)), 2)
                    End If
                End If
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                
                If vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', " & textcsb33 & ", " & DBSet(Agente, "N")
                Else
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                End If
                
                'departamento y transfer
                CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                '1 Julio 2009
                ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                 CadValues2 = CadValues2 & "," & CadenaDatosFiscales ' & ")"
                 
                If vParamAplic.ContabilidadNueva Then
                    vvIban = MiFormat(Iban, "") & MiFormat(Banco, "0000") & MiFormat(Sucursal, "0000") & MiFormat(DigControl, "00") & MiFormat(CuentaBan, "0000000000")
                
                    CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                Else
                     '[Monica]22/11/2013: tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                       CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                    Else
                       CadValues2 = CadValues2 & ")"
                    End If
                End If

                
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For I = 2 To rsVenci!numerove
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(Rs!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    End If
                        
                    CadValues = CadValues & ", " & cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & ", " & codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura2 - ImpCheque) / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    End If
                    
                    CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                    '1 Julio 2009
                    ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                    CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                    
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(Iban, "") & MiFormat(Banco, "0000") & MiFormat(Sucursal, "0000") & MiFormat(DigControl, "00") & MiFormat(CuentaBan, "0000000000")
                    
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                    Else
                         '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                        Else
                           CadValues2 = CadValues2 & ")"
                        End If
                    End If
                    
                Next I
                
                '--- Cheque regalo: laura 1/12/2006   y/o    Aportacion terminal
                'si hay cheque regalo insertar una linea más para la forma de pago correspondiente y el importe del cheque
                If ImpCheque > 0 Then
                
                    
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & "," & codmacta & ", " & vParamAplic.ForPagoChequeRegalo & ", "
                    
                    'FECHA VTO
                    CadValuesAuxConta = "primerve"
                    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", vParamAplic.ForPagoChequeRegalo, "N", CadValuesAuxConta)
                    FecVenci = CDate(FecFactu)
                    FecVenci = FecVenci + CByte(CadValuesAuxConta)
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                        MsgBox "FALTA cheque regalo con forma de pago no en EFECTIVO", vbInformation
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    
                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(ImpCheque, "N") & ", '" & CuentaPrev & "', " & textcsb33 & ", " & DBSet(Agente, "N")
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpCheque, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                    End If
                    'departamento
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL" ')"
                     
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(Iban, "") & MiFormat(Banco, "0000") & MiFormat(Sucursal, "0000") & MiFormat(DigControl, "00") & MiFormat(CuentaBan, "0000000000")
                    
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                    Else
                         '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                        Else
                           CadValues2 = CadValues2 & ")"
                        End If
                    End If
                End If
                
                
                'Aportacion al terminal
                'Si tiene cuenta aportacion entonces añadiremos un cobro
                If Me.Aportacion > 0 Then
                    If vParamAplic.ctaAportacion = "" Then
                        MsgBox "Error cta aportacion NULL", vbExclamation
                        Exit Function
                    End If
                        
                   'Montamos EL SQL para el cobro de la aportacion al termina
                    'ForPago
                    
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & ",'" & vParamAplic.ctaAportacion & "', " & ForPago & ", "
                    
                    
                    CadValuesAuxConta = "primerve"
                    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N", CadValuesAuxConta)
                    FecVenci = CDate(FecFactu)
                    FecVenci = FecVenci + CByte(CadValuesAuxConta)
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                        MsgBox "FALTA config con forma de pago no en EFECTIVO", vbInformation
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    
                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(Aportacion, "N") & ", '" & CuentaPrev & "', " & textcsb33 & ", " & DBSet(Agente, "N")
                    Else
                        CadValues2 = CadValues2 & DBSet(Aportacion, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                    End If
                    
                    'departamento
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL" ')"
                    
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(Iban, "") & MiFormat(Banco, "0000") & MiFormat(Sucursal, "0000") & MiFormat(DigControl, "00") & MiFormat(CuentaBan, "0000000000")
                    
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                    
                    Else
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                        Else
                           CadValues2 = CadValues2 & ")"
                        End If
                    End If
                End If
                
            End If
            Rs.Close
            
        Else
            'totalfac =0 and numerovtos >=1
            b = True
        End If
        
        Set Rs = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        Sql = "INSERT INTO svencicli (codtipom, numfactu, fecfactu, ordefect, fecefect, impefect)"
        Sql = Sql & " VALUES " & CadValues
        conn.Execute Sql
    End If
    
    
    'Grabar tabla scobro de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        '01/09/06
'        If (NumTicket = "") Or (NumTicket <> "" And TipForPago <> 0) Then
            If CuentaPrev <> "" Then
                'antes de grabar en la scobro comprobar que existe en conta.sforpa la
                'forma de pago de la factura. Sino existe insertarla
                'vemos si existe en la conta
                b = InsertarFormaPagoEnConta(ForPago, MenError)
                
                'si hay cheque regalo comprobar q existe su forma de pago en conta
                'sino insertarla
                If b And ImpCheque > 0 Then b = InsertarFormaPagoEnConta(vParamAplic.ForPagoChequeRegalo, MenError)
                
                
                
                If b Then
                
                    If vParamAplic.ContabilidadNueva Then
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, text41csb, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, "
                        Sql = Sql & "text33csb,agente,departamento,transfer "
                        'JULIO 2009
                        'nomclien,domclien,pobclien, cpclien,proclien
                        Sql = Sql & ",nomclien,domclien,pobclien, cpclien,proclien,nifclien,codpais,iban) " ')"
                        
                        Sql = Sql & " VALUES " & CadValues2
                    
                    Else
                
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl,text41csb ,text42csb, text43csb, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur,"
                        Sql = Sql & "digcontr, cuentaba,text33csb,agente,departamento,transfer "
                        'JULIO 2009
                        'nomclien,domclien,pobclien, cpclien,proclien
                        Sql = Sql & ",nomclien,domclien,pobclien, cpclien,proclien" ')"
                         '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           Sql = Sql & ",iban)"
                        Else
                           Sql = Sql & ")"
                        End If
                        
                        Sql = Sql & " VALUES " & CadValues2
                        
                    End If
                    ConnConta.Execute Sql
                End If
                
            Else
                'DAVID ####
                'ENERO 2008
                'Si no inserto en tesoreria por que ctaprvsta="" entonces dejo continuar
                b = True
        
            End If
    End If
    
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Insertar en Tesoreria: " & vbCrLf & Err.Description
    End If
    InsertarEnTesoreriaFACcli = b
End Function







'###### Laura 15/11/2006 Recuperar facturas ALZIRA
Public Function InsertarEnTesoreria_RecupFac(MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: aritaxi.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String, codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single
Dim I As Byte


    On Error GoTo EInsertarTesoreria

    b = False
    InsertarEnTesoreria_RecupFac = False
    CadValues = ""
    CadValues2 = ""

    'Obtener el Nº de Vencimientos de la forma de pago
    Sql = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago del cliente
            Sql = " SELECT  diapago1, diapago2, diapago3, mesnogir, diavtoat, codmacta "
            Sql = Sql & " FROM sclien "
            Sql = Sql & " WHERE codclien=" & Cliente
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            codmacta = DBSet(Rs!codmacta, "T")
         
            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
            
            If Not Rs.EOF Then
               cadValuesAux = "('" & codtipom & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
               CadValuesAux2 = "('" & LetraSerie & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "

              'Primer Vencimiento
              '------------------------------------------------------------
                I = 1
                'FECHA VTO
                FecVenci = CDate(FecFactu)
                FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                If TipForPago <> 0 Then
                    FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                Else
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                End If
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(Rs!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAux2 & I & ", "
                CadValues2 = CadValues2 & codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFac
                Else
                    ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFac Then
'                        If ImpVenci * rsVenci!numerove <> TotalFac2 Then
                            ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci!numerove), 2)
'                        End If
                    End If
                End If
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                'departamento
                CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ")"
                
                
              'Resto Vencimientos
              '--------------------------------------------------------------------
                For I = 2 To rsVenci!numerove
                  'FECHA Resto Vencimientos
                    FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(Rs!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    End If
                        
                    CadValues = CadValues & ", " & cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAux2 & I & ", " & codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                  'IMPORTE Resto de Vendimientos
                    ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") & ")"
                Next I
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        Sql = "INSERT INTO svenci (codtipom, numfactu, fecfactu, ordefect, fecefect, impefect)"
        Sql = Sql & " VALUES " & CadValues
        conn.Execute Sql
    End If
    
    

    b = True
    
EInsertarTesoreria:
    If Err.Number <> 0 Then b = False
    InsertarEnTesoreria_RecupFac = b
End Function
















Private Sub ReiniciarImportesFactura()
    Me.BaseImp = 0
    Me.BaseIVA1 = 0
    Me.BaseIVA2 = 0
    Me.BaseIVA3 = 0
    
    Me.BrutoFac = 0
    
    Me.ImpIVA1 = 0
    Me.ImpIVA2 = 0
    Me.ImpIVA3 = 0
    
    Me.PorceIVA1 = 0
    Me.PorceIVA2 = 0
    Me.PorceIVA3 = 0
    
    
    Me.ImpIVA1RE = 0
    Me.ImpIVA2RE = 0
    Me.ImpIVA3RE = 0
    
    Me.PorceIVA1RE = 0
    Me.PorceIVA2RE = 0
    Me.PorceIVA3RE = 0
    
    
    Me.TipoIVA1 = 0
    Me.TipoIVA2 = 0
    Me.TipoIVA3 = 0
    
    Me.ImpGnral = 0
    Me.ImpPPago = 0
    Me.TotalFac = 0
    

End Sub


Public Function CalcularDatosFactura(Cliente As Boolean, cadWHERE As String, NomTabla As String, NomTablaLin As String, Optional EsAnterior As Boolean) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim I As Integer

Dim Sql As String
Dim cadAux As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotImpIVA As Currency

Dim Impaux As Currency
Dim ImpIVA As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vBruto As Currency

Dim exento_IVA As Byte   ' 0.-IVA normal      1.- Recargo equivalencia    2.- Exento  3.-Exento INTRACOM
Dim conDesplaz As Boolean
    
    
Dim TipoIVAAplicable As Integer
Dim RE1 As String
Dim impRE As Currency
Dim TotRe As Currency


Dim ImporteParaComparar As Currency
Dim NueroIvas As Integer


    CalcularDatosFactura = False
    On Error GoTo ECalcular

    If codtipom <> "FAM" Then ReiniciarImportesFactura

   '------------------------------------------------------------------------
   'En estos dos procedimientos solo entra cuando va a generar la factura y
   'codTipoMov="FAV", cuando llamamos desde pedidos, albaranes,etc, para mostrar
   'el resumen de la factura codTipoMov="ALV","PEV",... y no entra
   '-------------------------------------------------------------------------
   
   'comprobar si hay que facturar DESPLAZAMIENTOS
   CalcularDatosFactura = CalcularDesplazamiento(cadWHERE, NomTabla)
   If Not CalcularDatosFactura Then Exit Function
   
   'comprobar si hay que facturar BONIFICACIONES
    CalcularDatosFactura = CalcularBonificacion(cadWHERE, NomTablaLin)
   If Not CalcularDatosFactura Then Exit Function
   '-------------------------------------------------------------------------
   
   'comprobar si el cliente de la factura esta EXENTO de IVA
   
   If codtipom = "ALZ" Or codtipom = "FAZ" Then
        'Albaranes de B. No llevan IVA
        exento_IVA = 2   'Exento IVA
   ElseIf Cliente Then
        exento_IVA = ClienteExentoIVA2
    Else
        'exento_IVA = vParamAplic.IVA_REA
        exento_IVA = 0
   End If
   
  
   If exento_IVA = 2 Or exento_IVA = 3 Then
        cadWHERE = Replace(cadWHERE, NomTabla, NomTablaLin)
        Sql = "SELECT sum(importel) as bruto"
        Sql = Sql & " FROM " & NomTablaLin
        Sql = Sql & " WHERE " & cadWHERE
   Else
    'Agrupar el importe bruto por tipos de iva
       If codtipom = "FAM" Then
            'para los mantenimientos obtenemos de la tabla de tipos de contrato(stipco.codartic)
            'el articulo que se va a utilizar para facturar los mantenimientos
    '        CadAux = DevuelveDesdeBDNew(conAri, "stipco", "codartic", "codtipco", TipCoMan, "T")
            If ArticFac <> "" Then 'tenemos el articulo
                Sql = "SELECT codartic,nomartic,sartic.codigiva, tiposiva.porceiva "
                Sql = Sql & " FROM sartic INNER JOIN conta" & vParamAplic.NumeroConta & ".tiposiva ON sartic.codigiva=conta" & vParamAplic.NumeroConta & ".tiposiva.codigiva"
                Sql = Sql & " WHERE codartic=" & DBSet(ArticFac, "T")
            End If
       Else
            cadWHERE = Replace(cadWHERE, NomTabla, NomTablaLin)
            Sql = "SELECT sartic.codigiva,sum(importel) as bruto"
            Sql = Sql & " FROM " & NomTablaLin & " LEFT OUTER JOIN sartic ON " & NomTablaLin & ".codartic=sartic.codartic"
            Sql = Sql & " WHERE " & cadWHERE
            Sql = Sql & " GROUP BY codigiva "
            Sql = Sql & " ORDER BY codigiva "
        End If
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenKeyset, adLockOptimistic, adCmdText
    NueroIvas = 0 'Numero IVAS
    ImporteParaComparar = 0
    
    'VEre cuantos IVAS tiene y el bruto. Para cuadrar los decimales
    If codtipom = "FAM" Then
        'No hago la comrobacion para cuadre decimales
        NueroIvas = 1
    Else
            
        While Not Rs.EOF
            TotBruto = DBLet(Rs!bruto, "N")
            ImporteParaComparar = ImporteParaComparar + TotBruto
            NueroIvas = NueroIvas + 1
            Rs.MoveNext
        Wend
    
    End If
    If NueroIvas > 0 Then
        TotImpIVA = CalcularPorcentaje(ImporteParaComparar, DtoPPago, 2)
        TotBruto = CalcularPorcentaje(ImporteParaComparar, DtoGnral, 2)
        ImporteParaComparar = ImporteParaComparar - TotImpIVA - TotBruto '(bruto - DtoPP -dtogral

        Rs.MoveFirst
    End If
    TotBruto = 0
    TotImpIVA = 0
    TotRe = 0
    I = 1
    conDesplaz = False

    While Not Rs.EOF
        'Aqui vamos acumulando la suma del importe bruto de las lineas
        If codtipom = "FAM" Then
            vBruto = BrutoFac
        Else
            vBruto = Rs!bruto
        End If
        
        TotBruto = TotBruto + vBruto
        ImpBImIVA = vBruto
        
        If exento_IVA < 2 Then     'Con IVA normal o RE
            'comprobar que si hay km de DESPLAZAMIENTO se añadan al tipo de IVA correspondiente
            If TotalKm > 0 Then
                If despIVA = Rs!codigiva Then
                    vBruto = vBruto + ImporteL
                    TotBruto = TotBruto + ImporteL
                    conDesplaz = True
                End If
            End If
        
            'comprobar si hay BONIFICACIONES que aplicar a la Factura
            If IVABonif1 = Rs!codigiva Then
                vBruto = vBruto - TotBonif1
                TotBruto = TotBruto - TotBonif1
            ElseIf IVABonif2 = Rs!codigiva Then
                vBruto = vBruto - TotBonif2
                TotBruto = TotBruto - TotBonif2
                
            ElseIf IVABonif3 = Rs!codigiva Then
                vBruto = vBruto - TotBonif3
                TotBruto = TotBruto - TotBonif3
            End If
            
        
            'Aplicarle el dto ppago
            '---- Laura: 05/10/2006
            Impaux = CalcularPorcentaje(vBruto, DtoPPago, 2)
'            ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoPPago)))
            '---- Laura: 27/09/2006
'            ImpAux = Round(ImpAux, 2)
            '----
            ImpBImIVA = vBruto - Impaux '(bruto - DtoPP)
            
            'Aplicarle el dto grnal
            '---- Laura: 05/10/2006
            Impaux = CalcularPorcentaje(vBruto, DtoGnral, 2)
            ImpBImIVA = ImpBImIVA - Impaux '(bruto - Dtogn)
            
            
            'NUEVO. 22 Mayo 2008
            'Si tiene mas de un tipo de iva, compruebo en la ultima de las bases
            'Si el total sera igual al total a comparar
            If I > 1 And I = NueroIvas Then
                'ImporteParaComparar
                If ImporteParaComparar - ImpBImIVA <> 0 Then
                    ImpBImIVA = ImporteParaComparar
                End If
            Else
                ImporteParaComparar = ImporteParaComparar - ImpBImIVA  'Para comparar en la ultima (si ha lugar)
            End If
            
            'Obtener el % de IVA
            cadAux = Rs!codigiva    'Por si no tiene  RE o si teniendo RE no tiene el correspondiente
            If exento_IVA = 1 Then cadAux = CStr(vParamAplic.DevuleveTipoIVA_RE(Rs!codigiva))
            TipoIVAAplicable = CInt(cadAux)
            
            '[Monica]
            'De David
            '
            If EsAnterior And exento_IVA < 2 Then
                    'JUNIO 2010
                    'Debido al cambio de IVA, los albaranes rectificativos podran ir en el IVA que llevaran en su momento
                    'Es decir, que aunque ahora sea al 18(a partir de Julio), para un albr rectifi podria ser
                    'con iva 16
                    cadAux = CStr(vParamAplic.DevuleveTipoIVA_Ant(exento_IVA = 1, Rs!codigiva))
                    TipoIVAAplicable = CInt(cadAux)
            End If
            
            
            'ANTES
            'cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(RS!Codigiva), "N")
            RE1 = "porcerec"
            cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", cadAux, "N", RE1)
            If exento_IVA <> 1 Then
                RE1 = "0"
                impRE = 0
            Else
                impRE = CalcularPorcentaje(ImpBImIVA, CCur(RE1), 2)
                TotRe = TotRe + impRE
            End If
            

            
            
            'aplicar el IVA a la base imponible de ese tipo
            '---- Laura: 05/10/2006
            ImpIVA = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
            
            'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
            'los vamos acumulando
            TotImpIVA = TotImpIVA + ImpIVA
            
            
            Select Case I
                Case 1  'IVA 1
                    TipoIVA1 = TipoIVAAplicable    'Antes Rs!codigiva
                    
                    BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA1 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA1 = ImpIVA
                    
                    'Recargo equivalencia
                    PorceIVA1RE = RE1
                    ImpIVA1RE = impRE
                     
                Case 2  'IVA 2
                    TipoIVA2 = TipoIVAAplicable     'Antes Rs!codigiva
                    
                    BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA2 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA2 = ImpIVA
    
                    'Recargo equivalencia
                    PorceIVA2RE = RE1
                    ImpIVA2RE = impRE
    
    
                Case 3  'IVA 3
                    TipoIVA3 = TipoIVAAplicable     'Antes Rs!codigiva
                    
                    BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA3 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA3 = ImpIVA
                    
                    'Recargo equivalencia
                    PorceIVA3RE = RE1
                    ImpIVA3RE = impRE
                    
            End Select
            I = I + 1
        Else 'EXENTO DE IVA
            'si hay que factura KM añadirlos al bruto
            If TotalKm > 0 Then vBruto = vBruto + ImporteL
            
            If exento_IVA = 2 Then
                'Exento normal
                If vParamAplic.IVA_Exento2 = 0 Then MsgBox "Falta configurar el IVA exento", vbExclamation
                TipoIVA1 = vParamAplic.IVA_Exento2
            Else
                'exento intracom
                If vParamAplic.IVA_Intracomunitario = 0 Then MsgBox "Falta configurar el IVA intracomunitario", vbExclamation
                TipoIVA1 = vParamAplic.IVA_Intracomunitario
                
            End If
        
                
                
            'DAVID.  20 Julio 2009
            Impaux = CalcularPorcentaje(vBruto, DtoPPago, 2)
            ImpIVA1 = Impaux
            Impaux = CalcularPorcentaje(vBruto, DtoGnral, 2)
            ImpIVA1 = ImpIVA1 + Impaux
                
            vBruto = vBruto - ImpIVA1
            ImpIVA1 = 0
                
            
                
                BaseIVA1 = vBruto 'BASE IMPONIBLE
                
                PorceIVA1 = 0 '% de IVA
                
                'Importe total con IVA
                ImpIVA1 = 0
                
                'Recargo equivalencia
                PorceIVA1RE = 0
                ImpIVA1RE = 0

        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    'si tiene desplazamiento comprobar que se ha añadido en el IVA
    'si no calcularlo
    If TotalKm > 0 And conDesplaz = False Then
    
        'FALTA##### con la conversion del Tipo de iva a RE
    
        TotBruto = TotBruto + ImporteL
        vBruto = ImporteL
        ImpBImIVA = ImporteL
        
        'Aplicarle el dto ppago
        '---- Laura: 05/10/2006
        Impaux = CalcularPorcentaje(vBruto, DtoPPago, 2)
'        ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoPPago)))
        ImpBImIVA = ImpBImIVA - Impaux '(bruto - DtoPP)
            
        'Aplicarle el dto grnal
        '---- Laura: 05/10/2006
        Impaux = CalcularPorcentaje(vBruto, DtoGnral, 2)
'        ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoGnral)))
        ImpBImIVA = ImpBImIVA - Impaux '(bruto - Dtogn)
            
        'Obtener el % de IVA
        cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(despIVA), "N")
            
        'cuota IVA: aplicar el IVA a la base imponible de ese tipo
        'Laura: 05/10/2006
        Impaux = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
'        ImpAux = CalcularDto(CStr(ImpBImIVA), cadAux)
        ImpIVA = Impaux
            
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotImpIVA = TotImpIVA + ImpIVA
        
        Select Case I
            Case 2 'IVA 2
                TipoIVA2 = despIVA
                
                BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA2 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA2 = ImpIVA
                
            Case 3 'IVA 3
                TipoIVA3 = Rs!codigiva
                
                BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA3 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA3 = ImpIVA
        End Select
    End If
    
    BrutoFac = TotBruto
    
    'Aplicarle el dto ppago
    '---- Laura: 05/10/2006
    ImpPPago = CalcularPorcentaje(TotBruto, DtoPPago, 2)
'    ImpPPago = CCur(CalcularDto(CStr(TotBruto), CStr(DtoPPago)))
    '---- Laura: 27/09/2006
'    ImpPPago = Round(ImpPPago, 2)
    '----
    
    'Aplicarle el dto general
    '---- Laura: 05/10/2006
    ImpGnral = CalcularPorcentaje(TotBruto, DtoGnral, 2)
'    ImpGnral = CCur(CalcularDto(CStr(TotBruto), CStr(DtoGnral)))
    '---- Laura: 27/09/2006
'    ImpGnral = Round(ImpGnral, 2)
    '----
    
    
    'Base Imponible
    BaseImp = TotBruto - ImpPPago - ImpGnral
    
    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA + TotRe
    
    CalcularDatosFactura = True
    
ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFactura = False
    Else
        CalcularDatosFactura = True
    End If
End Function



Public Function PasarAlbaranesAFactura2(Socio As Boolean, TipoAlb As String, cadSQL As String, textoCSB As String, Optional ErroresAux As String, Optional EstaRecu As Boolean) As Boolean
'IN ->  cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'       TextosCSB:  Para la tesoreria
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoAlbFac

    PasarAlbaranesAFactura2 = False
    b = False
    ErroresAux = ""
    '#Laura 14/11/2006 recuperar facturas de Alzira
    'si esta recuperando facturas
    EstaRecuFac = EstaRecu
    
    Select Case TipoAlb
        Case "ALV", "ALM": 'ALV: Albaranes venta a clientes
                    'ALM: Albaran mostrador
            codtipom = "FAV"
        Case "ALR": 'Albaranes de reparacion en clientes
            codtipom = "FAR"
        Case "ART": 'Albaranes de factura rectificativa
            codtipom = "FRT"
        Case "ARC": 'Albaranes de factura rectificativa de cuotas
            codtipom = "FRC"
        Case "ALS": 'Albaranes de servicios [SERVICIOS]
            codtipom = "FCN"
        Case "ALZ"  'Albaranes "presupuestos". Es decir, el "B"
            codtipom = "FAZ"
    End Select
      
     
    OpeFactu = PonerTrabajadorConectado("")
    If OpeFactu <> "" Then
        AlmFactu = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", OpeFactu, "N")
    Else
        AlmFactu = "1"
    End If
      
    'Aqui empieza transaccion
    conn.BeginTrans
    ConnConta.BeginTrans
   
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFactura(cadSQL)
    
    
    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    '#Laura: 14/11/2006 Recuperar facturas Alzira
    If b Then
        If EstaRecuFac = False Then
            If Not Socio Then
                MenError = "Error al actualizar en Tesoreria."
                b = InsertarEnTesoreria(textoCSB, MenError)
            End If
        Else
            MenError = "Error al actualizar en Tesoreria."
            b = InsertarEnTesoreria_RecupFac(MenError)
        End If
    End If
    
    'Eliminar los Albaranes incluidos en la Factura
    If b Then
        MenError = "Error al eliminar los albaranes."
        b = EliminarAlbaranes(cadSQL)
    End If
    
    'Si alguna linea de albaran tiene asociados num. serie actualizar
    'el campo numfactu de dichos num serie
    If b Then
        MenError = "Error al actualizar los Nº de serie."
        b = ActualizarNumSeries(cadSQL)
    End If
    
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        PasarAlbaranesAFactura2 = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarAlbaranesAFactura2 = False
        ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    End If
    Espera 0.4
End Function



Public Function PasarMtosAFactura(TipCo As String, Operador As String, mes As String, nMante As String, CenCoste As String) As Boolean
'IN -> cadSQL: cadena para seleccion de los Mantenimientos que vamos a Facturar
'Desde Mantenimientos Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String
Dim CC As String

    On Error GoTo ETraspasoFac

    PasarMtosAFactura = False
    b = False
      
    codtipom = "FAM"
    
    TipCoMan = TipCo
    OpeFactu = Operador
    MesFactu = mes
    nummante = nMante
    'CEntro de coste
    manCCost = ""
    
    
    ArticFac = DevuelveDesdeBDNew(conAri, "stipco", "codartic", "codtipco", TipCoMan, "T")
    If ArticFac = "" Then
        MenError = "El tipo de contrato no tiene articulo para facturar."
        Exit Function
    End If
      
    
    If OpeFactu <> "" Then
        'Cojo el del trabajador para no hacer el select otra vez
        manCCost = "codccost"
        AlmFactu = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", OpeFactu, "N", manCCost)
    Else
        AlmFactu = "1"
    End If
    
    'CEntro de coste
    
    If vEmpresa.TieneAnalitica Then
        'Tienen analitica
        Select Case vParamAplic.ModoAnalitica
        Case 0
            'Trabajador
            'YA lo he hecho arriba
            ' manCCost = devuelve......
        
        Case 1
            'Familia
            manCCost = DevuelveDesdeBD(conAri, "codccost", "sartic,sfamia", "sartic.codfamia=sfamia.codfamia AND sartic.codartic", ArticFac, "T")
        Case Else
            'proyecto
            manCCost = CenCoste
        
        End Select
        If manCCost = "" Then
            MsgBox "Centro de coste no especificado", vbExclamation
            Exit Function
        End If
    Else
        manCCost = ""
    End If
    'Aqui empieza transaccion
    conn.BeginTrans
    ConnConta.BeginTrans
    
       
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura"
    b = InsertarFactura("")
    
    
    If b Then
        'Insertar en la Tabla para detalle equipos, para reimprimir
       MenError = "Error al insertar detalle de equipos."
       InsertarDetEquip
    End If
    
    
    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    If b Then
        MenError = "Error al actualizar en Tesoreria"
        b = InsertarEnTesoreria("", MenError)
    End If
    
   
   
    'Actualizo en la tabla mantenimientos la fecha ultimio mantenimiento
    If b Then
        MenError = "Actualizar fecha ultimo mantenimiento"
        ArticFac = "UPDATE scaman SET ulmesfac = " & mes & " WHERE codclien = " & Cliente & " AND nummante= '" & nummante & "'"
        conn.Execute ArticFac
    End If
   
ETraspasoFac:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Pasando Mantenimientos a Factura.", Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        PasarMtosAFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarMtosAFactura = False
    End If
    Espera 0.2
End Function



Public Function PasarTicketAFactura(cadSQL As String, Optional ErroresAux As String, Optional NTicket As String, Optional NAlbTicket As String, Optional impRegalo As String) As Boolean
'IN -> cadSQL: cadena para seleccion de las ventas del ticket que vamos a Facturar
'Desde ventas Genera las Facturas de ticket correspondiente
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETrasTicFac

    OpeFactu = ErroresAux 'Aqui hemos pasado el trabajador
    NumTicket = NTicket
    NumAlbTicket = NAlbTicket '01/09/06 laura
    ImpCheque = CCur(ComprobarCero(impRegalo)) 'importe del cheque regalo (si hay)
    
    PasarTicketAFactura = False
    b = False
    ErroresAux = ""
      
      
    'por defecto se coge el almacen con el q trabaja el empleado
    If OpeFactu <> "" Then
        AlmFactu = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", OpeFactu, "N")
    Else
        AlmFactu = "1"
    End If
    
    
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFacturaTicket(cadSQL)
    
    
    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    If b Then
        MenError = "Error al actualizar en Tesoreria."
        b = InsertarEnTesoreria("", MenError)
    End If
    
    
    'Eliminar los Albaranes incluidos en la Factura
    If b Then
        MenError = "Error al eliminar la venta."
        b = EliminarVenta(cadSQL)
    End If
    
    
    
ETrasTicFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If Not b Then ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    PasarTicketAFactura = b
End Function



Private Function InsertarFactura(cadSQL As String, Optional cadNumFac As String) As Boolean
'Insertamos en las tablas de factura de cliente: scafac, scafac1, slifac
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean
Dim devuelve As String
Dim EsAnterior As Boolean

    InsertarFactura = False
    
    
    'Obtener el Contador de FACTURA
    '--------------------------------------------------------------
    Set vTipoMov = New CTiposMov
        
    If vTipoMov.Leer(codtipom) Then
        '#Laura: 14/11/2006 Recuperar facturas ALZIRA
        If EstaRecuFac = False Then
            'Comprobar si mientras tanto se incremento el contador de Facturas
            'para ello vemos si existe una Factura con ese contador y si existe la incrementamos
            Do
                NumFactu = vTipoMov.ConseguirContador(codtipom)
                devuelve = DevuelveDesdeBDNew(conAri, "scafac", "numfactu", "codtipom", codtipom, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFactu, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (codtipom)
                    NumFactu = vTipoMov.ConseguirContador(codtipom)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
        End If
    Else
        Exit Function
    End If
        
    LetraSerie = vTipoMov.LetraSerie 'la guardamos para INSERT en scobro de la Conta
    
    EsAnterior = False
    Dim FecAlbar As Date
    FecAlbar = DevuelveValor("select fechaalb from scaalb where " & cadSQL)
    If (codtipom = "FRC" And FecAlbar < vParamAplic.FecCambioIva) Or (codtipom = "FRT" And FecFactu < vParamAplic.FecCambioIva) Then EsAnterior = True
    
    'calcular el total de la factura antes de insertar
    b = CalcularDatosFactura(False, cadSQL, "scaalb", "slialb", EsAnterior)
    If Not b Then Exit Function
    
    '### Laura: 14/11/2006 Recuperar facturas Alzira
    devuelve = ""
    If EstaRecuFac = True Then
        'comprobar q el total de la factura a insertar coincide con el total de la
        'factura en conta.cabfact si esta existe en conta
        If vParamAplic.ContabilidadNueva Then
            devuelve = "SELECT COUNT(*) FROM factcli WHERE numserie=" & DBSet(Me.LetraSerie, "T")
            devuelve = devuelve & " AND numfactu=" & Me.NumFactu & " AND anofactu=" & Year(FecFactu)
            If (RegistrosAListar(devuelve, conConta) > 0) Then
                devuelve = DevuelveDesdeBDNew(conConta, "factcli", "totfactu", "numserie", Me.LetraSerie, "T", , "numfactu", Me.NumFactu, "N", "anofactu", Year(Me.FecFactu), "N")
                If devuelve <> "" Then
                    If CCur(devuelve) <> Me.TotalFac Then
                        If MsgBox("El total de la factura no coincide con el total en la Contabilidad" & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
                            b = False
                        End If
                    End If
                End If
                devuelve = "EXISTE"
            End If
        
        Else
            devuelve = "SELECT COUNT(*) FROM cabfact WHERE numserie=" & DBSet(Me.LetraSerie, "T")
            devuelve = devuelve & " AND codfaccl=" & Me.NumFactu & " AND anofaccl=" & Year(FecFactu)
            If (RegistrosAListar(devuelve, conConta) > 0) Then
                devuelve = DevuelveDesdeBDNew(conConta, "cabfact", "totfaccl", "numserie", Me.LetraSerie, "T", , "codfaccl", Me.NumFactu, "N", "anofaccl", Year(Me.FecFactu), "N")
                If devuelve <> "" Then
                    If CCur(devuelve) <> Me.TotalFac Then
                        If MsgBox("El total de la factura no coincide con el total en la Contabilidad" & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
                            b = False
                        End If
                    End If
                End If
                devuelve = "EXISTE"
            End If
        End If
        If Not b Then Exit Function
    End If
    '#####
    
    'Cabecera Factura
    b = InsertarCabeceraFactu
    If Not b Then Exit Function
    
    'Cabeceras Albaranes de la Factura
    b = InsertarCabAlbaranesFactu(cadSQL)
    If Not b Then Exit Function
    
    'Insertar lineas de la factura
    b = InsertarLineasFactu(cadSQL)
    If Not b Then Exit Function
    
    'Comprobar si los albaranes tienen km de desplazamiento e insertar
    'lineas de factura del desplazamiento e incrementar el bruto de la factura
    'Insertar lineas de desplazamiento
    b = InsertarDesplaz
    If Not b Then Exit Function
    
    'si hay bonificaciones en la factura
    If SQLBonif <> "" Then
        b = InsertarBonificaciones
        If Not b Then Exit Function
    End If
    
   
    '#Laura: 14/11/2006 Recuperar facturas Alzira
    If EstaRecuFac = True And devuelve = "EXISTE" Then
        b = MarcarContabilizada
        If Not b Then Exit Function
    End If
    
    
    'Incrementar contador del tipo de movimiento
    '-------------------------------------------------------------
    '#Laura: 14/11/2006 Recuperar facturas Alzira
    If EstaRecuFac = False Then vTipoMov.IncrementarContador (codtipom)
    Set vTipoMov = Nothing
      
    InsertarFactura = True
End Function




Private Function InsertarFacturaTicket(cadSQL As String) As Boolean
'Insertamos en las tablas de factura de cliente: scafac, scafac1, slifac
'Es un ticket de venta
Dim b As Boolean
Dim devuelve As String

    InsertarFacturaTicket = False
    
    
    'Calcular los campos de la factura y cuadrarla al
    'pasar las lineas sin IVA
    b = CalcularDatosFacturaTicket(cadSQL)
    If Not b Then Exit Function

    
    'Cabecera Factura
    b = InsertarCabeceraFactu
    If Not b Then Exit Function
    
    'Cabeceras Albaranes de la Factura
    b = InsertarCabAlbaranesFactu(cadSQL)
    If Not b Then Exit Function

    'Insertar lineas de la factura
    b = InsertarLineasFactu(cadSQL)
    If Not b Then Exit Function
    
       
'    'Incrementar contador del tipo de movimiento
'    '-------------------------------------------------------------
'    vTipoMov.IncrementarContador (codTipoM)
'    Set vTipoMov = Nothing
      
    InsertarFacturaTicket = True
End Function




Private Function InsertarCabeceraFactu() As Boolean
'Inserta la cabecera de la factura en la tabla: scafac
Dim Sql As String
Dim vClien As CCliente
Dim Nulo2 As String
Dim Nulo3 As String
Dim ctaDpto As Boolean
Dim Cad As String, cad2 As String

    On Error GoTo EInsertar

    
    'Insertar en la tabla cabecera de la factura de compras
    Sql = "INSERT INTO scafac (codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,coddirec,nomdirec,"
    Sql = Sql & "codagent,codforpa,dtoppago,dtognral,codbanco,codsucur,digcontr,cuentaba,brutofac,impdtopp,impdtogr,baseimp1,baseimp2,baseimp3, "
    Sql = Sql & "codigiv1,codigiv2,codigiv3,porciva1,porciva2,porciva3,imporiv1,imporiv2,imporiv3,totalfac,intconta "
            ' Añadimos IVA recargo de equivalencia                               02 Junio 08
    Sql = Sql & ",porciva1re,imporiv1re, porciva2re,imporiv2re, porciva3re,imporiv3re,aportacion,aridoc, iban)"
    
    
    
    Sql = Sql & " VALUES (" & DBSet(codtipom, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFactu, "F") & "," & DBSet(Cliente, "N") & ","
    Sql = Sql & DBSet(NombreClien, "T") & "," & DBSet(DomicilioClien, "T", "N") & "," & DBSet(CPostal, "T", "N") & "," & DBSet(Poblacion, "T", "N") & "," & DBSet(Provincia, "T", "N") & ","
    Sql = Sql & DBSet(NIF, "T", "N") & "," & DBSet(Telefono, "T") & "," & DBSet(DirDpto, "N", "S") & "," & DBSet(NombreDirDpto, "T") & "," & DBSet(Agente, "N") & ","
    Sql = Sql & ForPago & "," & DBSet(DtoPPago, "N") & "," & DBSet(DtoGnral, "N") & ","
    
    If codtipom = "FAV" Or codtipom = "FRT" Or codtipom = "FRC" Or codtipom = "FAR" Or codtipom = "FAS" Or codtipom = "FCN" Or codtipom = "FRQ" Then '[Monica]05/03/2013: añadido FCN
        'si son facturas de VENTA o Factura Rectificativa
        'Cargar los datos del banco del cliente
        'si se trabaja con departamentos y la factura va a un departamento
        'comprobar si tiene banco y asignar el del departamento
        ctaDpto = False
        If vParamAplic.Departamento Then
            If Me.DirDpto <> "" Then
                'comprobar que el dpto del cliente tiene cta bancaria
                cad2 = "digcontr"
                Cad = DevuelveDesdeBDNew(conAri, "sdirec", "cuentaba", "codclien", Me.Cliente, "N", cad2, "coddirec", Me.DirDpto, "N")
                If Cad <> "" And cad2 <> "" Then ctaDpto = True
            End If
        End If
        
        If ctaDpto Then 'cogemos la cta bancaria del departamento
            Me.CuentaBan = Cad
'            cad = DevuelveDesdeBDNew(conAri, "sdirec", "digcontr", "codclien", Me.Cliente, "N", , "coddirec", Me.DirDpto, "N")
            Me.DigControl = cad2
            cad2 = "codsucur"
            Cad = DevuelveDesdeBDNew(conAri, "sdirec", "codbanco", "codclien", Me.Cliente, "N", cad2, "coddirec", Me.DirDpto, "N")
            Me.Banco = Cad
            Me.Sucursal = cad2
        Else 'cuenta bancaria del cliente
            Set vClien = New CCliente
            If vClien.LeerDatos(Cliente, True) Then
                '[Monica]22/11/2013: iban
                Iban = vClien.Iban
                Banco = vClien.Banco
                Sucursal = vClien.Sucursal
                DigControl = vClien.DigControl
                CuentaBan = vClien.CuentaBan
            Else
                '[Monica]22/11/2013: iban
                Iban = ValorNulo
                Banco = ValorNulo
                Sucursal = ValorNulo
                DigControl = ValorNulo
                CuentaBan = ValorNulo
            End If
            Set vClien = Nothing
        End If
    End If
    
    Sql = Sql & DBSet(Banco, "N") & "," & DBSet(Sucursal, "N") & "," & DBSet(DigControl, "T") & "," & DBSet(CuentaBan, "T") & ","
    
    Sql = Sql & DBSet(BrutoFac, "N") & ","
    Sql = Sql & DBSet(ImpPPago, "N") & "," & DBSet(ImpGnral, "N") & ","
    Sql = Sql & DBSet(BaseIVA1, "N") & "," & DBSet(BaseIVA2, "N", "S") & "," & DBSet(BaseIVA3, "N", "S") & ","
    Nulo2 = "N"
    Nulo3 = "N"
    If BaseIVA2 = 0 Then Nulo2 = "S"
    If BaseIVA3 = 0 Then Nulo3 = "S"
    Sql = Sql & DBSet(TipoIVA1, "N") & "," & DBSet(TipoIVA2, "N", Nulo2) & "," & DBSet(TipoIVA3, "N", Nulo3) & ","
    Sql = Sql & DBSet(PorceIVA1, "N") & "," & DBSet(PorceIVA2, "N", Nulo2) & "," & DBSet(PorceIVA3, "N", Nulo3) & ","
     
    Sql = Sql & DBSet(ImpIVA1, "N") & "," & DBSet(ImpIVA2, "N", Nulo2) & "," & DBSet(ImpIVA3, "N", Nulo3) & ","
     
    Sql = Sql & DBSet(TotalFac, "N") & ",0"
    
    'AÑADIMOS EL recargo de equivalencia en el INSERT  porciva1re,imporiv1re .......
    Sql = Sql & "," & DBSet(PorceIVA1RE, "N", "S") & "," & DBSet(ImpIVA1RE, "N", "S")
    Sql = Sql & "," & DBSet(PorceIVA2RE, "N", "S") & "," & DBSet(ImpIVA2RE, "N", "S")
    Sql = Sql & "," & DBSet(PorceIVA3RE, "N", "S") & "," & DBSet(ImpIVA3RE, "N", "S") & ","
    
    'aportacion,aridoc
    Sql = Sql & DBSet(mAportacion, "N", "N") & ",NULL," ')"
    
    '[Monica]22/11/2013: iban
    Sql = Sql & DBSet(Iban, "T") & ")"
    
    conn.Execute Sql
    InsertarCabeceraFactu = True
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabeceraFactu = False
        MuestraError Err.Number, "Insertar cabecera factura.", Err.Description
    Else
        InsertarCabeceraFactu = True
    End If
End Function


Private Function InsertarCabAlbaranesFactu(cadSQL As String) As Boolean
'Insertar las cabeceras de los Albaranes de la factura en la tabla: scafac1
'NumTicket: para cuando de venta generamos Factura en numalbar guardamos el ticket que la genero
Dim Sql As String
Dim devuelve As String
Dim MantenimientoObservacion As String

    On Error GoTo EInsertarAlb

    devuelve = ""
    If codtipom = "FTI" Then
'        cadSQL = Replace(cadSQL, "sliven", "scaven")
    Else
        If InStr(cadSQL, "scaven") > 0 Then 'viene de generar factura en ticket
'            codTipoM = "FTI"
        Else
            cadSQL = Replace(cadSQL, "slialb", "scaalb")
        End If
    End If
    
    Sql = "INSERT INTO scafac1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,numpedcl,fecpedcl,sementre,numofert,fecofert,referenc,codenvio,codtraba,codtrab1,codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa)"
    
    If (codtipom = "FAV" And NumTicket <> "") Then
        'Facturas generadas desde una venta de tickets: en numalbar guardamos numticket
'        SQL = SQL & " VALUES('" & codTipoM & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "','FTI'," & NumTicket & ","
        '01/09/06 laura
        Sql = Sql & " VALUES('" & codtipom & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "','ATI'," & NumTicket & ","
        
        Sql = Sql & "'" & Format(FecFactu, FormatoFecha) & "'," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de pedido a nulo
        Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de oferta a nulo
        'obtenemos el codigo de envio de cliente
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codenvio", "codclien", Cliente, "N")
        Sql = Sql & DBSet(devuelve, "N") & ","
        'trabajadores
        Sql = Sql & DBSet(OpeFactu, "N") & "," & ValorNulo & "," & DBSet(OpeFactu, "N") & ","
        'observaciones
        '---- Laura 10/10/06: poner en la observacion 1 el valor de scaven.observa1
        devuelve = DevuelveDesdeBDNew(conAri, "scaven", "observa1", "numtermi", Me.NumTerminal, "N", , "numventa", Me.NumVenta, "N", "fecventa", Me.FecFactu, "F")
        Sql = Sql & DBSet(devuelve, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        
        Sql = Sql & NumTerminal & "," & NumVenta & ")"
        
    ElseIf (codtipom <> "FAM" And codtipom <> "FTI") And cadSQL <> "" Then 'Facturas de VEnta
        Sql = Sql & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & "codtipom as codtipoa, numalbar,fechaalb,numpedcl,fecpedcl,sementre,numofert,fecofert,referenc,codenvio,codtraba,codtrab1,codtrab2,observa01,observa02,observa03,observa04,observa05,numtermi,numventa "
        Sql = Sql & " FROM scaalb WHERE " & cadSQL
        
        
    Else 'Bonificaciones/Desplazamientos
        'Facturas de Mantenimiento
        'Facturas de Ticket del TPV
        Sql = Sql & " VALUES('" & codtipom & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "',"
'        If (codTipoM = "FAM" Or codTipoM = "FTI") Then
        '01/09/06 laura
        If (codtipom = "FAM") Then
            Sql = Sql & "'',0,"
        ElseIf codtipom = "FTI" Then
            Sql = Sql & "'ATI'," & NumAlbTicket & ","
        Else
            Sql = Sql & "'',9999999,"
        End If
        Sql = Sql & "'" & Format(FecFactu, FormatoFecha) & "'," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de pedido a nulo
        Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de oferta a nulo
        'obtenemos el codigo de envio de cliente
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codenvio", "codclien", Cliente, "N")
        Sql = Sql & DBSet(devuelve, "N") & ","
        'trabajadores
        Sql = Sql & DBSet(OpeFactu, "N") & "," & ValorNulo & "," & DBSet(OpeFactu, "N") & ","
        'observaciones 1
        If codtipom = "FTI" Then
            '---- Laura 19/12/06: poner en la observacion 1 el valor de scaven.observa1
            devuelve = DevuelveDesdeBDNew(conAri, "scaven", "observa1", "numtermi", Me.NumTerminal, "N", , "numventa", Me.NumVenta, "N", "fecventa", Me.FecFactu, "F")
        Else
            'La linea viene en el campo concefac que hemos guardado en observaciones
            If codtipom = "FAM" Then
                
                MantenimientoObservacion = DevuelveDesdeBDNew(conAri, "scaman", "tipopago", "codclien", Me.Cliente, "N", , "nummante", nummante, "T")
                If MantenimientoObservacion = "1" Then 'Trimestral
                    MantenimientoObservacion = "TRIMESTRAL DE " & Year(FecFactu)
                ElseIf MantenimientoObservacion = "2" Then 'Semestral
                    MantenimientoObservacion = "SEMESTRAL DE " & Year(FecFactu)
                ElseIf MantenimientoObservacion = "3" Then 'Anual
                    MantenimientoObservacion = "ANUAL " & Year(FecFactu)
                Else
                    'si es mensual: ej: "JUNIO de 2005" (ponemos el mes seleccionado para facturar)
                    MantenimientoObservacion = MonthName(CLng(MesFactu)) & " de " & Year(FecFactu)
                    MantenimientoObservacion = UCase(MantenimientoObservacion)
                End If
                                                            
                If mObservacion <> "" Then
                    devuelve = MantenimientoObservacion
                    
                Else
                    devuelve = ""
                    mObservacion = MantenimientoObservacion
                End If
                
            Else
                devuelve = ""
            End If
        End If
        Sql = Sql & DBSet(devuelve, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
        Sql = Sql & "," & DBSet(NumTerminal, "N", "S") & "," & DBSet(NumVenta, "N", "S") & ")"
    End If
    
    conn.Execute Sql

    InsertarCabAlbaranesFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabAlbaranesFactu = False
        MuestraError Err.Number, "Insertar cabecera albaranes factura.", Err.Description
    Else
        InsertarCabAlbaranesFactu = True
    End If
End Function


Private Function InsertarLineasFactu(cadSQL As String) As Boolean
'Insertar las lineas de los Albaranes de la factura en la tabla: slifac
Dim Sql As String
Dim Cad As String
Dim cad2 As String
Dim cArt As CArticulo

    On Error GoTo EInsertarLin

    '##########################################
    'Cambios para llevar la TRAZA con codprove
    '   si cojo slialb:  tengo la columna
    '   RESTO DE CASOS!!!!! FALTA###


    'Otra modificacion:   16 Mayo 08
    ' Se permiten dtos, con lo cual la donde ponia 0 as dto ya no vale
    

    Sql = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre"
    If NumTicket <> "" Then Sql = Sql & ",precioiv"
    Sql = Sql & ",preciomp,preciost,preciouc,codprovex,numlote,codccost)"     'ENERO 2008.  Traza sobre codprove.
                                                                            '[19/10/2009]: añadir centro coste
    
    If codtipom = "FAM" Then 'facturas de mantenimientos
        Sql = Sql & " VALUES('" & codtipom & "'," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','',0,1,"
            
        Sql = Sql & DBSet(AlmFactu, "N") & ","
            
        'Articulo para facturar mantenimientos de la tabla de tipos contrato: stipco.codartic
        Sql = Sql & DBSet(ArticFac, "T") & ", "
            
        'nombre del articulo para facturar,preciomp,preciost, preciouc
        If ArticFac <> "" Then
            Set cArt = New CArticulo
            If cArt.LeerDatos(ArticFac) Then
                Cad = cArt.Nombre   'Aunque luego machaco cad con la nueva amplicacion: contrato + numerocontrato
            End If
'            cad2 = "preciomp"
'            cad = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", ArticFac, "T", cad2)
        End If
        
        Cad = DevuelveDesdeBDNew(conAri, "stipco", "nomtipco", "codtipco", TipCoMan, "T")
        Cad = Cad & "  Nº: " & nummante
        
        Sql = Sql & DBSet(Cad, "T") & ","
            

        Cad = mObservacion
        
        Sql = Sql & DBSet(Cad, "T") & ","
            
        'cantidad,numbultos, precio, dtos
        Sql = Sql & "1,1," & DBSet(BrutoFac, "N") & ", 0,0, "
        'importe, origen precio
        Sql = Sql & DBSet(BrutoFac, "N") & ", 'M',"
        
        'guardar los precios del articulo
        Sql = Sql & DBSet(cArt.PrecioMedPon, "N") & "," & DBSet(cArt.PrecioStan, "N") & "," & DBSet(cArt.PrecioUltCom, "N") & ","
        'traza ENERO 2008                    'NUmero lote Junio2009
        Sql = Sql & DBSet(cArt.codProve, "N") & ",NULL,"
        'CEntro de coste
        Sql = Sql & DBSet(manCCost, "T", "S")
        Sql = Sql & ")"
        Set cArt = Nothing
        
    ElseIf (codtipom = "FTI") Then 'factura de ticket
        'se entra en esta opcion cuando desde el TPV generamos un ticket
        Sql = Sql & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu, 'ATI' as codtipoa," & NumAlbTicket & " as numalbar," & "numlinea,"
        Sql = Sql & AlmFactu & " as codalmac," & "sliven.codartic,sliven.nomartic," & ValorNulo & " as ampliaci,cantidad,1 as numbultos,precioar,"
        'David 16 Mayo 08
        '0 as dtoline1,0 as dtoline2, cantidad*precioar as importel,'' as origpre,precioiv "
        Sql = Sql & " dto1,dto2, implineareal,'' as origpre,precioiv "
        
         '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & ", sartic.preciomp,sartic.preciost,sartic.preciouc,sartic.codprove,NULL as numlote"
        '-- [19/10/2009] [LAURA] : añadir centro de coste
        Sql = Sql & ",codccost"
        '--
        Sql = Sql & " FROM sliven "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & " INNER JOIN sartic ON sliven.codartic=sartic.codartic "
        '--
        Sql = Sql & " WHERE " & Replace(cadSQL, "scaven", "sliven")
        
    ElseIf NumTicket <> "" Then 'Ticket de venta genera factura
        Sql = Sql & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu, 'ATI' as codtipoa," & NumTicket & " as numalbar," & "numlinea,"
        Sql = Sql & AlmFactu & " as codalmac," & "sliven.codartic,sliven.nomartic," & ValorNulo & " as ampliaci,cantidad,1 as numbultos,precioar,"
        
        'David 16 Mayo 08
        '0 as dtoline1,0 as dtoline2, cantidad*precioar as importel,'' as origpre,precioiv "
        Sql = Sql & " dto1,dto2, implineareal,'' as origpre,precioiv "
        
        
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & ", sartic.preciomp,sartic.preciost,sartic.preciouc,sartic.codprove,NULL as numlote"
        '-- [19/10/2009] [LAURA] : añadir centro de coste
        Sql = Sql & ",codccost"
        '--
        Sql = Sql & " FROM sliven "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & " INNER JOIN sartic ON sliven.codartic=sartic.codartic "
        '--
        Sql = Sql & " WHERE " & Replace(cadSQL, "scaven", "sliven")
    
    Else 'pasamos lineas de albaranes a lineas factura
        Sql = Sql & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & "codtipom as codtipoa, numalbar,numlinea,codalmac,slialb.codartic,slialb.nomartic,ampliaci,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & ", sartic.preciomp,sartic.preciost,sartic.preciouc,codprovex,numlote"   'ENERO 2008. Traza cond codprove
        '--
        ' -- [19/10/2009] [LAURA]: se añade el centro de coste
        Sql = Sql & ",codccost"
        
        Sql = Sql & " FROM slialb "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        Sql = Sql & " INNER JOIN sartic ON slialb.codartic=sartic.codartic "
        '--
        Sql = Sql & " WHERE " & Replace(cadSQL, "scaalb", "slialb")
    End If
    
    conn.Execute Sql

    InsertarLineasFactu = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasFactu = False
        MuestraError Err.Number, "Insertar lineas factura.", Err.Description
    Else
        InsertarLineasFactu = True
    End If
End Function


Private Function EliminarAlbaranes(cadSQL As String) As Boolean
'Eliminamos de las tablas de Albaranes: scaalb, slialb
Dim Sql As String

    On Error GoTo EEliminar

    EliminarAlbaranes = False
    
    'ELiminar lineas albaranes
    Sql = "DELETE FROM slialb "
    Sql = Sql & " WHERE " & Replace(cadSQL, "scaalb", "slialb")
    conn.Execute Sql
    
    Espera 0.1
    
    'Eliminar Cabeceras Albaranes
    Sql = "DELETE FROM scaalb "
    Sql = Sql & " WHERE " & Replace(cadSQL, "slialb", "scaalb")
    conn.Execute Sql
        
    EliminarAlbaranes = True

EEliminar:
    If Err.Number <> 0 Then
        EliminarAlbaranes = False
    Else
        EliminarAlbaranes = True
    End If
End Function



Private Function InsertarDetEquip() As Boolean
Dim Sql As String
Dim Cad As String

    'si son facturas de MAntenimiento y parametro de stipco: detallar equipos en la factura=true
    'guardar en una tabla el resumen por tipo de articulo de los Nº Series de equipos
    'asociados al MAntenimiento
    InsertarDetEquip = False
    
    On Error GoTo EDetalle
       
    Sql = DevuelveDesdeBDNew(conAri, "stipco", "detequip", "codtipco", TipCoMan, "T")
    
    If Sql = "1" Then 'Detalla Equipo
        Cad = ObtenerSQLcomponentes(" WHERE codclien=" & Cliente & " AND nummante=" & DBSet(nummante, "T"))
        Sql = "Select distinct '" & codtipom & "' as codtipom," & NumFactu & " as numfactu,'" & Format(FecFactu, FormatoFecha) & "' as fecfactu,"
        Cad = Replace(Cad, "Select distinct", Sql)
        Sql = "INSERT INTO srecom (codtipom,numfactu,fecfactu,codtipar,nomtipar,cantidad) "
        Sql = Sql & Cad

        conn.Execute Sql
    End If
    
EDetalle:
    If Err.Number <> 0 Then
        InsertarDetEquip = False
    Else
        InsertarDetEquip = True
    End If
End Function


'ANTES:  ClienteExentoIVA  as boolean
'Ahora:  ClienteExentoIVA2 as byte
'           0.- Exento
'           1.- RE
'           2.- EXNTO
Private Function ClienteExentoIVA2() As Byte
Dim vClien As CCliente

    Set vClien = New CCliente
    vClien.Codigo = Cliente
    ClienteExentoIVA2 = vClien.exentoIVA2
    Set vClien = Nothing
    
    'ClienteExentoIVA = b
End Function



Private Function InsertarDesplaz() As Boolean
'Si los albaranes a facturar tienen KM de desplazamiento añadirlos a la factura
'Añadir cantkm * preciokm al bruto de la factura
'Insertar linea en tabla slifac con el importe del desplazamiento
Dim Sql As String
Dim Cad As String
Dim b As Boolean

    On Error GoTo EDesplaz
    InsertarDesplaz = False
    
    If TotalKm > 0 Then
        'Insertar en scafac1 un albaran a 9999999, para el desplazamiento (y servira tb para bonificaciones)
        '-----------------------------------------------------------
        b = InsertarCabAlbaranesFactu("")
        If Not b Then Exit Function
            
               
        'Insertar linea de Desplazamiento en tabla: slifac (lineas factura)
        '------------------------------------------------------------------
        Sql = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
        Sql = Sql & "cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
        'codtipom,numfactu,fecfactu,codtipoa,numalbar
        Sql = Sql & " VALUES (" & DBSet(codtipom, "T") & "," & NumFactu & "," & DBSet(FecFactu, "F") & ",'',9999999,"
        'numlinea
        Sql = Sql & "1,"
        Sql = Sql & DBSet(AlmFactu, "N") & ","
            
        'codartic,nomartic,ampliaci
        Cad = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", codArtic, "T")
        Sql = Sql & DBSet(codArtic, "T") & "," & DBSet(Cad, "T") & "," & DBSet(Ampliacion, "T") & ","
        'cantidad,precioar,dtoline1,dtoline2,importel,origpre
        Sql = Sql & DBSet(TotalKm, "N") & "," & DBSet(PrecioKm, "N") & ",0,0," & DBSet(ImporteL, "N") & ",'M')"
        
        conn.Execute Sql
        
    End If
    InsertarDesplaz = True
     
EDesplaz:
    If Err.Number <> 0 Then
        InsertarDesplaz = False
    Else
        InsertarDesplaz = True
    End If
End Function


Private Function CalcularDesplazamiento(cadWHERE As String, NomTabla As String) As Boolean
'comprobar si los albaranes a facturar tienen DESPLAZAMIENTOS
'calcula en vbles: Totalkm, Ampliacion los valores para insertar en linea factura
Dim Rs As ADODB.Recordset
Dim Sql As String
    
    CalcularDesplazamiento = False
    On Error GoTo EDesp
    
    If codtipom = "FAV" Or codtipom = "FAR" Then
        'Obtenemos el total de km a facturar en los albaranes seleccionados
        Sql = "SELECT fechaalb,sum(cantidkm)" ',sum(sclien.kilometr) "
        Sql = Sql & " FROM " & NomTabla '& " INNER JOIN sclien ON " & NomTabla & ".codclien=sclien.codclien "
        Sql = Sql & " WHERE " & cadWHERE '& " AND facturkm=1 "
        Sql = Sql & " GROUP BY fechaalb"
        Sql = Sql & " ORDER BY fechaalb"
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        TotalKm = 0
        Ampliacion = ""
        While Not Rs.EOF
            If Rs.Fields(1).Value <> 0 Then
                TotalKm = TotalKm + Rs.Fields(1).Value
                Ampliacion = Ampliacion & Day(Rs.Fields(0).Value) & ", "
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        If Ampliacion <> "" Then 'quitar la ultima coma
            Ampliacion = Mid(Ampliacion, 1, Len(Ampliacion) - 2)
            Ampliacion = "DIAS: " & Ampliacion
        End If
        
        If TotalKm > 0 Then
            'obtenemos el precio del km de los parametros, y el cod Articulo
            codArtic = "codartid"
            Sql = DevuelveDesdeBDNew(conAri, "spara1", "preukmcl", "codigo", "1", "N", codArtic)
            PrecioKm = CCur(ComprobarCero(Sql))
            ImporteL = Round(TotalKm * PrecioKm, 2)
            
            'obtenemos el tipo IVA del articulo de desplazamiento
            Sql = DevuelveDesdeBDNew(conAri, "sartic", "codigiva", "codartic", codArtic, "T")
            despIVA = CInt(ComprobarCero(Sql))
        End If
    End If
    CalcularDesplazamiento = True
    
EDesp:
    If Err.Number <> 0 Then
        CalcularDesplazamiento = False
    Else
        CalcularDesplazamiento = True
    End If
End Function



Private Function CalcularBonificacion(cadWHERE As String, NomTablaLin As String) As Boolean
'comprobar los articulos de los albaranes a facturar tienen BONIFICACION
Dim Rs As ADODB.Recordset
Dim RSaux As ADODB.Recordset
Dim Sql As String, cadIVA As String
Dim vClien As CCliente
Dim EsExcluyente As Boolean
Dim Bonif1 As Currency, Bonif2 As Currency
Dim Canti1 As Single, Canti2 As Single
Dim numlinea As Integer

    CalcularBonificacion = False
    On Error GoTo EBonif
    
    If codtipom = "FAV" Then 'Factura de Venta
        'Obtener el cod. tarifa del cliente al que se va a facturar y
        'comprobar si para este tipo de tarifa se aplica Bonificacion
        Set vClien = New CCliente
        vClien.Codigo = Cliente
        SQLBonif = ""
        TotBonif1 = 0
        TotBonif2 = 0
        TotBonif3 = 0
        IVABonif1 = 0
        IVABonif2 = 0
        IVABonif3 = 0
            
        If vClien.Bonifica Then
            'Comprobar para los codartic de las lineas de albaran que vamos
            'a factura, para cuales hay bonificacion y la cantidad de la linea
            'cumple la cantidad para que se le aplica la bonificacion
            Sql = "SELECT slialb.codartic, sum(slialb.cantidad) as cantidad "
            Sql = Sql & "FROM " & NomTablaLin
            Sql = Sql & " INNER JOIN sbonif ON " & NomTablaLin & ".codartic=sbonif.codartic "
            Sql = Sql & " WHERE " & Replace(cadWHERE, "scaalb", NomTablaLin)
            Sql = Sql & " GROUP BY " & NomTablaLin & ".codartic "
            Sql = Sql & " order by " & NomTablaLin & ".codartic "
            
           
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            numlinea = 1
            
            While Not Rs.EOF 'para cada codartic
                'Para cada linea de slialb a la que se le aplica bonificacion
                'obtener la bonif. que se le va a aplicar

                'mirar primero si hay algun hastaca2<=cantidad
                '--------------------------------------------------
                Sql = "SELECT * from sbonif "
                Sql = Sql & " WHERE codartic=" & DBSet(Rs!codArtic, "T") '& " AND (" & RS!Cantidad & "<=sbonif.hastaca1)"
                Sql = Sql & " order by hastaca1 desc "
                Set RSaux = New ADODB.Recordset
                RSaux.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                Bonif1 = 0
                Bonif2 = 0
                If Not RSaux.EOF Then
'                    If (1 <= Rs!Cantidad) And (Rs!Cantidad <= RSaux!hastaca1) Then
                    If (0 < Rs!Cantidad) And (Rs!Cantidad <= RSaux!hastaca1) Then
                        Bonif1 = RSaux!impboni1
                        Canti1 = Rs!Cantidad
                    End If
                    If (RSaux!hastaca1 < Rs!Cantidad) And (Rs!Cantidad <= RSaux!hastaca2) Then
                        Bonif1 = RSaux!impboni1
                        Canti1 = Rs!Cantidad
                        Bonif2 = RSaux!impboni2
                        Canti2 = Rs!Cantidad
                    End If
                    
                    
                    'comprobar si la bonificacion es excluyente
                    EsExcluyente = CBool(RSaux!excluyen)
                    
                    If EsExcluyente Then
                    'se aplicara una bonificacion u otra
                       If Bonif2 > 0 Then Bonif1 = 0
                    End If
                    
                    If Bonif1 > 0 Then
                        numlinea = numlinea + 1
                        SQLBonif = SQLBonif & "('" & codtipom & "'," & NumFactu & "," & DBSet(FecFactu, "F") & ",'',9999999," & numlinea & "," & AlmFactu & "," & DBSet(RSaux!codarti1, "T") & ","
                        cadIVA = "codigiva"
                        Sql = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", RSaux!codarti1, "T", cadIVA)
                        SQLBonif = SQLBonif & DBSet(Sql, "T") & ",''," & DBSet(Canti1, "N") & "," & DBSet(-Bonif1, "N") & ",0,0," & DBSet(0 - Round(Canti1 * Bonif1, 2), "N") & ",'M'),"
                        If IVABonif1 = 0 Or IVABonif1 = CByte(cadIVA) Then
                            IVABonif1 = CByte(cadIVA)
                            TotBonif1 = TotBonif1 + Round(Canti1 * Bonif1, 2)
                        ElseIf IVABonif2 = 0 Or IVABonif2 = CByte(cadIVA) Then
                            IVABonif2 = CByte(cadIVA)
                            TotBonif2 = TotBonif2 + Round((Canti1 * Bonif1), 2)
                        ElseIf IVABonif3 = 0 Or IVABonif3 = CByte(cadIVA) Then
                            IVABonif3 = CByte(cadIVA)
                            TotBonif3 = TotBonif3 + Round(Canti1 * Bonif1, 2)
                        End If
                    End If
                    
                    If Bonif2 > 0 Then
'                        SQL = "(codtipom='" & codTipoM & "' and numfactu=" & NumFactu & " and fecfactu=" & DBSet(FecFactu, "F") & " and codtipoa='' and numalbar=0)"
                        numlinea = numlinea + 1
                        SQLBonif = SQLBonif & "('" & codtipom & "'," & NumFactu & "," & DBSet(FecFactu, "F") & ",'',9999999," & numlinea & "," & AlmFactu & "," & DBSet(RSaux!codarti2, "T") & ","
                        Sql = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", RSaux!codarti2, "T")
                        SQLBonif = SQLBonif & DBSet(Sql, "T") & ",''," & DBSet(Canti2, "N") & "," & DBSet(-Bonif2, "N") & ",0,0," & DBSet(0 - Round(Canti2 * Bonif2, 2), "N") & ",'M'),"
                        If IVABonif1 = 0 Or IVABonif1 = CByte(cadIVA) Then
                            IVABonif1 = CByte(cadIVA)
                            TotBonif1 = TotBonif1 + Round((Canti2 * Bonif2), 2)
                        ElseIf IVABonif2 = 0 Or IVABonif2 = CByte(cadIVA) Then
                            IVABonif2 = CByte(cadIVA)
                            TotBonif2 = TotBonif2 + Round((Canti2 * Bonif2), 2)
                        ElseIf IVABonif3 = 0 Or IVABonif3 = CByte(cadIVA) Then
                            IVABonif3 = CByte(cadIVA)
                            TotBonif3 = TotBonif3 + Round((Canti2 * Bonif2), 2)
                        End If
                    End If
'                    RSaux.MoveNext
                End If
                RSaux.Close
                Set RSaux = Nothing
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            If SQLBonif <> "" Then 'quitamos ultima coma
                SQLBonif = Mid(SQLBonif, 1, Len(SQLBonif) - 1)
            End If
        End If
    End If
    Set vClien = Nothing
    CalcularBonificacion = True
    
    
EBonif:
    If Err.Number <> 0 Then
        CalcularBonificacion = False
    Else
        CalcularBonificacion = True
    End If
End Function



Private Function InsertarBonificaciones() As Boolean
'Si los albaranes a facturar tienen lineas de articulos con bonificacion
'añadirlos a la factura
'Insertar lineas en tabla slifac con el importe de las bonificaciones
Dim Sql As String
Dim b As Boolean

    On Error GoTo EBonifica

    InsertarBonificaciones = False
    
    If SQLBonif > "" Then
        'Insertar en scafac1 un albaran a 0, para la bonificacion
        'si no se inserto antes para los desplazamientos
        '-----------------------------------------------------------
        Sql = "(codtipom='" & codtipom & "' and numfactu=" & NumFactu & " and fecfactu=" & DBSet(FecFactu, "F") & " and codtipoa='' and numalbar=9999999)"
        Sql = "Select count(*) from scafac1 where " & Sql
        If RegistrosAListar(Sql) = 0 Then
            b = InsertarCabAlbaranesFactu("")
            If Not b Then Exit Function
        End If
               
        'Insertar lineas de Desplazamiento en tabla: slifac (lineas factura)
        '------------------------------------------------------------------
        Sql = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
        Sql = Sql & "cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
        Sql = Sql & " VALUES " & SQLBonif
        conn.Execute Sql
        
    End If
    InsertarBonificaciones = True
     
EBonifica:
    If Err.Number <> 0 Then
        InsertarBonificaciones = False
    Else
        InsertarBonificaciones = True
    End If
End Function




Private Function ActualizarNumSeries(cadSQL As String) As Boolean
'Actualiza el campo "numfactu" de la tabla de num. serie: sserie
'Para ponerle la factura asociada
Dim Sql As String

    On Error GoTo EActualizar

    Sql = "UPDATE sserie SET numfactu= " & NumFactu
    Sql = Sql & " WHERE " & Replace(cadSQL, "scaalb", "sserie")
    conn.Execute Sql

    ActualizarNumSeries = True
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarNumSeries = False
    Else
        ActualizarNumSeries = True
    End If
End Function



Private Function CalcularDatosFacturaTicket(cadSQL As String) As Boolean
Dim b As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim I As Byte
Dim ImpTotal As Currency
Dim ImpDif As Currency

    On Error GoTo ETicket
    
'    'Bruto de la factura sera la suma de cantidad*precioar de las linea
'    'importe sin iva
'    SQL = "SELECT sum(cantidad*precioar) FROM sliven WHERE "
'    SQL = SQL & Replace(cadSQL, "scaven", "sliven")
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        BrutoFac = Round(RS.Fields(0).Value, 2)
'        BaseImp = BrutoFac
'    End If
'    RS.Close
'    Set RS = Nothing
    
    
    
    'desglosar las bases imponibles por tipo de IVA
'    SQL = "SELECT codigiva,sum(cantidad*precioar) as baseimp from sliven "
    '01/09/06
    Sql = "SELECT codigiva,sum(importel) as totalimp from sliven "
    
    Sql = Sql & " WHERE " & Replace(cadSQL, "scaven", "sliven")
    Sql = Sql & " group by codigiva"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    While Not Rs.EOF And I <= 3
        Sql = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(Rs!codigiva), "N")
        Select Case I
            Case 1
                TipoIVA1 = Rs.Fields(0).Value
                PorceIVA1 = CCur(Sql)
                BaseIVA1 = Round2(Rs.Fields(1).Value / (1 + (PorceIVA1 / 100)), 2)
'                SQL = CalcularDto(CStr(BaseIVA1), CStr(PorceIVA1))
'                ImpIVA1 = Round(CCur(SQL), 2)
                ImpIVA1 = Rs.Fields(1).Value - BaseIVA1
            Case 2
                TipoIVA2 = Rs.Fields(0).Value
                PorceIVA2 = CCur(Sql)
'                BaseIVA2 = RS.Fields(1).Value
'                SQL = CalcularDto(CStr(BaseIVA2), CStr(PorceIVA2))
'                ImpIVA2 = CCur(SQL)
                BaseIVA2 = Round2(Rs.Fields(1).Value / (1 + (PorceIVA2 / 100)), 2)
                ImpIVA2 = Rs.Fields(1).Value - BaseIVA2
            Case 3
                TipoIVA3 = Rs.Fields(0).Value
                PorceIVA3 = CCur(Sql)
'                BaseIVA3 = RS.Fields(1).Value
'                SQL = CalcularDto(CStr(BaseIVA3), CStr(PorceIVA3))
'                ImpIVA3 = CCur(SQL)
                BaseIVA3 = Round2(Rs.Fields(1).Value / (1 + (PorceIVA3 / 100)), 2)
                ImpIVA3 = Rs.Fields(1).Value - BaseIVA3
        End Select
        Rs.MoveNext
        I = I + 1
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    
    
    'comprobar que suma bases imp + imporiva = total factura
    ImpTotal = (BaseIVA1 + ImpIVA1) + (BaseIVA2 + ImpIVA2) + (BaseIVA3 + ImpIVA3)
    BrutoFac = BaseIVA1 + BaseIVA2 + BaseIVA3
    BaseImp = BrutoFac
    
    If Not (ImpTotal = TotalFac) Then
        ImpDif = TotalFac - ImpTotal
        ImpIVA1 = ImpIVA1 + ImpDif
'        If BaseIVA2 = 0 And BaseIVA3 = 0 Then ImpIVA1 = ImpIVA1 + ImpDif
'        ImpTotal = (BaseIVA1 + ImpIVA1) + (BaseIVA2 + ImpIVA2) + (BaseIVA3 + ImpIVA3)
    End If
    
    b = True
    
    
ETicket:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Calculando datos de ticket a factura.", Err.Description
        b = False
    End If
    CalcularDatosFacturaTicket = b
End Function


Private Function InsertarFormaPagoEnConta(nForPa As String, cadErr As String) As Boolean
Dim cadAux As String
Dim cadAux2 As String
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo ErrInsForpa
    InsertarFormaPagoEnConta = False
    
    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
    'forma de pago de la factura. Sino existe insertarla
    
    'vemos si existe en la conta
    If vParamAplic.ContabilidadNueva Then
        cadAux = DevuelveDesdeBDNew(conConta, "formapago", "codforpa", "codforpa", nForPa, "N")
    Else
        cadAux = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", nForPa, "N")
    End If
    'si no existe la forma de pago en conta, insertamos la de aritaxi
    If cadAux = "" Then
        Sql = "select * from sforpa where codforpa = " & DBSet(nForPa, "N")
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            'insertamos e sforpa de la CONTA
            If vParamAplic.ContabilidadNueva Then
                Sql = "INSERT INTO formapago(codforpa,nomforpa,tipforpa,numerove,primerve,restoven)"
            Else
                Sql = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
            End If
            Sql = Sql & " VALUES(" & nForPa & ", " & DBSet(Rs!nomforpa, "T") & ", " & DBSet(Rs!tipforpa, "N")
            If vParamAplic.ContabilidadNueva Then
                Sql = Sql & "," & DBSet(Rs!numerove, "N") & "," & DBSet(Rs!primerve, "N") & "," & DBSet(Rs!restoven, "N") & ")"
            Else
                Sql = Sql & ")"
            End If
            ConnConta.Execute Sql
            InsertarFormaPagoEnConta = True
        Else
            InsertarFormaPagoEnConta = False
        End If
        Rs.Close
        Set Rs = Nothing
        
    Else
        InsertarFormaPagoEnConta = True
    End If
    
    Exit Function
    
ErrInsForpa:
    InsertarFormaPagoEnConta = False
    cadErr = "Insertar forma de pago en Contablilidad: " & vbCrLf & Err.Description
End Function



'#### Laura 15/11/2006 Recuperar facturas ALZIRA
Public Function MarcarContabilizada(Optional cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    MarcarContabilizada = False
    
    Sql = "UPDATE scafac SET intconta=1 "
    Sql = Sql & " WHERE codtipom=" & DBSet(Me.codtipom, "T") & " AND numfactu=" & Me.NumFactu & " AND fecfactu=" & DBSet(Me.FecFactu, "F")
    conn.Execute Sql
    
    MarcarContabilizada = True
    Exit Function
    
EActualizar:
    MarcarContabilizada = False
    cadErr = Err.Description
End Function






Public Function InsertarEnTesoreriaCuotasSoc(MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: aritaxi.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String, codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAuxConta As String 'para insertar en conta.scobro
Dim CadValues3 As String
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single 'importe para insertar en la svenci
Dim ImpVenci2 As Single 'importe para insertar en conta.scobro
Dim I As Byte
Dim TotalFactura2 As Currency   'Por si acaso lleva aportacion al terminal
'1 Julio 2009. Los graba en scobro
Dim CadenaDatosFiscales As String

Dim CADENA As String
Dim vTextosCSB As String

Dim vvIban As String


    On Error GoTo EInsertarTesoreriaCuotasSoc

    b = False
    InsertarEnTesoreriaCuotasSoc = False

    vTextosCSB = ""
    If vParamAplic.ContabilidadNueva Then
        CADENA = mvarText41csb & " " & mvarText42csb & " " & mvarText43csb & " " & mvarText51csb & " " & mvarText52csb & " " & _
                mvarText53csb & " " & mvarText61csb & " " & mvarText62csb & " " & mvarText63csb & " " & mvarText71csb & " " & _
                mvarText72csb & " " & mvarText73csb & " " & mvarText81csb & " " & mvarText82csb & " " & mvarText83csb
        vTextosCSB = DBSet(CADENA, "T", "S")
    Else
        vTextosCSB = vTextosCSB & DBSet(mvarText41csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText42csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText43csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText51csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText52csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText53csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText61csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText62csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText63csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText71csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText72csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText73csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText81csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText82csb, "T", "S") & ","
        vTextosCSB = vTextosCSB & DBSet(mvarText83csb, "T", "S")
    End If
'    If vTextosCSB = "" Then
'        vTextosCSB = "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL"
'    Else
'        CadValues = DBSet(RecuperaValor(vTextosCSB, 1), "T", "S")
'        CadValues2 = DBSet(RecuperaValor(vTextosCSB, 2), "T", "S")
'        CadValues3 = DBSet(RecuperaValor(vTextosCSB, 3), "T", "S")
'
'        vTextosCSB = CadValues & "," & CadValues2 & "," & CadValues3
'    End If
    CadValues3 = ""
    CadValues = ""
    CadValues2 = ""

'    'campo para insertar en conta.scobro de Tesoreria
'    textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd mmm yyyy") & "'"


    'Datos fiscales en scobro     Julio 2009
    'nomclien,domclien,pobclien, cpclien,proclien
    CadenaDatosFiscales = DBSet(mNomClien, "T") & "," & DBSet(mDomClien, "T") & "," & DBSet(mPobClien, "T")
    CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(mCodpobla, "T") & "," & DBSet(mProClien, "T")
    
    If vParamAplic.ContabilidadNueva Then
        CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(mNIFClien, "T") & ",'ES'"
    End If
    
    
    'Obtener el Nº de Vencimientos de la forma de pago
    Sql = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 And CCur(TotalFac) <> 0 Then

            'Comporbamos si el importe es <>0
        
            'Obtener los dias de pago del cliente : el socio no tiene dias de pago
            Sql = " SELECT  0 diapago1, 0 diapago2, 0 diapago3, 0 mesnogir, 0 diavtoat, '' codmacta "
            Sql = Sql & " FROM sclien "
            Sql = Sql & " WHERE codclien=" & Cliente
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
'            codmacta = DBSet(RS!codmacta, "T")

            CADENA = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
            codmacta = vParamAplic.Raiz_CtaClien_Soc & Format(Cliente, CADENA)

'            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
            
            If Not Rs.EOF Then
                cadValuesAux = "('" & codtipom & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                CadValuesAuxConta = "('" & LetraSerie & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                '                    Añadire a la cadena fija esta los valores de textcsb41,txcs
                CadValuesAuxConta = CadValuesAuxConta & vTextosCSB & ","
                '-------- Primer Vencimiento
                I = 1
                'FECHA VTO
                FecVenci = CDate(FecFactu)
                '=== Laura 23/01/2007
                'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                '===
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                If TipForPago <> 0 Then
                    FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                Else
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                End If
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(Rs!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAuxConta & I & ", "
                CadValues2 = CadValues2 & DBSet(codmacta, "T") & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                TotalFactura2 = TotalFac - Aportacion
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFactura2
                    ImpVenci2 = TotalFactura2 - ImpCheque
                Else
                    
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura2 - ImpCheque) / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFactura2 Then
                        ImpVenci = Round(ImpVenci + (TotalFactura2 - ImpVenci * rsVenci!numerove), 2)
                    End If
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If (ImpVenci2 * rsVenci!numerove) + ImpCheque <> TotalFactura2 Then
                        ImpVenci2 = Round(ImpVenci2 + (TotalFactura2 - ImpCheque - (ImpVenci2 * rsVenci!numerove)), 2)
                    End If
                End If
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                If vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', '" & mvarText33csb & "', " & DBSet(Agente, "N")
                Else
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", '" & mvarText33csb & "', " & DBSet(Agente, "N")
                End If
                'departamento y transfer
                CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                '1 Julio 2009
                ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                 
                If vParamAplic.ContabilidadNueva Then
                    vvIban = MiFormat(Iban, "") & MiFormat(Banco, "0000") & MiFormat(Sucursal, "0000") & MiFormat(DigControl, "00") & MiFormat(CuentaBan, "0000000000")
                
                    CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                Else
                     '[Monica]22/11/2013: tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                       CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ") "
                    Else
                       CadValues2 = CadValues2 & ") "
                    End If
                End If
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For I = 2 To rsVenci!numerove
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(Rs!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    End If
                        
                    CadValues = CadValues & ", " & cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & ", " & codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura2 - ImpCheque) / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", '" & mvarText33csb & "', " & DBSet(Agente, "N") & ", "
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", '" & mvarText33csb & "', " & DBSet(Agente, "N") & ", "
                    End If
                    CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                    '1 Julio 2009
                    ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                    CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                    
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(Iban, "") & MiFormat(Banco, "0000") & MiFormat(Sucursal, "0000") & MiFormat(DigControl, "00") & MiFormat(CuentaBan, "0000000000")
                    
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                    Else
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                        Else
                           CadValues2 = CadValues2 & ")"
                        End If
                    End If
                Next I
                
                '--- Cheque regalo: laura 1/12/2006   y/o    Aportacion terminal
                'si hay cheque regalo insertar una linea más para la forma de pago correspondiente y el importe del cheque
                If ImpCheque > 0 Then
                
                    
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & "," & DBSet(codmacta, "T") & ", " & vParamAplic.ForPagoChequeRegalo & ", "
                    
                    'FECHA VTO
                    CadValuesAuxConta = "primerve"
                    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", vParamAplic.ForPagoChequeRegalo, "N", CadValuesAuxConta)
                    FecVenci = CDate(FecFactu)
                    FecVenci = FecVenci + CByte(CadValuesAuxConta)
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                        MsgBox "FALTA cheque regalo con forma de pago no en EFECTIVO", vbInformation
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    
                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(ImpCheque, "N") & ", '" & CuentaPrev & "',  '" & mvarText33csb & "', " & DBSet(Agente, "N")
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpCheque, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", '" & mvarText33csb & "', " & DBSet(Agente, "N")
                    End If
                    'departamento
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL" ')"
                    CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(Iban, "") & MiFormat(Banco, "0000") & MiFormat(Sucursal, "0000") & MiFormat(DigControl, "00") & MiFormat(CuentaBan, "0000000000")
                    
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                    
                    Else
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                        Else
                           CadValues2 = CadValues2 & ")"
                        End If
                    End If
                End If
                
                
                'Aportacion al terminal
                'Si tiene cuenta aportacion entonces añadiremos un cobro
                If Me.Aportacion > 0 Then
                    If vParamAplic.ctaAportacion = "" Then
                        MsgBox "Error cta aportacion NULL", vbExclamation
                        Exit Function
                    End If
                        
                   'Montamos EL SQL para el cobro de la aportacion al termina
                    'ForPago
                    
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & ",'" & vParamAplic.ctaAportacion & "', " & ForPago & ", "
                    
                    
                    CadValuesAuxConta = "primerve"
                    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N", CadValuesAuxConta)
                    FecVenci = CDate(FecFactu)
                    FecVenci = FecVenci + CByte(CadValuesAuxConta)
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                        MsgBox "FALTA config con forma de pago no en EFECTIVO", vbInformation
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    
                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(Aportacion, "N") & ", '" & CuentaPrev & "', '" & mvarText33csb & "', " & DBSet(Agente, "N")
                    Else
                        CadValues2 = CadValues2 & DBSet(Aportacion, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", '" & mvarText33csb & "', " & DBSet(Agente, "N")
                    End If
                    'departamento
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL" ')"
                    CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(Iban, "") & MiFormat(Banco, "0000") & MiFormat(Sucursal, "0000") & MiFormat(DigControl, "00") & MiFormat(CuentaBan, "0000000000")
                    
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                    Else
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ")"
                        Else
                           CadValues2 = CadValues2 & ")"
                        End If
                    End If
                End If
                
            End If
            Rs.Close
            
        Else
            'totalfac =0 and numerovtos >=1
            b = True
        End If
        
        Set Rs = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        Sql = "INSERT INTO svenci (codtipom, numfactu, fecfactu, ordefect, fecefect, impefect)"
        Sql = Sql & " VALUES " & CadValues
        conn.Execute Sql
    End If
    
    
    'Grabar tabla scobro de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        '01/09/06
'        If (NumTicket = "") Or (NumTicket <> "" And TipForPago <> 0) Then
            If CuentaPrev <> "" Then
                'antes de grabar en la scobro comprobar que existe en conta.sforpa la
                'forma de pago de la factura. Sino existe insertarla
                'vemos si existe en la conta
                b = InsertarFormaPagoEnConta(ForPago, MenError)
                
                'si hay cheque regalo comprobar q existe su forma de pago en conta
                'sino insertarla
                If b And ImpCheque > 0 Then b = InsertarFormaPagoEnConta(vParamAplic.ForPagoChequeRegalo, MenError)
                
                
                
                
                If b Then
                    If vParamAplic.ContabilidadNueva Then
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu,text41csb, "
                        Sql = Sql & "numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, "
                        Sql = Sql & "text33csb,agente,departamento,transfer "
                        'JULIO 2009
                        'nomclien,domclien,pobclien, cpclien,proclien
                        Sql = Sql & ",nomclien,domclien,pobclien, cpclien,proclien,nifclien,codpais,iban)" ')"
                    Else
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl,text41csb ,text42csb, text43csb, "
                        Sql = Sql & "text51csb ,text52csb, text53csb, text61csb ,text62csb, text63csb,"
                        Sql = Sql & "text71csb ,text72csb, text73csb, text81csb ,text82csb, text83csb, "
                        Sql = Sql & "numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur,"
                        Sql = Sql & "digcontr, cuentaba,text33csb,agente,departamento,transfer "
                        'JULIO 2009
                        'nomclien,domclien,pobclien, cpclien,proclien
                        Sql = Sql & ",nomclien,domclien,pobclien, cpclien,proclien" ')"
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           Sql = Sql & ",iban)"
                        Else
                           Sql = Sql & ")"
                        End If
                    End If
                    Sql = Sql & " VALUES " & CadValues2
                    ConnConta.Execute Sql
                End If
                
            Else
                'DAVID ####
                'ENERO 2008
                'Si no inserto en tesoreria por que ctaprvsta="" entonces dejo continuar
                b = True
        
        End If

    End If
    
EInsertarTesoreriaCuotasSoc:
    If Err.Number <> 0 Then
        b = False
        MenError = "Insertar en Tesoreria: " & vbCrLf & Err.Description
    End If
    InsertarEnTesoreriaCuotasSoc = b
End Function

