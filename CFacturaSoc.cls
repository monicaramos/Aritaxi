VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFacturaSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'COMPRAS
'FACTURA: se corresponde con la tabla aritaxi.sfactusoc (cabecera de facturas de Socios)
'-------------------------------------------------------------------------

'Variables locales que contienen valores de propiedad
Private mTipoMov As String ' codigo tipo de movimiento
Private mCodSocio As String 'Cod. socio
Private mNumFactu As Long 'Nº Factura
Private mFecFactu As String 'FEcha Factura
Private mFecRecep As String 'FEcha Recepcion

Private mCodmacta As String 'cuenta del proveedor para insertar in spagop

Private mNomSocio As String 'nombre del Socio
Private mDomSocio As String 'domicilio
Private mCodpobla As String 'cpostal
Private mPobSocio As String 'poblacion
Private mProSocio As String 'provincia
Private mNIFSocio As String 'NIF proveedor
Private mTelSocio As String 'telefono

Private mCodForpa As Integer 'Cod. Forma de pago
Private mCodTraba As String 'Cod. trabajador factura

Private mBrutofac As Currency 'total importe bruto de las lineas de la factura(sin dtos ni IVA)
Private mDtoPpago As Currency 'descuento pronto pago
Private mDtoGnral As Currency 'descuento general
Private mImpPPago As Currency 'Importe al aplicar al bruto el dto pronto pago
Private mImpGnral As Currency  'Importe al aplicar al bruto el dto general
Private mBaseImp As Currency  'Base imponible de la factura (bruto - dtopp - dtogn)

Private mBaseiva1 As Currency  'base imponible tipo IVA 1
Private mTipoiva1 As Byte   'Cod. tipo de iva 1
Private mPorciva1 As Currency    '% de iva 1

Private mPorcrec1 As Currency   '% de recargo 1

Private mImpoiva1 As Currency   'Importe de iva 1
Private mImporec1 As Currency   'Importe de recargo 1

Private mTotalfac As Currency  'total factura (con dtos y con IVA)

Private mBanPr As String 'Banco propio que ponemos como cuenta prevista de pago
Private mCtaPrev As String 'Cuenta prevista de pago


Private mConcepto As String 'Concepto de publicidad


'DAVID###
'[Monica]22/11/2013: iban
Private mCCC_Iban As String

Private mCCC_Entidad As String
Private mCCC_Oficina As String
Private mCCC_CC As String
Private mCCC_CTa As String

'retencion ----------------------------------------------------
Private mTipoRet As Byte   '0.- Retencion NORMAL
                           '1.- Reteccion coooperativas agricolas
Private mPorRet As Currency
Private mImpRet As Currency

'------------------------------------------------
'Propiedades del modulo CFacturaSoc
'------------------------------------------------

'**** Tipo de Movimiento de la Factura

Public Property Let tipoMov(ByVal vData As String)
     mTipoMov = vData
End Property

Public Property Get tipoMov() As String
     tipoMov = mTipoMov
End Property



'**** Socio de la Factura

Public Property Let Socio(ByVal vData As String)
     mCodSocio = vData
End Property

Public Property Get Socio() As String
     Socio = mCodSocio
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


'**** Fecha Recepcion de la Factura

Public Property Let FecRecep(ByVal vData As String)
     mFecRecep = vData
End Property

Public Property Get FecRecep() As String
     FecRecep = mFecRecep
End Property


'**** cta contable (codmacta) del Socio de la Factura

Public Property Let CtaSocio(ByVal vData As String)
     mCodmacta = vData
End Property

Public Property Get CtaSocio() As String
     CtaSocio = mCodmacta
End Property


'**** Nombre del Proveedor de la Factura

Public Property Let NombreSocio(ByVal vData As String)
     mNomSocio = vData
End Property

Public Property Get NombreSocio() As String
     NombreSocio = mNomSocio
End Property

'**** Domicilio del Proveedor de la Factura

Public Property Let DomicilioSocio(ByVal vData As String)
     mDomSocio = vData
End Property

Public Property Get DomicilioSocio() As String
     DomicilioSocio = mDomSocio
End Property


'**** CPostal del Proveedor de la Factura

Public Property Let CPostalSocio(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostalSocio() As String
     CPostalSocio = mCodpobla
End Property


'**** Poblacion del Proveedor de la Factura

Public Property Let PoblacionSocio(ByVal vData As String)
     mPobSocio = vData
End Property

Public Property Get PoblacionSocio() As String
     PoblacionSocio = mPobSocio
End Property

'**** Provincia del Proveedor de la Factura

Public Property Let ProvinciaSocio(ByVal vData As String)
     mProSocio = vData
End Property

Public Property Get ProvinciaSocio() As String
     ProvinciaSocio = mProSocio
End Property


'**** NIF del Proveedor de la Factura

Public Property Let nifSocio(ByVal vData As String)
     mNIFSocio = vData
End Property

Public Property Get nifSocio() As String
     nifSocio = mNIFSocio
End Property


'**** Telefono del Proveedor de la Factura

Public Property Let TelefSocio(ByVal vData As String)
     mTelSocio = vData
End Property

Public Property Get TelefSocio() As String
     TelefSocio = mTelSocio
End Property


'**** Trabajador de la Factura

Public Property Let Trabajador(ByVal vData As String)
     mCodTraba = vData
End Property

Public Property Get Trabajador() As String
     Trabajador = mCodTraba
End Property


'**** Forma Pago de la Factura

Public Property Let ForPago(ByVal vData As Integer)
     mCodForpa = vData
End Property

Public Property Get ForPago() As Integer
     ForPago = mCodForpa
End Property


'**** Importe Bruto Factura

Public Property Let BrutoFac(ByVal vData As Currency)
     mBrutofac = vData
End Property

Public Property Get BrutoFac() As Currency
     BrutoFac = mBrutofac
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

'**** Importe pronto pago de la Factura

Public Property Let ImpPPago(ByVal vData As Currency)
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

'**** Tipo iva 1 de la Factura

Public Property Let TipoIVA1(ByVal vData As Byte)
     mTipoiva1 = vData
End Property

Public Property Get TipoIVA1() As Byte
     TipoIVA1 = mTipoiva1
End Property

'**** % de iva 1 de la Factura

Public Property Let PorceIVA1(ByVal vData As Currency)
     mPorciva1 = vData
End Property

Public Property Get PorceIVA1() As Currency
     PorceIVA1 = mPorciva1
End Property

'**** Importe de IVA 1 de la Factura

Public Property Let ImpIVA1(ByVal vData As Currency)
     mImpoiva1 = vData
End Property

Public Property Get ImpIVA1() As Currency
     ImpIVA1 = mImpoiva1
End Property

'**** Total Factura

Public Property Let TotalFac(ByVal vData As Currency)
     mTotalfac = vData
End Property

Public Property Get TotalFac() As Currency
     TotalFac = mTotalfac
End Property


'**** Banco propio de pago de la Factura

Public Property Let BancoPr(ByVal vData As String)
     mBanPr = vData
End Property

Public Property Get BancoPr() As String
     BancoPr = mBanPr
End Property

'**** Cuenta Prevista de pago de la Factura
Public Property Let CuentaPrev(ByVal vData As String)
     mCtaPrev = vData
End Property
Public Property Get CuentaPrev() As String
     CuentaPrev = mCtaPrev
End Property

'**** Cuenta banacaria: IBAN
Public Property Let CCC_Iban(ByVal vData As String)
     mCCC_Iban = vData
End Property
Public Property Get CCC_Iban() As String
     CCC_Iban = mCCC_Iban
End Property




'**** Cuenta banacaria: ENTIDAD
Public Property Let CCC_Entidad(ByVal vData As String)
     mCCC_Entidad = vData
End Property
Public Property Get CCC_Entidad() As String
     CCC_Entidad = mCCC_Entidad
End Property

'**** Cuenta banacaria: OFICINA
Public Property Let CCC_Oficina(ByVal vData As String)
     mCCC_Oficina = vData
End Property
Public Property Get CCC_Oficina() As String
     CCC_Oficina = mCCC_Oficina
End Property

'**** Cuenta banacaria: Codigo control
Public Property Let CCC_CC(ByVal vData As String)
     mCCC_CC = vData
End Property
Public Property Get CCC_CC() As String
     CCC_CC = mCCC_CC
End Property

'Cuenta banacaria.  Cuenta
Public Property Let CCC_CTa(ByVal vData As String)
     mCCC_CTa = vData
End Property
Public Property Get CCC_CTa() As String
     CCC_CTa = mCCC_CTa
End Property


'Porcentaje y total retencion PARA liquidacion proveedores
'es decir, socios tipo cooperativas agrarias
Public Property Let PorRet(ByVal vData As Currency)
     mPorRet = vData
End Property

Public Property Get PorRet() As Currency
     PorRet = mPorRet
End Property

Public Property Let ImpRet2(ByVal vData As Currency)
     mImpRet = vData
End Property

Public Property Get ImpRet2() As Currency
     ImpRet2 = mImpRet
End Property


Public Property Let TipoRet(ByVal vData As Byte)
     mTipoRet = vData
End Property

Public Property Get TipoRet() As Byte
     TipoRet = mTipoRet
End Property


'**** Concepto de la Factura

Public Property Let Concepto(ByVal vData As String)
     mConcepto = vData
End Property

Public Property Get Concepto() As String
     Concepto = mConcepto
End Property


'------------------------------------------------
'Procedimientos del modulo CFacturaSoc
'------------------------------------------------
' Facturas de publicidad
Public Function LeerDatos3(vCodsocio As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla sfactusoc de publicidad
'Lee de la BD: Aritaxi, Tabla: sfactusoc
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim Mens As String
Dim Sql As String
Dim CADENA As String

    On Error GoTo ELeer
    
    LeerDatos3 = False
    Sql = "SELECT codSocio,numfactu,fecfactu,codforpa,baseiva1,totalfac FROM sfactusoc "
    Sql = Sql & " WHERE codSocio=" & vCodsocio
    Sql = Sql & " AND numfactu=" & vNumfactu & ""
    Sql = Sql & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"
    Sql = Sql & " AND codtipom ='FPS'"

    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If RS.EOF Then
        
    Else
        Mens = "Datos Factura"
        mCodSocio = RS!codSocio
        mNumFactu = CLng(RS!NumFactu)
        mFecFactu = CStr(RS!FecFactu)
        mCodForpa = CStr(RS!codforpa)
        mBrutofac = CStr(RS!BaseIVA1)
        mTotalfac = CStr(RS!TotalFac)
        'obterner la ctaCble del proveedor
        'CtaSocio = DevuelveDesdeBDNew(conAri, "sclien", "codmacta", "codclien", vCodSocio, "N")
        CADENA = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
        CtaSocio = vParamAplic.Raiz_Cta_Soc_publi & Format(mCodSocio, CADENA)
        
        
        'si venimos de proceso de insertar tenemos la cuenta prevista de pago que
        'se introdujo en el formulario de pasar albaranes a factura
        'si no venimos de moficicar y la leemos de spagos antes de actualizar
        If CuentaPrev = "" Then
            'leer la cuenta prevista de pagos de la tabla de contabilidad: conta.spagos
            'cuando vamos a borrar una factura porque modificarmos
            If vParamAplic.ContabilidadNueva Then
                Sql = DevuelveDesdeBDNew(conConta, "pagos", "ctabanc1", "codmacta", CtaSocio, "T", , "numfactu", vNumfactu, "N", "fecfactu", vFecfactu, "F")
            Else
                Sql = DevuelveDesdeBDNew(conConta, "spagop", "ctabanc1", "ctaprove", CtaSocio, "T", , "numfactu", vNumfactu, "N", "fecfactu", vFecfactu, "F")
            End If
            If Sql <> "" Then
                CuentaPrev = Sql
                LeerDatos3 = True
            Else
                Mens = "La cuenta prevista de pago para la factura no puede ser nula. "
                If vUsu.Nivel <= 1 Then
                    Mens = Mens & vbCrLf & "¿Continuar?"
                    If MsgBox(Mens, vbQuestion + vbYesNoCancel) = vbYes Then
                        CuentaPrev = "''"
                        LeerDatos3 = True
                    End If
                Else
                    MsgBox Mens, vbExclamation
                End If
            End If
        Else
            LeerDatos3 = True
        End If
        

    End If
    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    If Not RS Is Nothing Then Set RS = Nothing
    LeerDatos3 = False
End Function


' Facturas de liquidacion de socios
Public Function LeerDatosLiq(vCodsocio As String, vCodtipom As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla sfactusoc de publicidad
'Lee de la BD: Aritaxi, Tabla: sfactusoc
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim Mens As String
Dim Sql As String
Dim SQL2 As String
Dim Rs2 As ADODB.Recordset
Dim CADENA As String

    On Error GoTo ELeer
    
    LeerDatosLiq = False
    Sql = "SELECT codSocio,numfactu,fecfactu,codforpa,baseiva1,totalfac FROM sfactusoc "
    Sql = Sql & " WHERE codSocio=" & vCodsocio
    Sql = Sql & " AND numfactu=" & vNumfactu & ""
    Sql = Sql & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"
    Sql = Sql & " AND codtipom = " & DBSet(vCodtipom, "T")

    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If RS.EOF Then
        
    Else
        Mens = "Datos Factura"
        mCodSocio = RS!codSocio
        mNumFactu = CLng(RS!NumFactu)
        mFecFactu = CStr(RS!FecFactu)
        mCodForpa = CStr(RS!codforpa)
        mBrutofac = CStr(RS!BaseIVA1)
        mTotalfac = CStr(RS!TotalFac)
        mTipoMov = CStr(vCodtipom)
        
        SQL2 = "select iban, codbanco, codsucur, digcontr, cuentaba from sclien where codclien = " & DBSet(RS!codSocio, "N")
        Set Rs2 = New ADODB.Recordset
        Rs2.Open SQL2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not Rs2.EOF Then
            mCCC_Iban = DBLet(Rs2!Iban, "T")
            mCCC_Entidad = CStr(DBLet(Rs2!codbanco, "N"))
            mCCC_Oficina = CStr(DBLet(Rs2!codsucur, "N"))
            mCCC_CC = DBLet(Rs2!digcontr, "T")
            mCCC_CTa = DBLet(Rs2!cuentaba, "T")
        End If
        
        'obterner la ctaCble del proveedor
        CADENA = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
        CtaSocio = vParamAplic.Raiz_Cta_Soc_Liqui & Format(mCodSocio, CADENA)
        
        'si venimos de proceso de insertar tenemos la cuenta prevista de pago que
        'se introdujo en el formulario de pasar albaranes a factura
        'si no venimos de moficicar y la leemos de spagos antes de actualizar
        If CuentaPrev = "" Then
            'leer la cuenta prevista de pagos de la tabla de contabilidad: conta.spagos
            'cuando vamos a borrar una factura porque modificarmos
            
'[Monica]10/07/2012: tanto liquidaciones como rectificativas han de estar en la spagop
'            If vCodtipom = "FRL" Then
'                Sql = DevuelveDesdeBDNew(conConta, "scobro", "ctabanc1", "codmacta", CtaSocio, "T", , "codfaccl", CLng(vNumfactu), "N", "fecfaccl", vFecfactu, "F")
'            Else
        
            If vParamAplic.ContabilidadNueva Then
                Sql = DevuelveDesdeBDNew(conConta, "pagos", "ctabanc1", "ctaprove", CtaSocio, "T", , "numfactu", ObtenerLetraSerie(vCodtipom) & CLng(vNumfactu), "T", "fecfactu", vFecfactu, "F")
            Else
                Sql = DevuelveDesdeBDNew(conConta, "spagop", "ctabanc1", "ctaprove", CtaSocio, "T", , "numfactu", ObtenerLetraSerie(vCodtipom) & CLng(vNumfactu), "T", "fecfactu", vFecfactu, "F")
            End If
'            End If
            If Sql <> "" Then
                CuentaPrev = Sql
                LeerDatosLiq = True
            Else
                Mens = "La cuenta prevista de cobro para la factura no puede ser nula. "
                If vUsu.Nivel <= 1 Then
                    Mens = Mens & vbCrLf & "¿Continuar?"
                    If MsgBox(Mens, vbQuestion + vbYesNoCancel) = vbYes Then
                        CuentaPrev = "''"
                        LeerDatosLiq = True
                    End If
                Else
                    MsgBox Mens, vbExclamation
                End If
            End If
        Else
            LeerDatosLiq = True
        End If
    End If
    
    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    If Not RS Is Nothing Then Set RS = Nothing
    LeerDatosLiq = False
End Function

' Desmarcar llamadas de una factura de liquidacion
Public Function DesmarcarLLamadas(vCodsocio As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla sfactusoc de publicidad
'Lee de la BD: Aritaxi, Tabla: sfactusoc
'OUT: True si lee los datos correctamente
Dim Mens As String
Dim Sql As String
Dim SQL2 As String
Dim CADENA As String

    On Error GoTo ELeer
    
    DesmarcarLLamadas = False
    
    Sql = "SELECT fecha, hora, numeruve FROM sfactusoc_serv "
    Sql = Sql & " WHERE codSocio=" & vCodsocio
    Sql = Sql & " AND numfactu=" & vNumfactu & ""
    Sql = Sql & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"
    Sql = Sql & " AND codtipom ='FLI'"

    SQL2 = "update shilla set liquidadosocio = 0, abonados = 0 where (fecha, hora, numeruve) in (" & Sql & ")"

    conn.Execute SQL2


    DesmarcarLLamadas = True
    
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    DesmarcarLLamadas = False
End Function





'======================================================================
'GRABAR EN TESORERIA
'======================================================================

Public Function InsertarEnTesoreria(MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: aritaxi.svenci y en conta.scobros
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single
Dim i As Byte
Dim vvIban As String

    On Error GoTo EInsertarTesoreria

'    b = False
    InsertarEnTesoreria = False
    CadValues2 = ""

    'Obtener el Nº de Vencimientos de la forma de pago
    Sql = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago de la tabla de parametros: spara1
            Sql = " SELECT  diapago1, diapago2, diapago3,mesnogir "
            Sql = Sql & " FROM spara1 "
            Sql = Sql & " WHERE codigo=1"
            Set RS = New ADODB.Recordset
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not RS.EOF Then
               'vamos creando la cadena para insertar en spagosp de la CONTA
               
               CadValuesAux2 = "("
               If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & DBSet(SerieFraPro, "T") & ","
               CadValuesAux2 = CadValuesAux2 & "'" & CtaSocio & "', " & DBSet(ObtenerLetraSerie(mTipoMov) & NumFactu, "T") & ", '" & Format(FecFactu, FormatoFecha) & "', "
              
              'Primer Vencimiento
              '------------------------------------------------------------
              i = 1
              'FECHA VTO
              FecVenci = CDate(FecFactu)
              '=== Modificado: Laura 23/01/2007
'              FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
              FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
              '==================================
              'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
              FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))

              'Comprobar si  tiene mes a no girar
              FecVenci1 = FecVenci
              If DBSet(RS!mesnogir, "N") <> 0 Then
                  FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(RS!mesnogir, "N"), DBSet(0, "N"), RS!DiaPago1, RS!DiaPago2, RS!DiaPago3)
              End If
             
              CadValues2 = CadValuesAux2 & i
              CadValues2 = CadValues2 & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
              'IMPORTE del Vencimiento
              If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFac
              Else
                    ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFac Then
                        ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci!numerove), 2)
                    End If
              End If
              
              CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CuentaPrev, "T") & ","
              
              If Not vParamAplic.ContabilidadNueva Then
                    'David. Para que ponga la cuenta bancaria (SI LA tiene)
                    CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                    CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & ","
              End If

              'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
              Sql = "Factura num.: " & ObtenerLetraSerie(mTipoMov) & NumFactu & "-" & Format(FecFactu, "dd/mm/yyyy")
              CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
              Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
              CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
              
              If vParamAplic.ContabilidadNueva Then
                  vvIban = MiFormat(mCCC_Iban, "") & MiFormat(mCCC_Entidad, "0000") & MiFormat(mCCC_Oficina, "0000") & MiFormat(mCCC_CC, "00") & MiFormat(mCCC_CTa, "0000000000")
                  
                  CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ","
              
'nomprove , domprove, pobprove, cpprove, proprove, nifProve, codpais
                  CadValues2 = CadValues2 & DBSet(mNomSocio, "T") & "," & DBSet(mDomSocio, "T") & "," & DBSet(mPobSocio, "T") & "," & DBSet(mCodpobla, "T") & ","
                  CadValues2 = CadValues2 & DBSet(mProSocio, "T") & "," & DBSet(mNIFSocio, "T") & ",ES)"
              
              
              Else
                    '[Monica]22/11/2013: tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & "," & DBSet(mCCC_Iban, "T", "S") & ")"
                    Else
                        CadValues2 = CadValues2 & ")"
                    End If
              End If

              'Resto Vencimientos
              '--------------------------------------------------------------------
              For i = 2 To rsVenci!numerove
                 'FECHA Resto Vencimientos
                  '==== Modificado: Laura 23/01/2007
                  'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                  FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                  '==================================================
                  'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                  FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))

                  'Comprobar si tiene mes a no girar
                  FecVenci1 = FecVenci
                  If DBSet(RS!mesnogir, "N") <> 0 Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(RS!mesnogir, "N"), DBSet(0, "N"), RS!DiaPago1, RS!DiaPago2, RS!DiaPago3)
                  End If

                  CadValues2 = CadValues2 & ", " & CadValuesAux2 & i & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "

                  'IMPORTE Resto de Vendimientos
                  ImpVenci = Round(TotalFac / rsVenci!numerove, 2)

                  CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "',"
                  
                  If Not vParamAplic.ContabilidadNueva Then
                        'David. Para que ponga la cuenta bancaria (SI LA tiene)
                        CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                        CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & ","
                  End If

                  Sql = "Factura num.: " & ObtenerLetraSerie(mTipoMov) & NumFactu & "-" & Format(FecFactu, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                  Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
                  
                  If vParamAplic.ContabilidadNueva Then
                      vvIban = MiFormat(mCCC_Iban, "") & MiFormat(mCCC_Entidad, "0000") & MiFormat(mCCC_Oficina, "0000") & MiFormat(mCCC_CC, "00") & MiFormat(mCCC_CTa, "0000000000")
                      
                      CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ","
                  
    'nomprove , domprove, pobprove, cpprove, proprove, nifProve, codpais
                      CadValues2 = CadValues2 & DBSet(mNomProve, "T") & "," & DBSet(mDomProve, "T") & "," & DBSet(mPobprove, "T") & "," & DBSet(mCodpobla, "T") & ","
                      CadValues2 = CadValues2 & DBSet(mProprove, "T") & "," & DBSet(mNifProve, "T") & ",ES)"
                  
                  
                  Else
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & "," & DBSet(mCCC_Iban, "T", "S") & ")"
                        Else
                            CadValues2 = CadValues2 & ")"
                        End If
                  End If
              
              Next i
            End If
        End If
        RS.Close
        Set RS = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    'Grabar tabla spagop de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        'antes de grabar en la spago comprobar que existe en conta.sforpa la
        'forma de pago de la factura. Sino existe insertarla

        'vemos si existe en la conta
        If vParamAplic.ContabilidadNueva Then
            CadValuesAux2 = DevuelveDesdeBDNew(conConta, "formapago", "codforpa", "codforpa", ForPago, "N")
        Else
            CadValuesAux2 = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", ForPago, "N")
        End If
        'si no existe la forma de pago en conta, insertamos la de aritaxi
        If CadValuesAux2 = "" Then
'++
            Dim Sql8 As String
            Dim RS8 As ADODB.Recordset

            Sql8 = "select * from sforpa where codforpa = " & DBSet(ForPago, "N")
            Set RS8 = New ADODB.Recordset
            RS8.Open Sql8, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS8.EOF Then
                'insertamos e sforpa de la CONTA
                If vParamAplic.ContabilidadNueva Then
                    Sql8 = "INSERT INTO formapago(codforpa,nomforpa,tipforpa,numerove,primerve,restoven)"
                Else
                    Sql8 = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                End If
                Sql8 = Sql8 & " VALUES(" & ForPago & ", " & DBSet(RS!nomforpa, "T") & ", " & DBSet(RS!tipforpa, "N")
                If vParamAplic.ContabilidadNueva Then
                    Sql8 = Sql8 & "," & DBSet(RS!numerove, "N") & "," & DBSet(RS!primerve, "N") & "," & DBSet(RS!restoven, "N") & ")"
                Else
                    Sql8 = Sql8 & ")"
                End If
                ConnConta.Execute Sql8
            End If
            RS8.Close
            Set RS8 = Nothing
'++
        
        
        End If

        'Insertamos en la tabla spagop de la CONTA
        'SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1) "
        'David. Cuenta bancaria y descripcion textos
        If vParamAplic.ContabilidadNueva Then
            Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, " ' ) "
            Sql = Sql & "iban,nomprove,domprove,pobprove,cpprove,proprove,nifprove,codpais) "
        Else
            Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb " ' ) "
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

    b = True
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
    End If
    InsertarEnTesoreria = b
End Function

'Si es LiquidacionProveedores entonces
' es modificaion Terrasana. para cuando generamos las facturas de proveedores a nuestros socios
' con tipo iva REA
' La diferencia PPAL es en el Arimoney y que en la lineas de la factura , en la conta
' es el lugar donde van las retenciones
Public Function TraspasoAlbaranesAFactura(cadSQL As String, InsertarEnLaTesoreria As Boolean, MarcarAlbaranesComoContabilizados As Boolean, LiquidacionProveedores As Boolean) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String
Dim ImporteAuxiliarTesoreria As Currency
    
    On Error GoTo ETraspasoAlbFac
    TraspasoAlbaranesAFactura = False
    
    conn.BeginTrans
    ConnConta.BeginTrans
    

    
    cadSQL = Replace(cadSQL, "slialp", "scaalp")
    
    'Cuenta Prevista de Cobro de las Facturas
    CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", BancoPr, "N")
    
    
    'Insertar la Factura
    b = InsertarFactura(cadSQL, MenError)
    
    If b Then
        'Insertar en Tesoreria
        '#### Laura 15/11/2006 Recuperar facturas ALZIRA
        'If frmComFacturar.Check1(0).Value = 1 Then    'Lo cambia DAVID, para no pasar valores de formularios
        If InsertarEnLaTesoreria Then
            '####
            
            b = InsertarEnTesoreria(MenError)
            
        End If
        
        
'        '#### Laura 15/11/2006 Recuperar facturas ALZIRA
'        'If frmComFacturar.Check1(1).Value = 1 Then     'Lo cambia DAVID, para no pasar valores de formularios
'        If MarcarAlbaranesComoContabilizados Then
'            If B Then B = Me.MarcarContabilizada(MenError)
'        End If
'        '####
'
'
'        'Eliminar los Albaranes y sus lineas
'        If B Then
'            B = EliminarAlbaranes(cadSQL, MenError)
'        End If
    End If

     If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        TraspasoAlbaranesAFactura = True
        If Not LiquidacionProveedores Then MsgBox "La Factura de los Albaranes seleccionados se generó correctamente.", vbInformation
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        TraspasoAlbaranesAFactura = False
        MsgBox "ATENCIÓN:" & vbCrLf & "La Factura NO se generó correctamente!!!." & vbCrLf & MenError, vbInformation
    End If

    Exit Function


ETraspasoAlbFac:
    If Err.Number <> 0 Then
'        b = False
        conn.RollbackTrans
        ConnConta.RollbackTrans
        TraspasoAlbaranesAFactura = False
        MuestraError Err.Number, "Pasando Albaranes a Factura.(Recepción)", Err.Description
    End If
'    If b Then
'        Conn.CommitTrans
'        ConnConta.CommitTrans
'        TraspasoAlbaranesAFactura = True
'        MsgBox "La Factura de los Albaranes seleccionados se generó correctamente.", vbInformation
'    Else
'        Conn.RollbackTrans
'        ConnConta.RollbackTrans
'        TraspasoAlbaranesAFactura = False
'        MsgBox "ATENCIÓN:" & vbCrLf & "La Factura NO se generó correctamente!!!.", vbInformation
'    End If
End Function


Private Function InsertarFactura(cadSQL As String, cadError As String) As Boolean
'Insertamos en las tablas de factura de proveedor: scafpc, scafpa, slifpc
Dim b As Boolean
    
    InsertarFactura = False
    
    
    'Cabecera Factura
    cadError = ""
    b = InsertarCabeceraFactu(cadError)
    If Not b Then
'        cadError = "Error al insertar cabecera de factura"
        Exit Function
    End If
    
    InsertarFactura = True
End Function



Private Function InsertarCabeceraFactu(Optional cadErr As String) As Boolean
'Inserta la cabecera de la factura en la tabla: scafpc
Dim Sql As String
Dim vProve As CProveedor
Dim Nulo2 As String
Dim Nulo3 As String
Dim RS As ADODB.Recordset
    On Error GoTo EInsertar



    'Obtener los datos del socio: nombre, domicilio, NIF,....
    Sql = "select nomclien, domclien, codpobla, pobclien, proclien, nifclien, telclie1 "
    Sql = Sql & " from sclien where codclien= " & DBSet(Socio, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        NombreSocio = RS.Fields(0).Value
        DomicilioSocio = RS.Fields(1).Value
        CPostalSocio = RS.Fields(2).Value
        PoblacionSocio = RS.Fields(3).Value
        ProvinciaSocio = RS.Fields(4).Value
        nifSocio = RS.Fields(5).Value
        TelefSocio = RS.Fields(6).Value
        CtaSocio = "" 'vProve.CuentaCble
    End If
    
'    Set vProve = New CProveedor
'    If vProve.Existe(Proveedor) Then
'        vProve.LeerDatos (Proveedor)
'        '#### Leer estos datos de la tabla scaalpr y no de sprove
'        NombreProv = vProve.Nombre
'        DomicilioProv = vProve.Domicilio
'        CPostalProv = vProve.CPostal
'        PoblacionProv = vProve.Poblacion
'        ProvinciaProv = vProve.Provincia
'        NIFProv = vProve.NIF
'        TelefProv = vProve.TfnoAdmon
'        CtaProve = "" 'vProve.CuentaCble
'    End If

    'Insertar en la tabla cabecera de la factura de compras
    Sql = "INSERT INTO sfactusoc (codsocio,numfactu,fecfactu,fecrecep,nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,codforpa,codtraba,brutofac,dtoppago,dtognral,impppago,impgnral,baseiva1,baseiva2,baseiva3, "
    Sql = Sql & "tipoiva1,porciva1,impoiva1,totalfac,intconta"
    'RETENCION SOCIOS, y no socios: segun tiporet
    Sql = Sql & ",PorRet,ImpRet,tiporet"
    Sql = Sql & ") VALUES (" & Socio & "," & DBSet(NumFactu, "T") & "," & DBSet(FecFactu, "F") & "," & DBSet(FecRecep, "F") & ","
    Sql = Sql & DBSet(NombreSocio, "T", "N") & "," & DBSet(DomicilioSocio, "T", "N") & "," & DBSet(CPostalSocio, "T", "N") & "," & DBSet(PoblacionSocio, "T", "N") & "," & DBSet(ProvinciaSocio, "T", "N") & ","
    Sql = Sql & DBSet(nifSocio, "T", "N") & "," & DBSet(TelefSocio, "T", "N") & "," & ForPago & ","
    Sql = Sql & Trabajador & "," & DBSet(BrutoFac, "N", "N") & ","
    Sql = Sql & DBSet(DtoPPago, "N", "N") & "," & DBSet(DtoGnral, "N", "N") & ","
    Sql = Sql & DBSet(ImpPPago, "N", "N") & "," & DBSet(ImpGnral, "N", "N") & ","
    Sql = Sql & DBSet(BaseIVA1, "N", "N") & ","
    Nulo2 = "N"
    Nulo3 = "N"
    Sql = Sql & DBSet(TipoIVA1, "N", "N") & ","
    Sql = Sql & DBSet(PorceIVA1, "N", "N") & ","
    Sql = Sql & DBSet(ImpIVA1, "N", "N") & ","
    Sql = Sql & DBSet(TotalFac, "N", "N") & ",0,"
    'Retencion socios
    Sql = Sql & DBSet(PorRet, "N", "S") & "," & DBSet(ImpRet2, "N", "S") & "," & TipoRet & ")"
    
    conn.Execute Sql
    Set vProve = Nothing
    InsertarCabeceraFactu = True

EInsertar:
    If Err.Number <> 0 Then
        InsertarCabeceraFactu = False
        cadErr = "Error al insertar cabecera factura: " & Err.Description
    Else
        InsertarCabeceraFactu = True
    End If
End Function


Public Function CalcularDatosFactura(cadWHERE As String, NomTabla As String, NomTablaLin As String) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim RS As ADODB.Recordset
Dim i As Integer

Dim Sql As String
Dim cadAux As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotImpIVA As Currency

Dim Impaux As Currency
Dim ImpIVA As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

    CalcularDatosFactura = False
    On Error GoTo ECalcular

   
    cadWHERE = Replace(cadWHERE, NomTabla, NomTablaLin)
    Sql = "SELECT " & "sartic.codigiva,sum(importel) as bruto"
    Sql = Sql & " FROM " & NomTablaLin & " LEFT OUTER JOIN sartic ON " & NomTablaLin & ".codartic=sartic.codartic"
    Sql = Sql & " WHERE " & cadWHERE
    Sql = Sql & " GROUP BY codigiva "
    Sql = Sql & " ORDER BY codigiva "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    TotBruto = 0
    i = 1
    While Not RS.EOF
        'Aqui vamos acumulando la suma del importe bruto de las lineas
        TotBruto = TotBruto + RS!bruto
        ImpBImIVA = RS!bruto
        
        'Aplicarle el dto ppago
        '---- Laura: 24/10/2006
        Impaux = CalcularPorcentaje(RS!bruto, DtoPPago, 2)
        ImpBImIVA = ImpBImIVA - Impaux '(bruto - DtoPP)
        
        'Aplicarle el dto grnal
        '---- Laura: 24/10/2006
        Impaux = CalcularPorcentaje(RS!bruto, DtoGnral, 2)
        ImpBImIVA = ImpBImIVA - Impaux '(bruto - Dtogn)
        
        'Obtener el % de IVA
        'SI TIENE RETENCION ENTONCES el tipo de iva es el REA
        'FALTA###
        
        cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(RS!codigiva), "N")
        If cadAux = "" Then cadAux = "0"
        ImpIVA = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotImpIVA = TotImpIVA + ImpIVA
        
        
        Select Case i
            Case 1  'IVA 1
                TipoIVA1 = RS!codigiva
                
                BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA1 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA1 = ImpIVA
                
        End Select
        i = i + 1
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing

    'TOTALES
    BrutoFac = TotBruto
    
    'Aplicarle el dto ppago
'    ImpPPago = CCur(CalcularDto(CStr(TotBruto), CStr(DtoPPago)))
'    '---- Laura: 27/09/2006
'    ImpPPago = Round(ImpPPago, 2)
'    '----
    '---- Laura: 24/10/2006
    ImpPPago = CalcularPorcentaje(TotBruto, DtoPPago, 2)
    '----
    
    'Aplicarle el dto general
'    ImpGnral = CCur(CalcularDto(CStr(TotBruto), CStr(DtoGnral)))
'    '---- Laura: 27/09/2006
'    ImpGnral = Round(ImpGnral, 2)
'    '----
    '---- Laura: 24/10/2006
    ImpGnral = CalcularPorcentaje(TotBruto, DtoGnral, 2)
    '----
    
    'Base Imponible
    BaseImp = TotBruto - ImpPPago - ImpGnral
    
    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA
    If PorRet > 0 And ImpRet2 <> 0 Then TotalFac = TotalFac - ImpRet2
    
    CalcularDatosFactura = True
    
ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFactura = False
    Else
        CalcularDatosFactura = True
    End If
End Function

Public Function CalcularDatosFacturaLiq(cadWHERE As String, NomTabla As String, NomTablaLin As String) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim RS As ADODB.Recordset
Dim i As Integer

Dim Sql As String
Dim cadAux As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotImpIVA As Currency

Dim Impaux As Currency
Dim ImpIVA As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

    CalcularDatosFacturaLiq = False
    On Error GoTo ECalcular

    cadWHERE = Replace(cadWHERE, NomTabla, NomTablaLin)
    Sql = "SELECT sum(impventa) as bruto"
    Sql = Sql & " FROM " & NomTablaLin
    Sql = Sql & " WHERE " & cadWHERE
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Aqui vamos acumulando la suma del importe bruto de las lineas
    TotBruto = RS!bruto
    
    'Obtener el % de IVA
    'SI TIENE RETENCION ENTONCES el tipo de iva es el REA
    'FALTA###
    
    cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(vParamAplic.IVA_REA), "N")
    If cadAux = "" Then cadAux = "0"
    
    BaseIVA1 = Round2(TotBruto / (1 + (cadAux / 100)), 2)
    ImpIVA1 = Round2(TotBruto - BaseIVA1, 2)
    
    TipoIVA1 = vParamAplic.IVA_REA
    
    PorceIVA1 = cadAux '% de IVA
    
    ImpRet2 = Round2(TotBruto * vParamAplic.PorReten / 100, 2)
    
    RS.Close
    Set RS = Nothing

    BrutoFac = TotBruto
    PorRet = vParamAplic.PorReten
    
    'TOTAL de la factura
    TotalFac = BaseIVA1 + ImpIVA1
    If PorRet > 0 And ImpRet2 <> 0 Then TotalFac = TotalFac - ImpRet2
    
    CalcularDatosFacturaLiq = True
    
ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFacturaLiq = False
    Else
        CalcularDatosFacturaLiq = True
    End If
End Function




'#### Laura 15/11/2006 Recuperar facturas ALZIRA
Public Function MarcarContabilizada(Optional cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    MarcarContabilizada = False
    
    Sql = "UPDATE sfactusoc SET intconta=1 "
    Sql = Sql & " WHERE codsocio=" & Me.Socio & " AND numfactu=" & DBSet(Me.NumFactu, "T") & " AND fecfactu=" & DBSet(Me.FecFactu, "F")
    conn.Execute Sql
    
    MarcarContabilizada = True
    Exit Function
    
EActualizar:
    MarcarContabilizada = False
    cadErr = "Error al marcar como contabilizada:" & Err.Description
End Function




Public Function InsertarFacturaPublicidad(cadError As String) As Boolean
'Insertamos en las tablas de factura de socio: sfactusoc
Dim b As Boolean
    
    InsertarFacturaPublicidad = False
    
    'Cabecera Factura
    cadError = ""
    b = InsertarCabeceraFactuPubli(cadError)
    If Not b Then
        Exit Function
    End If
    
    InsertarFacturaPublicidad = True

End Function

Private Function InsertarCabeceraFactuPubli(Optional cadErr As String) As Boolean
Dim Sql As String

    On Error GoTo EInsertar


    InsertarCabeceraFactuPubli = False

    'sfactusoc
    Sql = "INSERT INTO sfactusoc (codtipom,codsocio,numfactu,fecfactu,concepto,importel,baseiva1,impoiva1,"
    Sql = Sql & "codiiva1,porciva1,totalfac,impreten,intconta,codforpa) VALUES (" & DBSet(tipoMov, "T") & "," & Socio & ","
    Sql = Sql & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "'," & DBSet(Concepto, "T") & ","
    Sql = Sql & TransformaComasPuntos(CStr(BaseIVA1)) & "," & TransformaComasPuntos(CStr(BaseIVA1)) & ","
    Sql = Sql & TransformaComasPuntos(CStr(ImpIVA1)) & "," & TransformaComasPuntos(CStr(TipoIVA1)) & "," & TransformaComasPuntos(CStr(PorceIVA1)) & ","
    Sql = Sql & TransformaComasPuntos(CStr(TotalFac)) & ",NULL,0," & ForPago & ")"
        
    conn.Execute Sql
    
    InsertarCabeceraFactuPubli = True
    Exit Function
    
EInsertar:
    If Err.Number <> 0 Then
        cadErr = "Error al insertar cabecera factura publicidad: " & Err.Description
    End If
End Function


Public Function InsertarEnTesoreriaCobro(vTextosCSB As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: aritaxi.svenci y en conta.scobros
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String, codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAuxConta As String
Dim CadValues3 As String
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single 'importe para insertar en la svenci
Dim ImpVenci2 As Single 'importe para insertar en conta.scobro
Dim i As Byte
Dim TotalFactura2 As Currency   'Por si acaso lleva aportacion al terminal
'1 Julio 2009. Los graba en scobro
Dim CadenaDatosFiscales As String

Dim CADENA As String

    On Error GoTo EInsertarTesoreria

    b = False
    InsertarEnTesoreriaCobro = False


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
    textcsb33 = "'FACTURA: " & ObtenerLetraSerie(tipoMov) & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd mmm yyyy") & "'"


    'Datos fiscales en scobro     Julio 2009
    'nomclien,domclien,pobclien, cpclien,proclien
    CadenaDatosFiscales = DBSet(NombreSocio, "T") & "," & DBSet(DomicilioSocio, "T") & "," & DBSet(PoblacionSocio, "T")
    CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(CPostalSocio, "T") & "," & DBSet(ProvinciaSocio, "T")
    
    If vParamAplic.ContabilidadNueva Then
        CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(nifSocio, "T") & ",'ES'"
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
            Sql = Sql & " WHERE codclien=" & Socio
            Set RS = New ADODB.Recordset
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
'            codmacta = DBSet(RS!codmacta, "T")

            CADENA = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
            
            
            '[Monica] 30/05/2011: Añadida la condicion del tipo de movimiento
            codmacta = vParamAplic.Raiz_Cta_Soc_Liqui & Format(Socio, CADENA)
'            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
            
            If Not RS.EOF Then
                cadValuesAux = "('" & tipoMov & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                CadValuesAuxConta = "('" & ObtenerLetraSerie(tipoMov) & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                '                    Añadire a la cadena fija esta los valores de textcsb41,txcs
                CadValuesAuxConta = CadValuesAuxConta & vTextosCSB & ","
                '-------- Primer Vencimiento
                i = 1
                'FECHA VTO
                FecVenci = CDate(FecFactu)
                '=== Laura 23/01/2007
                'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                '===
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(RS!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(RS!mesnogir, "N"), DBLet(RS!DiaVtoAt, "N"), DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & i & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAuxConta & i & ", "
                CadValues2 = CadValues2 & codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                TotalFactura2 = TotalFac
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFactura2
                    ImpVenci2 = TotalFactura2
                Else
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura2) / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFactura2 Then
                        ImpVenci = Round(ImpVenci + (TotalFactura2 - ImpVenci * rsVenci!numerove), 2)
                    End If
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If (ImpVenci2 * rsVenci!numerove) <> TotalFactura2 Then
                        ImpVenci2 = Round(ImpVenci2 + (TotalFactura2 - (ImpVenci2 * rsVenci!numerove)), 2)
                    End If
                End If
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', "
                If Not vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & DBSet(CCC_Entidad, "N") & ", " & DBSet(CCC_Oficina, "N") & ", " & DBSet(CCC_CC, "T") & ", " & DBSet(CCC_CTa, "T") & ", "
                End If
                CadValues2 = CadValues2 & textcsb33 & ", " & DBSet(1, "N")
                
                'departamento y transfer
                CadValues2 = CadValues2 & ",NULL,NULL"
                '1 Julio 2009
                ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                 
                If vParamAplic.ContabilidadNueva Then
                    vvIban = MiFormat(mCCC_Iban, "") & MiFormat(mCCC_Entidad, "0000") & MiFormat(mCCC_Oficina, "0000") & MiFormat(mCCC_CC, "00") & MiFormat(mCCC_CTa, "0000000000")
                    
                    CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                
                Else
                    '[Monica]22/11/2013: tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                       CadValues2 = CadValues2 & "," & DBSet(mCCC_Iban, "T", "S") & ")"
                    Else
                       CadValues2 = CadValues2 & ")"
                    End If
                End If
                
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For i = 2 To rsVenci!numerove
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(RS!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(RS!mesnogir, "N"), DBLet(RS!DiaVtoAt, "N"), DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                    End If
                        
                    CadValues = CadValues & ", " & cadValuesAux & i & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & i & ", " & codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura2) / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & textcsb33 & ", " & DBSet(1, "N") & ", "
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & DBSet(CCC_Entidad, "N") & ", " & DBSet(CCC_Oficina, "N") & ", " & DBSet(CCC_CC, "T") & ", " & DBSet(CCC_CTa, "T") & ", " & textcsb33 & ", " & DBSet(1, "N") & ", "
                    End If
                    CadValues2 = CadValues2 & "NULL,NULL"
                    '1 Julio 2009
                    ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                    CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                    
                    
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(mCCC_Iban, "") & MiFormat(mCCC_Entidad, "0000") & MiFormat(mCCC_Oficina, "0000") & MiFormat(mCCC_CC, "00") & MiFormat(mCCC_CTa, "0000000000")
                        
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                    Else
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & "," & DBSet(mCCC_Iban, "T", "S") & ")"
                        Else
                           CadValues2 = CadValues2 & ")"
                        End If
                    End If
                Next i
                
            End If
            RS.Close
            
        Else
            'totalfac =0 and numerovtos >=1
            b = True
        End If
        
        Set RS = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
'    If CadValues <> "" Then
'        Sql = "INSERT INTO svenci (codtipom, numfactu, fecfactu, ordefect, fecefect, impefect)"
'        Sql = Sql & " VALUES " & CadValues
'        conn.Execute Sql
'    End If
    
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
                
                If b Then
                    If vParamAplic.ContabilidadNueva Then
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu,text41csb, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, "
                        Sql = Sql & "text33csb,agente,departamento,transfer "
                        'JULIO 2009
                        'nomclien,domclien,pobclien, cpclien,proclien
                        Sql = Sql & ",nomclien,domclien,pobclien, cpclien,proclien, nifclien,codpais" ')"
                        
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           Sql = Sql & ",iban)"
                        Else
                           Sql = Sql & ")"
                        End If
                    
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
        MenError = "Insertar en Cobro Tesoreria: " & vbCrLf & Err.Description
    End If
    InsertarEnTesoreriaCobro = b
End Function


Private Function InsertarFormaPagoEnConta(nForPa As String, cadErr As String) As Boolean
Dim cadAux As String
Dim cadAux2 As String
Dim Sql As String

    On Error GoTo ErrInsForpa
    InsertarFormaPagoEnConta = False
    
    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
    'forma de pago de la factura. Sino existe insertarla
    
    If vParamAplic.ContabilidadNueva Then
        'vemos si existe en la conta
        cadAux = DevuelveDesdeBDNew(conConta, "formapago", "codforpa", "codforpa", nForPa, "N")
    Else
        'vemos si existe en la conta
        cadAux = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", nForPa, "N")
    End If
    
    'si no existe la forma de pago en conta, insertamos la de aritaxi
    If cadAux = "" Then
        Sql = "select * from sforpa where codforpa = " & DBSet(nForPa, "N")
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            'insertamos e sforpa de la CONTA
            If vParamAplic.ContabilidadNueva Then
                Sql = "INSERT INTO formapago(codforpa,nomforpa,tipforpa,numerove,primerve,restoven)"
            Else
                Sql = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
            End If
            Sql = Sql & " VALUES(" & nForPa & ", " & DBSet(RS!nomforpa, "T") & ", " & DBSet(RS!tipforpa, "N")
            If vParamAplic.ContabilidadNueva Then
                Sql = Sql & "," & DBSet(RS!numerove, "N") & "," & DBSet(RS!primerve, "N") & "," & DBSet(RS!restoven, "N") & ")"
            Else
                Sql = Sql & ")"
            End If
            ConnConta.Execute Sql
            InsertarFormaPagoEnConta = True
        Else
            InsertarFormaPagoEnConta = False
        End If
        RS.Close
        Set RS = Nothing
        
    Else
        InsertarFormaPagoEnConta = True
    End If
    
    
    Exit Function
    
ErrInsForpa:
    InsertarFormaPagoEnConta = False
    cadErr = "Insertar forma de pago en Contablilidad: " & vbCrLf & Err.Description
End Function

