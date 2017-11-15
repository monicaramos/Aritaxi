VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ==== Modificaciones:  =====
' ---- [14/09/2009] (LAURA)  --> se añade el parámetro "dtoxcanti": lleva dto por cantidad (S/N)
' ---- [14/09/2009] (LAURA)  --> se añade el parámetro "creatarifart": crear articulos en todas las tarifas (S/N)
'============================

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros SPARA1
Private mvarTipoDtos As Byte      'Tipo de Descuentos: 0= Aditivo 1=Sobre Resto
Private mvarOrdenDtos As Byte   'Orden de Dtos para obtener los dtos a aplicar al precio segun la tabla: sdtofm
                                'Si F=0004 y M=0004 y existe dtos para (F,M)=(0000,0004) y para (F,M)=(0004,0000) siendo 0000=todas las marcas/todas las familias
                                'si ordendto=0(Fam/Marca) aplicaremos el dto (F,M)=(0004,0000)
                                'si ordendto=1 (Marca/Fam) aplicaremos el dto (F,M)=(0000,0004)
                                
Private mvarCodTarif As Integer   'Cod. Tarifa de PVP
Private mvarCajaComp As Boolean      'Cajas completas precios (S/N)
Private mvarPreuKmCli As Single 'Precio Km Desplazamiento Clientes
Private mvarPreuKmTec As Single 'Precio Km Desplazamiento Técnicos
Private mvarnomGerente As String  'Nombre del Director Gerente
Private mvarnomAdmin As String  'Nombre del Responsable Administración
Private mvarHayMante As Boolean   'Si hay Mantenimiento (S/N)
Private mvarHayRepar As Boolean   'Si hay Reparaciones (S/N)
Private mvarHayFrecu As Boolean   'Si hay Frecuencias (S/N)
Private mvarHayServicio As Boolean   'Si hay Servicios (S/N)
Private mvarHayDepar As Boolean   'Si hay Departamento o Direccion
Private mvarCtrStock As Boolean   'Si hay Control de Stock (S/N)
Private mvarInventar As Boolean  'Si se realiza Inventario ordenando listados por Proveedor o no
Private mvarHaynserie As Boolean 'Si hay Nº de Serie en Compras (S/N)
Private mvarDiasSiMan As Integer 'Nº de dias si hay mantenimiento
Private mvarDiasNoMan As Integer 'Nº de dias si NO hay mantenimiento

'Datos Facturacion
Private mvarCodArticDes As String 'Articulo para facturar desplazamientos
Private mvarDiaPago1 As Byte 'Dia de pago 1 para compras
Private mvarDiaPago2 As Byte 'Dia de pago 2 para compras
Private mvarDiaPago3 As Byte 'Dia de pago 3 para compras
Private mvarMesnogir As Byte 'Mes a no girar para compras
Private mvarCodForpa As String 'forma de pago para cheque regalo
Private mvarConFactuPubli As String 'concepto facturación publicidad

'Datos de Internet
Private mvarDireMail As String 'Direccion email
Private mvarSMTPhost As String 'Host SMTP para envio de email
Private mvarSMTPuser As String 'Usuario SMTP para envio de email
Private mvarSMTPpass As String 'Password SMTP para envio de email
Private mvarWebSoporte As String 'Web de soporte
Private mvarMailSoporte As String 'Mail de soporte
Private mvarWebVersion As String 'Version de la web

'Datos de la Contabilidad
Private mvarSerConta As String 'Servidor BD de Contabilidad
Private mvarUsuConta As String 'Usuario BD de Contabilidad
Private mvarPasConta As String 'Password BD de Contabilidad
Private mvarNumConta As Byte 'Numero BD de Contabilidad
Private mvarRaiz_Cta_Soc_publi As String 'raiz de la cuenta socio de publicidad
Private mvarRaiz_Cta_Soc_liqui As String 'raiz de la cuenta socio de liquidacion
Private mvarRaiz_CtaClien_Soc As String 'raiz de la cuenta cliente del socio para cuotas
Private mvarRaiz_Cta_Reten_Soc As String 'raiz de la cuenta retencion socio de liquidacion
Private mvarCtaLiquidacion As String 'Cuenta base de liquidacion socios

Private mvarRaiz_Cta_Soc_Equip As String 'raiz de la cuenta socio de facturas de venta

Private mvarCtaTituloAlta As String 'Cuenta titulo alta socio
Private mvarCtaGastoAlta As String 'Cuenta gastos alta socio
Private mvarImpTituloAlta As Currency 'Importe titulo alta socio
Private mvarImpGastoAlta As Currency 'Importe gastos alta socio
Private mvarRaiz_CtaAltaSoc As String 'raiz de la cuenta de alta del socio


'Valores por defecto.   Toca meter un variable pq en el matenimiento (en el form) no deja modificar si no es con el objeto
'-----------------
'
'      Clientes
Private mvarDefActiv As Integer
Private mvarDefEnvio As Integer
Private mvarDefZona As Integer
Private mvarDefSitu As Integer
Private mvarDefTarifa As Integer
Private mvarDefAgente As Integer
Private mvarPresinchofe As Single
Private mvarPreconchofe As Single
Private mvarPreporserv As Single
Private mvarPreporAlq As Single

'Facturacion recargas telefonia
Private mvarCodarticTfnia As String

Private mvarArtAlquiler As String 'Articulo para alquiler de equipos
Private mvarArtServCuotas As String 'Articulo de Servicios para cuotas
Private mvarArtCuotaSinChofer As String
Private mvarArtCuotaConChofer As String
Private mvarArtCuotaExtraor As String

Private mvarArtRetenciones As String 'Articulo para la factura de retenciones

Private mvarArtSuplidos As String 'Articulo de suplidos para factura de liquidacion de socios y de cliente



    'aVISOS  al entrar
    Private mvaravipedcli  As Integer
    Private mvaravipedpro As Integer
    Private mvaravialbcli As Integer
    Private mvaravialbpro As Integer
    Private mvaravimanteni As Integer
    Private mvaravirepara As Integer
    Private mvaraviavisos As Integer
    Private mvarTieneAvisos As Boolean 'De este no hay valor en BD. Simplememnte si alguno de ellos tiene valor, el tendra valor


'Si LLEVA REcargo de equivalencia
Private mvarRecargoEquivalencia As Boolean
Private mvarTipoIVAre1 As Integer
Private mvarTipoIVAre2 As Integer
Private mvarTipoIVAre3 As Integer
Private mvarIVA1 As Integer
Private mvarIVA2 As Integer
Private mvarIVA3 As Integer


Private mvarTipoIVAAntre1 As Integer
Private mvarTipoIVAAntre2 As Integer
Private mvarIVAAnt1 As Integer
Private mvarIVAAnt2 As Integer
Private mvarFecCambioIva As Date


'Retencion y REA
Private mvarIVAREA As Integer
Private mvarCtaReten As String
Private mvarPorReten As Currency

Private mvarIVA_Exento As Integer
Private mvarIVA_Intracomunitario As Integer

'Tickets TPV agrupados a contabilidad
Private mvarContabilizarTicketAgrupados As Boolean

'Albaranes y facturas en B
Private mvarContabilidadB As Byte

'Cueenta para la aportacion al terminal en RADIVAL
Private mvarctaaportacion As String

'Produccion
Private mvarProduccion As Boolean

'Descriptores
Private mvarDescriptores As Boolean

'Que llevara al campo observaciones de la factura cli/pro
Private mvarObsFactura As Byte

'Artículo de gastos de admon
Private mvarArtGastosAdmon As String


'FONTENAS
Private mvarArtPortes As String
Private mvarImporteMinimo As Currency
Private mvarAbonoKilos As Currency


'Produccion
Private mvarComponentePorcentaje As Boolean   'Si el componente entra como porcentaje o como cantidad


' ---- [14/09/2009] (LAURA)
Private mvarDtoxCanti As Boolean 'Lleva dtos por cantidad para facturacion: Pedidos, Ofertas, Albaranes
Private mvarCreaTarifArt2 As Byte 'crea articulo en todas las tarifas
'                               David cambia a byte
'                               0.- NO crea
'                               1.- Solo en tarifa ppal
'                               2.- En todas tarifas



' ---- [19/10/2009] [LAURA]: añadir campo modo analitica
Private mvarModoAnalit As Byte 'modo analitica: 0=trabajador, 1=Familia, 2=Proyecto

Private mvarEnvioDesdeOutlook As Boolean  'Si envia por outlook



'Exe nuestro k llama a el Outlook
Private mvarExeEnvioMail As String

'Dias de garantia despues de efectuar un reparacion
Private mvarDiasGarantia As Integer


'De momento solo tiene el DE leer
'Si lleva el trozo del CRM
Private mvarTieneCRM As Boolean

Private mvarCooperativa As Byte ' Cooperativa 0= Teletaxi 1=TaxiVip
Private mvarPrecioxDistancia As Single 'Precio por distancia
Private mvarPrecioxTpoEspera As Single 'Precio por tpo de espera

Private mvarPathFacturaE As String  'FacturaE

'[Monica]28/09/2012: Impresora de tarjetas
Private mvarImpresoraTarjetas As String  'impresora de tarjetas

Private Sql As String  'Para la sentencia
Private Mens As String    'Para los mensajes


Private mvarContabilidadNueva As Boolean




'------------------------------------------------
'Propiedades del modulo CParamAplic
'------------------------------------------------

Public Property Let TipoDtos(ByVal vData As Byte)
     mvarTipoDtos = vData
End Property

Public Property Get TipoDtos() As Byte
     TipoDtos = mvarTipoDtos
End Property

Public Property Let Raiz_Cta_Soc_publi(ByVal vData As String)
    mvarRaiz_Cta_Soc_publi = vData
End Property
Public Property Get Raiz_Cta_Soc_publi() As String
    Raiz_Cta_Soc_publi = mvarRaiz_Cta_Soc_publi
End Property

Public Property Let Raiz_Cta_Soc_Liqui(ByVal vData As String)
    mvarRaiz_Cta_Soc_liqui = vData
End Property
Public Property Get Raiz_Cta_Soc_Liqui() As String
    Raiz_Cta_Soc_Liqui = mvarRaiz_Cta_Soc_liqui
End Property

Public Property Let Raiz_Cta_Soc_Equip(ByVal vData As String)
    mvarRaiz_Cta_Soc_Equip = vData
End Property
Public Property Get Raiz_Cta_Soc_Equip() As String
    Raiz_Cta_Soc_Equip = mvarRaiz_Cta_Soc_Equip
End Property


Public Property Let Raiz_Cta_Reten_Soc(ByVal vData As String)
    mvarRaiz_Cta_Reten_Soc = vData
End Property

Public Property Get Raiz_Cta_Reten_Soc() As String
    Raiz_Cta_Reten_Soc = mvarRaiz_Cta_Reten_Soc
End Property

Public Property Let Raiz_CtaClien_Soc(ByVal vData As String)
    mvarRaiz_CtaClien_Soc = vData
End Property
Public Property Get Raiz_CtaClien_Soc() As String
    Raiz_CtaClien_Soc = mvarRaiz_CtaClien_Soc
End Property


Public Property Get CtaLiquidacion() As String
     CtaLiquidacion = mvarCtaLiquidacion
End Property

Public Property Let CtaLiquidacion(ByVal vData As String)
     mvarCtaLiquidacion = vData
End Property


' Alta de socios
Public Property Get CtaTituloAlta() As String
     CtaTituloAlta = mvarCtaTituloAlta
End Property

Public Property Let CtaTituloAlta(ByVal vData As String)
     mvarCtaTituloAlta = vData
End Property

Public Property Get CtaGastoAlta() As String
     CtaGastoAlta = mvarCtaGastoAlta
End Property

Public Property Let CtaGastoAlta(ByVal vData As String)
     mvarCtaGastoAlta = vData
End Property

Public Property Get ImpTituloAlta() As Currency
     ImpTituloAlta = mvarImpTituloAlta
End Property

Public Property Let ImpTituloAlta(ByVal vData As Currency)
     mvarImpTituloAlta = vData
End Property

Public Property Get ImpGastoAlta() As Currency
     ImpGastoAlta = mvarImpGastoAlta
End Property

Public Property Let ImpGastoAlta(ByVal vData As Currency)
     mvarImpGastoAlta = vData
End Property

Public Property Let Raiz_CtaAltaSoc(ByVal vData As String)
    mvarRaiz_CtaAltaSoc = vData
End Property

Public Property Get Raiz_CtaAltaSoc() As String
    Raiz_CtaAltaSoc = mvarRaiz_CtaAltaSoc
End Property



Public Property Let OrdenDtos(ByVal vData As Byte)
     mvarOrdenDtos = vData
End Property

Public Property Get OrdenDtos() As Byte
     OrdenDtos = mvarOrdenDtos
End Property

Public Property Let CodTarifa(ByVal vData As Integer)
     mvarCodTarif = vData
End Property

Public Property Get CodTarifa() As Integer
     CodTarifa = mvarCodTarif
End Property
Public Property Let ConFactuPubli(ByVal vData As String)
    mvarConFactuPubli = vData
End Property
Public Property Get ConFactuPubli() As String
    ConFactuPubli = mvarConFactuPubli
End Property
Public Property Let CajasCompletas(ByVal vData As Boolean)
     mvarCajaComp = vData
End Property

Public Property Get CajasCompletas() As Boolean
     CajasCompletas = mvarCajaComp
End Property


Public Property Let PrecioKmClientes(ByVal vData As Single)
     mvarPreuKmCli = vData
End Property

Public Property Get PrecioKmClientes() As Single
     PrecioKmClientes = mvarPreuKmCli
End Property

Public Property Let PrecioCuotaSinChofe(ByVal vData As Single)
    mvarPresinchofe = vData
End Property
Public Property Get PrecioCuotaSinChofe() As Single
    PrecioCuotaSinChofe = mvarPresinchofe
End Property

Public Property Let PrecioCuotaConChofe(ByVal vData As Single)
    mvarPreconchofe = vData
End Property
Public Property Get PrecioCuotaConChofe() As Single
    PrecioCuotaConChofe = mvarPreconchofe
End Property

Public Property Let PrecioPorServicio(ByVal vData As Single)
    mvarPreporserv = vData
End Property
Public Property Get PrecioPorServicio() As Single
    PrecioPorServicio = mvarPreporserv
End Property

Public Property Let PrecioPorAlquiler(ByVal vData As Single)
    mvarPreporAlq = vData
End Property
Public Property Get PrecioPorAlquiler() As Single
    PrecioPorAlquiler = mvarPreporAlq
End Property




Public Property Let PrecioKmTecnicos(ByVal vData As Single)
     mvarPreuKmTec = vData
End Property

Public Property Get PrecioKmTecnicos() As Single
     PrecioKmClientes = mvarPreuKmTec
End Property


Public Property Let NomGerente(ByVal vData As String)
     mvarnomGerente = vData
End Property


Public Property Get NomGerente() As String
     NomGerente = mvarnomGerente
End Property


Public Property Let NomRespAdmin(ByVal vData As String)
     mvarnomAdmin = vData
End Property

Public Property Get NomRespAdmin() As String
     NomRespAdmin = mvarnomAdmin
End Property


Public Property Let Mantenimientos(ByVal vData As Boolean)
     mvarHayMante = vData
End Property

Public Property Get Mantenimientos() As Boolean
     Mantenimientos = mvarHayMante
End Property


Public Property Let Reparaciones(ByVal vData As Boolean)
     mvarHayRepar = vData
End Property

Public Property Get Reparaciones() As Boolean
     Reparaciones = mvarHayRepar
End Property

Public Property Let Frecuencias(ByVal vData As Boolean)
'    If vData Then
'        MsgBox "Lo sentimos todavía no puede activar las frecuencias. En próximas versiones de la aplicación será posible", vbExclamation
'        vData = False
'    End If
     mvarHayFrecu = vData
End Property

Public Property Get Servicios() As Boolean
     Servicios = mvarHayServicio
End Property

Public Property Let Servicios(ByVal vData As Boolean)
    Dim Mens As String
    If vData Then
        If MovimientosParaServicios() Then
            vData = True
        Else
            Mens = "Para poder activar los servicios es necesario que los movimientos ALS y FAS estén dados de alta." & vbCrLf & _
                    "Délos de alta ahora y luego vuelva a intentar activar los servicios. Gracias."
            MsgBox Mens, vbExclamation
            vData = False
        End If
    End If
     mvarHayServicio = vData
End Property

Public Property Get Frecuencias() As Boolean
     Frecuencias = mvarHayFrecu
End Property

Public Property Let Departamento(ByVal vData As Boolean)
     mvarHayDepar = vData
End Property

Public Property Get Departamento() As Boolean
     Departamento = mvarHayDepar
End Property


Public Property Let ControlStock(ByVal vData As Boolean)
     mvarCtrStock = vData
End Property

Public Property Get ControlStock() As Boolean
     ControlStock = mvarCtrStock
End Property

Public Property Let InventarioxProv(ByVal vData As Boolean)
     mvarInventar = vData
End Property

Public Property Get InventarioxProv() As Boolean
     InventarioxProv = mvarInventar
End Property

Public Property Let NumSeries(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarHaynserie = vData
End Property

Public Property Get NumSeries() As Boolean
     NumSeries = mvarHaynserie
End Property


Public Property Let DiasSiMante(ByVal vData As Integer)
     mvarDiasSiMan = vData
End Property

Public Property Get DiasSiMante() As Integer
     DiasSiMante = mvarDiasSiMan
End Property


Public Property Let DiasNoMante(ByVal vData As Integer)
     mvarDiasNoMan = vData
End Property

Public Property Get DiasNoMante() As Integer
     DiasNoMante = mvarDiasNoMan
End Property


Public Property Let ArticServ(ByVal vData As String)
     mvarCodArticDes = vData
End Property

Public Property Get ArticServ() As String
     ArticServ = mvarCodArticDes
End Property


Public Property Let DiaPago1(ByVal vData As Byte)
     mvarDiaPago1 = vData
End Property

Public Property Get DiaPago1() As Byte
     DiaPago1 = mvarDiaPago1
End Property

Public Property Let DiaPago2(ByVal vData As Byte)
     mvarDiaPago2 = vData
End Property

Public Property Get DiaPago2() As Byte
     DiaPago2 = mvarDiaPago2
End Property

Public Property Let DiaPago3(ByVal vData As Byte)
     mvarDiaPago3 = vData
End Property

Public Property Get DiaPago3() As Byte
     DiaPago3 = mvarDiaPago3
End Property


Public Property Let MesNoGirar(ByVal vData As Byte)
     mvarMesnogir = vData
End Property

Public Property Get MesNoGirar() As Byte
     MesNoGirar = mvarMesnogir
End Property

'forma de pago de los cheques regalo
Public Property Let ForPagoChequeRegalo(ByVal vData As String)
     mvarCodForpa = vData
End Property

Public Property Get ForPagoChequeRegalo() As String
     ForPagoChequeRegalo = mvarCodForpa
End Property




Public Property Let DireMail(ByVal vData As String)
     mvarDireMail = vData
End Property

Public Property Get DireMail() As String
     DireMail = mvarDireMail
End Property

Public Property Let SMTPhost(ByVal vData As String)
     mvarSMTPhost = vData
End Property

Public Property Get SMTPhost() As String
     SMTPhost = mvarSMTPhost
End Property

Public Property Let SMTPuser(ByVal vData As String)
     mvarSMTPuser = vData
End Property

Public Property Get SMTPuser() As String
     SMTPuser = mvarSMTPuser
End Property

Public Property Let SMTPpass(ByVal vData As String)
     mvarSMTPpass = vData
End Property

Public Property Get SMTPpass() As String
     SMTPpass = mvarSMTPpass
End Property

Public Property Let WebSoporte(ByVal vData As String)
     mvarWebSoporte = vData
End Property

Public Property Get WebSoporte() As String
     WebSoporte = mvarWebSoporte
End Property

Public Property Let MailSoporte(ByVal vData As String)
     mvarMailSoporte = vData
End Property

Public Property Get MailSoporte() As String
     MailSoporte = mvarMailSoporte
End Property

Public Property Let WebVersion(ByVal vData As String)
     mvarWebVersion = vData
End Property

Public Property Get WebVersion() As String
     WebVersion = mvarWebVersion
End Property



'servidor conta
Public Property Let ServidorConta(ByVal vData As String)
     mvarSerConta = vData
End Property

Public Property Get ServidorConta() As String
     ServidorConta = mvarSerConta
End Property



Public Property Let UsuarioConta(ByVal vData As String)
     mvarUsuConta = vData
End Property

Public Property Get UsuarioConta() As String
     UsuarioConta = mvarUsuConta
End Property


Public Property Let PasswordConta(ByVal vData As String)
     mvarPasConta = vData
End Property

Public Property Get PasswordConta() As String
     PasswordConta = mvarPasConta
End Property


Public Property Let NumeroConta(ByVal vData As Byte)
     mvarNumConta = vData
End Property

Public Property Get NumeroConta() As Byte
     NumeroConta = mvarNumConta
End Property




Public Property Let PorDefecto_Activ(ByVal vData As Integer)
     mvarDefActiv = vData
End Property

Public Property Get PorDefecto_Activ() As Integer
     PorDefecto_Activ = mvarDefActiv
End Property
Public Property Let PorDefecto_Envio(ByVal vData As Integer)
     mvarDefEnvio = vData
End Property

Public Property Get PorDefecto_Envio() As Integer
     PorDefecto_Envio = mvarDefEnvio
End Property
Public Property Let PorDefecto_Zona(ByVal vData As Integer)
     mvarDefZona = vData
End Property

Public Property Get PorDefecto_Zona() As Integer
     PorDefecto_Zona = mvarDefZona
End Property

Public Property Let ArtCuotaSinChofer(ByVal vData As String)
     mvarArtCuotaSinChofer = vData
End Property

Public Property Get ArtCuotaSinChofer() As String
     ArtCuotaSinChofer = mvarArtCuotaSinChofer
End Property

Public Property Let ArtCuotaConChofer(ByVal vData As String)
     mvarArtCuotaConChofer = vData
End Property

Public Property Get ArtCuotaConChofer() As String
     ArtCuotaConChofer = mvarArtCuotaConChofer
End Property

Public Property Let ArtCuotaExtraor(ByVal vData As String)
    mvarArtCuotaExtraor = vData
End Property

Public Property Get ArtCuotaExtraor() As String
    ArtCuotaExtraor = mvarArtCuotaExtraor
End Property

Public Property Let PorDefecto_Situ(ByVal vData As Integer)
     mvarDefSitu = vData
End Property

Public Property Get PorDefecto_Situ() As Integer
     PorDefecto_Situ = mvarDefSitu
End Property

Public Property Let PorDefecto_Tarifa(ByVal vData As Integer)
     mvarDefTarifa = vData
End Property

Public Property Get PorDefecto_Tarifa() As Integer
     PorDefecto_Tarifa = mvarDefTarifa
End Property
Public Property Let PorDefecto_Agente(ByVal vData As Integer)
     mvarDefAgente = vData
End Property

Public Property Get PorDefecto_Agente() As Integer
     PorDefecto_Agente = mvarDefAgente
End Property



Public Property Let CodarticTfnia(ByVal vData As String)
     mvarCodarticTfnia = vData
End Property

Public Property Get CodarticTfnia() As String
     CodarticTfnia = mvarCodarticTfnia
End Property


Public Property Let ArtAlquiler(ByVal vData As String)
     mvarArtAlquiler = vData
End Property

Public Property Get ArtAlquiler() As String
     ArtAlquiler = mvarArtAlquiler
End Property

Public Property Let ArtServCuotas(ByVal vData As String)
     mvarArtServCuotas = vData
End Property

Public Property Get ArtServCuotas() As String
     ArtServCuotas = mvarArtServCuotas
End Property





Public Property Let avipedcli(ByVal vData As Integer)
     mvaravipedcli = vData
End Property

Public Property Get avipedcli() As Integer
     avipedcli = mvaravipedcli
End Property
    
Public Property Let avipedpro(ByVal vData As Integer)
     mvaravipedpro = vData
End Property

Public Property Get avipedpro() As Integer
     avipedpro = mvaravipedpro
End Property

Public Property Let avialbcli(ByVal vData As Integer)
     mvaravialbcli = vData
End Property

Public Property Get avialbcli() As Integer
     avialbcli = mvaravialbcli
End Property
Public Property Let avialbpro(ByVal vData As Integer)
     mvaravialbpro = vData
End Property

Public Property Get avialbpro() As Integer
     avialbpro = mvaravialbpro
End Property
    

Public Property Let avimanteni(ByVal vData As Integer)
     mvaravimanteni = vData
End Property

Public Property Get avimanteni() As Integer
     avimanteni = mvaravimanteni
End Property
    
Public Property Let avirepara(ByVal vData As Integer)
     mvaravirepara = vData
End Property

Public Property Get avirepara() As Integer
     avirepara = mvaravirepara
End Property

Public Property Let aviavisos(ByVal vData As Integer)
     mvaraviavisos = vData
End Property

Public Property Get aviavisos() As Integer
     aviavisos = mvaraviavisos
End Property

'Solo tiene GET
Public Property Get TieneAvisos() As Integer
     TieneAvisos = mvarTieneAvisos
End Property



'---------------------------------------
'RECARGO EQUIVALENCIA
Public Property Get TipoIVAre1() As Integer
     TipoIVAre1 = mvarTipoIVAre1
End Property

Public Property Let TipoIVAre1(ByVal vData As Integer)
     mvarTipoIVAre1 = vData
End Property
Public Property Get TipoIVAre2() As Integer
     TipoIVAre2 = mvarTipoIVAre2
End Property

Public Property Let TipoIVAre2(ByVal vData As Integer)
     mvarTipoIVAre2 = vData
End Property

Public Property Get TipoIVAre3() As Integer
     TipoIVAre3 = mvarTipoIVAre3
End Property

Public Property Let TipoIVAre3(ByVal vData As Integer)
     mvarTipoIVAre3 = vData
End Property

Public Property Get TipoIVA1() As Integer
     TipoIVA1 = mvarIVA1
End Property

Public Property Let TipoIVA1(ByVal vData As Integer)
     mvarIVA1 = vData
End Property
Public Property Get TipoIVA2() As Integer
     TipoIVA2 = mvarIVA2
End Property

Public Property Let TipoIVA2(ByVal vData As Integer)
     mvarIVA2 = vData
End Property

Public Property Get TipoIVA3() As Integer
     TipoIVA3 = mvarIVA3
End Property

Public Property Let TipoIVA3(ByVal vData As Integer)
     mvarIVA3 = vData
End Property

Public Property Get RecargoEquivalencia() As Integer
     RecargoEquivalencia = mvarRecargoEquivalencia
End Property


'---------------------------------------
'ANTIGUOS IVAS
'---------------------------------------
Public Property Get TipoIVAAntre1() As Integer
     TipoIVAAntre1 = mvarTipoIVAAntre1
End Property

Public Property Let TipoIVAAntre1(ByVal vData As Integer)
     mvarTipoIVAAntre1 = vData
End Property
Public Property Get TipoIVAAntre2() As Integer
     TipoIVAAntre2 = mvarTipoIVAAntre2
End Property

Public Property Let TipoIVAAntre2(ByVal vData As Integer)
     mvarTipoIVAAntre2 = vData
End Property


Public Property Get TipoIVAAnt1() As Integer
     TipoIVAAnt1 = mvarIVAAnt1
End Property

Public Property Let TipoIVAAnt1(ByVal vData As Integer)
     mvarIVAAnt1 = vData
End Property

Public Property Get TipoIVAAnt2() As Integer
     TipoIVAAnt2 = mvarIVAAnt2
End Property

Public Property Let TipoIVAAnt2(ByVal vData As Integer)
     mvarIVAAnt2 = vData
End Property

Public Property Get FecCambioIva() As Date
     FecCambioIva = mvarFecCambioIva
End Property

Public Property Let FecCambioIva(ByVal vData As Date)
     mvarFecCambioIva = vData
End Property




'RETENCION Y REGIMEN ESPECIAL AGRARIO
'---------------------------------------
Public Property Get IVA_REA() As Integer
     IVA_REA = mvarIVAREA
End Property

Public Property Let IVA_REA(ByVal vData As Integer)
     mvarIVAREA = vData
End Property

Public Property Get PorReten() As Currency
     PorReten = mvarPorReten
End Property

Public Property Let PorReten(ByVal vData As Currency)
     mvarPorReten = vData
End Property

Public Property Get CtaReten() As String
     CtaReten = mvarCtaReten
End Property

Public Property Let CtaReten(ByVal vData As String)
     mvarCtaReten = vData
End Property


'IVA exento
Public Property Get IVA_Exento2() As Integer
     IVA_Exento2 = mvarIVA_Exento
End Property

Public Property Let IVA_Exento2(ByVal vData As Integer)
     mvarIVA_Exento = vData
End Property



'IVA intracomunitario
Public Property Get IVA_Intracomunitario() As Integer
     IVA_Intracomunitario = mvarIVA_Intracomunitario
End Property

Public Property Let IVA_Intracomunitario(ByVal vData As Integer)
     mvarIVA_Intracomunitario = vData
End Property



'Contabilizacioon del Ticket del TPV a grupado
Public Property Get ContabilizarTicketAgrupados() As Boolean
     ContabilizarTicketAgrupados = mvarContabilizarTicketAgrupados
End Property

Public Property Let ContabilizarTicketAgrupados(ByVal vData As Boolean)
     mvarContabilizarTicketAgrupados = vData
End Property


'Numero contabilidad para las albaranes /facturas en B
Public Property Get ContabilidadB() As Byte
     ContabilidadB = mvarContabilidadB
End Property

Public Property Let ContabilidadB(ByVal vData As Byte)
     mvarContabilidadB = vData
End Property


'APortacion al terminal. Ej. Facturas de mostrador en RADIVAL
Public Property Get ctaAportacion() As String
     ctaAportacion = mvarctaaportacion
End Property

Public Property Let ctaAportacion(ByVal vData As String)
     mvarctaaportacion = vData
End Property


'Tiene produccion
Public Property Get Produccion() As Boolean
     Produccion = mvarProduccion
End Property

Public Property Let Produccion(ByVal vData As Boolean)
     mvarProduccion = vData
End Property

'Tiene Descriptores
Public Property Get Descriptores() As Boolean
     Descriptores = mvarDescriptores
End Property

Public Property Let Descriptores(ByVal vData As Boolean)
     mvarDescriptores = vData
End Property


'Campo observaciones en factura
Public Property Get ObsFactura() As Byte
     ObsFactura = mvarObsFactura
End Property

Public Property Let ObsFactura(ByVal vData As Byte)
     mvarObsFactura = vData
End Property


'Articulo de Gastos de administracion
Public Property Get ArtGastosAdmon() As String
     ArtGastosAdmon = mvarArtGastosAdmon
End Property

Public Property Let ArtGastosAdmon(ByVal vData As String)
     mvarArtGastosAdmon = vData
End Property



Public Property Get ArtPortes() As String
     ArtPortes = mvarArtPortes
End Property

Public Property Let ArtPortes(ByVal vData As String)
     mvarArtPortes = vData
End Property

Public Property Get ImporteMinimo() As Currency
     ImporteMinimo = mvarImporteMinimo
End Property

Public Property Let ImporteMinimo(ByVal vData As Currency)
     mvarImporteMinimo = vData
End Property
Public Property Get AbonoKilos() As Currency
     AbonoKilos = mvarAbonoKilos
End Property

Public Property Let AbonoKilos(ByVal vData As Currency)
     mvarAbonoKilos = vData
End Property


Public Property Get ComponentePorcentaje() As Boolean
     ComponentePorcentaje = mvarComponentePorcentaje
End Property

Public Property Let ComponentePorcentaje(ByVal vData As Boolean)
     mvarComponentePorcentaje = vData
End Property


' ---- [14/09/2009] (LAURA)
Public Property Get DtoxCantidad() As Boolean
     DtoxCantidad = mvarDtoxCanti
End Property

Public Property Let DtoxCantidad(ByVal vData As Boolean)
     mvarDtoxCanti = vData
End Property

' ---- [14/09/2009] (LAURA)
'David cambia ENERO 2010
'                               David cambia a byte
'                               0.- NO crea
'                               1.- Solo en tarifa ppal
'                               2.- En todas tarifas

Public Property Get CreaTarifasArticulo() As Byte
     CreaTarifasArticulo = mvarCreaTarifArt2
End Property

Public Property Let CreaTarifasArticulo(ByVal vData As Byte)
     mvarCreaTarifArt2 = vData
End Property


' ---- [19/10/2009] [LAURA]: añadir campo modo analitica  0=trabajador, 1=Familia, 2=Proyecto
Public Property Get ModoAnalitica() As Byte
     ModoAnalitica = mvarModoAnalit
End Property

Public Property Let ModoAnalitica(ByVal vData As Byte)
     mvarModoAnalit = vData
End Property



' ---- [31/12/2009] (DAVID)
Public Property Get EnvioDesdeOutlook() As Boolean
     EnvioDesdeOutlook = mvarEnvioDesdeOutlook
End Property

Public Property Let EnvioDesdeOutlook(ByVal vData As Boolean)
     mvarEnvioDesdeOutlook = vData
End Property



'Febrero 2010
Public Property Get ExeEnvioMail() As String
     ExeEnvioMail = mvarExeEnvioMail
End Property

Public Property Let ExeEnvioMail(ByVal vData As String)
     mvarExeEnvioMail = vData
End Property

Public Property Get DiasGarantia() As Integer
     DiasGarantia = mvarDiasGarantia
End Property

Public Property Let DiasGarantia(ByVal vData As Integer)
     mvarDiasGarantia = vData
End Property




' ---- [13/04/2010] (DAVID) De momento NO tiene LET
Public Property Get TieneCRM() As Boolean
     TieneCRM = mvarTieneCRM
End Property

'Public Property Let TieneCRM(ByVal vData As Boolean)
'     mvarTieneCRM = vData
'End Property

Public Property Get Cooperativa() As Byte
     Cooperativa = mvarCooperativa
End Property

Public Property Let Cooperativa(ByVal vData As Byte)
     mvarCooperativa = vData
End Property

Public Property Let PrecioxDistancia(ByVal vData As Single)
     mvarPrecioxDistancia = vData
End Property

Public Property Get PrecioxDistancia() As Single
     PrecioxDistancia = mvarPrecioxDistancia
End Property

Public Property Let PrecioxTpoEspera(ByVal vData As Single)
     mvarPrecioxTpoEspera = vData
End Property

Public Property Get PrecioxTpoEspera() As Single
     PrecioxTpoEspera = mvarPrecioxTpoEspera
End Property

'[Monica]05/09/2012: Añadimos el path en parametros
'Facturacion Electrónica FacturaE
Public Property Get PathFacturaE() As String
     PathFacturaE = mvarPathFacturaE
End Property

Public Property Let PathFacturaE(ByVal vData As String)
     mvarPathFacturaE = vData
End Property

'[Monica]28/09/2012: Impresora de Tarjetas
Public Property Let ImpresoraTarjetas(ByVal vData As String)
     mvarImpresoraTarjetas = vData
End Property

Public Property Get ImpresoraTarjetas() As String
     ImpresoraTarjetas = mvarImpresoraTarjetas
End Property

Public Property Let ArtRetenciones(ByVal vData As String)
     mvarArtRetenciones = vData
End Property

Public Property Get ArtRetenciones() As String
     ArtRetenciones = mvarArtRetenciones
End Property


'[Monica]13/11/2017: articulo de suplidos
Public Property Let ArtSuplidos(ByVal vData As String)
     mvarArtSuplidos = vData
End Property

Public Property Get ArtSuplidos() As String
     ArtSuplidos = mvarArtSuplidos
End Property


'[Monica]02/05/2017: si hay contabilidad nueva
Public Property Get ContabilidadNueva() As Boolean
     ContabilidadNueva = mvarContabilidadNueva
End Property



'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------

'****   LEER  ***
Public Function Leer() As Byte
Dim RS As ADODB.Recordset
Dim Precio As Currency
On Error GoTo Err3
    
    Sql = "SELECT * FROM spara1 "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic

    If RS.EOF Then
        Leer = 1
    Else
        mvarTipoDtos = RS!TipoDtos
        mvarOrdenDtos = RS!ordendto
        mvarCodTarif = RS!codTarif
        mvarCajaComp = RS!cajacomp  'Cajas completas precios (S/N)
        mvarPreuKmCli = DBLet(RS!preukmcl, "N") 'Precio Km Desplazamiento Clientes
        mvarPreuKmTec = DBLet(RS!preukmtc, "N") 'Precio Km Desplazamiento Técnicos
        mvarnomGerente = DBLet(RS!nomgeren, "T") 'Nombre del Director Gerente
        mvarnomAdmin = DBLet(RS!nomadmin, "T") 'Nombre del Responsable Administración
        mvarHayMante = RS!haymante  'Si hay Mantenimiento (S/N)
        mvarHayRepar = RS!hayrepar  'Si hay Reparaciones (S/N)
        mvarHayFrecu = RS!hayfrecu  'Si hay Frecuencias (S/N)
        mvarHayServicio = RS!hayservicio  'Si hay Servicios (S/N)
        mvarHayDepar = RS!haydepar  'Si hay Departamento o Direccion
        mvarCtrStock = RS!CtrStock  'Si hay Control de Stock (S/N)
        mvarInventar = RS!inventar
        mvarHaynserie = RS!haynserie 'Si hay Nº Serie en Compras (S/N)
        mvarDiasSiMan = RS!diassiman 'Dias reparacion si hay Mantenimiento
        mvarDiasNoMan = RS!diasnoman 'Dias reparacion si NO hay Mantenimiento
        Me.ForPagoChequeRegalo = DBLet(RS!codforpa, "T")
        mvarRaiz_Cta_Soc_publi = DBLet(RS!Raiz_Cta_Soc_publi, "T") 'raiz cuenta socios publicidad
        mvarRaiz_Cta_Soc_liqui = DBLet(RS!Raiz_Cta_Soc_Liqui, "T") 'raiz cuenta socios liquidacion
        mvarRaiz_Cta_Soc_Equip = DBLet(RS!Raiz_Cta_Soc_Equip, "T") 'raiz cuenta socios facturas de venta(equipamiento)
        mvarRaiz_CtaClien_Soc = DBLet(RS!Raiz_CtaClien_Soc, "T") 'raiz cuenta cliente de socios(cuotas)
        mvarRaiz_Cta_Reten_Soc = DBLet(RS!Raiz_CtaReten, "T") 'raiz cuenta retencion socios liquidacion
        mvarCtaLiquidacion = DBLet(RS!CtaLiquidacion, "T")  'cuenta base de liquidacion de socios
                
        ' Datos de alta de socio
        mvarCtaTituloAlta = DBLet(RS!CtaTituloAlta, "T")  'cuenta titulo alta de socios
        mvarCtaGastoAlta = DBLet(RS!CtaGastoAlta, "T")  'cuenta gasto alta de socios
        mvarImpTituloAlta = DBLet(RS!ImpTituloAlta, "N")  'importe titulo alta de socios
        mvarImpGastoAlta = DBLet(RS!ImpGastoAlta, "N")  'importe gasto alta de socios
        mvarRaiz_CtaAltaSoc = DBLet(RS!Raiz_CtaAltaSoc, "T") 'raiz cuenta alta de socios(cuotas)
                
                
        'Datos Facturacion
        mvarCodArticDes = DBLet(RS!codartid, "T") 'articulo para facturar mantenimientos
        mvarDiaPago1 = DBLet(RS!DiaPago1, "N") 'dia pago 1 para compras
        mvarDiaPago2 = DBLet(RS!DiaPago2, "N") 'dia pago 2 para compras
        mvarDiaPago3 = DBLet(RS!DiaPago3, "N") 'dia pago 3 para compras
        mvarMesnogir = DBLet(RS!mesnogir, "N") 'mes a no girar para compras
        mvarConFactuPubli = DBLet(RS!ConFactuPubli, "T") 'concepto factura publicidad
        
        'Datos de Internet
        mvarDireMail = DBLet(RS!DireMail, "T") 'Direccion de email
        mvarSMTPhost = DBLet(RS!SMTPhost, "T") 'SMTP host para el envio de email
        mvarSMTPuser = DBLet(RS!SMTPuser, "T") 'SMTP usuario para el envio de email
        mvarSMTPpass = DBLet(RS!SMTPpass, "T") 'SMTP password para el envio de email
        mvarWebSoporte = DBLet(RS!WebSoporte, "T") 'Web de soporte
        mvarMailSoporte = DBLet(RS!MailSoporte, "T") 'email de soporte
        mvarWebVersion = DBLet(RS!WebVersion, "T") 'Version de la Web
        
        
        'Datos Contabilidad
        mvarSerConta = DBLet(RS!serconta, "T") 'servidor conta
        mvarUsuConta = DBLet(RS!UsuConta, "T") 'usuario conta
        mvarPasConta = DBLet(RS!pasconta, "T") 'password
        mvarNumConta = DBLet(RS!NumConta, "N") 'numero conta
        
        
        'Valores por defecto
        mvarDefActiv = DBLet(RS!defactividad, "N") 'numero conta
        mvarDefEnvio = DBLet(RS!defenvio, "N") 'numero conta
        mvarDefZona = DBLet(RS!defzona, "N") 'numero conta
        mvarDefSitu = DBLet(RS!defstituacion, "N") 'numero conta
        mvarDefTarifa = DBLet(RS!deftarifa, "N") 'numero conta
        mvarDefAgente = DBLet(RS!defagente, "N") 'numero conta

        mvaravipedcli = DBLet(RS!avipedcli, "N")
        mvaravipedpro = DBLet(RS!avipedpro, "N")
        mvaravialbcli = DBLet(RS!avialbcli, "N")
        mvaravialbpro = DBLet(RS!avialbpro, "N")
        mvaravimanteni = DBLet(RS!avimanteni, "N")
        mvaravirepara = DBLet(RS!avirepara, "N")
        mvaraviavisos = DBLet(RS!aviavios, "N")
        PonVariableTienAvisos

        
        mvarTipoIVAre1 = DBLet(RS!IVAre1eq, "N")
        mvarTipoIVAre2 = DBLet(RS!IVAre2eq, "N")
        mvarTipoIVAre3 = DBLet(RS!IVAre3eq, "N")
        mvarRecargoEquivalencia = (mvarTipoIVAre1 > 0 Or mvarTipoIVAre2 > 0 Or mvarTipoIVAre3 > 0)
        
        mvarIVA1 = DBLet(RS!IVAre1, "N")
        mvarIVA2 = DBLet(RS!IVAre2, "N")
        mvarIVA3 = DBLet(RS!IVAre3, "N")


    '[Monica]11/10/2012: Antiguos tipos de iva
        mvarTipoIVAAntre1 = DBLet(RS!IVAAnt1eq, "N")
        mvarTipoIVAAntre2 = DBLet(RS!IVAAnt2eq, "N")
        
        mvarIVAAnt1 = DBLet(RS!IVAAnt1, "N")
        mvarIVAAnt2 = DBLet(RS!IVAAnt2, "N")

        mvarFecCambioIva = DBLet(RS!FecCambioIva, "F")

    '
        'El articulo de telefonia no lo pongo como parametros
        'codartictel
        mvarCodarticTfnia = DBLet(RS!codartictel, "T") 'numero conta
                
        mvarArtAlquiler = DBLet(RS!ArtAlquiler, "T") 'articulo para alquiler de equipos
        mvarArtServCuotas = DBLet(RS!ArtServCuotas, "T") 'articulo de servicios para cuotas normales
        mvarArtCuotaSinChofer = DBLet(RS!artcuotasinchof, "T") 'articulo para cuota normal sin chofer
        mvarArtCuotaConChofer = DBLet(RS!artcuotaconchof, "T") 'articulo para cuota normal con chofer
        mvarArtCuotaExtraor = DBLet(RS!artcuotaext, "T") 'articulo para cuota extraordinaria
        
        Precio = DevuelveValor("select preciove from sartic where codartic = " & DBSet(RS!artcuotasinchof, "T"))
        mvarPresinchofe = DBLet(Precio, "N")
        
        Precio = DevuelveValor("select preciove from sartic where codartic = " & DBSet(RS!artcuotaconchof, "T"))
        mvarPreconchofe = DBLet(Precio, "N")
        
        Precio = DevuelveValor("select preciove from sartic where codartic = " & DBSet(RS!ArtServCuotas, "T"))
        mvarPreporserv = DBLet(Precio, "N")
        
        Precio = DevuelveValor("select preciove from sartic where codartic = " & DBSet(RS!ArtAlquiler, "T"))
        mvarPreporAlq = DBLet(Precio, "N") ' precio por alquiler
                
        'Retencion
        mvarIVAREA = DBLet(RS!IVA_REA, "N")
        mvarPorReten = DBLet(RS!PorReten, "N")
        mvarCtaReten = DBLet(RS!CtaReten, "T") 'cuenta retencion proveedores
                
                
        'IVA exento
        mvarIVA_Exento = DBLet(RS!ivaexento, "N")
        
        
        'Contabilidad
        mvarContabilidadB = DBLet(RS!conta_B, "N")
        
        'Contabiliza los tickets resumidos
        mvarContabilizarTicketAgrupados = (DBLet(RS!conttickagrupado, "N") = 1)
        
        'Cta aportacion
        mvarctaaportacion = DBLet(RS!ctaAportacion, "T")
        
        'produccion
        mvarProduccion = DBLet(RS!Produccion, "N") = 1
        'descriptores
        mvarDescriptores = DBLet(RS!Descriptores, "N") = 1
        'obs factura
        mvarObsFactura = DBLet(RS!ObsFactura, "N")
        '
        mvarIVA_Intracomunitario = DBLet(RS!IvaIntracom, "N")
        'Articulo de gastos de administracion
        mvarArtGastosAdmon = DBLet(RS!ArtReciclado, "T")

        mvarArtPortes = DBLet(RS!ArticuloPortes, "T")
        mvarImporteMinimo = DBLet(RS!impminped, "N")
        mvarAbonoKilos = DBLet(RS!AbonoKilos, "N")
    
        mvarComponentePorcentaje = DBLet(RS!compoporcen, "N") = 1

        ' ---- [14/09/2009] (LAURA)
        Me.DtoxCantidad = IIf(RS!dtoxcanti = 1, True, False)
        Me.CreaTarifasArticulo = DBLet(RS!creatarifart, "N")
        ' ----
        
        ' ---- [19/10/2009] [LAURA]: añadir campo modo analitica
        Me.ModoAnalitica = RS!modanalitica
        
        
        'David 31/12/09
        Me.EnvioDesdeOutlook = DBLet(RS!EnvioDesdeOutlook, "N") = 1
        
        'David Febrero 2010
        ExeEnvioMail = DBLet(RS!arigesmail, "T")
        DiasGarantia = DBLet(RS!diasgaranrepa, "N")
        
        mvarTieneCRM = DBLet(RS!TieneCRM, "N") = 1
        mvarCooperativa = DBLet(RS!Cooperativa, "N")
        mvarPrecioxDistancia = DBLet(RS!PrecioxDistancia, "N")
        mvarPrecioxTpoEspera = DBLet(RS!PrecioxTpoEspera, "N")
        
        'FacturaE
        mvarPathFacturaE = DBLet(RS!PathFacturaE, "T")
        'Impresora de Tarjetas
        mvarImpresoraTarjetas = DBLet(RS!impretarjeta, "T")
        
        
        mvarArtRetenciones = DBLet(RS!ArtRetenciones, "T") 'articulo de retenciones
        
        '[Monica]02/05/2017: nueva contabilidad
        
        mvarContabilidadNueva = (DBLet(RS!nuevacontabilidad, "N") = 1)
        
        '[Monica]13/11/2017: articulo de suplidos
        mvarArtSuplidos = DBLet(RS!ArtSuplidos, "T") ' articulo de suplidos
        
        
        Leer = 0
        
        ComprobarProgramaEnvioMail
        
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error(Leer Parametros)" & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Set RS = Nothing
    Leer = 1
End Function


Public Function Modificar(Codigo As Byte) As Byte
'Modifica los parametros de la Aplicación
'Modifica la Tabla: spara1, BD: Aritaxi
On Error GoTo EModificar

    Sql = "UPDATE spara1 SET "
    Sql = Sql & " tipodtos= " & mvarTipoDtos
    Sql = Sql & ", ordendto= " & mvarOrdenDtos
    Sql = Sql & ", codtarif= " & mvarCodTarif
    Sql = Sql & ", cajacomp= " & DBSet(mvarCajaComp, "B")
    Sql = Sql & ", preukmcl= " & DBSet(mvarPreuKmCli, "N")
'    SQL = SQL & ", presinchofe= " & DBSet(mvarPresinchofe, "N")
'    SQL = SQL & ", preconchofe= " & DBSet(mvarPreconchofe, "N")
'    SQL = SQL & ", preporserv= " & DBSet(mvarPreporserv, "N")
'    SQL = SQL & ", preporalq= " & DBSet(mvarPreporAlq, "N")
    Sql = Sql & ", preukmtc= " & DBSet(mvarPreuKmTec, "N")
    Sql = Sql & ", nomgeren= " & DBSet(mvarnomGerente, "T")
    Sql = Sql & ", nomadmin= " & DBSet(mvarnomAdmin, "T")
    Sql = Sql & ", haymante = " & DBSet(mvarHayMante, "B")
    Sql = Sql & ", hayrepar = " & DBSet(mvarHayRepar, "B")
    Sql = Sql & ", hayfrecu = " & DBSet(mvarHayFrecu, "B")
    Sql = Sql & ", hayservicio = " & DBSet(mvarHayServicio, "B")
    Sql = Sql & ", haydepar = " & DBSet(mvarHayDepar, "B")
    Sql = Sql & ", ctrstock = " & DBSet(mvarCtrStock, "B")
    Sql = Sql & ", inventar= " & DBSet(mvarInventar, "B")
    Sql = Sql & ", haynserie= " & DBSet(mvarHaynserie, "B")
    Sql = Sql & ", diassiman= " & DBSet(mvarDiasSiMan, "N")
    Sql = Sql & ", diasnoman= " & DBSet(mvarDiasNoMan, "N")
    Sql = Sql & ", codartid= " & DBSet(mvarCodArticDes, "T")
    Sql = Sql & ", diapago1= " & DBSet(mvarDiaPago1, "N")
    Sql = Sql & ", diapago2= " & DBSet(mvarDiaPago2, "N")
    Sql = Sql & ", diapago3= " & DBSet(mvarDiaPago3, "N")
    Sql = Sql & ", mesnogir= " & DBSet(mvarMesnogir, "N")
    Sql = Sql & ", codforpa= " & DBSet(mvarCodForpa, "N", "S")
    Sql = Sql & ", diremail= " & DBSet(mvarDireMail, "T")
    Sql = Sql & ", smtphost= " & DBSet(mvarSMTPhost, "T")
    Sql = Sql & ", smtpuser= " & DBSet(mvarSMTPuser, "T")
    Sql = Sql & ", smtppass= " & DBSet(mvarSMTPpass, "T")
    Sql = Sql & ", websoporte= " & DBSet(mvarWebSoporte, "T")
    Sql = Sql & ", mailsoporte= " & DBSet(mvarMailSoporte, "T")
    Sql = Sql & ", webversion= " & DBSet(mvarWebVersion, "T")
    Sql = Sql & ", serconta= " & DBSet(mvarSerConta, "T")
    Sql = Sql & ", usuconta= " & DBSet(mvarUsuConta, "T")
    Sql = Sql & ", pasconta= " & DBSet(mvarPasConta, "T")
    Sql = Sql & ", numconta= " & DBSet(mvarNumConta, "N", "S")
    Sql = Sql & ", confactupubli= " & DBSet(mvarConFactuPubli, "T")
    
    'Valore por defecto
    Sql = Sql & ", defactividad= " & DBSet(mvarDefActiv, "N", "N")
    Sql = Sql & ", defenvio= " & DBSet(mvarDefEnvio, "N", "N")
    Sql = Sql & ", defzona= " & DBSet(mvarDefZona, "N", "N")
    Sql = Sql & ", defstituacion= " & DBSet(mvarDefSitu, "N", "N")
    Sql = Sql & ", deftarifa= " & DBSet(mvarDefTarifa, "N", "N")
    Sql = Sql & ", defagente= " & DBSet(mvarDefAgente, "N", "N")
    'Telefonia
    Sql = Sql & ", codartictel= " & DBSet(mvarCodarticTfnia, "T", "S")
    ' Articulo de alquiler de equipos
    Sql = Sql & ", artalquiler= " & DBSet(mvarArtAlquiler, "T")
    ' Articulo de servicios de cuotas
    Sql = Sql & ", artservcuotas= " & DBSet(mvarArtServCuotas, "T")
    ' Articulo de cuotas normal sin chofer
    Sql = Sql & ", artcuotasinchof= " & DBSet(mvarArtCuotaSinChofer, "T", "N")
    ' Articulo de cuotas normal con chofer
    Sql = Sql & ", artcuotaconchof= " & DBSet(mvarArtCuotaConChofer, "T", "N")
    ' Articulo de cuotas extraordinarias
    Sql = Sql & ", artcuotaext= " & DBSet(mvarArtCuotaExtraor, "T", "N")
      
      
    'Avisos
    Sql = Sql & ", avipedcli= " & DBSet(mvaravipedcli, "N", "S")
    Sql = Sql & ", avipedpro= " & DBSet(mvaravipedpro, "N", "S")
    Sql = Sql & ", avialbcli= " & DBSet(mvaravialbcli, "N", "S")
    Sql = Sql & ", avialbpro= " & DBSet(mvaravialbpro, "N", "S")
    Sql = Sql & ", avimanteni= " & DBSet(mvaravimanteni, "N", "S")
    Sql = Sql & ", avirepara= " & DBSet(mvaravirepara, "N", "S")
    Sql = Sql & ", aviavios= " & DBSet(mvaraviavisos, "N", "S")
      
    PonVariableTienAvisos

    Sql = Sql & ", IVAre1eq= " & DBSet(mvarTipoIVAre1, "N", "S")
    Sql = Sql & ", IVAre2eq= " & DBSet(mvarTipoIVAre2, "N", "S")
    Sql = Sql & ", IVAre3eq= " & DBSet(mvarTipoIVAre3, "N", "S")
    mvarRecargoEquivalencia = (mvarTipoIVAre1 > 0 Or mvarTipoIVAre2 > 0 Or mvarTipoIVAre3 > 0)
    
    Sql = Sql & ", IVAre1= " & DBSet(mvarIVA1, "N", "S")
    Sql = Sql & ", IVAre2= " & DBSet(mvarIVA2, "N", "S")
    Sql = Sql & ", IVAre3= " & DBSet(mvarIVA3, "N", "S")

    '[Monica]11/10/2012: antiguos ivas
    Sql = Sql & ", IVAAnt1eq= " & DBSet(mvarTipoIVAAntre1, "N", "S")
    Sql = Sql & ", IVAAnt2eq= " & DBSet(mvarTipoIVAAntre2, "N", "S")
    
    Sql = Sql & ", IVAant1= " & DBSet(mvarIVAAnt1, "N", "S")
    Sql = Sql & ", IVAant2= " & DBSet(mvarIVAAnt2, "N", "S")



    'Retencion y REA
    Sql = Sql & ", IVA_REA = " & DBSet(mvarIVAREA, "N", "S")
    Sql = Sql & ", raiz_CtaReten = " & DBSet(mvarRaiz_Cta_Reten_Soc, "T", "S")
    Sql = Sql & ", CtaReten = " & DBSet(mvarCtaReten, "T", "S")
    Sql = Sql & ", PorReten = " & DBSet(mvarPorReten, "N", "S")
    Sql = Sql & ", CtaLiquidacion = " & DBSet(mvarCtaLiquidacion, "T", "S")


    ' datos de alta de socio
    Sql = Sql & ", Ctatituloalta  = " & DBSet(mvarCtaTituloAlta, "T", "S")
    Sql = Sql & ", Ctagastoalta  = " & DBSet(mvarCtaGastoAlta, "T", "S")
    Sql = Sql & ", imptituloalta  = " & DBSet(mvarImpTituloAlta, "N", "S")
    Sql = Sql & ", impgastoalta  = " & DBSet(mvarImpGastoAlta, "N", "S")
    


    'IVA exento
    Sql = Sql & ", ivaexento = " & DBSet(mvarIVA_Exento, "N", "S")

    'Tickets agrupados
    Sql = Sql & ", conttickagrupado = " & DBSet(mvarContabilizarTicketAgrupados, "B")

    'Contabilidad "B". Radival
    Sql = Sql & ", conta_B  = " & DBSet(mvarContabilidadB, "N")

    'Ctaa aportacion termial
    Sql = Sql & ", ctaaportacion  = " & DBSet(mvarctaaportacion, "T")
    'Produccion
    Sql = Sql & ", Produccion  = " & DBSet(mvarProduccion, "B")
    'Descriptores
    Sql = Sql & ", descriptores = " & DBSet(mvarDescriptores, "B")
    'Observaciones factura proveedor/cliente
    Sql = Sql & ", ObsFactura  = " & DBSet(mvarObsFactura, "N")
    Sql = Sql & ", IvaIntracom  = " & DBSet(mvarIVA_Intracomunitario, "N")

    'Articulo de Gastos de Administracion
    Sql = Sql & ", ArtReciclado  = " & DBSet(mvarArtGastosAdmon, "T", "S")
    

    'FONTENAS. Portes
    Sql = Sql & ", ArticuloPortes  = " & DBSet(mvarArtPortes, "T", "S")
    Sql = Sql & ", impminped  = " & DBSet(mvarImporteMinimo, "N", "S")
    Sql = Sql & ", abonokilos  = " & DBSet(mvarAbonoKilos, "N", "S")
    Sql = Sql & ", compoporcen = " & DBSet(mvarComponentePorcentaje, "B")
    
    ' ---- [14/09/2009] (LAURA)
    Sql = Sql & ", dtoxcanti = " & DBSet(Me.DtoxCantidad, "B")
    Sql = Sql & ", creatarifart = " & DBSet(Me.CreaTarifasArticulo, "B")
    '-----------------
    
    ' ---- [19/10/2009] [LAURA]: añadir campo modo analitica
    Sql = Sql & ", modanalitica = " & DBSet(Me.ModoAnalitica, "N")
    
    
    ' ---- [31/12/2009] [DAVID]
    Sql = Sql & ", EnvioDesdeOutlook = " & DBSet(Me.EnvioDesdeOutlook, "B")
    
    
    ' ---- [02/02/2010] [DAVID]
    Sql = Sql & ", arigesmail = " & DBSet(ExeEnvioMail, "T", "S")
    ' ---- [10/02/2010] [DAVID]
    Sql = Sql & ", diasgaranrepa = " & DBSet(Me.DiasGarantia, "N", "S")
    'cuenta raiz socio publicidad
    Sql = Sql & ",Raiz_Cta_Soc_publi= " & DBSet(mvarRaiz_Cta_Soc_publi, "T")
    
    'cuenta raiz socio liquidacion
    Sql = Sql & ",Raiz_Cta_Soc_liqui= " & DBSet(mvarRaiz_Cta_Soc_liqui, "T")
    
    'cuenta raiz cta cliente socio (cuotas)
    Sql = Sql & ",Raiz_Ctaclien_Soc= " & DBSet(mvarRaiz_CtaClien_Soc, "T")
    
    'cuenta raiz cta cliente socio (cuotas)
    Sql = Sql & ",raiz_ctaaltasoc= " & DBSet(mvarRaiz_CtaAltaSoc, "T")
    
    'cuenta raiz socio facturas de venta (equipamiento)
    Sql = Sql & ",Raiz_Cta_Soc_equip= " & DBSet(mvarRaiz_Cta_Soc_Equip, "T")
    
    ' cooperativa VIP
    Sql = Sql & ",cooperativa= " & DBSet(mvarCooperativa, "N")
    ' precio por distancia VIP
    Sql = Sql & ",precioxdistancia= " & DBSet(mvarPrecioxDistancia, "N")
    ' precio por tpo de espera VIP
    Sql = Sql & ",precioxtpoespera= " & DBSet(mvarPrecioxTpoEspera, "N")
    
    'FacturaE
    Sql = Sql & ", pathfacturae = " & DBSet(mvarPathFacturaE, "T")
    
    'Impresora de Tarjetas
    Sql = Sql & ", impretarjeta = " & DBSet(mvarImpresoraTarjetas, "T")
    
    ' Articulo de retenciones
    Sql = Sql & ", artretenciones= " & DBSet(mvarArtRetenciones, "T")
    
    ' Articulo de suplidos
    Sql = Sql & ", artsuplidos= " & DBSet(mvarArtSuplidos, "T")
    
    
    'where
    Sql = Sql & " WHERE codigo =" & Codigo & ";"

    conn.Execute Sql



    'Desde los ivas de retencion, hay una tabla(sivatra) que tiene la relacion entre
    ' iva sin retencion, iva con retencion
    'Entocnes cuando modifico, borro las entradas y las vuevlo a generer
    CargarTablaTraduccionIVAS
    
    
    'Todo bien
    Modificar = 1

    Exit Function

EModificar:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 0
End Function

Private Sub PonVariableTienAvisos()
    mvarTieneAvisos = (mvaravipedcli + mvaravipedpro + mvaravialbcli + mvaravialbpro + mvaravimanteni _
        + mvaravirepara + mvaraviavisos > 0)
End Sub


Private Function MovimientosParaServicios() As Boolean
    '-- Esta función comprueba si los movimientos ALS y FAS están creados
    '   este es un requisito previo si se desa activar la opción de servicios
    Dim Sql As String
    Dim RS As ADODB.Recordset
    Sql = "select count(*) from stipom where codtipom = 'ALS' or codtipom = 'FAS'"
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly
    If RS.Fields(0) = 2 Then
        MovimientosParaServicios = True
    Else
        MovimientosParaServicios = False
    End If
End Function



Public Function DevuleveTipoIVA_RE(TipoIVAOrigen As Integer) As Integer
Dim devuelve As Integer
    
    'IVA con recargo de equivalencia
    devuelve = TipoIVAOrigen
    If TipoIVAOrigen = TipoIVA1 Then
        devuelve = TipoIVAre1
    ElseIf TipoIVAOrigen = vParamAplic.TipoIVA2 Then
        devuelve = TipoIVAre2
    ElseIf TipoIVAOrigen = TipoIVA3 Then
        devuelve = TipoIVAre3
    Else
        MsgBox "NO se ha encontrado la correspondencia para el IVA RE: " & TipoIVAOrigen, vbExclamation
    End If
    DevuleveTipoIVA_RE = devuelve
End Function


Public Function DevuleveTipoIVA_Ant(ConRE As Boolean, TipoIVAOrigen As Integer) As Integer
Dim devuelve As Integer
    
    'IVA con recargo de equivalencia
    If ConRE Then
        devuelve = TipoIVAOrigen
        If TipoIVAOrigen = TipoIVAre1 Then
            devuelve = TipoIVAAntre1
        ElseIf TipoIVAOrigen = vParamAplic.TipoIVAre2 Then
            devuelve = TipoIVAAntre2
        Else
            MsgBox "NO se ha encontrado la correspondencia para el IVA Antiguo: " & TipoIVAOrigen, vbExclamation
        End If
    Else
        devuelve = TipoIVAOrigen
        If TipoIVAOrigen = TipoIVA1 Then
            devuelve = TipoIVAAnt1
        ElseIf TipoIVAOrigen = vParamAplic.TipoIVA2 Then
            devuelve = TipoIVAAnt2
        Else
            MsgBox "NO se ha encontrado la correspondencia para el IVA Antiguo: " & TipoIVAOrigen, vbExclamation
        End If
    End If
    DevuleveTipoIVA_Ant = devuelve
End Function







Private Sub CargarTablaTraduccionIVAS()
Dim RS As ADODB.Recordset

    On Error GoTo ECargarTablaTraduccionIVAS
    If ConnConta Is Nothing Then
        MsgBox "Reinicie ARITAXI y vuelva a modificar los parametros para cargar los tipos de IVA", vbExclamation
        Exit Sub
    End If
    
    Sql = "DELETE FROM sivatra"
    conn.Execute Sql
    
    
    Set RS = New ADODB.Recordset
    Sql = "select codigiva from tiposiva"
    RS.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Sql = "INSERT INTO sivatra VALUES (" & RS!Codigiva & "," & RS!Codigiva & ")"
        
        conn.Execute Sql
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    'Ahora acutalizo la relacion
    If mvarTipoIVAre1 > 0 Then
        Sql = "UPDATE sivatra set codigiva2=" & mvarTipoIVAre1 & " WHERE codigiva = " & mvarIVA1
        conn.Execute Sql
    End If
    
    If mvarTipoIVAre2 > 0 Then
        Sql = "UPDATE sivatra set codigiva2=" & mvarTipoIVAre2 & " WHERE codigiva = " & mvarIVA2
        conn.Execute Sql
    End If
        
    If mvarTipoIVAre3 > 0 Then
        Sql = "UPDATE sivatra set codigiva2=" & mvarTipoIVAre3 & " WHERE codigiva = " & mvarIVA3
        conn.Execute Sql
    End If
        
    Exit Sub
ECargarTablaTraduccionIVAS:
    MuestraError Err.Number, Err.Description, "Cargar Tabla Traduccion IVAS"
End Sub



Public Sub ComprobarProgramaEnvioMail()

On Error GoTo EC
    If Me.ExeEnvioMail <> "" Then
        If Dir(App.Path & "\" & Me.ExeEnvioMail, vbArchive) = "" Then
            MsgBox "No existe en este PC el archivo: " & Me.ExeEnvioMail, vbInformation
            ExeEnvioMail = ""
        End If
    End If
    Exit Sub
EC:
    ExeEnvioMail = ""
End Sub
