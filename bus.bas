Attribute VB_Name = "bus"
Option Explicit


Public Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Global I&, J&, K&                             ' Contadores
Global msg$, MsgErr$, NumErr&                 ' Variables de control de error
Global cont%, Opc%, Skn$, SknDir$             ' Otros contadores
Public Tmp%, m_hMod&



Public vUsu As Usuario  'Datos usuario
Public vEmpresa As Cempresa 'Los datos de la empresa
Public vParam As Cparametros  'Parametros Generales de la Empresa (nombre, direc.,...
Public vParamAplic As CParamAplic 'Parametros Aplicaci�n
Public vConfig As Configuracion 'Parametros Configuracion

Public vParamTPV As CParamTPV 'Parametros para el TPV

'LOG de acciones relevantes
Public LOG As cLOG   'Se instancia , se ejecuta LOG.insertar y se elimina :LOG=nothing   Ver ejemplo borre facturas


Public Const NumeroDeDecimales = 2
Public Const PrecioDecimales = 5   'Para ir poniendolo poco a poco

'Formato de fecha
Public FormatoFecha As String
Public FormatoFechaHora As String
Public FormatoImporte As String 'Decimal(12,2)
Public FormatoPrecio As String 'Decimal(10,4)
Public FormatoPrecio2 As String 'Por si podemops parametrizarlo mas adelante
Public FormatoHora As String
Public FormatoHora2 As String

Public FormatoCantidad As String 'Decimal(10,2)
Public FormatoCantidad2 As String 'Decimal(8,2)
Public FormatoDescuento As String 'Decimal(4,2)
Public FormatoKms As String 'Decimal(8,4)
Public FormatoPorcen As String 'Decimal(5,2)

Public CadenaDesdeOtroForm As String


'Conexi�n a la BD Aritaxi de la empresa
Public conn As ADODB.Connection

'Conexi�n a la BD de Usuarios
Public ConnUsuarios As ADODB.Connection

'Conexi�n a la BD de Contabilidad
Public ConnConta As ADODB.Connection

'Que conexion a base de datos se va a utilizar
Public Const conAri As Byte = 1 'Si conAri entonces trabajaremos con conexion conn a la BD ARITAXI
Public Const conConta As Byte = 2 'Si conConta entonces trabajaremos con conexion connConta a la BD CONTA

Public Const vbLightBlue = &HFEEFDA
Public Const vbErrorColor = &HDFE1FF      '&HFFFFC0
Public Const vbMoreLightBlue = &HFEFBD8   ' azul clarito






'Para las formas de pago.  David
Public Const vbFPTransferencia = 1
Public Const vbCrearNuevaCta = "### CREAR CTA CONTAB. ###"


'Global para n� de registro eliminado
Public NumRegElim  As Long

'Para algunos campos de texto suletos controlarlos
'Public miTag As CTag

'Variable para saber si se ha actualizado algun asiento
'Public AlgunAsientoActualizado As Boolean
'Public TieneIntegracionesPendientes As Boolean

'Public miRsAux As ADODB.Recordset

Public AnchoLogin As String  'Para fijar los anchos de columna


'Variables para la nueva forma de leer la scryst
Public pImprimeDirecto As Boolean
Public pPdfRpt As String

'[Monica]28/09/2012: tema de la impresora por defecto de tarjetas
Public ImpresoraDefecto As String

Public teclaBuscar As Integer 'llamada desde prismaticos

Public Const SerieFraPro = "1"

'Calendar
Public TextosLabelEspanol As String
Public UltimaLecturaReminders As Single


Public DireccionAyuda As String

Public ComoValidados As Boolean
Public ContinuarValidar As Boolean



'Inicio Aplicaci�n
Public Sub Main()
Dim T1 As Single

'        frmPrueba.Show vbModal


       Load frmIdentifica
       CadenaDesdeOtroForm = ""
       
       'Necesitaremos el archivo arifon.dat
       
       Set vEmpresa = Nothing
       frmIdentifica.Show vbModal
               
               
               
               
       If CadenaDesdeOtroForm = "" Then
            'NO se ha identificado
            Set conn = Nothing
            End
       End If

'        BuscaEmpresas
       NumeroEmpresaMemorizar True
       LeerDatosEmpresa
       
       If vEmpresa Is Nothing Then Exit Sub
       
       

       
'       CadenaDesdeOtroForm = ""
'       frmLogin.Show vbModal
'
'        If CadenaDesdeOtroForm = "" Then
'            'No ha seleccionado ninguna empresa
'            Set conn = Nothing
'            End
'            Exit Sub
'        End If
        Screen.MousePointer = vbHourglass
'
'        LeerEmpresa 'Carga los Datos de la empresa
'        Carga los Datos B�sicos de la empresa

        'Cerramos la conexion con BD: Usuarios
        conn.Close

        'Abre la conexi�n a BDatos:Aritaxi
        If AbrirConexion() = False Then
            MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
            End
        Else
            'Carga Parametros Generales y Contables de la empresa
            LeerParametros
        End If

        'Abrir conexi�n a la BDatos de Contabilidad para acceder a
        'Tablas: Cuentas, Tipos IVA
        If AbrirConexionConta(False) = False Then
            MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
            End
        End If


        'Carga los Niveles de cuentas de Contabilidad de la empresa y las fechasINICIO y FIN
        LeerNivelesEmpresa

'        'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
        GestionaPC
        
               
        
        'Otras acciones
        OtrasAcciones
         
        frmppal.Show
        
        '[Monica]23/01/2018: fras pendientes de contabilizar
'        If vParamAplic.ContabilidadNueva And (vUsu.Nivel = 0 Or vUsu.Nivel = 1) Then FrasPendientesContabilizar
        If (vUsu.Nivel = 0 Or vUsu.Nivel = 1) Then FrasPendientesContabilizar

'ANTES
'Exit Sub

'       Load frmInicio
'       frmInicio.Show
'       frmInicio.Refresh
'       T1 = Timer
'       Set vConfig = New Configuracion
'       If vConfig.Leer = 1 Then
'            vConfig.SERVER = InputBox("Servidor: ")
'            vConfig.User = InputBox("Usuario: ")
'            vConfig.password = InputBox("Password: ")
''            vConfig.Integraciones = InputBox("Path integraciones: ")
'            vConfig.Grabar
'            MsgBox "Reinicie la contabilidad", vbCritical
'            End
'            Exit Sub
'       End If
 
'
'        If AbrirConexion() = False Then
'            MsgBox "La apliacaci�n no puede continuar sin acceso a los datos. ", vbCritical
'            End
'        End If
        
'        'La llave
'        Load frmLLave
'        If Not frmLLave.ActiveLock1.RegisteredUser Then
'            'No ESTA REGISTRADO
'            frmLLave.Show vbModal
'        Else
'            Unload frmLLave
'        End If
        
        
'
'        'Que se vea un momentito
'        T1 = Timer - T1
'        If T1 < 0.5 Then
'            T1 = 0.5 - T1
'            espera T1
'        End If
        
'        'Descargamos inicio
'        Unload frmInicio
'
'
'        CadenaDesdeOtroForm = ""
'        frmLogin.Show vbModal
'        If vUsu Is Nothing Then
'            'Esto significa que no se ha identifcado como usuario
'            'luego finaliza la aplicacion
'            End
'        End If

'        Screen.MousePointer = vbHourglass

        'Cerramos la conexion
'        Conn.Close

'
'        If AbrirConexion() = False Then
'            MsgBox "La apliacaci�n no puede continuar sin acceso a los datos. ", vbCritical
'            End
'        End If
        
'        LeerEmpresaParametros
        
        'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
'        GestionaPC
        
        'Otras acciones
'        OtrasAcciones
         
'        frmPpal.Show
End Sub


Public Sub FrasPendientesContabilizar()
Dim Sql As String
Dim SQL2 As String
Dim SqlBd As String
Dim SQLinsert As String
Dim RsBd As ADODB.Recordset
Dim BBDD As String

Dim frmMens As frmMensajes

    On Error GoTo eFrasPendientesContabilizar

'[Monica]29/06/2017. quitamos el control de campa�as, en cada campa�a miramos lo que hay
'    If vParamAplic.Cooperativa <> 12 And vParamAplic.Cooperativa <> 9 And vParamAplic.Cooperativa <> 14 Then
'        If vEmpresa.BDAriagro <> "ariagro" Then Exit Sub
'    End If


    '[Monica]18/06/2018: solo en el caso de que sean radio y tele sale el mensaje. En el caso de cordoba no tienen facuras
    If vParamAplic.Cooperativa <> 0 And vParamAplic.Cooperativa <> 3 Then Exit Sub


    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    BBDD = vEmpresa.BDAritaxi
    
    SQLinsert = "insert into " & BBDD & ".tmpinformes (codusu, nombre1,codigo1,nombre2,fecha1, nombre3,importe1) "

    ' fav : facturas de venta socios
    Sql = " select " & vUsu.Codigo & ",'Venta Socios' tipofact, 0, concat(scafac.codtipom,right(concat('0000000',numfactu),7)),fecfactu, nomclien nomsocio, totalfac  from " & BBDD & ".scafac where intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and scafac.codtipom = 'FAV' "
    Sql = Sql & " union "
    
    ' fce : facturas cuotas extraordinarias
    Sql = Sql & " select " & vUsu.Codigo & ",'Cuotas Extraordinarias' tipofact,1, concat(scafac.codtipom,right(concat('0000000',numfactu),7)),fecfactu, nomclien nomsocio, totalfac from " & BBDD & ".scafac where intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and scafac.codtipom = 'FCE' "
    
    ' fcn: facturas cuotas normales
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",'Cuotas Socios' tipofact,2, concat(scafac.codtipom,right(concat('0000000',numfactu),7)),fecfactu, nomclien nomsocio, totalfac from " & BBDD & ".scafac  where intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and scafac.codtipom = 'FCN' "
    
    ' frc: facturas rectificativas de cuotas
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",'Rectificativa Cuotas' tipofact,3, concat(scafac.codtipom,right(concat('0000000',numfactu),7)),fecfactu, nomclien nomsocio, totalfac from " & BBDD & ".scafac  where intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and scafac.codtipom = 'FRC' "
    
    
    ' fac: facturas servicios cliente
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",'Servicios Clientes' tipofact,4, concat(scafaccli.codtipom,right(concat('0000000',numfactu),7)),fecfactu, nomclien nomsocio, totalfac from " & BBDD & ".scafaccli  where intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and scafaccli.codtipom = 'FAC' "
    
    ' fpc: facturas publicidad cliente
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",'Publicidad Clientes' tipofact,5, concat(scafaccli.codtipom,right(concat('0000000',numfactu),7)),fecfactu, nomclien nomsocio, totalfac from " & BBDD & ".scafaccli  where intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and scafaccli.codtipom = 'FPC' "
    ' frn: facturas rectificativas cliente
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",'Rectificativas Servicios' tipofact,6, concat(scafaccli.codtipom,right(concat('0000000',numfactu),7)),fecfactu, nomclien nomsocio, totalfac from " & BBDD & ".scafaccli  where intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and scafaccli.codtipom = 'FRN' "
    
    ' frp: facturas rectificativas publicidad cliente
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",'Rectif.Publicidad Cliente' tipofact,7, concat(scafaccli.codtipom,right(concat('0000000',numfactu),7)),fecfactu, nomclien nomsocio, totalfac from " & BBDD & ".scafaccli  where intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and scafaccli.codtipom = 'FRP' "
    
    ' fvc: facturas venta cliente
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",'Venta Cliente' tipofact,8, concat(scafaccli.codtipom,right(concat('0000000',numfactu),7)),fecfactu, nomclien nomsocio, totalfac from " & BBDD & ".scafaccli  where intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and scafaccli.codtipom = 'FVC' "
    
    ' fli: facturas liquidacion socio
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",'Liquidacion Socio' tipofact,9, concat(sfactusoc.codtipom,right(concat('0000000',numfactu),7)),fecfactu, sclien.nomclien nomsocio, totalfac from " & BBDD & ".sfactusoc, " & BBDD & ".sclien  where sfactusoc.codsocio = sclien.codclien and intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and sfactusoc.codtipom = 'FLI' "
    
    ' fps: facturas publicidad socio
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",'Publicidad Socio' tipofact,10, concat(sfactusoc.codtipom,right(concat('0000000',numfactu),7)),fecfactu, sclien.nomclien nomsocio, totalfac from " & BBDD & ".sfactusoc, " & BBDD & ".sclien  where sfactusoc.codsocio = sclien.codclien and intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and sfactusoc.codtipom = 'FPS' "
    
    ' frl: facturas rectificativas liquidacion socio
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",'Rectificativas Liquidacion' tipofact,11, concat(sfactusoc.codtipom,right(concat('0000000',numfactu),7)),fecfactu, sclien.nomclien nomsocio, totalfac from " & BBDD & ".sfactusoc, " & BBDD & ".sclien  where sfactusoc.codsocio = sclien.codclien and intconta = 0 "
    If vEmpresa.TieneSII Then
        Sql = Sql & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    Sql = Sql & " and sfactusoc.codtipom = 'FRL' "
    
    conn.Execute SQLinsert & Sql


    
    Sql = "select codusu,nombre1,codigo1,nombre2,nombre3,fecha1,importe1 from tmpinformes where codusu = " & vUsu.Codigo '& " order by 6,5 "
    
    If TotalRegistrosConsulta(Sql) > 0 Then
    
        If MsgBox("Tiene facturas pendientes de integrar a contabilidad." & vbCrLf & vbCrLf & "         �Desea verlas ahora? " & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    
            Set frmMens = New frmMensajes
            
            frmMens.OpcionMensaje = 31
            frmMens.CADENA = Sql
            frmMens.Show vbModal
        
            Set frmMens = Nothing
            
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
eFrasPendientesContabilizar:
    MuestraError Err.Number, "Facturas Pendientes de Integrar a Contabilidad", Err.Description
End Sub




Public Function LeerDatosEmpresa()
 'Crea instancia de la clase Cempresa con los valores en
 'Tabla: AritaxiEmpresa
 'BDatos: Usuarios
 
        Set vEmpresa = New Cempresa
        If vEmpresa.LeerDatos = 1 Then
            MsgBox "No se han podido cargar datos empresa (BD:usuarios). Debe configurar la aplicaci�n.", vbExclamation
            Set vEmpresa = Nothing
        End If
            
End Function


Public Function LeerNivelesEmpresa()
 'Crea instancia de la clase Cempresa con los valores en
 'Tabla: Empresa
 'BDatos: Conta
 
        If vEmpresa.LeerNiveles = 1 Then
            MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicaci�n.", vbExclamation
'            Set vEmpresa = Nothing
        End If
        
End Function


Public Function LeerParametros()
'Crea instancia de la clase CParametros con los valores en
'Tabla: sparam
'BDatos: Aritaxi
 Dim devuelve As String
 
    'Parametros Generales
    Set vParam = New Cparametros
    If vParam.Leer() = 1 Then
        devuelve = "No se han podido cargar los Par�metros Generales.(sparam)" & vbCrLf
        MsgBox devuelve & " Debe configurar la aplicaci�n.", vbExclamation
        Set vParam = Nothing
    End If
        
    'Parametros Aplicacion
    Set vParamAplic = New CParamAplic
    If vParamAplic.Leer() = 1 Then
        devuelve = "No se han podido cargar los Par�metros de la Aplicaci�n.(spara1)" & vbCrLf
        MsgBox devuelve & "Debe configurar la aplicaci�n.", vbExclamation
        Set vParamAplic = Nothing
    End If
    
    
    'Pondra la tb ppal
    vUsu.LeerTabPorDefecto
    
End Function


'/////////////////////////////////////////////////////////////////
'// Se trata de identificar el PC en la BD. Asi conseguiremos tener
'// los nombres de los PC para poder asignarles un codigo
'// UNa vez asignado el codigo  se lo sumaremos (x 1000) al codusu
'// con lo cual el usuario sera distinto( aunque sea con el mismo codigo de entrada)
'// dependiendo desde k PC trabaje

Public Sub GestionaPC()
Dim miRsAux As ADODB.Recordset

CadenaDesdeOtroForm = ComputerName
If CadenaDesdeOtroForm <> "" Then
    'conAri=1: conexion a BD Aritaxi
    FormatoFecha = DevuelveDesdeBD(conAri, "codpc", "usuarios.pcs", "nompc", CadenaDesdeOtroForm, "T")
    If FormatoFecha = "" Then
        NumRegElim = 0
        FormatoFecha = "Select max(codpc) from usuarios.pcs"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open FormatoFecha, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            NumRegElim = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        NumRegElim = NumRegElim + 1
        If NumRegElim > 9999 Then
            MsgBox "Error en numero de PC's activos. Demasiados PC en BD. Llame a soporte t�cnico.", vbCritical
            End
        End If
        FormatoFecha = "INSERT INTO usuarios.pcs (codpc, nompc) VALUES (" & NumRegElim & ", '" & CadenaDesdeOtroForm & "')"
        conn.Execute FormatoFecha
    End If
End If

End Sub


Private Sub OtrasAcciones()
On Error Resume Next

    FormatoFecha = "yyyy-mm-dd"
    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"  'Decimal(12,2)
    FormatoHora = "hh:mm:ss"
    FormatoHora2 = "hh:mm"
    
    
    
    'Por si paraemtrizamos la ampliacion
    FormatoPrecio = "###,##0.0000"  'Decimal(10,4)
    FormatoPrecio2 = "###,##0." & String(PrecioDecimales, "0") 'Decimal(10,4)
    
    
    
    'Por si acasomcambaimos la aplicacion los numeros de decimales
    'ANTES
    'FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    'FormatoCantidad2 = "###,##0.00"   'Decimal(8,2)
    'Ahora
    FormatoCantidad = "##,###,##0." & String(NumeroDeDecimales, "0")
    FormatoCantidad2 = "###,##0." & String(NumeroDeDecimales, "0")
    
    FormatoDescuento = "#0.00" 'Decima(4,2)
    FormatoKms = "#,##0.00##" 'Decimal(8,4)
    FormatoPorcen = "##0.00" 'Decima(5,2)
    
    teclaBuscar = 43
    
    
    'Borramos uno de los archivos temporales
    If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
    
    
    'Borramos tmp bloqueos
    'Borramos temporal
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    NumRegElim = Len(CadenaDesdeOtroForm)
    If NumRegElim = 0 Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = " WHERE codusu = " & vUsu.Codigo
    End If
    conn.Execute "Delete from zbloqueos " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
End Sub


'Usuario As String, Pass As String --> Directamente el usuario
Public Function AbrirConexion() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set conn = Nothing
    Set conn = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

'        cad = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=accUPVMED"
'        cad = cad & ";UID=" & Usuario
'        cad = cad & ";PWD=" & Pass
'        Conn.ConnectionString = cad
    
    'cad = "DSN=plannertours;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=plannertours;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    
    '---- Laura: 17/10/2006
'    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vAritaxi;DATABASE=" & vUsu.CadenaConexion
    '[Monica]21/04/2015: no abria correctamente la conexion, cambiada la de arriba por esta
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER
    
    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password
    Cad = Cad & ";Persist Security Info=true"
    
    conn.ConnectionString = Cad
    conn.Open
    conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n BD:Aritaxi.", Err.Description
End Function





Public Function AbrirConexionUsuarios() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion


    AbrirConexionUsuarios = False
    Set conn = Nothing
    Set conn = New Connection
    'Conn.CursorLocation = adUseClient
    conn.CursorLocation = adUseServer
    'Cad = "DSN=vUsuarios;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=usuarios;"
    'Cad = Cad & "SERVER=" & vConfig.SERVER & ";UID=" & vConfig.User & ";PASSWORD=" & vConfig.password & ";PORT=3306;OPTION=3;STMT=;"

    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=usuarios;SERVER=" & vConfig.SERVER

    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password
    Cad = Cad & ";OPTION=3;STMT=;Persist Security Info=true"

    conn.ConnectionString = Cad
    conn.Open
    AbrirConexionUsuarios = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n usuarios.", Err.Description
End Function



Public Function AbrirConexionConta(ContabilidadEnB As Boolean) As Boolean
'Abre

Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionConta = False
    Set ConnConta = Nothing
    Set ConnConta = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
    '[Monica]02/05/2017: nueva contabilidad
    If vParamAplic.ContabilidadNueva Then
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=ariconta"
    Else
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta"
    End If
                        
    If ContabilidadEnB Then
        Cad = Cad & vParamAplic.ContabilidadB
    Else
        Cad = Cad & vParamAplic.NumeroConta
    End If
    Cad = Cad & ";SERVER=" & vParamAplic.ServidorConta & ";"
    Cad = Cad & ";UID=" & vParamAplic.UsuarioConta
    Cad = Cad & ";PWD=" & vParamAplic.PasswordConta
    '---- Laura: 29/09/2006
    Cad = Cad & ";PORT=3306;OPTION=3;STMT="
    '----
    Cad = Cad & ";Persist Security Info=true"
    ConnConta.ConnectionString = Cad
    ConnConta.Open
    ConnConta.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionConta = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n.", Err.Description
End Function



Public Function CerrarConexionConta()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnConta.Close
   If Err.Number <> 0 Then Err.Clear
End Function




'Para las cosas que tengan que ver con aridoc
'Utilizaremos la conexion de conta
Public Function AbrirConexionAridoc() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionAridoc = False
    Set ConnConta = Nothing
    Set ConnConta = New Connection
    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= Aridoc;DATABASE=Aridoc"
    'Cad = Cad & ";UID=" & vConfig.User
    'Cad = Cad & ";PWD=" & vConfig.password
    Cad = Cad & ";Persist Security Info=true"
    
    ConnConta.ConnectionString = Cad
    ConnConta.Open
    ConnConta.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionAridoc = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n BD:Aridoc.", Err.Description
End Function



Public Function Conexion_Aridoc_(abrir As Boolean) As Boolean
Dim Bien As Boolean
    Conexion_Aridoc_ = False
    CerrarConexionConta
    If abrir Then
        Bien = AbrirConexionAridoc()
    Else
        'Reabrimos la conexion conta
        Bien = AbrirConexionConta(False)
    End If
    If Not Bien Then
        If Not abrir Then
            MsgBox "EL PRORGRAMA FINALIZARA", vbExclamation
            End
        End If
    Else
        Conexion_Aridoc_ = True
    End If
    
End Function














'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=1"
End Sub


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosComas(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ".")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "," & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaPuntosComas = CADENA
End Function


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(CADENA As String) As String
Dim I As Integer
    Do
        I = InStr(1, CADENA, ",")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "." & Mid(CADENA, I + 1)
        End If
    Loop Until I = 0
    TransformaComasPuntos = CADENA
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosHoras(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ".")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & ":" & Mid(CADENA, I + 1)
        End If
    Loop Until I = 0
    TransformaPuntosHoras = CADENA
End Function


Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = "0"
                Case "F"    'Fecha
                     '==David
'                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function



Public Function DBLetMemo(vData As Variant) As String
    On Error Resume Next
    
    DBLetMemo = vData
    
'    If IsNull(DBLetMemo) Then DBLetMemo = ""
    
    If Err.Number <> 0 Then
        Err.Clear
        DBLetMemo = ""
    End If
End Function




Public Function DBSet(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
'Tipos
'       T
'       N
'       F
'       H
'       FH
'       B
'       S   single O DOUBLE. sINGLE DE MOMENTO.    MAYO 2009
Dim Cad As String
Dim ValorNumericoCero As Boolean

    On Error GoTo Error1

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If
    
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        Cad = (CStr(vData))
                        NombreSQL Cad
                        DBSet = "'" & Cad & "'"
                    End If
                    
                Case "N", "S"   'Numero  y  SINGLE
                    
                    If CStr(vData) = "" Then
                        ValorNumericoCero = True
                    
                    Else
                        If Tipo = "S" Then
                            ValorNumericoCero = CSng(vData) = 0
                        Else
                            ValorNumericoCero = CCur(vData) = 0
                        End If
                    End If
                    
                    If ValorNumericoCero Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        If Tipo = "N" Then
                            Cad = CStr(ImporteFormateado(CStr(vData)))
                        Else
                            'Sngle
                            Cad = CStr(ImporteFormateadoSingle(CStr(vData)))
                        End If
                        DBSet = TransformaComasPuntos(Cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If

                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If

                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                
                Case "HH" 'Hora hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm") & "'"
                    End If
                
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
Error1:
    If Err.Number <> 0 Then MuestraError Err.Number, "Formato para la BD.", Err.Description
End Function





Public Function DBSetDavid(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim Cad As String
    On Error GoTo Error1

        If IsNull(vData) Then
            'Aqui esta la modificacion de David
            'DBSet = ValorNulo
            vData = ""
            If Tipo = "" Then DBSetDavid = ValorNulo
            'Exit Function
        End If
    
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSetDavid = "''"
                        Else
                            DBSetDavid = ValorNulo
                        End If
                    Else
                        Cad = (CStr(vData))
                        NombreSQL Cad
                        DBSetDavid = "'" & Cad & "'"
                    End If
                    
                Case "N"    'Numero
                    If CStr(vData) = "" Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSetDavid = ValorNulo
                            Else
                                DBSetDavid = 0
                            End If
                        Else
                            DBSetDavid = 0
                        End If
                    ElseIf CCur(vData) = 0 Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSetDavid = ValorNulo
                            Else
                                DBSetDavid = 0
                            End If
                        Else
                            DBSetDavid = 0
                        End If
                    Else
                        Cad = CStr(ImporteFormateado(CStr(vData)))
                        DBSetDavid = TransformaComasPuntos(Cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSetDavid = ValorNulo
                        Else
                            DBSetDavid = "'1900-01-01'"
                        End If
                    Else
                        DBSetDavid = "'" & Format(vData, FormatoFecha) & "'"
                    End If

                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSetDavid = ValorNulo
                    Else
                        DBSetDavid = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If

                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSetDavid = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                
                Case "B"  'Boolean
                    If vData Then
                        DBSetDavid = 1
                    Else
                        DBSetDavid = 0
                    End If
            End Select
        End If
Error1:
    If Err.Number <> 0 Then MuestraError Err.Number, "Formato para la BD.(DBSetDav)", Err.Description
End Function





'Public Function FechaCorrecta(vFecha As Date) As Byte
''--------------------------------------------------------
''   Dada una fecha dira si pertenece o no
''   al intervalo de fechas que maneja la apliacion
''   Resultados:
''       0 .- A�o actual
''       1 .- Siguiente
''       2 .- Anterior al inicio
''       3 .- Posterior al fin
''--------------------------------------------------------
'    FechaCorrecta = 2
'    If vFecha >= vParam.fechaini Then
'        If vFecha <= vParam.fechafin Then
'            FechaCorrecta = 0
'        Else
'            'Compruebo si el a�o siguiente
'            If vFecha <= DateAdd("yyyy", 1, vParam.fechafin) Then
'                FechaCorrecta = 1
'            Else
'                FechaCorrecta = 3
'            End If
'        End If
'    End If
'End Function


Public Sub MuestraError(numero As Long, Optional CADENA As String, Optional Desc As String)
    Dim Cad As String
    Dim Aux As String
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    Cad = "Se ha producido un error: " & vbCrLf
    If CADENA <> "" Then
        Cad = Cad & vbCrLf & CADENA & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If conn.Errors.Count > 0 Then
        ControlamosError Aux
        conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then Cad = Cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then Cad = Cad & "N�mero: " & numero & vbCrLf & "Descripci�n: " & Error(numero)
    MsgBox Cad, vbExclamation
End Sub


Public Function Espera(Segundos As Single)
Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function


Public Function RellenaCodigoCuenta(vCodigo As String) As String
'Rellena con ceros hasta poner una cuenta.
'Ejemplo: 43.1 --> 430000001
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim Cad As String

    RellenaCodigoCuenta = vCodigo
    If Len(vCodigo) > vEmpresa.DigitosUltimoNivel Then Exit Function
    
    I = 0: cont = 0
    Do
        I = I + 1
        I = InStr(I, vCodigo, ".")
        If I > 0 Then
            If cont > 0 Then cont = 1000
            cont = cont + I
        End If
    Loop Until I = 0

    'Habia mas de un punto
    If cont > 1000 Or cont = 0 Then Exit Function

    'Cambiamos el punto por 0's  .-Utilizo la variable maximocaracteres, para no tener k definir mas
    I = Len(vCodigo) - 1 'el punto lo quito
    J = vEmpresa.DigitosUltimoNivel - I
    Cad = ""
    For I = 1 To J
        Cad = Cad & "0"
    Next I

    Cad = Mid(vCodigo, 1, cont - 1) & Cad
    Cad = Cad & Mid(vCodigo, cont + 1)
    RellenaCodigoCuenta = Cad
End Function



Public Function DevuelveDesdeBD(vBD As Byte, kCampo As String, KTabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef otroCampo As String) As String
    Dim Rs As Recordset
    Dim Cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    Cad = "Select " & kCampo
    If otroCampo <> "" Then Cad = Cad & ", " & otroCampo
    Cad = Cad & " FROM " & KTabla
    Cad = Cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        Cad = Cad & ValorCodigo
    Case "T", "F"
        Cad = Cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
'    Debug.Print cad
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    If vBD = 1 Then 'BD 1: Aritaxi
        Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else    'BD 2: Conta
        Rs.Open Cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    
    If Not Rs.EOF Then
        DevuelveDesdeBD = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function


'Este metodo sustituye a DevuelveDesdeBD
'Funciona para claves primarias formadas por 2 campos
Public Function DevuelveDesdeBDNew(vBD As Byte, KTabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim Rs As Recordset
Dim Cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    Cad = "Select " & kCampo
    If otroCampo <> "" Then Cad = Cad & ", " & otroCampo
    Cad = Cad & " FROM " & KTabla
    If Kcodigo1 <> "" Then
        Cad = Cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            Cad = Cad & Val(valorCodigo1)
        Case "T"
            Cad = Cad & DBSet(valorCodigo1, "T")
        Case "F"
            Cad = Cad & "'" & valorCodigo1 & "'"
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If
    
    If KCodigo2 <> "" Then
        Cad = Cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            Cad = Cad & DBSet(ValorCodigo2, "T")
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        Cad = Cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo3)
            End If
        Case "T"
            Cad = Cad & "'" & ValorCodigo3 & "'"
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    If vBD = conAri Then 'BD 1: Aritaxi
        Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else    'BD 2: Conta
        Rs.Open Cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    
    If Not Rs.EOF Then
        DevuelveDesdeBDNew = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
EDevuelveDesdeBDnew:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function


'Obvio
Public Function EsCuentaUltimoNivel(cuenta As String) As Boolean
    EsCuentaUltimoNivel = (Len(cuenta) = vEmpresa.DigitosUltimoNivel)
End Function


Public Function CuentaCorrectaUltimoNivel(ByRef cuenta As String, ByRef devuelve As String) As Boolean
'Comprueba si es numerica
Dim Sql As String
Dim otroCampo As String

CuentaCorrectaUltimoNivel = False
If cuenta = "" Then
    devuelve = "Cuenta vacia"
    Exit Function
End If

If Not IsNumeric(cuenta) Then
    devuelve = "La cuenta debe de ser num�rica: " & cuenta
    Exit Function
End If

'Rellenamos si procede
cuenta = RellenaCodigoCuenta(cuenta)

'==========
If Not EsCuentaUltimoNivel(cuenta) Then
    devuelve = "No es cuenta de �ltimo nivel: " & cuenta
    Exit Function
End If
'==================

otroCampo = "apudirec"
'BD 2: conexion a BD Conta
Sql = DevuelveDesdeBD(conConta, "nommacta", "cuentas", "codmacta", cuenta, "T", otroCampo)
If Sql = "" Then
    devuelve = "No existe la cuenta : " & cuenta
    CuentaCorrectaUltimoNivel = True
    Exit Function
End If

'Llegados aqui, si que existe la cuenta
If otroCampo = "S" Then 'Si es apunte directo
    CuentaCorrectaUltimoNivel = True
    devuelve = Sql
Else
    devuelve = "No es apunte directo: " & cuenta
End If

End Function

'-------------------------------------------------------------------------
'
'   Es la misma solo k no si no existe cuenta no da error
'Public Function CuentaCorrectaUltimoNivelSIN(ByRef Cuenta As String, ByRef devuelve As String) As Byte
''Comprueba si es numerica
'Dim SQL As String
'
'CuentaCorrectaUltimoNivelSIN = 0
'If Cuenta = "" Then
'    devuelve = "Cuenta vacia"
'    Exit Function
'End If
'If Not IsNumeric(Cuenta) Then
'    devuelve = "La cuenta debe de ser num�rica: " & Cuenta
'    Exit Function
'End If
'
''Rellenamos si procede
'Cuenta = RellenaCodigoCuenta(Cuenta)
'
'CuentaCorrectaUltimoNivelSIN = 1
'If Not EsCuentaUltimoNivel(Cuenta) Then
'    SQL = "No es cuenta de �ltimo nivel"
'Else
'    'BD 2: conexion a BD Conta
'    SQL = DevuelveDesdeBD(2, "nommacta", "cuentas", "codmacta", Cuenta, "T")
'    If SQL = "" Then
'        SQL = "No existe la cuenta  "
'    Else
'        CuentaCorrectaUltimoNivelSIN = 2
'    End If
'End If
'
''Llegados aqui, si que existe la cuenta
'devuelve = SQL
'End Function


'Devuelve, para un nivel determinado, cuantos digitos tienen las cuentas
' a ese nivel
'Public Function DigitosNivel(numnivel As Integer) As Integer
'    Select Case numnivel
'    Case 1
'        DigitosNivel = vEmpresa.numdigi1
'
'    Case 2
'        DigitosNivel = vEmpresa.numdigi2
'
'    Case 3
'        DigitosNivel = vEmpresa.numdigi3
'
'    Case 4
'        DigitosNivel = vEmpresa.numdigi4
'
'    Case 5
'        DigitosNivel = vEmpresa.numdigi5
'
'    Case 6
'        DigitosNivel = vEmpresa.numdigi6
'
'    Case 7
'        DigitosNivel = vEmpresa.numdigi7
'
'    Case 8
'        DigitosNivel = vEmpresa.numdigi8
'
'    Case 9
'        DigitosNivel = vEmpresa.numdigi9
'
'    Case 10
'        DigitosNivel = vEmpresa.numdigi10
'
'    Case Else
'        DigitosNivel = -1
'    End Select
'End Function


'Public Function NivelCuenta(CodigoCuenta As String) As Integer
'Dim lon As Integer
'Dim niv As Integer
'Dim I As Integer
'    NivelCuenta = -1
'    lon = Len(CodigoCuenta)
'    I = 0
'    Do
'       I = I + 1
'       niv = DigitosNivel(I)
'       If niv > 0 Then
'            If niv = lon Then
'                NivelCuenta = I
'                I = 11 'para salir del bucle
'            End If
'        Else
'            I = 11 'salimos pq ya no hay nveles para las cuentas de longitud lon
'        End If
'    Loop Until I > 10
'End Function


'Public Function ExistenSubcuentas(ByRef Cuenta As String, Nivel As Integer) As Boolean
'Dim I As Integer
'Dim b As Boolean
'Dim Cad As String
'
'    I = DigitosNivel(Nivel)
'    Cad = Mid(Cuenta, 1, I)
'    Cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cad, "T")
'    If Cad = "" Then
'        'NO existe la subcuenta de nivel N
'        'salimos
'        ExistenSubcuentas = False
'        Exit Function
'    End If
'    If Nivel > 1 Then
'        ExistenSubcuentas = ExistenSubcuentas(Cuenta, Nivel - 1)
'    Else
'        ExistenSubcuentas = True
'    End If
'End Function


'Public Function CreaSubcuentas(ByRef Cuenta, HastaNivel As Integer, TEXTO As String) As Boolean
'Dim I As Integer
'Dim J As Integer
'Dim Cad As String
'Dim Cta As String
'
'On Error GoTo ECreaSubcuentas
'CreaSubcuentas = False
'For I = 1 To HastaNivel
'    J = DigitosNivel(I)
'    Cta = Mid(Cuenta, 1, J)
'    Cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cta, "T")
'    If Cad = "" Then
'        'CreaCuenta
'        Cad = "INSERT INTO cuentas (codmacta, nommacta, apudirec, model347, razosoci, "
'        Cad = Cad & " dirdatos, codposta, despobla, desprovi, nifdatos, maidatos, webdatos,"
'        Cad = Cad & " obsdatos) VALUES ("
'        Cad = Cad & " '" & Cta
'        Cad = Cad & " ', '" & TEXTO
'        Cad = Cad & " ', "
'        Cad = Cad & " 'N', 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
'        Conn.Execute Cad
'    End If
'Next I
'CreaSubcuentas = True
'Exit Function
'ECreaSubcuentas:
'    MuestraError Err.Number, "Creando subcuentas", Err.Description
'End Function




Public Function CambiarBarrasPATH2(ParaGuardarBD As Boolean, CADENA) As String
Dim I As Integer
Dim CH As String
Dim Ch2 As String

If ParaGuardarBD Then
    CH = "\"
    Ch2 = "/"
Else
    CH = "/"
    Ch2 = "\"
End If
I = 0
Do
    I = I + 1
    I = InStr(1, CADENA, CH)
    If I > 0 Then CADENA = Mid(CADENA, 1, I - 1) & Ch2 & Mid(CADENA, I + 1)
Loop Until I = 0
CambiarBarrasPATH2 = CADENA
End Function


Public Function ImporteSinFormato(CADENA As String) As String
Dim I As Integer
    'Quitamos puntos
    Do
        I = InStr(1, CADENA, ".")
        If I > 0 Then CADENA = Mid(CADENA, 1, I - 1) & Mid(CADENA, I + 1)
    Loop Until I = 0
    ImporteSinFormato = TransformaPuntosComas(CADENA)
End Function




'Public Sub SaldoHistorico(Cuenta As String)
'Dim RS As Recordset
'Dim SQL As String
'Dim RC2 As String
'    Screen.MousePointer = vbHourglass
'    SQL = "Select Sum(timporteD),sum(timporteH) from hlinapu"
'    SQL = SQL & " WHERE codmacta='" & Cuenta & "'"
'    SQL = SQL & " AND fechaent>='" & Format(vParam.fechaini, FormatoFecha) & "' AND punteada "
'    Set RS = New ADODB.Recordset
'    RC2 = Cuenta & "|"
'    'PUNTEADO
'    RS.Open SQL & "='S';", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not RS.EOF Then
'       RC2 = RC2 & Format(RS.Fields(0), FormatoImporte) & "|"
'       RC2 = RC2 & Format(RS.Fields(1), FormatoImporte) & "|"
'    Else
'        RC2 = RC2 & "||"
'    End If
'    RS.Close
'    'SIN puntear
'    RS.Open SQL & "<>'S';", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not RS.EOF Then
'       RC2 = RC2 & Format(RS.Fields(0), FormatoImporte) & "|"
'       RC2 = RC2 & Format(RS.Fields(1), FormatoImporte) & "|"
'    Else
'        RC2 = RC2 & "||"
'    End If
'    RS.Close
'    Set RS = Nothing
'    'Mostramos la ventanita de mesaje
'    frmMensajes.Opcion = 1
'    frmMensajes.Parametros = RC2
'    frmMensajes.Show vbModal
'
'End Sub

'Lo que hace es comprobar que si la resolucion es mayor
'que 800x600 lo pone en el 400
Public Sub AjustarPantalla(ByRef formulario As Form)
'    If Screen.Width > 13000 Then
'        formulario.Top = 400
'        formulario.Left = 400
'    Else
'        formulario.Top = 0
'        formulario.Left = 0
'    End If
'    formulario.Width = 12000
'    formulario.Height = 9000
End Sub


'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256,98
'   Tiene que venir num�rico
Public Function ImporteFormateado(Importe As String) As Currency
Dim I As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateado = Importe
    End If
End Function
Public Function ImporteFormateadoSingle(Importe As String) As Single
Dim I As Integer

    If Importe = "" Then
        ImporteFormateadoSingle = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateadoSingle = Importe
    End If
End Function

Public Function ComprobarEmpresaBloqueada(codusu As Long, ByRef Empresa As String) As Boolean
End Function


Public Function Bloquear_DesbloquearBD(Bloquear As Boolean) As Boolean

On Error GoTo EBLo
    Bloquear_DesbloquearBD = False
    If Bloquear Then
        CadenaDesdeOtroForm = "INSERT INTO usuarios.vBloqBD (codusu, conta) VALUES (" & vUsu.Codigo & ",'" & vUsu.CadenaConexion & "')"
    Else
        CadenaDesdeOtroForm = "DELETE FROM  usuarios.vBloqBD WHERE codusu =" & vUsu.Codigo & " AND conta = '" & vUsu.CadenaConexion & "'"
    End If
    conn.Execute CadenaDesdeOtroForm
    Bloquear_DesbloquearBD = True
    Exit Function
EBLo:
    'MuestraError Err.Number, "Bloq. BD"
    Err.Clear
End Function


Public Function OtrosPCsContraContabiliad() As String
Dim MiRS As Recordset
Dim Cad As String
Dim Equipo As String
Dim EquipoConBD As Boolean


    Set MiRS = New ADODB.Recordset
    EquipoConBD = (vUsu.PC = vConfig.SERVER Or LCase(vConfig.SERVER = "localhost"))
    Cad = "show processlist"
    MiRS.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not MiRS.EOF
        If UCase(MiRS.Fields(3)) = UCase(vUsu.CadenaConexion) Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vUsu.PC Then
                    
                    NumRegElim = 0
                    If Equipo <> "LOCALHOST" Then
                        'Si no es localhost
                        NumRegElim = 1
                    Else
                        'HAy un proceso de loclahost. Luego, si el equipo no tiene la BD
                        If Not EquipoConBD Then NumRegElim = 1
                    End If

                    
                    'Si hay que insertar
                    If NumRegElim = 1 Then
                        If InStr(1, Cad, Equipo & "|") = 0 Then Cad = Cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraContabiliad = Cad

End Function


Public Function EsNumerico(Texto As String) As Boolean
Dim I As Integer
Dim C As Integer
Dim L As Integer
Dim Cad As String
Dim b As Boolean
    
    EsNumerico = False
    b = True
    Cad = ""
    If Not IsNumeric(Texto) Then
        Cad = "El campo debe ser num�rico"
        b = False
        '======= A�ade Laura
        'formato: (.25)
        I = InStr(1, Texto, ".")
        If I = 1 Then
            If IsNumeric(Mid(Texto, 2, Len(Texto))) Then b = True
        End If
        '======================
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            I = InStr(L, Texto, ".")
            If I > 0 Then
                L = I + 1
                C = C + 1
            End If
        Loop Until I = 0
        If C > 1 Then
            Cad = "Numero de puntos incorrecto"
            b = False
        End If
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                I = InStr(L, Texto, ",")
                If I > 0 Then
                    L = I + 1
                    C = C + 1
                End If
            Loop Until I = 0
            If C > 1 Then
                Cad = "Numero incorrecto"
                b = False
            End If
        End If
    End If
    If Not b Then
        MsgBox Cad, vbExclamation
    Else
        EsNumerico = b
    End If
End Function







'==== Laura==
'Public Function EsPorcentajeOK(ByRef T As TextBox) As Boolean
'Dim cad As String
'Dim OK As Boolean
'
'    cad = TransformaPuntosComas(T.Text)
'
'    OK = False
'    If InStr(1, cad, ",") = 0 Then 'No hay decimales
'        If Len(T.Text) = 5 Then
'            cad = Mid(cad, 1, 2) & "," & Mid(cad, 3, 2)
'            OK = True
'        Else
'            If Len(T.Text) = 4 Then cad = Mid(cad, 1, 2) & "," & Mid(cad, 3, 2)
'            OK = True
'        End If
'    ElseIf InStr(1, cad, ",") = 1 Or InStr(1, cad, ",") = 2 Or InStr(1, cad, ",") = 3 Then 'Hay punto
'        OK = True
'    End If
'    If OK Then T.Text = cad
'    EsPorcentajeOK = OK
''    If IsDate(Cad) Then
''        EsFechaOK = True
''        T.Text = Format(Cad, "dd/mm/yyyy")
''    Else
''        EsFechaOK = False
''    End If
'
'End Function
'============




'Devuelve si hay archivos
'                                                        Llevara la forma: 01, 02  para la empresa 1 o 2..
'Public Function BuscarIntegraciones(Errores As Boolean, Empresa As String) As Boolean
'Dim cad As String
'On Error GoTo Ebuscarintegraciones
'
'    BuscarIntegraciones = False
'    If vConfig.Integraciones = "" Then Exit Function
'
'    cad = vConfig.Integraciones
'    If Right(cad, 1) <> "\" Then cad = cad & "\"
'    If Dir(cad, vbDirectory) = "" Then
'        MsgBox "Carpeta de errores no encontrada: " & vConfig.Integraciones, vbExclamation
'        Exit Function
'    End If
'
'    If Errores Then
'        cad = vConfig.Integraciones & "\ERRORES"
'    Else
'        cad = vConfig.Integraciones & "\INTEGRA"
'    End If
'
'    'Facturas clientes
'    If Dir(cad & "\FRACLI\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If
'
'    'Facturas Proveedores
'    If Dir(cad & "\FRAPRO\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If
'
'    'Asientos al diario
'    If Dir(cad & "\ASIDIA\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If
'
'    'Asientos al historico
'    If Dir(cad & "\ASIHCO\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If
'
'    Exit Function
'Ebuscarintegraciones:
'    MuestraError Err.Number, Err.Description, "Buscar archivos integraciones" & vbCrLf
'End Function


'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef CADENA As String)
Dim J As Integer
Dim I As Integer
Dim Aux As String

    J = 1
    '-- (RAFA/ALZIRA) 07052006
    Do
        I = InStr(J, CADENA, "\")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
    

    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
    
End Sub

Public Function DevNombreSQL(CADENA As String) As String
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSQL = CADENA
End Function



'Para los balnces
'Public Function FechaInicioIGUALinicioEjerecicio(FecIni As Date, EjerciciosCerrados1 As Boolean) As Byte
'Dim Fecha As Date
'Dim Salir As Boolean
'Dim I As Integer
'On Error GoTo EfechaInicioIGUALinicioEjerecicio
'
'    FechaInicioIGUALinicioEjerecicio = 1
'    If EjerciciosCerrados1 Then
'        I = -1 'En ejercicios cerrados emp�zamos mirando un a�o por debajo fecini
'    Else
'        I = 1
'    End If
'    Fecha = DateAdd("yyyy", I, vParam.fechaini)
'    Salir = False
'    While Not Salir
'        If FecIni = Fecha Then
'            'Fecha inicio del listado contiene es fecha incio ejercicio
'            FechaInicioIGUALinicioEjerecicio = 0
'            Salir = True
'        Else
'            If FecIni < Fecha Then
'                Fecha = DateAdd("yyyy", -1, Fecha)
'            Else
'                Salir = True
'            End If
'        End If
'    Wend
'
'    Exit Function
'EfechaInicioIGUALinicioEjerecicio:
'    Err.Clear  'No tiene importancia
'End Function





'Public Function DevuelveDigitosNivelAnterior() As Integer
'Dim J As Integer
'    DevuelveDigitosNivelAnterior = 3
'    If vEmpresa Is Nothing Then Exit Function
'    If vEmpresa.numnivel < 2 Then Exit Function
'    J = vEmpresa.numnivel - 1
'    J = DigitosNivel(J)
'    If J < 3 Then J = 3
'    DevuelveDigitosNivelAnterior = J
'End Function



'--------------------------------------------------------------------------
' Los numeros vendran formateados o sin formatear, pero siempre viene texto
'
Public Function CadenaCurrency(Texto As String, ByRef Importe As Currency) As Boolean
Dim I As Integer
On Error GoTo ECadenaCurrency
    
    Importe = 0
    CadenaCurrency = False
    If Not IsNumeric(Texto) Then Exit Function
    I = InStr(1, Texto, ",")
    If I = 0 Then
        'Significa k el numero no esta  formateado y como mucho tiene punto
        Importe = CCur(TransformaPuntosComas(Texto))
    Else
        Importe = ImporteFormateado(Texto)
    End If
    
    CadenaCurrency = True
    
    Exit Function
ECadenaCurrency:
    Err.Clear
End Function


Public Sub CommitConexion()
On Error Resume Next
    conn.Execute "Commit"
    If Err.Number <> 0 Then Err.Clear
End Sub






'------------------------------------------------------------------
'   Comprobara si una daterminada fecha esta o no en los ejercicios
'   contables (actual y siguiente)
'   Dando un O: SI. Correcto. Ok
'            1: Inferior
'            2: Superior

Public Function EsFechaOKConta(Fecha As Date) As Byte
Dim F2 As Date

    If vEmpresa.FechaIni > Fecha Then
        EsFechaOKConta = 1
    Else
        F2 = DateAdd("yyyy", 1, vEmpresa.FechaFin)
        If Fecha > F2 Then
            EsFechaOKConta = 2
        Else
            'OK. Dentro de los ejercicios contables
            EsFechaOKConta = 0
        End If
    End If

End Function



'--------------------------------------------------------------------
'-------------------------------------------------------------------
'Para el envio de los mails
Public Function PrepararCarpetasEnvioMail(Optional NoBorrar As Boolean) As Boolean
    On Error GoTo EPrepararCarpetasEnvioMail
    PrepararCarpetasEnvioMail = False

    If Dir(App.Path & "\temp", vbDirectory) = "" Then
        MkDir App.Path & "\temp"
    Else
        If Not NoBorrar Then
            If Dir(App.Path & "\temp\*.*", vbArchive) <> "" Then Kill App.Path & "\temp\*.*"
        End If
    End If


    PrepararCarpetasEnvioMail = True
    Exit Function
EPrepararCarpetasEnvioMail:
    MuestraError Err.Number, "", "Preparar Carpetas temporal para envio eMail. Borrando tmp "
End Function



Public Function TieneAvisosPendientes() As Boolean
Dim CW As String
Dim F As Date
    On Error GoTo ETieneAvisosPendientes
    TieneAvisosPendientes = False
    
    
    'Alabaranes clientes
    If vParamAplic.avialbcli > 0 Then
        DoEvents
        F = DateAdd("d", -vParamAplic.avialbcli, Now)
        CW = " fechaalb <= '" & Format(F, FormatoFecha) & "'"
        If HayRegParaInforme("scaalb", CW, True) Then
            'No hace falta que siga puesto que si que hay alertar
            TieneAvisosPendientes = True
            Exit Function
        End If
    End If
    
    
    'Albaranes proveedores
    If vParamAplic.avialbpro > 0 Then
        DoEvents
        F = DateAdd("d", -vParamAplic.avialbpro, Now)
        CW = " fechaalb <= '" & Format(F, FormatoFecha) & "'"
        If HayRegParaInforme("scaalp", CW, True) Then
            'No hace falta que siga puesto que si que hay alertar
            TieneAvisosPendientes = True
            Exit Function
        End If
    End If
    
    
    'Pedidos proveedor
    '
    If vParamAplic.avipedpro > 0 Then
        DoEvents
        F = DateAdd("d", -vParamAplic.avipedpro, Now)
        CW = " fecpedpr <= '" & Format(F, FormatoFecha) & "'"
        If HayRegParaInforme("scappr", CW, True) Then
            'No hace falta que siga puesto que si que hay alertar
            TieneAvisosPendientes = True
            Exit Function
        End If
    End If
    
    Exit Function
ETieneAvisosPendientes:
    MuestraError Err.Number, Err.Description
End Function

'--------------------  ELIMINAR ARTICULO
Public Function SePuedeEliminarArticulo(ByVal Articulo As String, ByRef L1 As Label) As String
On Error GoTo Salida
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim I As Integer
Dim C As String
Dim NT As Integer

    SePuedeEliminarArticulo = ""
    Set Rs = New ADODB.Recordset
    Articulo = "'" & DevNombreSQL(Articulo) & "'"
    
    
    'Clientes
    DevuelveTablasBorre 0, C, Sql, NT
    For I = 1 To NT
        L1.Caption = RecuperaValor(Sql, I) & " (Clientes)"
        L1.Refresh
        If TieneDatosSQLCount(Rs, "SELECT count(*) from " & RecuperaValor(C, I) & " where codartic = " & Articulo, 0) Then
            SePuedeEliminarArticulo = SePuedeEliminarArticulo & "    -" & L1.Caption & vbCrLf
            
        End If
    Next I
    If SePuedeEliminarArticulo <> "" Then SePuedeEliminarArticulo = SePuedeEliminarArticulo & vbCrLf & vbCrLf
    
    'Si llega aqui comprobamos en  proveedores
    'PROVEEDORES
    DevuelveTablasBorre 1, C, Sql, NT
    For I = 1 To NT
        L1.Caption = RecuperaValor(Sql, I) & " (Proveedores)"
        L1.Refresh
        If TieneDatosSQLCount(Rs, "SELECT count(*) from " & RecuperaValor(C, I) & " where codartic = " & Articulo, 0) Then
            SePuedeEliminarArticulo = SePuedeEliminarArticulo & "    -" & L1.Caption & vbCrLf
        
        End If
    Next I
    If SePuedeEliminarArticulo <> "" Then SePuedeEliminarArticulo = SePuedeEliminarArticulo & vbCrLf
    
    'Varios
    DevuelveTablasBorre 2, C, Sql, NT
    For I = 1 To NT
        L1.Caption = RecuperaValor(Sql, I) & " (Varios)"
        L1.Refresh
        If TieneDatosSQLCount(Rs, "SELECT count(*) from " & RecuperaValor(C, I) & " where codartic = " & Articulo, 0) Then
            SePuedeEliminarArticulo = SePuedeEliminarArticulo & "    -" & L1.Caption & vbCrLf
            
        End If
    Next I
    
        
        
    'Si es articulo de parametros
    C = ""
    Sql = vbCrLf & Space(10)
    With vParamAplic
        If DBSet(.ArticServ, "T") = Articulo Then C = C & Sql & "Servicios"
        If DBSet(.ArtPortes, "T") = Articulo Then C = C & Sql & "Portes"
        If DBSet(.ArtGastosAdmon, "T") = Articulo Then C = C & Sql & "Gastos de Admon" '"Tasa reciclado"
        If DBSet(.CodarticTfnia, "T") = Articulo Then C = C & Sql & "Telefonia"
    End With
    If C <> "" Then
        C = " -Parametros " & C
        SePuedeEliminarArticulo = SePuedeEliminarArticulo & C
    End If
    
    
    
    
Salida:
    If Err.Number <> 0 Then
        SePuedeEliminarArticulo = "Error: " & Err.Description
        Err.Clear
    End If
End Function



Private Function TieneDatosSQLCount(ByRef Rs As ADODB.Recordset, vSQL As String, IndexdelCount As Integer) As Boolean
    TieneDatosSQLCount = False
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(IndexdelCount)) Then If Rs.Fields(IndexdelCount) > 0 Then TieneDatosSQLCount = True
    End If
        
    Rs.Close

End Function



Public Function EliminarArticulo(ByVal codArtic As String, L1 As Label) As Boolean
Dim NT As Integer
Dim Tablas As String
Dim Dsc As String

    On Error GoTo EEliminarArticulo
    
    EliminarArticulo = False
    
    codArtic = " WHERE codartic = '" & DevNombreSQL(codArtic) & "'"
    
    'Borraremos de tablas que se inserta autmaticamente
    'Ejm: slistas, precios especiales......
    DevuelveTablasBorre 3, Tablas, Dsc, NT
    Do
'        Debug.Print RecuperaValor(Tablas, NT)
        L1.Caption = RecuperaValor(Dsc, NT)
        L1.Refresh
        conn.Execute "DELETE FROM " & RecuperaValor(Tablas, NT) & codArtic
        Debug.Print "DELETE FROM " & RecuperaValor(Tablas, NT) & codArtic
        NT = NT - 1
    Loop Until NT = 0
    
    
    
    'BORRAMOS EL ARTICULO
    L1.Caption = Mid(codArtic, 19)
    L1.Refresh
    conn.Execute "DELETE FROM sartic " & codArtic
    
    EliminarArticulo = True
    
    Exit Function
EEliminarArticulo:
    MuestraError Err.Number, Err.Description
End Function


'Opcion
'   0- Clientes
'   1- Proveedores
'   2- Varios
'   ---------
'   3.- Tabas que cuando eliminen el articulo tendre que borrar yo
Public Sub DevuelveTablasBorre(Opcion As Byte, ByRef Tablas As String, ByRef Descripcion As String, ByRef NumeroTablas As Integer)

    If Opcion = 0 Then
        'CLIENTES
        Tablas = "slhalb|slialb|slifac|slirep|"
        Descripcion = "Hco albaranes|Albaranes|Facturas|Reparaciones|"
        NumeroTablas = 4
    ElseIf Opcion = 1 Then
        'PROVEEDRORES
        Tablas = "slhalp|slhppr|slialp|slifpc|slippr|"
        Descripcion = "Hco albaranes|Hco pedidos|Albaranes|Facturas|Pedidos|"
        NumeroTablas = 5
        
        
    ElseIf Opcion = 2 Then
        'VARIOS
        Tablas = "slhmov|sarti2|slimov|smoval|sserie|shinve|"
        Descripcion = "Hco Lineas Movimientos Almacen|Instalaciones|"
        Descripcion = Descripcion & "Lin mov almacen|Mov almacen|N� serie|Hco inventario|"
        NumeroTablas = 6
        If vParamAplic.Produccion Then
            Tablas = Tablas & "sarti1|"
            Descripcion = Descripcion & "Artic. produccion|"
            NumeroTablas = NumeroTablas + 1
        End If
        
    Else
        'Tablas que al eliminar el articulo voy a tener que borrar
        'Esta salmac. Antes de lanzar el proceso hay que comprobar que la suma de stock es CERO
        '---- [29/09/2009] LAURA: a�adir tablas sarti1,sarti2,sarti3 para eliminar
        Tablas = "slisp1|slispr|sbonif|slist1|slista|spree1|sprees|spromo|salmac|sarti1|sarti2|sarti3|"
        Descripcion = "Precios proveedor|cab. precios provee.|bonificacion facturas|"
        Descripcion = Descripcion & "Hco tarifas|Tarifas|Hco precios especiales|Precios especiales|Promociones|Articulos x Almacen|"
        Descripcion = Descripcion & "L�n. Componentes|L�n. control instalaciones|L�n. codigos EAN|"
        NumeroTablas = 12
        
    End If
    
End Sub



Public Sub MostrarCadenasConexion()
Dim Cad As String
Dim cadCon As String
Dim I As Integer
Dim Propiedades() As String
Dim cadBD As String, cadDSN As String, cadSERVER As String

    On Error GoTo ErrCadCon
    
    Cad = "CONEXIONES BASES DE DATOS " & UCase(App.Title) & vbCrLf & vbCrLf
    
    '---  conexion ARITAXI  ---
    cadCon = conn.Properties("Extended Properties").Value
    Propiedades = Split(cadCon, ";")
    
    '- coger las propiedades q nos interesan
    For I = 0 To UBound(Propiedades)
        If InStr(1, Propiedades(I), "DATABASE=") > 0 Then
            cadBD = Propiedades(I)
        ElseIf InStr(1, Propiedades(I), "DSN=") > 0 Then
            cadDSN = Propiedades(I)
         ElseIf InStr(1, Propiedades(I), "SERVER=") > 0 Then
            cadSERVER = Propiedades(I)
        End If
    Next I
    
    Cad = Cad & "Conexi�n: " & Replace(cadBD, "DATABASE=", "") & vbCrLf
    Cad = Cad & "----------------------------------------   " & vbCrLf
    Cad = Cad & cadDSN & vbCrLf
    Cad = Cad & cadSERVER & vbCrLf
    Cad = Cad & cadBD & vbCrLf & vbCrLf
    
    
    '---  conexion CONTABILIDAD  ---
    cadCon = ConnConta.Properties("Extended Properties").Value
    Propiedades = Split(cadCon, ";")
    cadBD = ""
    cadDSN = "DSN="
    cadSERVER = ""
    
    '- coger las propiedade q nos interesan
    For I = 0 To UBound(Propiedades)
        If InStr(1, Propiedades(I), "DATABASE=") > 0 Then
            cadBD = Propiedades(I)
        ElseIf InStr(1, Propiedades(I), "DSN=") > 0 Then
            cadDSN = Propiedades(I)
         ElseIf InStr(1, Propiedades(I), "SERVER=") > 0 Then
            cadSERVER = Propiedades(I)
        End If
    Next I
    
    Cad = Cad & "Conexi�n: " & Replace(cadBD, "DATABASE=", "") & vbCrLf
    Cad = Cad & "----------------------------------------   " & vbCrLf
    Cad = Cad & cadDSN & vbCrLf
    Cad = Cad & cadSERVER & vbCrLf
    Cad = Cad & cadBD & vbCrLf & vbCrLf
    

    MsgBox Cad, vbInformation
    Exit Sub
    
ErrCadCon:
    MuestraError Err.Number, "Mostrar cadenas conexi�n.", Err.Description
End Sub




Public Function ejecutar(ByRef Sql As String, OcultarMsg As Boolean) As Boolean
    On Error Resume Next
    conn.Execute Sql
    If Err.Number <> 0 Then
        If Not OcultarMsg Then MuestraError Err.Number, Err.Description, Sql
        ejecutar = False
    Else
        ejecutar = True
    End If
End Function


Public Function EjecutarSQL(CadenaSQL As String) As Boolean
    On Error Resume Next
    conn.Execute CadenaSQL
    If Err.Number <> 0 Then
         
         MuestraError Err.Number, "Error ejecutando SQL: " & vbCrLf & CadenaSQL, Err.Description
         EjecutarSQL = False
    Else
         EjecutarSQL = True
    End If
    
End Function

'???????????????????


'Lo que haremos aqui es ver, o guardar, el ultimo numero de empresa
'a la que ha entrado, y el usuario
Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
Dim NF As Integer
Dim C1 As String
Dim Cad As String

On Error GoTo ENumeroEmpresaMemorizar


    If Leer Then
        If CadenaDesdeOtroForm <> "" Then
            'Ya estabamos trabajando con la aplicacion
            
'            If Not (vEmpresa Is Nothing) Then
'                 For NF = 1 To Me.lw1.ListItems.Count
'                    If lw1.ListItems(NF).Text = vEmpresa.nomempre Then
'                        Set lw1.SelectedItem = lw1.ListItems(NF)
'                        Exit For
'                    End If
'                Next NF
'            End If
'
'                'El tercer pipe, si tiene es el ancho col1
'                Cad = AnchoLogin
'                C1 = RecuperaValor(Cad, 3)
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 4360
'                End If
'                lw1.ColumnHeaders(1).Width = NF
'                'El cuarto pipe si tiene es el ancho de col2
'                C1 = RecuperaValor(Cad, 4)
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 1400
'                End If
'                lw1.ColumnHeaders(2).Width = NF
'
'
'            CadenaDesdeOtroForm = ""
'            Exit Sub
        End If
    End If
    Cad = App.Path & "\ultempre.dat"
    If Leer Then
        If Dir(Cad) <> "" Then
            NF = FreeFile
            Open Cad For Input As #NF
            Line Input #NF, Cad
            Close #NF
            Cad = Trim(Cad)
            If Cad <> "" Then
                vUsu.CadenaConexion = RecuperaValor(Cad, 2)
            End If
'                'El primer pipe es el usuario. Como ya no lo necesito, no toco nada
'
'                C1 = RecuperaValor(Cad, 2)
'                'el segundo es el
'                If C1 <> "" Then
'                    For NF = 1 To Me.lw1.ListItems.Count
'                        If lw1.ListItems(NF).Text = C1 Then
'                            Set lw1.SelectedItem = lw1.ListItems(NF)
'                            lw1.SelectedItem.EnsureVisible
'                            Exit For
'                        End If
'                    Next NF
'                End If
'
'                'El tercer pipe, si tiene es el ancho col1
'                C1 = RecuperaValor(Cad, 3)
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 4360
'                End If
'                lw1.ColumnHeaders(1).Width = NF
'                'El cuarto pipe si tiene es el ancho de col2
'                C1 = RecuperaValor(Cad, 4)
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 1400
'                End If
'                lw1.ColumnHeaders(2).Width = NF
'            End If
        End If
    Else 'Escribir
        NF = FreeFile
        Open Cad For Output As #NF
        Cad = "NO ncesito|"
'        If lw1.ListItems.Count > 0 Then
'            Cad = Cad & lw1.SelectedItem.Text
'        Else
'            Cad = Cad & " "
'        End If
'        Cad = Cad & "|" & Int(Round(lw1.ColumnHeaders(1).Width, 2)) & "|" & Int(Round(lw1.ColumnHeaders(2).Width, 2)) & "|"
'        AnchoLogin = Cad
        Print #NF, Cad
        Close #NF
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub


