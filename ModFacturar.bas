Attribute VB_Name = "ModFacturar"
Option Explicit

'===================================================================================
'Modulo para el traspaso de registros de cabecera y lineas de las tablas de ALBARAN
'A las tablas del FACTURACION
' o para pasar de las tablas de Mantenimientos a tablas de FACTURACION
'====================================================================================

'operador del albaran para facturas de Mantenimientos
Private OpeFactu As String
Private MesFactu As String 'mes a facturar para Mantenimientos
Private TipCoMan As String 'tipo de contrato del mantenimiento

'Variables comunes en Albaranes para la cabecera de la FACTURA
Private LetraSer As String

Private TipoAlb As String
Private TipoFac As String

'Variable con la WHERE que selecciona todos los Albaranes que forma parte de la Factura
Private cadW As String


Dim Errores As String
Dim ErroresAux As String


Public Function TraspasoAlbaranesFacturas(cadSQL As String, cadWHERE As String, FechaFact As String, banPr As String, ByRef PBar1 As ProgressBar, ByRef LblBar As Label, ImprimeLasFacturasGeneradas As Boolean, ByRef vTipom As String, TextosCSB As String, Socio As Boolean) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'      FechaFact: Fecha de la Factura
'      BanPr: Cod. de Banco Propio
'      Pbar1:  Una progressbar. Se puede mandar un NOTHING, y no pasa nada. Si no se manda
'              es que estamos en un proceso corto o que no necesitabaos un pb1, con lo cual NO muestro el PB1
'      Imprime: Si despues de generarlo los imprime
'
'       vTipom:  Que tipo de albaran es, para luego la impresion saber que factura imprime
'      TextosCSB:  Si lleva llevara 3 lineas para meter ent tesoreria

'Desde Albaranes Genera las Facturas correspondientes
Dim RSalb As ADODB.Recordset 'Ordenados por: tipofac,clien,dpto,forma pago, dtoppago, dtognral
Dim b As Boolean
Dim SQL As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim antClien As Long
Dim antDirec As Long
Dim antForpa As Integer
Dim antDtoPP As Single, antDtoGn As Single

'direc/dpto actual para controlar el valor nulo
Dim actDirec As Long

'Concatenamos todas las facturas generadas para listarlas en el informe
Dim ListFactu As String
Dim vFactu As CFactura
Dim Inc As Integer
Dim condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura

'Por si no mando una progressbar, que no de errores
Dim PgbVisible As Boolean

    On Error GoTo ETraspasoAlbFac

    TraspasoAlbaranesFacturas = False

    ListFactu = ""
        
    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("VENFAC") 'facturas de venta
    If Not BloqueoManual("VENFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los albaranes que vamos a facturar (cabeceras y lineas)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    SQL = " (scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien ) INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar "
    If Not BloqueaRegistro(SQL, cadWHERE) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        Exit Function
    End If
    
   
    'Inicializar la Progress Bar
    PgbVisible = False
    If Not (PBar1 Is Nothing) Then
        If PBar1.visible Then PgbVisible = True
    End If
    If PgbVisible Then
        If InStr(1, cadSQL, "sclien") Then
            SQL = Replace(cadSQL, "scaalb.*, sclien.periodof", "count(*)") 'si hay INNER JOIN con sclien
        Else
            SQL = Replace(cadSQL, "*", "count(*)") 'si NO hay INNER JOIN con sclien
        End If
        
        
        Set RSalb = New ADODB.Recordset
        RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RSalb.EOF Then
            CargarProgresNew PBar1, CInt(RSalb.Fields(0))
            LblBar.Caption = "Inicializando el proceso..."
            LblBar.Refresh
            
        End If
        RSalb.Close
        Set RSalb = Nothing
    End If
    
        
    'EMPEZAMOS LA FACTURA
    Set vFactu = New CFactura
    vFactu.FecFactu = FechaFact 'Fecha para las Facturas

    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = banPr
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", banPr, "N")
    
    'comprobar que la cuenta prevista de cobro tiene valor
    b = (vFactu.CuentaPrev <> "")
    If Not b Then
        Set vFactu = Nothing
        'Desbloqueamos ya no estamos facturando
        DesBloqueoManual ("VENFAC")
        TerminaBloquear
        MsgBox "La cta. prevista de cobro debe tener valor.", vbExclamation
        Exit Function
    End If
    
       
        
    'Marcar Albaranes que se van a Facturar
    '----------------------------------------
    SQL = cadSQL & " ORDER BY scaalb.tipofact, scaalb.codclien, scaalb.coddirec, codforpa, dtoppago, dtognral "
    Set RSalb = New ADODB.Recordset
    RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    'Agrupar los Albaranes posibles en una misma Factura
    'Calcular y Grabar Factura en la Tabla de Facturas
    'Albaran(scaalb, slialb) -> Factura (scafac,scafac1,slifac)
    '----------------------------------------------------
    'Se factura por cliente y departamento
    'Agrupar albaranes en 1 factura por : tipofact,codclien,coddirec,codforpa,dtoppago, dtognral
    antClien = 0 'cliente
    antDirec = 0 'direccion/departamento
    antForpa = 0 'forma de pago
    antDtoPP = 0 'dto pronto pago
    antDtoGn = 0 'dto general
    
    cadW = ""
    Errores = ""
    Inc = 0
    
    While Not RSalb.EOF
        TipoAlb = RSalb!codtipom
        Inc = Inc + 1
        If IsNull(RSalb!CodDirec) Then
            actDirec = -1
        Else
            actDirec = DBLet(RSalb!CodDirec, "N")
        End If
        
        If RSalb!TipoFact = 1 Then 'tipofact=1 "FACTURA x ALBARAN"
        '---------------------------------------------------------
            'frmListadoPed.lblProgess(0).Caption = "Facturando: Facturas individuales"
            LblBar.Caption = "Facturando: Facturas individuales"
            LblBar.Refresh
            If cadW <> "" Then 'Facturacion pendiente
                cadW = cadW & ")) "
                If Not vFactu.PasarAlbaranesAFactura2(False, TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
                    If b Then b = False
                    AnyadirAvisos ErroresAux
                Else 'a�adirlo a la lista de facturas a imprimir
                    If ListFactu = "" Then
                        ListFactu = vFactu.NumFactu
                    Else
                        ListFactu = ListFactu & "," & vFactu.NumFactu
                    End If
                End If
                If PgbVisible Then
                    IncrementarProgresNew PBar1, Inc - 1
                    LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien
                    LblBar.Refresh
                End If
                Espera 0.2
                'Empezamos una nueva Factura
                cadW = ""
            End If
            
            'Los Albaranes que tengan tipofact=1 "factura x Albaran" generar una factura
            'para cada uno de ellos
            cadW = " scaalb.codtipom='" & RSalb!codtipom & "' AND scaalb.numalbar=" & RSalb!NumAlbar
            
            'Generar una Factura nueva
            vFactu.Cliente = RSalb!CodClien
            vFactu.NombreClien = RSalb!nomclien
            vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
            vFactu.CPostal = DBLet(RSalb!codpobla, "T")
            vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
            vFactu.Provincia = DBLet(RSalb!proclien, "T")
            vFactu.NIF = DBLet(RSalb!nifClien, "T")
            vFactu.Telefono = DBLet(RSalb!telclien, "T")
            vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
            vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
            vFactu.Agente = RSalb!codagent
            vFactu.ForPago = RSalb!codforpa
            vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
            vFactu.DtoPPago = CCur(RSalb!DtoPPago)
            vFactu.DtoGnral = CCur(RSalb!DtoGnral)

                
                
            If Not vFactu.PasarAlbaranesAFactura2(Socio, TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
                If b Then b = False
                AnyadirAvisos ErroresAux
            Else 'a�adirlo a la lista de facturas a imprimir
                If ListFactu = "" Then
                    ListFactu = vFactu.NumFactu
                Else
                    ListFactu = ListFactu & "," & vFactu.NumFactu
                End If
            End If
            If PgbVisible Then
                Inc = 1 '1 albaran x factura
                LblBar.Caption = "Cliente: " & Format(RSalb!CodClien, "000000") & " - " & RSalb!nomclien
                LblBar.Refresh
                IncrementarProgresNew PBar1, Inc
                Inc = 0
            End If
            Espera 0.2
                
            cadW = ""
            
        Else 'tipofac=0 "factura COLECTIVA"
        '----------------------------------------------------------
            'Seleccionar todos los Albaranes pertenecientes a un mismo Cliente,Departamento
            'Los que tengan tipofac=0 "factura colectiva" agruparlos en una misma factura
            'para la misma Forma de PAgo, mismo dtoppago y mismo dtognral
             
             '-- David.      Esta linea da error si no viene de frmlistadoped
             'frmListadoPed.lblProgess(0).Caption = "Facturando: Facturas colectivas"
             LblBar.Caption = "Facturando: Facturas colectivas"
             LblBar.Refresh
             '---- Laura: 06/10/2006
             'Comprobar si es Departamento o Direccion (segun paramatro)
             If vParamAplic.Departamento Then
                'agrupar tb por departamento
                condicion = (antClien <> RSalb!CodClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral)
             Else
                condicion = (antClien <> RSalb!CodClien) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral)
             End If
             
'             If (antClien <> RSalb!CodClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral) Then
             If condicion Then
             '-----
                If cadW <> "" Then 'Facturacion PEndiente
                    cadW = cadW & ")) "
                    If Not vFactu.PasarAlbaranesAFactura2(Socio, TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
                        If b Then b = False
                        AnyadirAvisos ErroresAux
                    Else 'a�adirlo a la lista de facturas a imprimir
                        If ListFactu = "" Then
                            ListFactu = vFactu.NumFactu
                        Else
                            ListFactu = ListFactu & "," & vFactu.NumFactu
                        End If
                    End If
                    If PgbVisible Then
                        LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien
                        LblBar.Refresh
                        IncrementarProgresNew PBar1, Inc
                        Inc = 0
                    End If
                    Espera 0.2
                    
                    'Empezamos una nueva Factura
                    cadW = ""
                End If
                'Generar una Factura nueva
                vFactu.Cliente = RSalb!CodClien
                vFactu.NombreClien = RSalb!nomclien
                vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
                vFactu.CPostal = DBLet(RSalb!codpobla, "T")
                vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
                vFactu.Provincia = DBLet(RSalb!proclien, "T")
                vFactu.NIF = DBLet(RSalb!nifClien, "T")
                vFactu.Telefono = DBLet(RSalb!telclien, "T")
                vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
                vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
                vFactu.Agente = RSalb!codagent
                vFactu.ForPago = RSalb!codforpa
                vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
                vFactu.DtoPPago = CCur(RSalb!DtoPPago)
                vFactu.DtoGnral = CCur(RSalb!DtoGnral)
                vFactu.Aportacion = 0
                If RSalb!codtipom = "ALM" Then vFactu.Aportacion = DBLet(RSalb!Aportacion, "N")
                cadW = " (scaalb.codtipom='" & RSalb!codtipom & "' AND scaalb.numalbar IN (" & RSalb!NumAlbar
            Else
                cadW = cadW & ", " & RSalb!NumAlbar
            End If
        
            'Guardamos datos del registro anterior
            antClien = RSalb!CodClien
'            antDirec = DBLet(RSalb!CodDirec, "N")
            antDirec = actDirec
            antForpa = RSalb!codforpa
            antDtoPP = RSalb!DtoPPago
            antDtoGn = RSalb!DtoGnral
        End If
        RSalb.MoveNext
    Wend
    RSalb.Close
    Set RSalb = Nothing
        
    'Facturar la ultima Factura generada del blucle
    If cadW <> "" Then
        cadW = cadW & "))"
        If PgbVisible Then LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
        
        If Not vFactu.PasarAlbaranesAFactura2(Socio, TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
            If b Then b = False
            AnyadirAvisos "Error Facturando el Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien & vbCrLf & ErroresAux
        Else 'a�adirlo a la lista de facturas a imprimir
            If ListFactu = "" Then
                ListFactu = vFactu.NumFactu
            Else
                ListFactu = ListFactu & "," & vFactu.NumFactu
            End If
        End If
        If PgbVisible Then
'            LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
            IncrementarProgresNew PBar1, Inc
        End If
        Espera 0.2
    End If
    
    TipoFac = vFactu.codtipom
    Set vFactu = Nothing
    TraspasoAlbaranesFacturas = True
    
    If b Then
        LblBar.Caption = "Proceso finalizado correctamente."
        MsgBox "Las Facturas de los Albaranes seleccionados se generaron correctamente.", vbInformation
    Else
        LblBar.Caption = "Proceso finalizado con errores."
        SQL = "ATENCI�N:" & vbCrLf
        MsgBox SQL & "No todas las Facturas se generaron correctamente!!!.", vbExclamation
        If Errores <> "" Then MostrarAvisos
    End If
    
    Espera 0.2
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("VENFAC")
    TerminaBloquear
    
    
    If ImprimeLasFacturasGeneradas Then
        If ListFactu <> "" Then
            
            ImprimirFacturas ListFactu, FechaFact, , DevuelveTipoDocumentoFactura(vTipom)
        End If
    End If
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando Albaranes", Err.Description
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        TerminaBloquear
    End If
End Function




'#Laura: 14/11/2006 Recuperar facturas Alzira
Public Function TraspasoAlbaranesFacturas_RecuperaFac(cadSQL As String, cadWHERE As String, FechaFact As String, banPr As String, NumFac As String, ByRef PBar As ProgressBar, ByRef LblBar As Label) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'      FechaFact: Fecha de la Factura
'      BanPr: Cod. de Banco Propio
'Desde Albaranes Genera las Facturas correspondientes
Dim RSalb As ADODB.Recordset 'Ordenados por: tipofac,clien,dpto,forma pago, dtoppago, dtognral
Dim b As Boolean
Dim SQL As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim antClien As Long
Dim antDirec As Long
Dim antForpa As Byte
Dim antDtoPP As Single, antDtoGn As Single

'direc/dpto actual para controlar el valor nulo
Dim actDirec As Long

'Concatenamos todas las facturas generadas para listarlas en el informe
Dim ListFactu As String
Dim vFactu As CFactura
Dim Inc As Integer
Dim condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura



    On Error GoTo ETraspasoAlbFac

    TraspasoAlbaranesFacturas_RecuperaFac = False

    ListFactu = ""
        
    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("VENFAC") 'facturas de venta
    If Not BloqueoManual("VENFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los albaranes que vamos a facturar (cabeceras y lineas)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    'SQL = " (scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien ) INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar "
    SQL = " (scaalb  INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar) "
    If Not BloqueaRegistro(SQL, cadWHERE) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        Exit Function
    End If
    
   
    'Inicializar la Progress Bar
    If PBar.visible Then
        If InStr(1, cadSQL, "sclien") Then
            SQL = Replace(cadSQL, "scaalb.*, sclien.periodof", "count(*)") 'si hay INNER JOIN con sclien
        Else
            SQL = Replace(cadSQL, "*", "count(*)") 'si NO hay INNER JOIN con sclien
        End If
        
        
        Set RSalb = New ADODB.Recordset
        RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RSalb.EOF Then
            CargarProgresNew PBar, CInt(RSalb.Fields(0))
            LblBar.Caption = "Inicializando el proceso..."
        End If
        RSalb.Close
        Set RSalb = Nothing
    End If
    
        
    'EMPEZAMOS LA FACTURA
    Set vFactu = New CFactura
    vFactu.FecFactu = FechaFact 'Fecha para las Facturas
    vFactu.NumFactu = NumFac
    
    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = banPr
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", banPr, "N")
    
    'comprobar que la cuenta prevista de cobro tiene valor
    b = (vFactu.CuentaPrev <> "")
    If Not b Then
        Set vFactu = Nothing
        'Desbloqueamos ya no estamos facturando
        DesBloqueoManual ("VENFAC")
        TerminaBloquear
        MsgBox "La cta. prevista de cobro debe tener valor.", vbExclamation
        Exit Function
    End If
    
       
        
    'Marcar Albaranes que se van a Facturar
    '----------------------------------------
    SQL = cadSQL & " ORDER BY scaalb.tipofact, scaalb.codclien, scaalb.coddirec, codforpa, dtoppago, dtognral "
    Set RSalb = New ADODB.Recordset
    RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    'Agrupar los Albaranes posibles en una misma Factura
    'Calcular y Grabar Factura en la Tabla de Facturas
    'Albaran(scaalb, slialb) -> Factura (scafac,scafac1,slifac)
    '----------------------------------------------------
    'Se factura por cliente y departamento
    'Agrupar albaranes en 1 factura por : tipofact,codclien,coddirec,codforpa,dtoppago, dtognral
    antClien = 0 'cliente
    antDirec = 0 'direccion/departamento
    antForpa = 0 'forma de pago
    antDtoPP = 0 'dto pronto pago
    antDtoGn = 0 'dto general
    
    cadW = ""
    Errores = ""
    Inc = 0
    
    While Not RSalb.EOF
        TipoAlb = RSalb!codtipom
        Inc = Inc + 1
        If IsNull(RSalb!CodDirec) Then
            actDirec = -1
        Else
            actDirec = DBLet(RSalb!CodDirec, "N")
        End If
        
        If RSalb!TipoFact = 1 Then 'tipofact=1 "FACTURA x ALBARAN"
        '---------------------------------------------------------
            frmListadoPed.lblProgess(0).Caption = "Facturando: Facturas individuales"
            
            If cadW <> "" Then 'Facturacion pendiente
                cadW = cadW & ")) "
                'le pasamos el parametro de estaRecuperando a true
                If Not vFactu.PasarAlbaranesAFactura2(False, TipoAlb, cadW, "", ErroresAux, True) Then
                    If b Then b = False
                    AnyadirAvisos ErroresAux
                Else 'a�adirlo a la lista de facturas a imprimir
                    If ListFactu = "" Then
                        ListFactu = vFactu.NumFactu
                    Else
                        ListFactu = ListFactu & "," & vFactu.NumFactu
                    End If
                End If
                If PBar.visible Then
                    LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien
                    IncrementarProgresNew PBar, Inc - 1
                End If
                Espera 0.2
                'Empezamos una nueva Factura
                cadW = ""
            End If
            
            'Los Albaranes que tengan tipofact=1 "factura x Albaran" generar una factura
            'para cada uno de ellos
            cadW = " scaalb.codtipom='" & RSalb!codtipom & "' AND scaalb.numalbar=" & RSalb!NumAlbar
            
            'Generar una Factura nueva
            vFactu.Cliente = RSalb!CodClien
            vFactu.NombreClien = RSalb!nomclien
            vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
            vFactu.CPostal = DBLet(RSalb!codpobla, "T")
            vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
            vFactu.Provincia = DBLet(RSalb!proclien, "T")
            vFactu.NIF = DBLet(RSalb!nifClien, "T")
            vFactu.Telefono = DBLet(RSalb!telclien, "T")
            vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
            vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
            vFactu.Agente = RSalb!codagent
            vFactu.ForPago = RSalb!codforpa
            vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
            vFactu.DtoPPago = CCur(RSalb!DtoPPago)
            vFactu.DtoGnral = CCur(RSalb!DtoGnral)
                
            If Not vFactu.PasarAlbaranesAFactura2(False, TipoAlb, cadW, "", ErroresAux, True) Then
                If b Then b = False
                AnyadirAvisos ErroresAux
            Else 'a�adirlo a la lista de facturas a imprimir
                If ListFactu = "" Then
                    ListFactu = vFactu.NumFactu
                Else
                    ListFactu = ListFactu & "," & vFactu.NumFactu
                End If
            End If
            If PBar.visible Then
                Inc = 1 '1 albaran x factura
                LblBar.Caption = "Cliente: " & Format(RSalb!CodClien, "000000") & " - " & RSalb!nomclien
                IncrementarProgresNew PBar, Inc
                Inc = 0
            End If
            Espera 0.2
                
            cadW = ""
            
        Else 'tipofac=0 "factura COLECTIVA"
        '----------------------------------------------------------
            'Seleccionar todos los Albaranes pertenecientes a un mismo Cliente,Departamento
            'Los que tengan tipofac=0 "factura colectiva" agruparlos en una misma factura
            'para la misma Forma de PAgo, mismo dtoppago y mismo dtognral
             
             frmListadoPed.lblProgess(0).Caption = "Facturando: Facturas colectivas"
             
             '---- Laura: 06/10/2006
             'Comprobar si es Departamento o Direccion (segun paramatro)
             If vParamAplic.Departamento Then
                'agrupar tb por departamento
                condicion = (antClien <> RSalb!CodClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral)
             Else
                condicion = (antClien <> RSalb!CodClien) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral)
             End If
             
'             If (antClien <> RSalb!CodClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral) Then
             If condicion Then
             '-----
                If cadW <> "" Then 'Facturacion PEndiente
                    cadW = cadW & ")) "
                    If Not vFactu.PasarAlbaranesAFactura2(False, TipoAlb, cadW, "", ErroresAux, True) Then
                        If b Then b = False
                        AnyadirAvisos ErroresAux
                    Else 'a�adirlo a la lista de facturas a imprimir
                        If ListFactu = "" Then
                            ListFactu = vFactu.NumFactu
                        Else
                            ListFactu = ListFactu & "," & vFactu.NumFactu
                        End If
                    End If
                    If PBar.visible Then
                        LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien
                        IncrementarProgresNew PBar, Inc
                        Inc = 0
                    End If
                    Espera 0.2
                    
                    'Empezamos una nueva Factura
                    cadW = ""
                End If
                'Generar una Factura nueva
                vFactu.Cliente = RSalb!CodClien
                vFactu.NombreClien = RSalb!nomclien
                vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
                vFactu.CPostal = DBLet(RSalb!codpobla, "T")
                vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
                vFactu.Provincia = DBLet(RSalb!proclien, "T")
                vFactu.NIF = DBLet(RSalb!nifClien, "T")
                vFactu.Telefono = DBLet(RSalb!telclien, "T")
                vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
                vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
                vFactu.Agente = RSalb!codagent
                vFactu.ForPago = RSalb!codforpa
                vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
                vFactu.DtoPPago = CCur(RSalb!DtoPPago)
                vFactu.DtoGnral = CCur(RSalb!DtoGnral)
                
                cadW = " (scaalb.codtipom='" & RSalb!codtipom & "' AND scaalb.numalbar IN (" & RSalb!NumAlbar
            Else
                cadW = cadW & ", " & RSalb!NumAlbar
            End If
        
            'Guardamos datos del registro anterior
            antClien = RSalb!CodClien
'            antDirec = DBLet(RSalb!CodDirec, "N")
            antDirec = actDirec
            antForpa = RSalb!codforpa
            antDtoPP = RSalb!DtoPPago
            antDtoGn = RSalb!DtoGnral
        End If
        RSalb.MoveNext
    Wend
    RSalb.Close
    Set RSalb = Nothing
        
    'Facturar la ultima Factura generada del blucle
    If cadW <> "" Then
        cadW = cadW & "))"
        If PBar.visible Then LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
        'ponemos a true el campo de EstaRecuperando facturas
        If Not vFactu.PasarAlbaranesAFactura2(False, TipoAlb, cadW, "", ErroresAux, True) Then
            If b Then b = False
            AnyadirAvisos "Error Facturando el Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien & vbCrLf & ErroresAux
        Else 'a�adirlo a la lista de facturas a imprimir
            If ListFactu = "" Then
                ListFactu = vFactu.NumFactu
            Else
                ListFactu = ListFactu & "," & vFactu.NumFactu
            End If
        End If
        If PBar.visible Then
'            LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
            IncrementarProgresNew PBar, Inc
        End If
        Espera 0.2
    End If
    
    TipoFac = vFactu.codtipom
    Set vFactu = Nothing
    TraspasoAlbaranesFacturas_RecuperaFac = True
    
    If b Then
        LblBar.Caption = "Proceso finalizado correctamente."
        MsgBox "Las Facturas de los Albaranes seleccionados se generaron correctamente.", vbInformation
    Else
        LblBar.Caption = "Proceso finalizado con errores."
        SQL = "ATENCI�N:" & vbCrLf
        MsgBox SQL & "No todas las Facturas se generaron correctamente!!!.", vbExclamation
        If Errores <> "" Then MostrarAvisos
    End If
    
    Espera 0.2
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("VENFAC")
    TerminaBloquear
    
'    If ListFactu <> "" Then
'        ImprimirFacturas ListFactu, FechaFact
'    End If

ETraspasoAlbFac:
    If Err.Number <> 0 Then
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        TerminaBloquear
        TraspasoAlbaranesFacturas_RecuperaFac = False
        MuestraError Err.Number, "Facturando Albaranes", Err.Description
    End If
End Function





Private Sub AnyadirAvisos(Donde As String)
    Errores = Errores & vbCrLf & vbCrLf & Donde & vbCrLf
End Sub



Private Sub MostrarAvisos()
    frmMensajes.vCampos = Errores
    frmMensajes.OpcionMensaje = 13
    frmMensajes.Show vbModal
End Sub


'========================================================

Public Function ComprobarFechaVenci(FechaVenci As Date, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim newFecha As Date
Dim b As Boolean

'=== Modificada Laura: 23/01/2007
    On Error GoTo ErrObtFec
    b = False
    
    '--- comprobar que tiene dias de pago para obtener nueva fecha
    If Not (Dia1 > 0 Or Dia2 > 0 Or Dia3 > 0) Then
        'si no tiene dias de pago la fecha es OK y fin
        ComprobarFechaVenci = FechaVenci
        Exit Function
    End If
        
    
    '--- Obtener nueva fecha del vencimiento
    newFecha = FechaVenci
    
    Do
        'si dia de la fecha vencimiento es uno de los 3 dias de pagos fecha es OK
        If Day(newFecha) = Dia1 Or Day(newFecha) = Dia2 Or Day(newFecha) = Dia3 Then
'            newFecha = CStr(newFecha)
            b = True
        Else
            'mientras esta en el mismo mes vamos aumentando dias hasta encontrar un dia de pago
            newFecha = DateAdd("d", 1, CDate(newFecha))
        End If
    Loop Until b = True Or Year(newFecha) = Year(FechaVenci) + 3
    
    ComprobarFechaVenci = newFecha
    Exit Function
    
ErrObtFec:
    MuestraError Err.Number, "Obtener Fecha vencimiento seg�n dias de pago.", Err.Description
End Function





Public Function ComprobarFechaVenci_old(FechaVenci As Date, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim fechaV As Date
'Dim cadDias As String
Dim F As String

    fechaV = FechaVenci
    If Dia1 <> 0 Or Dia2 <> 0 Or Dia3 <> 0 Then
        OrdenarDias Dia1, Dia2, Dia3
        If Dia1 >= Day(fechaV) Then
            fechaV = Format(Dia1 & "/" & Month(fechaV) & "/" & Year(fechaV), "dd/mm/yyyy")
        Else
            If Dia2 >= Day(fechaV) Then
                fechaV = Format(Dia2 & "/" & Month(fechaV) & "/" & Year(fechaV), "dd/mm/yyyy")
            Else
                If Dia3 >= Day(fechaV) Then
                    fechaV = Format(Dia3 & "/" & Month(fechaV) & "/" & Year(fechaV), "dd/mm/yyyy")
                
                Else
                    'coger el primero del mes siguiente
                    If Dia1 <> 0 Then
                        F = Dia1 & "/"
                        
                    ElseIf Dia2 <> 0 Then
                        F = Dia2 & "/"
'                        fechaV = Format(Dia2 & "/" & Month(fechaV) + 1 & "/" & Year(fechaV), "dd/mm/yyyy")
                    ElseIf Dia3 <> 0 Then
                        F = Dia3 & "/"
'                        fechaV = Format(Dia3 & "/" & Month(fechaV) + 1 & "/" & Year(fechaV), "dd/mm/yyyy")
                    End If
                    If Month(fechaV) + 1 < 13 Then
                        F = F & Month(fechaV) + 1 & "/" & Year(fechaV)
                    Else
                        F = F & "01/" & Year(fechaV) + 1
                    End If
                    fechaV = Format(F, "dd/mm/yyyy")
                End If
            End If
        End If

    End If
    ComprobarFechaVenci_old = fechaV
End Function





Private Sub OrdenarDias(Dia1 As Byte, Dia2 As Byte, Dia3 As Byte)
'Entran los dias desordenados: dia1=10, dia2=5, dia3=20
'devuelve los dias ordenados: dia1=5, dia2=10, dia3=20
Dim diaAux As Byte

    On Error GoTo EOrdenar

    If Dia1 < Dia2 And Dia1 < Dia3 Then
        'dia 1 es el menor
        If Dia2 > Dia3 Then
            diaAux = Dia2
            Dia2 = Dia3
            Dia3 = diaAux
        End If
    ElseIf Dia2 < Dia3 Then
        'dia2 es el menor
        diaAux = Dia1
        Dia1 = Dia2
        If diaAux < Dia3 Then
            Dia2 = diaAux
        Else
            Dia2 = Dia3
            Dia3 = diaAux
        End If
    Else
        'dia3 es el menor
        diaAux = Dia1
        Dia1 = Dia3
        If diaAux < Dia2 Then
            Dia3 = Dia2
            Dia2 = diaAux
        Else
            Dia3 = diaAux
        End If
    End If

EOrdenar:
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function ComprobarMesNoGira(FecVenci As Date, MesNG As Byte, DiaVtoAt As Byte, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim F As String
Dim diaPago As Byte

    If Month(FecVenci) = MesNG Then
        '### LAURA 14/08/2008
'        If DiaVtoAt > 0 Then
'            F = DiaVtoAt & "/"
'        Else
'            F = Day(FecVenci) & "/"
'        End If
        
'        If Month(FecVenci) + 1 < 13 Then
'            F = F & Month(FecVenci) + 1 & "/" & Year(FecVenci)
'        Else
'            F = F & "01/" & Year(FecVenci) + 1
'        End If

        If DiaVtoAt > 0 Then
            'si tiene dia de vto atrasado a ese dia del mes siguiente
            'al mes a no girar
            F = DiaVtoAt & "/"
            F = F & Month(FecVenci) & "/" & Year(FecVenci)
            F = DateAdd("m", 1, F)
        Else
            'si no tiene dia de vto atrasado el primer dia de pago
            'del mes siguiente si tiene o sino el siguiente mes del
            'vencimiento obtenido
            If Dia1 > 0 Or Dia2 > 0 Or Dia3 > 0 Then
                'tiene dias de pago: el menor dia del mes siguiente
                diaPago = Dia1
                If (diaPago = 0) Or ((Dia2 < diaPago) And Dia2 <> 0) Then diaPago = Dia2
                If (diaPago = 0) Or ((Dia3 < diaPago) And Dia3 <> 0) Then diaPago = Dia3
                
                F = diaPago & "/"
                F = F & Month(FecVenci) & "/" & Year(FecVenci)
            Else
                'no tiene dias de pago: al mes siguiente
                F = Day(FecVenci) & "/"
                F = F & Month(FecVenci) & "/" & Year(FecVenci)
            End If
            
            F = DateAdd("m", 1, F)
        End If
        '###
        
        FecVenci = Format(F, "dd/mm/yyyy")
    End If
    
    ComprobarMesNoGira = FecVenci
End Function





Public Sub ImprimirHojaExpedicion(OpcionListado As Byte, NumAlb As String, tipMov As String, Optional fecAlb As String)
Dim cadFormula As String
Dim cadParam As String
Dim cadSelect As String 'select para insertar en tabla temporal
Dim numParam As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim ImpresionDirecta As Boolean
'Dim codClien As String
'Dim EsHistorico As Boolean
Dim NombreTabla As String
Dim NomTablaLineas As String

    If NumAlb = "" Then
        MsgBox "Debe seleccionar un Albaran para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
'    EsHistorico = (fecAlb <> "")
    
'    If EsHistorico <> "" Then 'es historico
        NombreTabla = "scaalb"
        NomTablaLineas = "slialb" 'Tabla lineas de Albaranes
'    Else
'        NombreTabla = "schalb"
'        NomTablaLineas = "slhalb"
'    End If
    
    
    '===================================================
    '============ PARAMETROS ===========================
'    If (OpcionListado = 45) Then
'        If EsInformePortes Then
            'Es el de portes
             indRPT = 34
'        Else
'            If hcoCodTipoM = "ALZ" Then
'                indRPT = 29   'Albaranes B
'            Else
'                If EsHistorico Then
'                    indRPT = 11 'Hist. Albaranes clientes
'                Else
'                    indRPT = 10 'Albaran Clientes
'                End If
'            End If
'        End If
'    End If
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, ImpresionDirecta, pPdfRpt) Then Exit Sub
   
    'A�adir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    cadParam = cadParam & "pCodUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
    
    'PORTES
    cadParam = cadParam & "vPortes=""" & vParamAplic.ArtPortes & """|"
    numParam = numParam + 1
    
'    'PUNTO VERDE
'    cadParam = cadParam & "PuntoVerde=""" & vParamAplic.ArtReciclado & """|"
'    numParam = numParam + 1
    
'    'Si se imprimen importes y/o
'    devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", codClien, "N")
'    If devuelve = "" Then devuelve = "0"
'    ' 0 "Todo"
'    ' 1 "Cantidad y Precio"
'    ' 2 "Cantidad"
'    cadParam = cadParam & "Albarcon=" & devuelve & "|"
'    numParam = numParam + 1
    
    
    'Nombre fichero .rpt a Imprimir
    If Not ImpresionDirecta Then
        frmImprimir.NombreRPT = nomDocu
        frmImprimir.NombrePDF = pPdfRpt
    End If
        
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion N� de Albaran
    '---------------------------------------------------
    If NumAlb <> "" Then
        '- Cod Tipo Movimiento
        devuelve = "{" & NombreTabla & ".codtipom}='" & tipMov & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        '- N� Albaran
        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(NumAlb)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        cadSelect = cadFormula
        
'        If EsHistorico <> "" Then 'historico
'            'El campo fecha tambien es clave primaria
'            devuelve = fecAlb
'            devuelve = "{" & NombreTabla & ".fechaalb}=Date(" & Year(devuelve) & "," & Month(devuelve) & "," & Day(devuelve) & ")"
'            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
'
'            devuelve = "{" & NombreTabla & ".fechaalb}='" & Format(fecAlb, FormatoFecha) & "'"
'            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
'        End If
    End If
   
'    '=========================================================================
'    'Aqui sabemos que valor tiene CodClien y a�adimos a los parametros el tipo de IVA
'    'que se aplica a ese cliente
'    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", codClien, "N")
'    If devuelve <> "" Then
'        cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
'        numParam = numParam + 1
'    End If

        
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    devuelve = NombreTabla & " INNER JOIN " & NomTablaLineas & " ON "
    devuelve = devuelve & NombreTabla & ".codtipom=" & NomTablaLineas & ".codtipom AND " & NombreTabla & ".numalbar= " & NomTablaLineas & ".numalbar "
'    If EsHistorico Then devuelve = devuelve & " AND " & NombreTabla & ".fechaalb= " & NomTablaLineas & ".fechaalb "
    If Not HayRegParaInforme(devuelve, cadSelect) Then Exit Sub
    
    
    If ImpresionDirecta Then
        'Imrpimie directamente. Tipo 4tonda.  -----------
        If MsgBox("�Imprimir el albar�n?", vbQuestion + vbYesNo) = vbYes Then ImprimirDirectoAlb cadSelect
    Else
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = OpcionListado
            If indRPT = 34 Then
                .Titulo = "Portes albaran "
'            Else
'                .Titulo = "Albaran de Cliente"
            End If
            .ConSubInforme = True
            .Show vbModal
        End With
    End If
End Sub





'FormatoFactura:
'               0.- Normal
'               1.- TPV
'               2.- Factura "B"
'               3.- Facturas de cuotas socios

Public Sub ImprimirFacturas(listaF As String, fechaF As String, Optional SQL As String, Optional FormatoFactura As Byte)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NombreTabla As String
Dim ImprimeDirecto As Boolean


    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    NombreTabla = "scafac"

    '===================================================
    '============ PARAMETROS ===========================
    If FormatoFactura = 0 Then
        indRPT = 12 'Facturas Clientes  NORMAL
        
    ElseIf FormatoFactura = 1 Then
        indRPT = 18 'FACTURAS TPV
    
    ElseIf FormatoFactura = 2 Then
        indRPT = 30 'FACTURAS "B"
        
    ElseIf FormatoFactura = 3 Then
        indRPT = 49 'FACTURAS DE CUOTAS DE SOCIOS
    
    End If
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, ImprimeDirecto, pPdfRpt) Then
        Exit Sub
    End If


'--[Monica]: cambiado por gastos de administracion
'    'PUNTO VERDE
'    '--------------------------------------------------------------------------
'    cadParam = cadParam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
'    numParam = numParam + 1
    

    'Nombre fichero .rpt a Imprimir
    If Not ImprimeDirecto Then
        frmImprimir.NombreRPT = nomDocu
        frmImprimir.NombrePDF = pPdfRpt
    End If

    If SQL <> "" Then
        'Llamo desde el menu de Reimprimir facturas y tengo construida la
        'cadena de seleccion D/H tipoMov, D/H NumFactu, D/H fecfactu
        cadSelect = SQL
        cadFormula = listaF
        cadParam = cadParam & fechaF
        numParam = numParam + 1
    Else
        'Llama desde PasarAlbaranes a  Facturas y al terminar las imprime
        '===================================================
        '================= FORMULA =========================
        'Cadena para seleccion N� de Factura
        '---------------------------------------------------
        'Cod Tipo Movimiento
        devuelve = "({" & NombreTabla & ".codtipom}='" & TipoFac & "') "
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
        'N� Factura
        devuelve = "({" & NombreTabla & ".numfactu} IN [" & listaF & "])"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
        'fecha factu
        devuelve = "(year({" & NombreTabla & ".fecfactu}) = " & Year(fechaF) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub

        cadSelect = cadFormula


    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub


     If ImprimeDirecto Then
         'Abrire un formulario por si acaso quieren cancelar la impresion. Ya que al ser
         'directa puede tardar mucho, haberse equivocado ......
        CadenaDesdeOtroForm = cadSelect
        frmVarios.Opcion = 0
        frmVarios.Show vbModal
        'Ha terminado la reimpresion
        
     Else
         With frmImprimir
                .FormulaSeleccion = cadFormula
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = 53
                .Titulo = ""
                .Show vbModal
        End With
    End If
End Sub


Public Sub ImprimirFacturasClien(listaF As String, fechaF As String, Optional SQL As String, Optional FormatoFactura As Byte)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NombreTabla As String
Dim ImprimeDirecto As Boolean


    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    NombreTabla = "scafaccli"

    '===================================================
    '============ PARAMETROS ===========================
    If FormatoFactura = 0 Then
        indRPT = 12 'Facturas Clientes  NORMAL
        
    ElseIf FormatoFactura = 1 Then
        indRPT = 18 'FACTURAS TPV
    
    ElseIf FormatoFactura = 2 Then
        indRPT = 30 'FACTURAS "B"
        
    ElseIf FormatoFactura = 3 Then
        indRPT = 49 'FACTURAS DE CUOTAS DE SOCIOS
    
    ElseIf FormatoFactura = 4 Then
        indRPT = 54 'FACTURAS RECTIFICATIVAS DE SERVICIOS DE SOCIOS
        
    ElseIf FormatoFactura = 5 Then
        indRPT = 47  'FACTURAS RECTIFICATIVAS DE PUBLICIDAD DE CLIENTES
    
    End If
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, ImprimeDirecto, pPdfRpt) Then
        Exit Sub
    End If


'--[Monica]: cambiado por gastos de administracion
'    'PUNTO VERDE
'    '--------------------------------------------------------------------------
'    cadParam = cadParam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
'    numParam = numParam + 1
    

    'Nombre fichero .rpt a Imprimir
    If Not ImprimeDirecto Then
        frmImprimir.NombreRPT = nomDocu
        frmImprimir.NombrePDF = pPdfRpt
    End If

    If SQL <> "" Then
        'Llamo desde el menu de Reimprimir facturas y tengo construida la
        'cadena de seleccion D/H tipoMov, D/H NumFactu, D/H fecfactu
        cadSelect = SQL
        cadFormula = listaF
        cadParam = cadParam & fechaF
        numParam = numParam + 1
    Else
        'Llama desde PasarAlbaranes a  Facturas y al terminar las imprime
        '===================================================
        '================= FORMULA =========================
        'Cadena para seleccion N� de Factura
        '---------------------------------------------------
        'Cod Tipo Movimiento
        devuelve = "({" & NombreTabla & ".codtipom}='" & TipoFac & "') "
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
        'N� Factura
        devuelve = "({" & NombreTabla & ".numfactu} IN [" & listaF & "])"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
        'fecha factura
        devuelve = "(year({" & NombreTabla & ".fecfactu}) = " & Year(fechaF) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub

        cadSelect = cadFormula


    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub


     If ImprimeDirecto Then
         'Abrire un formulario por si acaso quieren cancelar la impresion. Ya que al ser
         'directa puede tardar mucho, haberse equivocado ......
        CadenaDesdeOtroForm = cadSelect
        frmVarios.Opcion = 0
        frmVarios.Show vbModal
        'Ha terminado la reimpresion
        
     Else
         With frmImprimir
                .FormulaSeleccion = cadFormula
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = 53
                .Titulo = ""
                .Show vbModal
        End With
    End If
End Sub






Public Function TraspasoMtosAFacturas(cadSQL As String, cadSel As String, FechaFact As String, OpeFact As String, banPr As String, MesFact As String, ByRef Lbl As Label, CentroCoste As String) As Boolean      'Fecha de la factura, Operador
'IN -> cadSQL: cadena para seleccion de los mantenimientos que vamos a Facturar
'      FechaFact: Fecha de la Factura
'      OpeFact: Operador Factura
'
'   CentroCoste.   Si tiene analitica y el modoanalitica es por poryecto, es un dato del formulario
'
'Desde Mantenimientos Genera las Facturas correspondientes
Dim RSmto As ADODB.Recordset 'Ordenados por: clien,dpto,forma pago, dtoppago, dtognral
Dim b As Boolean
Dim SQL As String

Dim vClien As CCliente 'aqui cargamos los datos del cliente del mantenimiento para grabar en scafac
Dim vFactu As CFactura

Dim ListFactu As String
Dim Conta2 As Long

    On Error GoTo ETraspasoMtoFac


    TraspasoMtosAFacturas = False
    
    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("VENFAC") 'facturas de mantenimiento
    If Not BloqueoManual("VENFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los mantenimientos que vamos a facturar (cabeceras y lineas)
'    SQL = " (scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien ) INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar "
    SQL = " scaman "
    
    If Not BloqueaRegistro(SQL, cadSel) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        Exit Function
    End If
    
    
    
    
    
        
    'EMPEZAMOS LA FACTURA
    Set vFactu = New CFactura
    vFactu.FecFactu = FechaFact 'Fecha para las Facturas

    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = banPr
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", banPr, "N")
    
    OpeFactu = OpeFact 'operador de la factura de mantenimiento
    MesFactu = MesFact 'mes a factura para los mantenimientos
    
    b = True
    
    'Marcar Mantenimientos que se van a Facturar
    '----------------------------------------
    
    SQL = cadSQL & " ORDER BY scaman.codclien, scaman.coddirec, scaman.nummante "
    Set RSmto = New ADODB.Recordset
    Conta2 = InStr(1, cadSQL, " FROM ")
    ListFactu = "Select count(*) " & Mid(cadSQL, Conta2)
    
    
    
    RSmto.Open ListFactu, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Lbl.Tag = RSmto.Fields(0)
    RSmto.Close
    
    
    
    Conta2 = 0
    ListFactu = ""
    RSmto.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    'Le pongo                KEYSET      pq quiero contar los registros
    'Cada MAntenimiento genera una factura
    'Calcular y Grabar Factura en las Tablas de Facturas
    '---    -------------------------------------------------
     While Not RSmto.EOF
            
           Conta2 = Conta2 + 1
           Lbl.Caption = Conta2 & " de " & Lbl.Tag
           Lbl.Refresh
            
            If (RSmto.RecordCount Mod 10) = 9 Then DoEvents
        'para cada mantenimiento de la tabla scaman seleccionado para facturar
        vFactu.BrutoFac = CCur(RSmto!Importe)
        'tipo de contrato del mantenimientos
        TipCoMan = RSmto!codtipco
        
        'Datos de la Cabecera: Insertar en scafac
        '-----------------------------------------
        Set vClien = New CCliente
        If vClien.LeerDatos(RSmto!CodClien) Then
            'Datos cliente
            vFactu.Cliente = RSmto!CodClien
            vFactu.NombreClien = vClien.Nombre
            vFactu.DomicilioClien = vClien.Domicilio
            vFactu.CPostal = vClien.CPostal
            vFactu.Poblacion = vClien.Poblacion
            vFactu.Provincia = vClien.Provincia
            vFactu.NIF = vClien.NIF
            vFactu.Telefono = vClien.TfnoClien
            vFactu.DirDpto = DBLet(RSmto!CodDirec, "T")
            vFactu.NombreDirDpto = DBLet(RSmto!nomdirec, "T")
            vFactu.Agente = vClien.Agente
            'forma de pago del mantenimiento
            vFactu.ForPago = RSmto!codforpa
            vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSmto!codforpa, "N")
            
            vFactu.DtoGnral = 0
            vFactu.DtoPPago = 0
            vFactu.Banco = DBLet(vClien.Banco, "N")
            vFactu.Sucursal = DBLet(vClien.Sucursal, "N")
            vFactu.DigControl = DBLet(vClien.DigControl, "T")
            vFactu.CuentaBan = DBLet(vClien.CuentaBan, "T")
            
            vFactu.Observacion = DBLet(RSmto!concefac, "T")
                
            
            
            
            If Not vFactu.PasarMtosAFactura(TipCoMan, OpeFactu, MesFactu, RSmto!nummante, CentroCoste) Then
                If b Then b = False
            Else
                vClien.ActualizaUltFecMovim (FechaFact)
                
                
                'a�adirlo a la lista de facturas a imprimir
                If ListFactu = "" Then
                    ListFactu = vFactu.NumFactu
                Else
                    ListFactu = ListFactu & "," & vFactu.NumFactu
                End If
            End If
        End If
        Set vClien = Nothing
        RSmto.MoveNext
    Wend
    
    RSmto.Close
    Set RSmto = Nothing
    
    Set vFactu = Nothing
    Lbl.Caption = "Finalizando proceso"
    Lbl.Refresh
    If b Then
        MsgBox "Las Facturas de los Mantenimientos seleccionados se generaron correctamente.", vbInformation
    Else
        SQL = "ATENCI�N:" & vbCrLf
        MsgBox SQL & "No todas las Facturas se generaron correctamente!!!.", vbInformation
    End If
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("VENFAC")
    TerminaBloquear
    
    If ListFactu <> "" Then
        Lbl.Caption = "Imprimiend"
        Lbl.Refresh
        ImprimirFacturaMan 53, ListFactu, FechaFact
    End If
    
    
ETraspasoMtoFac:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando Mantenimientos", Err.Description
    End If
End Function




Private Sub ImprimirFacturaMan(OpcionListado As Byte, ListFactu As String, FecFactu As String)
'Imprime una factura de Mantenimiento
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NombreTabla As String
    
    NombreTabla = "scafac"
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    If (OpcionListado = 53) Then indRPT = 12 'Facturas Clientes
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt) Then
        Exit Sub
    End If
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    frmImprimir.NombrePDF = pPdfRpt
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion N� de Factura
    '---------------------------------------------------
    'Cod Tipo Movimiento
    devuelve = "{" & NombreTabla & ".codtipom}='FAM'"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    cadSelect = cadFormula
    
    'N� Factura
    devuelve = "{" & NombreTabla & ".numfactu} IN [" & ListFactu & "]"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "{" & NombreTabla & ".numfactu} IN (" & ListFactu & ")"
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
    'Fecha Factura
    devuelve = "year({" & NombreTabla & ".fecfactu})=" & Year(FecFactu)
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    'Fecha Factura en cadSelect
'        devuelve = "{" & NombreTabla & ".fecfactu}= '" & Format(FecFactu, FormatoFecha) & "'"
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
   
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = OpcionListado
            .Titulo = ""
            .Show vbModal
    End With
End Sub






'Ventas de TICKET
'=================================================================
Public Function EliminarVenta(cadSQL As String) As Boolean
'Eliminamos de las tablas de ventas: scaven, sliven
Dim SQL As String

    On Error GoTo EElimVen

    EliminarVenta = False
    
    
    'ELiminar lineas venta
    SQL = "DELETE FROM sliven "
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaven", "sliven")
    conn.Execute SQL
    
'    Espera 0.1
    
    'Eliminar Cabeceras venta
    SQL = "DELETE FROM scaven "
    SQL = SQL & " WHERE " & Replace(cadSQL, "sliven", "scaven")
    conn.Execute SQL
        
    EliminarVenta = True

EElimVen:
    If Err.Number <> 0 Then
        MsgBox Err.Number, "Eliminar venta.", Err.Description
        EliminarVenta = False
    Else
        EliminarVenta = True
    End If
End Function




Private Function DevuelveTipoDocumentoFactura(ByRef TipoAlbaran As String) As Byte
    DevuelveTipoDocumentoFactura = 0
    If TipoAlbaran <> "" Then
        If TipoAlbaran = "ATI" Then
            'Factura de tickets
            TipoAlbaran = 1
            DevuelveTipoDocumentoFactura = 1
        Else
            If TipoAlbaran = "ALZ" Then
                TipoAlbaran = 2
                DevuelveTipoDocumentoFactura = 2
            Else
                If TipoAlbaran = "ALS" Or TipoAlbaran = "ARC" Then
                    TipoAlbaran = 3
                    DevuelveTipoDocumentoFactura = 3
                Else
                    If TipoAlbaran = "ARN" Then
                        TipoAlbaran = 4
                        DevuelveTipoDocumentoFactura = 4
                    Else
                        If TipoAlbaran = "ARP" Then
                            TipoAlbaran = 5
                            DevuelveTipoDocumentoFactura = 5
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Function




Public Function TraspasoAlbaranesFacturasCli(cadSQL As String, cadWHERE As String, FechaFact As String, banPr As String, ByRef PBar1 As ProgressBar, ByRef LblBar As Label, ImprimeLasFacturasGeneradas As Boolean, ByRef vTipom As String, TextosCSB As String, Socio As Boolean) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'      FechaFact: Fecha de la Factura
'      BanPr: Cod. de Banco Propio
'      Pbar1:  Una progressbar. Se puede mandar un NOTHING, y no pasa nada. Si no se manda
'              es que estamos en un proceso corto o que no necesitabaos un pb1, con lo cual NO muestro el PB1
'      Imprime: Si despues de generarlo los imprime
'
'       vTipom:  Que tipo de albaran es, para luego la impresion saber que factura imprime
'      TextosCSB:  Si lleva llevara 3 lineas para meter ent tesoreria

'Desde Albaranes Genera las Facturas correspondientes
Dim RSalb As ADODB.Recordset 'Ordenados por: tipofac,clien,dpto,forma pago, dtoppago, dtognral
Dim b As Boolean
Dim SQL As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim antClien As Long
Dim antDirec As Long
Dim antForpa As Integer
Dim antDtoPP As Single, antDtoGn As Single

'direc/dpto actual para controlar el valor nulo
Dim actDirec As Long

'Concatenamos todas las facturas generadas para listarlas en el informe
Dim ListFactu As String
Dim vFactu As CFacturaCli
Dim Inc As Integer
Dim condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura

'Por si no mando una progressbar, que no de errores
Dim PgbVisible As Boolean

    On Error GoTo ETraspasoAlbFac

    TraspasoAlbaranesFacturasCli = False

    ListFactu = ""
        
    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("VENFAC") 'facturas de venta
    If Not BloqueoManual("VENFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los albaranes que vamos a facturar (cabeceras y lineas)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    SQL = " (scaalbcli INNER JOIN scliente ON scaalbcli.codclien=scliente.codclien ) INNER JOIN slialbcli ON scaalbcli.codtipom=slialbcli.codtipom AND scaalbcli.numalbar=slialbcli.numalbar "
    If Not BloqueaRegistro(SQL, cadWHERE) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        Exit Function
    End If
    
   
    'Inicializar la Progress Bar
    PgbVisible = False
    If Not (PBar1 Is Nothing) Then
        If PBar1.visible Then PgbVisible = True
    End If
    If PgbVisible Then
        If InStr(1, cadSQL, "scliente") Then
            SQL = Replace(cadSQL, "scaalbcli.*, scliente.periodof", "count(*)") 'si hay INNER JOIN con sclien
        Else
            SQL = Replace(cadSQL, "*", "count(*)") 'si NO hay INNER JOIN con sclien
        End If
        
        
        Set RSalb = New ADODB.Recordset
        RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RSalb.EOF Then
            CargarProgresNew PBar1, CInt(RSalb.Fields(0))
            LblBar.Caption = "Inicializando el proceso..."
            LblBar.Refresh
            
        End If
        RSalb.Close
        Set RSalb = Nothing
    End If
    
        
    'EMPEZAMOS LA FACTURA
    Set vFactu = New CFacturaCli
    vFactu.FecFactu = FechaFact 'Fecha para las Facturas

    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = banPr
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", banPr, "N")
    
    'comprobar que la cuenta prevista de cobro tiene valor
    b = (vFactu.CuentaPrev <> "")
    If Not b Then
        Set vFactu = Nothing
        'Desbloqueamos ya no estamos facturando
        DesBloqueoManual ("VENFAC")
        TerminaBloquear
        MsgBox "La cta. prevista de cobro debe tener valor.", vbExclamation
        Exit Function
    End If
    
       
        
    'Marcar Albaranes que se van a Facturar
    '----------------------------------------
    SQL = cadSQL & " ORDER BY scaalbcli.tipofact, scaalbcli.codclien, scaalbcli.coddirec, codforpa, dtoppago, dtognral "
    Set RSalb = New ADODB.Recordset
    RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    'Agrupar los Albaranes posibles en una misma Factura
    'Calcular y Grabar Factura en la Tabla de Facturas
    'Albaran(scaalb, slialb) -> Factura (scafac,scafac1,slifac)
    '----------------------------------------------------
    'Se factura por cliente y departamento
    'Agrupar albaranes en 1 factura por : tipofact,codclien,coddirec,codforpa,dtoppago, dtognral
    antClien = 0 'cliente
    antDirec = 0 'direccion/departamento
    antForpa = 0 'forma de pago
    antDtoPP = 0 'dto pronto pago
    antDtoGn = 0 'dto general
    
    cadW = ""
    Errores = ""
    Inc = 0
    
    While Not RSalb.EOF
        TipoAlb = RSalb!codtipom
        Inc = Inc + 1
        If IsNull(RSalb!CodDirec) Then
            actDirec = -1
        Else
            actDirec = DBLet(RSalb!CodDirec, "N")
        End If
        
        If RSalb!TipoFact = 1 Then 'tipofact=1 "FACTURA x ALBARAN"
        '---------------------------------------------------------
            'frmListadoPed.lblProgess(0).Caption = "Facturando: Facturas individuales"
            LblBar.Caption = "Facturando: Facturas individuales"
            LblBar.Refresh
            If cadW <> "" Then 'Facturacion pendiente
                cadW = cadW & ")) "
                If Not vFactu.PasarAlbaranesAFactura2(False, TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
                    If b Then b = False
                    AnyadirAvisos ErroresAux
                Else 'a�adirlo a la lista de facturas a imprimir
                    If ListFactu = "" Then
                        ListFactu = vFactu.NumFactu
                    Else
                        ListFactu = ListFactu & "," & vFactu.NumFactu
                    End If
                End If
                If PgbVisible Then
                    IncrementarProgresNew PBar1, Inc - 1
                    LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien
                    LblBar.Refresh
                End If
                Espera 0.2
                'Empezamos una nueva Factura
                cadW = ""
            End If
            
            'Los Albaranes que tengan tipofact=1 "factura x Albaran" generar una factura
            'para cada uno de ellos
            cadW = " scaalbcli.codtipom='" & RSalb!codtipom & "' AND scaalbcli.numalbar=" & RSalb!NumAlbar
            
            'Generar una Factura nueva
            vFactu.Cliente = RSalb!CodClien
            vFactu.NombreClien = RSalb!nomclien
            vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
            vFactu.CPostal = DBLet(RSalb!codpobla, "T")
            vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
            vFactu.Provincia = DBLet(RSalb!proclien, "T")
            vFactu.NIF = DBLet(RSalb!nifClien, "T")
            vFactu.Telefono = DBLet(RSalb!telclien, "T")
            vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
            vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
            vFactu.Agente = RSalb!codagent
            vFactu.ForPago = RSalb!codforpa
            vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
            vFactu.DtoPPago = CCur(RSalb!DtoPPago)
            vFactu.DtoGnral = CCur(RSalb!DtoGnral)

                
                
            If Not vFactu.PasarAlbaranesAFactura2(Socio, TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
                If b Then b = False
                AnyadirAvisos ErroresAux
            Else 'a�adirlo a la lista de facturas a imprimir
                If ListFactu = "" Then
                    ListFactu = vFactu.NumFactu
                Else
                    ListFactu = ListFactu & "," & vFactu.NumFactu
                End If
            End If
            If PgbVisible Then
                Inc = 1 '1 albaran x factura
                LblBar.Caption = "Cliente: " & Format(RSalb!CodClien, "000000") & " - " & RSalb!nomclien
                LblBar.Refresh
                IncrementarProgresNew PBar1, Inc
                Inc = 0
            End If
            Espera 0.2
                
            cadW = ""
            
        Else 'tipofac=0 "factura COLECTIVA"
        '----------------------------------------------------------
            'Seleccionar todos los Albaranes pertenecientes a un mismo Cliente,Departamento
            'Los que tengan tipofac=0 "factura colectiva" agruparlos en una misma factura
            'para la misma Forma de PAgo, mismo dtoppago y mismo dtognral
             
             '-- David.      Esta linea da error si no viene de frmlistadoped
             'frmListadoPed.lblProgess(0).Caption = "Facturando: Facturas colectivas"
             LblBar.Caption = "Facturando: Facturas colectivas"
             LblBar.Refresh
             '---- Laura: 06/10/2006
             'Comprobar si es Departamento o Direccion (segun paramatro)
             If vParamAplic.Departamento Then
                'agrupar tb por departamento
                condicion = (antClien <> RSalb!CodClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral)
             Else
                condicion = (antClien <> RSalb!CodClien) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral)
             End If
             
'             If (antClien <> RSalb!CodClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral) Then
             If condicion Then
             '-----
                If cadW <> "" Then 'Facturacion PEndiente
                    cadW = cadW & ")) "
                    If Not vFactu.PasarAlbaranesAFactura2(Socio, TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
                        If b Then b = False
                        AnyadirAvisos ErroresAux
                    Else 'a�adirlo a la lista de facturas a imprimir
                        If ListFactu = "" Then
                            ListFactu = vFactu.NumFactu
                        Else
                            ListFactu = ListFactu & "," & vFactu.NumFactu
                        End If
                    End If
                    If PgbVisible Then
                        LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien
                        LblBar.Refresh
                        IncrementarProgresNew PBar1, Inc
                        Inc = 0
                    End If
                    Espera 0.2
                    
                    'Empezamos una nueva Factura
                    cadW = ""
                End If
                'Generar una Factura nueva
                vFactu.Cliente = RSalb!CodClien
                vFactu.NombreClien = RSalb!nomclien
                vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
                vFactu.CPostal = DBLet(RSalb!codpobla, "T")
                vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
                vFactu.Provincia = DBLet(RSalb!proclien, "T")
                vFactu.NIF = DBLet(RSalb!nifClien, "T")
                vFactu.Telefono = DBLet(RSalb!telclien, "T")
                vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
                vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
                vFactu.Agente = RSalb!codagent
                vFactu.ForPago = RSalb!codforpa
                vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
                vFactu.DtoPPago = CCur(RSalb!DtoPPago)
                vFactu.DtoGnral = CCur(RSalb!DtoGnral)
                vFactu.Aportacion = 0
                
                '[Monica]07/05/2015: faltan los datos del banco
                Dim Sqldatos As String
                Dim RsDatos As ADODB.Recordset
                Sqldatos = "select codbanco, codsucur, digcontr, cuentaba, iban from scliente where codclien = " & DBSet(RSalb!CodClien, "N")
                
                Set RsDatos = New ADODB.Recordset
                RsDatos.Open Sqldatos, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RsDatos.EOF Then
                    vFactu.Banco = Format(DBLet(RsDatos!codbanco, "N"), "0000")
                    vFactu.Sucursal = Format(DBLet(RsDatos!codsucur, "N"), "0000")
                    vFactu.DigControl = DBLet(RsDatos!digcontr, "T")
                    vFactu.CuentaBan = DBLet(RsDatos!cuentaba, "T")
                    vFactu.Iban = DBLet(RsDatos!Iban, "T")
                End If
                
                
                If RSalb!codtipom = "ALM" Then vFactu.Aportacion = DBLet(RSalb!Aportacion, "N")
                cadW = " (scaalbcli.codtipom='" & RSalb!codtipom & "' AND scaalbcli.numalbar IN (" & RSalb!NumAlbar
            Else
                cadW = cadW & ", " & RSalb!NumAlbar
            End If
        
            'Guardamos datos del registro anterior
            antClien = RSalb!CodClien
'            antDirec = DBLet(RSalb!CodDirec, "N")
            antDirec = actDirec
            antForpa = RSalb!codforpa
            antDtoPP = RSalb!DtoPPago
            antDtoGn = RSalb!DtoGnral
        End If
        RSalb.MoveNext
    Wend
    RSalb.Close
    Set RSalb = Nothing
        
    'Facturar la ultima Factura generada del blucle
    If cadW <> "" Then
        cadW = cadW & "))"
        If PgbVisible Then LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
        
        If Not vFactu.PasarAlbaranesAFactura2(Socio, TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
            If b Then b = False
            AnyadirAvisos "Error Facturando el Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien & vbCrLf & ErroresAux
        Else 'a�adirlo a la lista de facturas a imprimir
            If ListFactu = "" Then
                ListFactu = vFactu.NumFactu
            Else
                ListFactu = ListFactu & "," & vFactu.NumFactu
            End If
        End If
        If PgbVisible Then
'            LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
            IncrementarProgresNew PBar1, Inc
        End If
        Espera 0.2
    End If
    
    TipoFac = vFactu.codtipom
    Set vFactu = Nothing
    TraspasoAlbaranesFacturasCli = True
    
    If b Then
        LblBar.Caption = "Proceso finalizado correctamente."
        MsgBox "Las Facturas de los Albaranes seleccionados se generaron correctamente.", vbInformation
    Else
        LblBar.Caption = "Proceso finalizado con errores."
        SQL = "ATENCI�N:" & vbCrLf
        MsgBox SQL & "No todas las Facturas se generaron correctamente!!!.", vbExclamation
        If Errores <> "" Then MostrarAvisos
    End If
    
    Espera 0.2
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("VENFAC")
    TerminaBloquear
    
    
    If ImprimeLasFacturasGeneradas Then
        If ListFactu <> "" Then
            
            ImprimirFacturasClien ListFactu, FechaFact, , DevuelveTipoDocumentoFactura(vTipom)
        End If
    End If
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando Albaranes", Err.Description
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        TerminaBloquear
    End If
End Function



