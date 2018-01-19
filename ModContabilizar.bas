Attribute VB_Name = "ModContabilizar"
Option Explicit


'===================================================================================
'CONTABILIZAR FACTURAS:
'Modulo para el traspaso de registros de cabecera y lineas de tablas de FACTURACION
'A las tablas de FACTURACION de Contabilidad
'====================================================================================

Private DtoGnral As Currency
Private DtoPPago As Currency
Private BaseImp As Currency
Private IvaImp As Currency

Private TotalFac As Currency
Private CCoste2 As String

Private vCCos As Byte
    'Para cuando pasamos en la contabilizacion de las facturas
    'Sera 2:    tiene mas de un centro de coste. Habra que agrupar por CC
    '     1:  o solo es un trabajador o tienen el mismo CC, con lo cual no hace falta agrupar por CC
    '     0:  no habra CC.  Si vpara.. tieneanalitica = false

Private conCtaAlt As Boolean 'el cliente utiliza cuentas alternativas

'Para pasar a contabilidad facturas de proveedor
Private AnyoFacPr As Integer 'a�o factura proveedor, es el ano de fecha_recepcion

'Modificacion Centro de coste.
'La factura cogera el Centro de coste del trabajador del albaran


Private vTipoIva(2) As Currency
Private vPorcIva(2) As Currency
Private vPorcRec(2) As Currency
Private vBaseIva(2) As Currency
Private vImpIva(2) As Currency
Private vImpRec(2) As Currency

'llevara: codmacta_proveedor | impo_retencion |
Private DatosRetencion As String
Private DatosAportacion As String
Private FechaRecepcion As String


Public Function CrearTMPFacturas(cadTabla As String, cadWHERE As String) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    SQL = "CREATE TEMPORARY TABLE tmpFactu ( "
    If cadTabla = "scafaccli" Or cadTabla = "scafac" Then
        SQL = SQL & "codtipom char(3) NOT NULL default '',"
        SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        If cadTabla = "sfactusoc" Then
            SQL = SQL & "codtipom char(3) NOT NULL default '',"
            SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
            SQL = SQL & "codsocio mediumint(7) unsigned NOT NULL default '0',"
        Else
            SQL = SQL & "codprove int(6) unsigned NOT NULL default '0',"
            SQL = SQL & "numfactu varchar(10)  NOT NULL ,"
        End If
    End If
    SQL = SQL & "fecfactu date NOT NULL default '0000-00-00') "
    conn.Execute SQL
     
     
    If cadTabla = "scafaccli" Or cadTabla = "scafac" Then
        SQL = "SELECT codtipom, numfactu, fecfactu"
    Else
        If cadTabla = "sfactusoc" Then
            SQL = "SELECT codtipom, numfactu, codsocio, fecfactu "
        Else
            SQL = "SELECT codprove, numfactu, fecfactu"
        End If
    End If
    SQL = SQL & " FROM " & cadTabla
    SQL = SQL & " WHERE " & cadWHERE
    
    'DAVID###
    'Si son de proveedores el orden es MUY importante para
    'que vayan ordenaditas por fecha recepcion
    'ademas, por si tiene mas de una por prove a�ado los dos campos
    If cadTabla = "sprove" Then SQL = SQL & " ORDER BY fecrecep,codprove,numfactu"

    
    SQL = " INSERT INTO tmpFactu " & SQL
    conn.Execute SQL

    CrearTMPFacturas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        MuestraError Err.Number, "", Err.Description
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpFactu;"
        conn.Execute SQL
    End If
End Function


Public Sub BorrarTMPFacturas()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpFactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub InsertarTMPErrFac(MenError As String, cadWHERE As String)
Dim SQL As String

    On Error Resume Next
    SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    SQL = SQL & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    SQL = SQL & " WHERE " & Replace(cadWHERE, "scafpc", "tmpFactu")
    conn.Execute SQL
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function CrearTMPErrFact(cadTabla As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    SQL = "CREATE TEMPORARY TABLE tmpErrFac ( "
    If cadTabla = "scafac" Or cadTabla = "sfactusoc" Then
        SQL = SQL & "codtipom char(3) NOT NULL default '',"
        SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        SQL = SQL & "codprove int(6) unsigned NOT NULL default '0',"
        SQL = SQL & "numfactu varchar(10) NOT NULL ,"
    End If
    SQL = SQL & "fecfactu date NOT NULL default '0000-00-00', "
    SQL = SQL & "error varchar(200) NULL )"
    conn.Execute SQL
     
     CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpErrFac;"
        conn.Execute SQL
    End If
End Function


Public Sub BorrarTMPErrFact()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpErrFac;"
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function ComprobarLetraSerie(cadTabla As String) As Boolean
'Para Facturas VENTA a clientes
'Comprueba que la letra del serie del tipo de movimiento es  correcta
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetra

    ComprobarLetraSerie = False
    
    'Comprobar que existe la letra de serie en contabilidad
    If cadTabla = "scafaccli" Or cadTabla = "sfactusoc" Or cadTabla = "scafac" Then
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
        SQL = "Select distinct tiporegi from contadores"
        Set RSconta = New ADODB.Recordset
        RSconta.Open SQL, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
        If RSconta.EOF Then
            RSconta.Close
            Set RSconta = Nothing
            Exit Function
        End If
            
    
        'obtenemos los distintos tipos de movimiento que vamos a contabilizar
        'de las facturas seleccionadas
        SQL = "select distinct " & cadTabla & ".codtipom from " & cadTabla
        SQL = SQL & " INNER JOIN tmpFactu ON " & cadTabla & ".codtipom=tmpFactu.codtipom AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
'        SQL = SQL & cadWHERE
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        b = True
        While Not Rs.EOF And b
            'comprobar que todas las letras serie existen en Aritaxi
            SQL = "letraser"
            devuelve = DevuelveDesdeBDNew(conAri, "stipom", "codtipom", "codtipom", Rs!codtipom, "T", SQL)
            If devuelve = "" Then
                b = False
                Cad = Rs!codtipom & " en BD de Gesti�n."
            ElseIf SQL <> "" Then
                'comprobar que todas las letras serie existen en la contabilidad
                devuelve = "tiporegi= " & DBSet(SQL, "T")
                RSconta.MoveFirst
                RSconta.Find (devuelve), , adSearchForward
                If RSconta.EOF Then
                    'no encontrado
                    b = False
                    Cad = SQL & " en BD de Contabilidad."
                End If
            End If
            If b Then Cad = Cad & DBSet(Rs!codtipom, "T") & ","
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        RSconta.Close
        Set RSconta = Nothing
        
        If Not b Then 'Hay algun movimiento que no existe
            devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
            devuelve = devuelve & "Consulte con el administrador."
            MsgBox devuelve, vbExclamation
            Exit Function
        End If
        
        'Todos los Tipo de movimiento existen
        If Cad <> "" Then
            Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitamos ult. coma
        
            'miramos si hay algun movimiento de factura que la letra serie sea nulo
            SQL = "select count(*) from stipom "
            SQL = SQL & "where codtipom IN (" & Cad & ") and (isnull(letraser) or letraser='')"
            If RegistrosAListar(SQL) > 0 Then
                SQL = "Hay algun tipo de movimiento de Facturaci�n que no tiene letra serie." & vbCrLf
                SQL = SQL & "Comprobar en la tabla de tipos de movimiento: " & Cad
                MsgBox SQL, vbExclamation
                Exit Function
            End If
        End If
        ComprobarLetraSerie = True
    Else
        ComprobarLetraSerie = True
    End If

EComprobarLetra:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Letra Serie", Err.Description
    End If
End Function

'###### ESTE YA NO SE UTILIZA
'Public Function ComprobarNumFacturas(cadTabla As String, cadWConta) As Boolean
''Comprobar que no exista ya en la contabilidad un n� de factura para la fecha que
''vamos a contabilizar
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
'Dim b As Boolean
'
'    On Error GoTo ECompFactu
'
'    ComprobarNumFacturas = False
'
'    SQL = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
'    SQL = SQL & " WHERE " & cadWConta
'
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Not RSconta.EOF Then
'        'Seleccionamos las distintas facturas que vamos a facturar
'        SQL = "SELECT DISTINCT " & cadTabla & ".codtipom,letraser,scafac.numfactu,scafac.fecfactu "
'        SQL = SQL & " FROM (" & cadTabla & " INNER JOIN stipom ON " & cadTabla & ".codtipom=stipom.codtipom) "
'        SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
''        SQL = SQL & " WHERE " & cadWHERE
'
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        b = True
'        While Not RS.EOF And b
'            SQL = "(numserie= " & DBSet(RS!LetraSer, "T") & " AND codfaccl=" & DBSet(RS!NumFactu, "N") & " AND anofaccl=" & Year(RS!FecFactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
'                b = False
'                SQL = "          N� Fac.: " & Format(RS!NumFactu, "0000000") & vbCrLf
'                SQL = SQL & "          Fecha: " & RS!FecFactu
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'
'        If Not b Then
'            SQL = "Ya existe la factura: " & vbCrLf & SQL
'            SQL = "Comprobando N� Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
'
'            MsgBox SQL, vbExclamation
'            ComprobarNumFacturas = False
'        Else
'            ComprobarNumFacturas = True
'        End If
'    Else
'        ComprobarNumFacturas = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
'
'ECompFactu:
'     If Err.Number <> 0 Then
'        MuestraError Err.Number, "Comprobar N� Facturas", Err.Description
'    End If
'End Function


Public Function ComprobarNumFacturas_new(cadTabla As String, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un n� de factura para la fecha que
'vamos a contabilizar
Dim SQL As String
Dim SQLconta As String
Dim Rs As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas_new = False
    
'    SQLconta = "SELECT numserie,codfaccl,anofaccl FROM cabfact "

     If vParamAplic.ContabilidadNueva Then
        SQLconta = "SELECT count(*) FROM factcli WHERE "
     Else
        SQLconta = "SELECT count(*) FROM cabfact WHERE "
     End If
'    SQLconta = SQLconta & " WHERE (" & cadWConta & ") "

    
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText

'    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        SQL = "SELECT DISTINCT " & cadTabla & ".codtipom,letraser," & cadTabla & ".numfactu," & cadTabla & ".fecfactu "
        SQL = SQL & " FROM (" & cadTabla & " INNER JOIN stipom ON " & cadTabla & ".codtipom=stipom.codtipom) "
        SQL = SQL & " INNER JOIN tmpFactu ON " & cadTabla & ".codtipom=tmpFactu.codtipom AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "

        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF And b
            If vParamAplic.ContabilidadNueva Then
                SQL = "(numserie= " & DBSet(Rs!LetraSer, "T") & " AND numfactu=" & DBSet(Rs!NumFactu, "N") & " AND anofactu=" & Year(Rs!FecFactu) & ")"
            Else
                SQL = "(numserie= " & DBSet(Rs!LetraSer, "T") & " AND codfaccl=" & DBSet(Rs!NumFactu, "N") & " AND anofaccl=" & Year(Rs!FecFactu) & ")"
            End If
'            If SituarRSetMULTI(RSconta, SQL) Then
            SQL = SQLconta & SQL
            If RegistrosAListar(SQL, conConta) Then
                b = False
                SQL = "          Letra Serie: " & DBSet(Rs!LetraSer, "T") & vbCrLf
                SQL = SQL & "          N� Fac.: " & Format(Rs!NumFactu, "0000000") & vbCrLf
                SQL = SQL & "          Fecha: " & Format(Rs!FecFactu, "dd/mm/yyyy")
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            SQL = "Ya existe la factura: " & vbCrLf & SQL
            SQL = "Comprobando N� Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
            
            MsgBox SQL, vbExclamation
            ComprobarNumFacturas_new = False
        Else
            ComprobarNumFacturas_new = True
        End If
'    Else
'        ComprobarNumFacturas_new = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
    Exit Function
    
ECompFactu:
     If Err.Number <> 0 Then
        ComprobarNumFacturas_new = False
        MuestraError Err.Number, "Comprobar N� Facturas", Err.Description
    End If
End Function




'###### ESTE YA NO SE UTILIZA
'Public Function ComprobarCtaContable(cadTabla As String, Opcion As Byte) As Boolean
''Comprobar que todas las ctas contables de los distintos clientes de las facturas
''que vamos a contabilizar existan en la contabilidad
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
'Dim b As Boolean
'Dim cadG As String
'
'    On Error GoTo ECompCta
'
'    ComprobarCtaContable = False
'
'    If Opcion = 3 Then 'si hay analitica comprobar que todas las cuentas
'                        'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
'        cadG = "grupovta"
'        SQL = DevuelveDesdeBDNew(conConta, "parametros", "grupogto", "", "", "", cadG)
'        If SQL <> "" And cadG <> "" Then
'            SQL = " AND (codmacta like '" & SQL & "%' OR codmacta like '" & cadG & "%')"
'        ElseIf SQL <> "" Then
'            SQL = " AND (codmacta like '" & SQL & "%')"
'        ElseIf cadG <> "" Then
'            SQL = " AND (codmacta like '" & cadG & "%')"
'        End If
'        cadG = SQL
'    End If
'
'    SQL = "SELECT codmacta FROM cuentas "
'    SQL = SQL & " WHERE apudirec='S'"
'    If cadG <> "" Then SQL = SQL & cadG
'
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText
'
'    If Not RSconta.EOF Then
'        If Opcion = 1 Then
'            If cadTabla = "scafac" Then
'                'Seleccionamos los distintos clientes,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafac.codclien, sclien.codmacta "
'                SQL = SQL & " FROM (scafac INNER JOIN sclien ON scafac.codclien=sclien.codclien) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
'            Else
'                'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafpc.codprove, sprove.codmacta "
'                SQL = SQL & " FROM (scafpc INNER JOIN sprove ON scafpc.codprove=sprove.codprove) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
'            End If
'
'        ElseIf Opcion = 2 Or Opcion = 3 Then
'            SQL = "SELECT distinct "
'            If Opcion = 2 Then SQL = SQL & " sartic.codfamia,"
'            If cadTabla = "scafac" Then
'                SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((slifac "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifac.codtipom=tmpFactu.codtipom AND slifac.numfactu=tmpFactu.numfactu AND slifac.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifac.codartic=sartic.codartic) "
'            Else
'                SQL = SQL & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
'            End If
'            SQL = SQL & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
'        End If
'
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        b = True
'        While Not RS.EOF And b
'            SQL = "codmacta= " & DBSet(RS!Codmacta, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
'                b = False 'no encontrado
'                If Opcion = 1 Then
'                    If cadTabla = "scafac" Then
'                        SQL = RS!Codmacta & " del Cliente " & Format(RS!CodClien, "000000")
'                    Else
'                        SQL = RS!Codmacta & " del Proveedor " & Format(RS!codProve, "000000")
'                    End If
'                ElseIf Opcion = 2 Then
'                    SQL = RS!Codmacta & " de la familia " & Format(RS!codfamia, "0000")
'                ElseIf Opcion = 3 Then
'                    SQL = RS!Codmacta
'                End If
'            End If
'
'            If Opcion = 2 Then
'                'Comprobar que ademas de existir la cuenta de ventas exista tambien
'                'la cuenta ABONO ventas
'                SQL = "codmacta= " & DBSet(RS!ctaabono, "T")
'                RSconta.MoveFirst
'                RSconta.Find (SQL), , adSearchForward
'                If RSconta.EOF Then
'                    b = False 'no encontrado
'                    SQL = RS!ctaabono & " de la familia " & Format(RS!codfamia, "0000")
'                End If
'            End If
'
'            'comprobar cuentas alternativas solo para facturacion a clientes
'            If cadTabla = "scafac" Then
'                If Opcion = 2 Then
'                    ' Comprobar cuenta venta alternativa
'                    If DBLet(RS!ctavent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!ctavent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctavent1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta venta alternativa."
'                    End If
'                End If
'                If Opcion = 2 Then
'                    ' Comprobar cuenta de abono alternativa
'                    If DBLet(RS!abovent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!abovent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctaabon1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta abono alternativa."
'                    End If
'                End If
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'
'        If Not b Then
'            If Opcion <> 3 Then
'                SQL = "No existe la cta contable " & SQL
'            Else
'                SQL = "La cuenta " & SQL & " no es del nivel correcto."
'            End If
'            SQL = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & SQL
'
'            MsgBox SQL, vbExclamation
'            ComprobarCtaContable = False
'        Else
'            ComprobarCtaContable = True
'        End If
'    Else
'        ComprobarCtaContable = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
'
'ECompCta:
'     If Err.Number <> 0 Then
'        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
'    End If
'End Function






Public Function ComprobarCtaContable_new(cadTabla As String, Opcion As Byte, Optional Equipamiento As Boolean) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad

'NEUVO MARZO 2009
'COmprobaremos que no esten bloqueadas
Dim cContabF As CControlFacturaContab
Dim QueCuentasSon As String
Dim CtaBloq As Collection

Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Ic As Integer
'Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim SQLcuentas As String
    
    On Error GoTo ECompCta

    ComprobarCtaContable_new = False
    
    cadG = ""
    
    
    If Opcion = 3 Then
            'si hay analitica comprobar que todas las cuentas
            'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
            cadG = "grupovta"
            SQL = DevuelveDesdeBDNew(conConta, "parametros", "grupogto", "", "", "", cadG)
            If SQL <> "" And cadG <> "" Then
                SQL = " AND (codmacta like '" & SQL & "%' OR codmacta like '" & cadG & "%')"
            ElseIf SQL <> "" Then
                SQL = " AND (codmacta like '" & SQL & "%')"
            ElseIf cadG <> "" Then
                SQL = " AND (codmacta like '" & cadG & "%')"
            End If
            cadG = SQL
    End If
    
    
    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG
    
    If Opcion = 1 Then
        If cadTabla = "scafaccli" Then
            'Seleccionamos los distintos clientes,cuentas que vamos a facturar
            
            SQL = "SELECT DISTINCT scafaccli.codclien, scliente.codmacta "
            SQL = SQL & " FROM (scafaccli INNER JOIN scliente ON scafaccli.codclien=scliente.codclien) "
            SQL = SQL & " INNER JOIN tmpFactu ON scafaccli.codtipom=tmpFactu.codtipom AND scafaccli.numfactu=tmpFactu.numfactu AND scafaccli.fecfactu=tmpFactu.fecfactu "
        Else
            If cadTabla = "scafac" Then
                Dim CADENA1 As String
                Dim LCad As Integer
                
                CADENA1 = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
                LCad = Len(CADENA1)
            
                If Equipamiento Then
                    SQL = "SELECT DISTINCT scafac.codclien, concat('" & vParamAplic.Raiz_Cta_Soc_Equip & "',right(concat('" & CADENA1 & "',scafac.codclien)," & LCad & ")) codmacta "
                Else
                    ' cuotas
                    SQL = "SELECT DISTINCT scafac.codclien, concat('" & vParamAplic.Raiz_CtaClien_Soc & "',right(concat('" & CADENA1 & "',scafac.codclien)," & LCad & ")) codmacta "
                End If
                SQL = SQL & " FROM (scafac INNER JOIN sclien ON scafac.codclien=sclien.codclien) "
                SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
            Else
                'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
                SQL = "SELECT DISTINCT scafpc.codprove, sprove.codmacta "
                SQL = SQL & " FROM (scafpc INNER JOIN sprove ON scafpc.codprove=sprove.codprove) "
                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
            End If
        End If
    
    ElseIf Opcion = 2 Or Opcion = 3 Then
        SQL = "SELECT distinct "
        If Opcion = 2 Then SQL = SQL & " sartic.codfamia,"
        If cadTabla = "scafaccli" Then
            SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((slifaccli "
            SQL = SQL & " INNER JOIN tmpFactu ON slifaccli.codtipom=tmpFactu.codtipom AND slifaccli.numfactu=tmpFactu.numfactu AND slifaccli.fecfactu=tmpFactu.fecfactu) "
            SQL = SQL & "INNER JOIN sartic ON slifaccli.codartic=sartic.codartic) "
        Else
            If cadTabla = "scafac" Then
                SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((slifac "
                SQL = SQL & " INNER JOIN tmpFactu ON slifac.codtipom=tmpFactu.codtipom AND slifac.numfactu=tmpFactu.numfactu AND slifac.fecfactu=tmpFactu.fecfactu) "
                SQL = SQL & "INNER JOIN sartic ON slifac.codartic=sartic.codartic) "
            Else
                SQL = SQL & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
                SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
                SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
            End If
        End If
        SQL = SQL & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
        
        
    ElseIf Opcion = 5 Then ' comprobamos la cta de ventas de la familia del articulo de telefonia
        SQL = "SELECT sfamia.ctacompr as codmacta from sfamia inner join sartic on sfamia.codfamia = sartic.codfamia and sartic.codartic = " & DBSet(vParamAplic.CodarticTfnia, "T")
    
    ElseIf Opcion = 4 Then
        'opcion para la contabilizacion de tickets AGRUPADA  FTG
        
        
        Set Rs = New ADODB.Recordset
      
        
        cadG = "select sfactik.* from sfactik ,scafac where sfactik.numfacFTG=scafac.numfactu and sfactik.fecfacftg=scafac.fecfactu AND  codtipom='FTG' "
        Rs.Open cadG, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        cadG = ""
        Do
            cadG = cadG & "," & Rs!NumFactu
            Rs.MoveNext
        Loop Until Rs.EOF
        Rs.Close
        Set Rs = Nothing
        cadG = Mid(cadG, 2)
        'Monto el SELECT , igual que el de arriba, pero partiendo de los FTIs
         SQL = "SELECT distinct  sartic.codfamia, sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1"
         SQL = SQL & " from (slifac   INNER JOIN sartic ON slifac.codartic=sartic.codartic)  LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia"
         SQL = SQL & " WHERE  codtipom='FTI' and numfactu IN (" & cadG & ")"
         cadG = ""
         'Fuerzo para que haga las mismas comprobaciones que si fuera la opcion 2
         Opcion = 2
         
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    b = True
    QueCuentasSon = ""

    While Not Rs.EOF And b
        SQL = SQLcuentas & " AND codmacta= " & DBSet(Rs!codmacta, "T")
        
        'Para comporbar si estan bloqueadas
        QueCuentasSon = QueCuentasSon & ", '" & Rs!codmacta & "'"
        
        
        If Not (RegistrosAListar(SQL, conConta) > 0) Then
        'si no lo encuentra
            b = False 'no encontrado
            If Opcion = 1 Then
                If cadTabla = "scafaccli" Or cadTabla = "scafac" Then
                    SQL = Rs!codmacta & " del Cliente " & Format(Rs!CodClien, "000000")
                Else
                    SQL = Rs!codmacta & " del Proveedor " & Format(Rs!codProve, "000000")
                End If
            ElseIf Opcion = 2 Or Opcion = 5 Then
                SQL = Rs!codmacta & " de la familia " & Format(Rs!Codfamia, "0000")
            ElseIf Opcion = 3 Then
                SQL = Rs!codmacta
            End If
        End If
        
        
        If Opcion = 2 Or Opcion = 3 Then
            'Comprobar que ademas de existir la cuenta de ventas exista tambien
            'la cuenta ABONO ventas (sfamia.aboventa)
            '---------------------------------------------
            SQL = SQLcuentas & " AND codmacta= " & DBSet(Rs!ctaabono, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
            If Not (RegistrosAListar(SQL, conConta) > 0) Then
                b = False 'no encontrado
                If Opcion = 2 Then
                    SQL = Rs!ctaabono & " de la familia " & Format(Rs!Codfamia, "0000")
                ElseIf Opcion = 3 Then
                    SQL = Rs!ctaabono
                End If
            End If
            
            
            'comprobar cuentas alternativas solo para facturacion a CLIENTES
            '----------------------------------------------------------------
            If cadTabla = "scafaccli" Or cadTabla = "scafac" Then
                ' Comprobar cuenta VENTA alternativa
                If DBLet(Rs!ctavent1, "T") <> "" Then
                    SQL = SQLcuentas & " AND codmacta= " & DBSet(Rs!ctavent1, "T")
'                    RSconta.MoveFirst
'                    RSconta.Find (SQL), , adSearchForward
'                    If RSconta.EOF Then
                    If Not (RegistrosAListar(SQL, conConta) > 0) Then
                        b = False 'no encontrado
                        If Opcion = 2 Then
                            SQL = Rs!ctavent1 & " de la familia " & Format(Rs!Codfamia, "0000")
                        ElseIf Opcion = 3 Then
                            SQL = Rs!ctavent1
                        End If
                    End If
                Else
                    b = False
                    SQL = " o la familia no tiene asignada cuenta venta alternativa."
                End If
                
                ' Comprobar cuenta de ABONO alternativa
                If DBLet(Rs!abovent1, "T") <> "" Then
                    SQL = SQLcuentas & " AND codmacta= " & DBSet(Rs!abovent1, "T")
'                    RSconta.MoveFirst
'                    RSconta.Find (SQL), , adSearchForward
'                    If RSconta.EOF Then
                    If Not (RegistrosAListar(SQL, conConta) > 0) Then
                        b = False 'no encontrado
                        If Opcion = 2 Then
                            SQL = Rs!abovent1 & " de la familia " & Format(Rs!Codfamia, "0000")
                        ElseIf Opcion = 3 Then
                            SQL = Rs!abovent1
                        End If
                    End If
                Else
                    b = False
                    SQL = " o la familia no tiene asignada cuenta abono alternativa."
                End If
            End If
            
        End If
        
        Rs.MoveNext
    Wend
    
    
        
        
        
        If Not b Then
            If Opcion <> 3 Then
                SQL = "No existe la cta contable " & SQL
            Else
                SQL = "La cuenta " & SQL & " no es del nivel correcto. (Familias de art�culos)."
            End If
            SQL = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & SQL
            
            MsgBox SQL, vbExclamation
            ComprobarCtaContable_new = False
        Else
        
            'MARZO 2010
            'Para ver si estanbloqueadas las cuentas
            SQL = ""
            If QueCuentasSon <> "" Then
                QueCuentasSon = Mid(QueCuentasSon, 2)
                Set cContabF = New CControlFacturaContab
                cContabF.CuentasBloqueadas ConnConta, QueCuentasSon, Now, CtaBloq
                If CtaBloq.Count > 0 Then
                    'EXISTEN CUENTAS BLOQUEADAS
                    For Ic = 1 To CtaBloq.Count
                        QueCuentasSon = CtaBloq.Item(Ic)
                        SQL = SQL & RecuperaValor(QueCuentasSon, 1) & "   " & RecuperaValor(QueCuentasSon, 2) & vbCrLf
                    Next
                    SQL = "Cuentas bloqueadas en contabilidad: " & vbCrLf & String(30, "=") & vbCrLf & SQL
                    MsgBox SQL, vbExclamation
                Else
                    SQL = ""
                End If
                Set cContabF = Nothing
            End If
            If SQL = "" Then
                ComprobarCtaContable_new = True
            Else
                ComprobarCtaContable_new = False
            End If
        End If
        
        
        
        
    Exit Function
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function







Public Function ComprobarTiposIVA(cadTabla As String) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    SQL = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For i = 1 To 3
            If cadTabla = "scafaccli" Then
                SQL = "SELECT DISTINCT scafaccli.codigiv" & i
                SQL = SQL & " FROM scafaccli "
                SQL = SQL & " INNER JOIN tmpFactu ON scafaccli.codtipom=tmpFactu.codtipom AND scafaccli.numfactu=tmpFactu.numfactu AND scafaccli.fecfactu=tmpFactu.fecfactu "
                SQL = SQL & " WHERE not isnull(codigiv" & i & ")"
'                SQL = SQL & " WHERE " & " codigiv" & i & " <> 0 "
            ElseIf cadTabla = "scafac" Then
                SQL = "SELECT DISTINCT scafac.codigiv" & i
                SQL = SQL & " FROM scafac "
                SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
                SQL = SQL & " WHERE not isnull(codigiv" & i & ")"
            ElseIf cadTabla = "sfactusoc" Then
                If i = 1 Then
                    SQL = "SELECT DISTINCT sfactusoc.codiiva" & i
                    SQL = SQL & " FROM sfactusoc "
                    SQL = SQL & " INNER JOIN tmpFactu ON sfactusoc.codtipom=tmpFactu.codtipom AND sfactusoc.numfactu=tmpFactu.numfactu AND sfactusoc.fecfactu=tmpFactu.fecfactu "
                    SQL = SQL & " WHERE not isnull(codiiva" & i & ")"
                Else
                    Exit Function
                End If
            Else
                SQL = "SELECT DISTINCT scafpc.tipoiva" & i
                SQL = SQL & " FROM " & cadTabla
                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
                SQL = SQL & " WHERE not isnull(tipoiva" & i & ")"
'                SQL = SQL & " WHERE " & " tipoiva" & i & " <> 0 "
            End If
'            SQL = SQL & " WHERE " & cadWHERE & " AND codigiv" & i & " <> 0 "

            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF And b
                SQL = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (SQL), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    SQL = "Tipo de IVA: " & Rs.Fields(0)
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not b Then
                SQL = "No existe el " & SQL
                SQL = "Comprobando Tipos de IVA en contabilidad..." & vbCrLf & vbCrLf & SQL
            
                MsgBox SQL, vbExclamation
                ComprobarTiposIVA = False
                Exit For
            Else
                ComprobarTiposIVA = True
            End If
        Next i
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function

'La comprobacion del centro de coste ha cambiado
'El centro de coste lo cojera de CADA factura donde tiene
'un trabajador asignado. Luego ya no necesito cadCC
'Comprobaremos:
'           que todas las facturas el trabajador asignado tiene CC
'           y que es distintos, puesto que si es el mismo CC no hare la fiesta

'Noviembe 2009
'
'   Analitica
'    vParamAplic.ModoAnalitica modo analitica: 0=trabajador, 1=Familia, 2=Proyecto
'
'       Es decir, todas las lineas traen el centro de coste asociado, con lo cual,
'   la opcion de comprobar coste ser�:
Public Function ComprobarCCoste(cadSQL As String, Clientes As Boolean) As Byte
Dim SQL As String
Dim i As Integer
Dim C As String
Dim Errores As String
    On Error GoTo ECCoste

    ComprobarCCoste = 0
    Set miRsAux = New ADODB.Recordset
    
    
    'anterior a la introducion de modoanalitica
    '
    'If Clientes Then
    '    SQL = "select codccost from scafac , scafac1, straba "
    '    SQL = SQL & " WHERE scafac.codtipom=scafac1.codtipom and scafac.numfactu=scafac1.numfactu and"
    '    SQL = SQL & " scafac.fecfactu=scafac1.fecfactu  and scafac1.codtraba=straba.codtraba"
    '
    'Else
    '    'PROVEEDORES
    '    SQL = "select codccost from scafpc ,scafpa, straba WHERE"
    '    SQL = SQL & " scafpc.codProve = scafpa.codProve And scafpc.NumFactu = scafpa.NumFactu And"
    '    SQL = SQL & " scafpc.FecFactu = scafpa.FecFactu AND codtrab2=straba.codtraba"
    '
    '
    'End If
    
    
    'AHORA
    If Clientes Then
        SQL = "select codccost from slifac where (codtipom,numfactu,fecfactu) "
        SQL = SQL & " in ( select codtipom,numfactu,fecfactu from scafac "
        If cadSQL <> "" Then SQL = SQL & " WHERE " & cadSQL
        SQL = SQL & ") GROUP BY codccost"
    
    Else
        SQL = "select codccost from slifpc where (codprove,numfactu,fecfactu) in ("
        SQL = SQL & "select codprove,numfactu,fecfactu from scafpc "
        If cadSQL <> "" Then SQL = SQL & " WHERE " & cadSQL
        SQL = SQL & ") GROUP BY codccost"
    

        
    End If
    
    Errores = ""  'De momento NO HAY ERRORES
    miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    SQL = ""
    While Not miRsAux.EOF
        If IsNull(miRsAux.Fields(0)) Then
            'MAL MAL. NO puede ser NULO
            Errores = Errores & "  ***  Lineas sin centro de coste asginado" & vbCrLf & vbCrLf
        Else
            SQL = SQL & DevNombreSQL(miRsAux.Fields(0)) & "|"
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If SQL <> "" Then
        
    
            
            While SQL <> ""
                i = InStr(1, SQL, "|")
                If i = 0 Then
                    'MSGBOX ALGO HA PASADO
                    Errores = Errores & " Sin asignar | en contabilizacion con centros de coste  " & vbCrLf
                    SQL = ""
                Else
                    C = Mid(SQL, 1, i - 1)
                    
                    C = DevuelveDesdeBD(conConta, "codccost", "cabccost", "codccost", C, "T")
                    If C = "" Then
                        'ERROR EN CC. NO EXISTE
                        Errores = Errores & " - " & Mid(SQL, 1, i - 1) & "       no existe  " & vbCrLf
                    End If
                    SQL = Mid(SQL, i + 1)
                End If
            Wend
    
            If Errores <> "" Then
                MsgBox Errores, vbExclamation
                
            Else
                ComprobarCCoste = 2
            End If
    Else
            ComprobarCCoste = 0
            If Errores <> "" Then
                Errores = "Errores en CC. No deberia continuar. " & vbCrLf & Errores & "�Continuar?"
                If MsgBox(Errores, vbQuestion + vbYesNo) = vbYes Then ComprobarCCoste = 1
                

            End If
    End If
    
    
    'ANTES
    'If SQL <> "" Then
    '    If Len(SQL) = 1 Then
    '        'Todos los CEntros de coste son el mismo. Con lo cual NO hara falta agrupar por trabajador
    '        ComprobarCCoste = 1
    '    Else
    '        'Tiene CC distintos. SI agruparemos por Trabajador
    '        ComprobarCCoste = 2
    '    End If
    'End If
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Cento de Coste", Err.Description
    End If
    Set miRsAux = Nothing
End Function


'Ccoste
'   0: No tendra analitica
'   1: Solo hay un CC que tratar. NO agruparemos por trabajador
'   2: Mas de un CC. Agruparemos por trabajador
'

Public Function PasarFactura(cadWHERE As String, CodCCost As Byte, EsContabilizacionAgrupadaTickets As Boolean, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' aritaxi.scafac --> conta.cabfact
' aritaxi.slifac --> conta.linfact
'Actualizar la tabla aritaxi.scafac.inconta=1 para indicar que ya esta contabilizada

'EsContabilizacionAgrupadaTickets:  La diferencia es en las lineas de la factura.
'                                   Si false: procedimeineto normal
'                                       true: Las lineas hare los select de otra forma
Dim b As Boolean
Dim cadMen As String
Dim SQL As String
Dim ErrorContab As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFact(cadWHERE, cadMen, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    vCCos = CodCCost
    If b Then
 
        'Insertar lineas de Factura en la Conta
        If EsContabilizacionAgrupadaTickets Then
            'Tickets agrupados
            b = InsertarLinFact_TicketsAgrupados("scafaccli", cadWHERE, cadMen, False)
        Else
            'Normal. Esta es la forma NORMAL NORMAL de hacerlo
            If vParamAplic.ContabilidadNueva Then
                b = InsertarLinFact_NUEVOContaNueva("scafaccli", cadWHERE, cadMen, False)
            Else
                b = InsertarLinFact_NUEVO("scafaccli", cadWHERE, cadMen, False)
            End If
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

'[Monica]18/02/2011: No contabilizamos las facturas
        If b Then
            If vParamAplic.ContabilidadNueva Then
                vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
            End If
        End If
            

        If b Then
            'Poner intconta=1 en aritaxi.scafac
            b = ActualizarCabFact("scafaccli", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    
EContab:
    
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    

    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFactura = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFactura = False
        'Inserto en errores, DESPUES del rollback. Si no no lo refleja, y al hacer el rollback
        'tira atras la insercion
        SQL = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        SQL = SQL & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        SQL = SQL & " WHERE " & Replace(cadWHERE, "scafac", "tmpFactu")
        conn.Execute SQL
        
    End If
        

        
    
End Function


Public Function PasarFacturaSOC(cadWHERE As String, CodCCost As Byte, EsContabilizacionAgrupadaTickets As Boolean, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' aritaxi.scafac --> conta.cabfact
' aritaxi.slifac --> conta.linfact
'Actualizar la tabla aritaxi.scafac.inconta=1 para indicar que ya esta contabilizada

'EsContabilizacionAgrupadaTickets:  La diferencia es en las lineas de la factura.
'                                   Si false: procedimeineto normal
'                                       true: Las lineas hare los select de otra forma
Dim b As Boolean
Dim cadMen As String
Dim SQL As String
Dim ErrorContab As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFactSOC(cadWHERE, cadMen, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    vCCos = CodCCost
    If b Then
 
        'Insertar lineas de Factura en la Conta
        If EsContabilizacionAgrupadaTickets Then
            'Tickets agrupados
            b = InsertarLinFact_TicketsAgrupados("scafac", cadWHERE, cadMen, False)
        Else
            'Normal. Esta es la forma NORMAL NORMAL de hacerlo
            If vParamAplic.ContabilidadNueva Then
                b = InsertarLinFact_NUEVOContaNueva("scafac", cadWHERE, cadMen, False)
            Else
                b = InsertarLinFact_NUEVO("scafac", cadWHERE, cadMen, False)
            End If
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

'[Monica]18/02/2011: No contabilizamos las facturas
'        If vContaFra.RealizarContabilizacion Then
'            ErrorContab = vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
'            vContaFra.AnyadeElError ErrorContab
'        End If
        If b Then
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
        End If

        If b Then
            'Poner intconta=1 en aritaxi.scafac
            b = ActualizarCabFact("scafac", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    
EContab:
    
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    

    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaSOC = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaSOC = False
        'Inserto en errores, DESPUES del rollback. Si no no lo refleja, y al hacer el rollback
        'tira atras la insercion
        SQL = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        SQL = SQL & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        SQL = SQL & " WHERE " & Replace(cadWHERE, "scafac", "tmpFactu")
        conn.Execute SQL
        
    End If
        

        
    
End Function




Public Function PasarFacturaCuotas(cadWHERE As String, CodCCost As Byte, EsContabilizacionAgrupadaTickets As Boolean, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' aritaxi.scafac --> conta.cabfact
' aritaxi.slifac --> conta.linfact
'Actualizar la tabla aritaxi.scafac.inconta=1 para indicar que ya esta contabilizada

'EsContabilizacionAgrupadaTickets:  La diferencia es en las lineas de la factura.
'                                   Si false: procedimeineto normal
'                                       true: Las lineas hare los select de otra forma
Dim b As Boolean
Dim cadMen As String
Dim SQL As String
Dim ErrorContab As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFactCuota(cadWHERE, cadMen, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    vCCos = CodCCost
    If b Then
 
        'Insertar lineas de Factura en la Conta
        If EsContabilizacionAgrupadaTickets Then
            'Tickets agrupados
            b = InsertarLinFact_TicketsAgrupados("scafac", cadWHERE, cadMen, False)
        Else
            'Normal. Esta es la forma NORMAL NORMAL de hacerlo
            If vParamAplic.ContabilidadNueva Then
                b = InsertarLinFact_NUEVOContaNueva("scafac", cadWHERE, cadMen, False)
            Else
                b = InsertarLinFact_NUEVO("scafac", cadWHERE, cadMen, False)
            End If
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen
        If b Then
            If vParamAplic.ContabilidadNueva Then
                vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
            End If
        End If

        If b Then
            'Poner intconta=1 en aritaxi.scafac
            b = ActualizarCabFact("scafac", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    
EContab:
    
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    

    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaCuotas = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaCuotas = False
        'Inserto en errores, DESPUES del rollback. Si no no lo refleja, y al hacer el rollback
        'tira atras la insercion
        SQL = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        SQL = SQL & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        SQL = SQL & " WHERE " & Replace(cadWHERE, "scafac", "tmpFactu")
        conn.Execute SQL
        
    End If
        

        
    
End Function





Private Function InsertarCabFact(cadWHERE As String, cadErr As String, ByRef vCF As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim SQL2 As String

Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim FraRectifica As String
Dim i As Integer
Dim CadenaInsertFaclin2 As String



    On Error GoTo EInsertar
    
    Set Rs = New ADODB.Recordset
    
    
    FraRectifica = ""
    If InStr(1, cadWHERE, "'FRT'") > 0 Then
        '�Voy a intentar sacar le numero de factura a la que rectifica. Sera de laobservacion
        SQL = Replace(cadWHERE, "scafac.", "scafac1.")
        Cad = "select observa1 from scafac1 where " & SQL
        Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            If Not IsNull(Rs!observa1) Then
                'Tiene valor, Vere si el valor es del tipo:
                ' RECTIFICA A FACTURA: A, 2600007, 30/1/2006
                Cad = CStr(Rs!observa1)
                i = InStr(1, Cad, "TURA:")  '.. FACTURA: A, 2600007 ....
                If i > 0 Then
                    Cad = Mid(Cad, i + 5)
                    i = InStr(1, Cad, ",")
                    If i > 0 Then
                        'La letra
                        SQL = Trim(Mid(Cad, 1, i - 1))
                        Cad = Mid(Cad, i + 1)
                        'Busco el NUMERO DE factura
                        i = InStr(1, Cad, ",")
                        If i > 0 Then
                            Cad = Mid(Cad, 1, i - 1)
                            If IsNumeric(Cad) Then
                                'Biennnnnnnnnnnnnnn
                                'Ya tengo el numero de factura
                                SQL = SQL & Cad
                            Else
                                SQL = ""
                            End If
                            FraRectifica = SQL
                        End If 'De buscando letra
                    End If 'De buscando n�fac
                End If 'RECTIFICA A FACTURA: A, 2600007, 30/1/2006
            End If
        End If
        Rs.Close
        Cad = ""
        
    End If
    SQL = " SELECT stipom.letraser,numfactu,fecfactu, scliente.codmacta,scliente.cliabono,year(fecfactu) as anofaccl,"
    SQL = SQL & "scafaccli.dtoppago,scafaccli.dtognral,baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,imporiv1,imporiv2,imporiv3,"
    SQL = SQL & "totalfac,codigiv1,codigiv2,codigiv3,aportacion "
    
    'Cuando MIS facfuras llevan recargo equivalencia
    SQL = SQL & ",porciva1re,porciva2re,porciva3re,imporiv1re,imporiv2re,imporiv3re,tipoiva"
    
    SQL = SQL & ",scafaccli.nomclien,scafaccli.domclien,scafaccli.codpobla,scafaccli.pobclien,scafaccli.proclien,scafaccli.nifclien,scafaccli.codforpa "
    
    
    SQL = SQL & " FROM (" & "scafaccli inner join " & "stipom on scafaccli.codtipom=stipom.codtipom) "
    SQL = SQL & "INNER JOIN " & "scliente ON scafaccli.codclien=scliente.codclien "
    SQL = SQL & " WHERE " & cadWHERE
    
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        vCF.NumeroFactura = DBLet(Rs!NumFactu)
        vCF.Anofac = Year(DBLet(Rs!FecFactu))
        vCF.Serie = DBLet(Rs!LetraSer)
    
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = Rs!DtoPPago
        DtoGnral = Rs!DtoGnral
        BaseImp = Rs!baseimp1 + CCur(DBLet(Rs!baseimp2, "N")) + CCur(DBLet(Rs!baseimp3, "N"))
        IvaImp = DBLet(Rs!imporiv1, "N") + DBLet(Rs!imporiv2, "N") + DBLet(Rs!imporiv3, "N")
        
        '---- Laura 10/10/2006:  a�adir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        DatosAportacion = ""
        If Rs!Aportacion > 0 Then
            'Deberia dar error si vparam.ctaaportacion=""
            DatosAportacion = Rs!codmacta & "|" & Rs!Aportacion & "|"
        Else
            
        End If
        '----
        conCtaAlt = Rs!cliabono
        
        
        'Guardamos los valores de la factura que estoy integrando
        If vCF.RealizarContabilizacion Then vCF.FijarNumeroFactura Rs!NumFactu, Year(Rs!FecFactu), Rs!LetraSer
        
        SQL = ""
        SQL = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!codmacta, "T") & "," & Year(Rs!FecFactu) & ","
        
        'MAYO 2009
        'Si es una factura rectificativa, y hemos encontrado
        ' a k factura rectifica entonces meto esto, sino sigue como antes
        If FraRectifica = "" Then
            Select Case vParamAplic.ObsFactura
            Case 0
                'Vacio
                SQL = SQL & ValorNulo
            Case 1
                'N� Factura
                SQL = SQL & "'" & DevNombreSQL("N/Fra " & Rs!NumFactu) & "'"
            Case 2
                'Fecha integracion
                SQL = SQL & "'" & Format(Now, FormatoFecha) & "'"
            End Select
            
        Else
            SQL = SQL & "'" & FraRectifica & "'"
        End If
        
        '## LAURA (25/07/2008)
        Nulo2 = "N"
        Nulo3 = "N"
        If DBLet(Rs!codigiv2, "N") = 0 Then Nulo2 = "S"
        If DBLet(Rs!codigiv3, "N") = 0 Then Nulo3 = "S"
        
        If Not vParamAplic.ContabilidadNueva Then
            'Abril
            SQL = SQL & "," & DBSet(Rs!baseimp1, "N") & "," & DBSetDavid(Rs!baseimp2, "N", Nulo2) & "," & DBSetDavid(Rs!baseimp3, "N", Nulo3) & ","
            
            'SQL = SQL & DBSet(RS!porciva1, "N") & "," & DBSet(RS!porciva2, "N", Nulo2) & "," & DBSet(RS!porciva3, "N", Nulo3)
            SQL = SQL & DBSet(Rs!porciva1, "N") & "," & DBSetDavid(Rs!porciva2, "N", Nulo2) & "," & DBSetDavid(Rs!porciva3, "N", Nulo3)
            
            SQL = SQL & "," & DBSet(Rs!porciva1re, "N", "S") & "," & DBSet(Rs!porciva2re, "N", "S") & "," & DBSet(Rs!porciva3re, "N", "S")
            
            SQL = SQL & "," & DBSet(Rs!imporiv1, "N", "N") & "," & DBSetDavid(Rs!imporiv2, "N", Nulo2) & "," & DBSetDavid(Rs!imporiv3, "N", Nulo3)
            
            'ANTES
            'SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & "," & DBSet(Rs!imporiv1re, "N", "S") & "," & DBSet(Rs!imporiv2re, "N", "S") & "," & DBSet(Rs!imporiv3re, "N", "S") & ","
            
            SQL = SQL & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!codigiv1, "N") & "," & DBSet(Rs!codigiv2, "N", Nulo2) & "," & DBSet(Rs!codigiv3, "N", Nulo3) & ","
            
            'INTRACOM
            If Rs!TipoIVA = 3 Then
                'Tipo de iva intrcomunitatro
                SQL = SQL & "1"
            Else
                SQL = SQL & "0"
            End If
            
            SQL = SQL & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(Rs!FecFactu, "F")
            Cad = Cad & "(" & SQL & ")"
    '        RS.MoveNext
    
            'Insertar en la contabilidad
            SQL = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            SQL = SQL & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            SQL = SQL & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
            SQL = SQL & " VALUES " & Cad
            ConnConta.Execute SQL
        Else
        
            SQL = SQL & ","
            'If FraRectifica <> "" Then
            If InStr(1, cadWHERE, "'FRT'") > 0 Then
                SQL = SQL & "'D',"
            Else
                SQL = SQL & "'0',"
            End If
            
            SQL = SQL & "0," & DBSet(Rs!codforpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
            SQL = SQL & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!FecFactu, "F") & ","
            SQL = SQL & DBSet(Rs!nomclien, "T") & "," & DBSet(Rs!domclien, "T") & "," & DBSet(Rs!codpobla, "T") & "," & DBSet(Rs!pobclien, "T") & ","
            SQL = SQL & DBSet(Rs!proclien, "T") & "," & DBSet(Rs!nifClien, "T") & ",'ES',1"
            
            Cad = "(" & SQL & ")"
        
            SQL = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            SQL = SQL & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            SQL = SQL & "codpais,codagente)"
            SQL = SQL & " VALUES " & Cad
            ConnConta.Execute SQL
            
            CadenaInsertFaclin2 = ""

            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            SQL2 = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
            SQL2 = SQL2 & "1," & DBSet(Rs!baseimp1, "N") & "," & Rs!codigiv1 & "," & DBSet(Rs!porciva1, "N") & ","
            SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!imporiv1, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & SQL2 & ")"
        
            'para las lineas
            vTipoIva(0) = Rs!codigiv1
            vPorcIva(0) = Rs!porciva1
            vPorcRec(0) = 0
            vImpIva(0) = Rs!imporiv1
            vImpRec(0) = 0
            vBaseIva(0) = Rs!baseimp1
            
            vTipoIva(1) = 0: vTipoIva(2) = 0
            
            If Not IsNull(Rs!porciva2) Then
                SQL2 = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
                SQL2 = SQL2 & "2," & DBSet(Rs!baseimp2, "N") & "," & Rs!codigiv2 & "," & DBSet(Rs!porciva2, "N") & ","
                SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!imporiv2, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                vTipoIva(1) = Rs!codigiv2
                vPorcIva(1) = Rs!porciva2
                vPorcRec(1) = 0
                vImpIva(1) = Rs!imporiv2
                vImpRec(1) = 0
                vBaseIva(1) = Rs!baseimp2
            End If
            If Not IsNull(Rs!porciva3) Then
                SQL2 = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
                SQL2 = SQL2 & "3," & DBSet(Rs!baseimp3, "N") & "," & Rs!codigiv3 & "," & DBSet(Rs!porciva3, "N") & ","
                SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!imporiv3, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                vTipoIva(2) = Rs!codigiv3
                vPorcIva(2) = Rs!porciva3
                vPorcRec(2) = 0
                vImpIva(2) = Rs!imporiv3
                vImpRec(2) = 0
                vBaseIva(2) = Rs!baseimp3
            End If
    
            SQL = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            SQL = SQL & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnConta.Execute SQL
        
        End If
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFact = False
        cadErr = Err.Description
    Else
        InsertarCabFact = True
    End If
End Function


Private Function InsertarCabFactSOC(cadWHERE As String, cadErr As String, ByRef vCF As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim SQL2 As String

Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim FraRectifica As String
Dim i As Integer
Dim CadenaInsertFaclin2 As String

    On Error GoTo EInsertar
    
    Set Rs = New ADODB.Recordset
    
    
    FraRectifica = ""
    If InStr(1, cadWHERE, "'FRT'") > 0 Then
        '�Voy a intentar sacar le numero de factura a la que rectifica. Sera de laobservacion
        SQL = Replace(cadWHERE, "scafac.", "scafac1.")
        Cad = "select observa1 from scafac1 where " & SQL
        Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            If Not IsNull(Rs!observa1) Then
                'Tiene valor, Vere si el valor es del tipo:
                ' RECTIFICA A FACTURA: A, 2600007, 30/1/2006
                Cad = CStr(Rs!observa1)
                i = InStr(1, Cad, "TURA:")  '.. FACTURA: A, 2600007 ....
                If i > 0 Then
                    Cad = Mid(Cad, i + 5)
                    i = InStr(1, Cad, ",")
                    If i > 0 Then
                        'La letra
                        SQL = Trim(Mid(Cad, 1, i - 1))
                        Cad = Mid(Cad, i + 1)
                        'Busco el NUMERO DE factura
                        i = InStr(1, Cad, ",")
                        If i > 0 Then
                            Cad = Mid(Cad, 1, i - 1)
                            If IsNumeric(Cad) Then
                                'Biennnnnnnnnnnnnnn
                                'Ya tengo el numero de factura
                                SQL = SQL & Cad
                            Else
                                SQL = ""
                            End If
                            FraRectifica = SQL
                        End If 'De buscando letra
                    End If 'De buscando n�fac
                End If 'RECTIFICA A FACTURA: A, 2600007, 30/1/2006
            End If
        End If
        Rs.Close
        Cad = ""
        
    End If
    Dim CADENA1 As String
    Dim LCad As Integer
    
    CADENA1 = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
    LCad = Len(CADENA1)

    '[Monica] 30/05/2011: si es una factura rectificativa o de venta
    If InStr(1, cadWHERE, "'FRT'") > 0 Or (InStr(1, cadWHERE, "'FCN'") = 0 And InStr(1, cadWHERE, "'FCE'") = 0) Then
        SQL = " SELECT stipom.letraser,numfactu,fecfactu,concat('" & vParamAplic.Raiz_Cta_Soc_Equip & "',right(concat('" & CADENA1 & "',scafac.codclien)," & LCad & ")) codmacta  ,0 cliabono,year(fecfactu) as anofaccl,"
    Else
        SQL = " SELECT stipom.letraser,numfactu,fecfactu,concat('" & vParamAplic.Raiz_CtaClien_Soc & "',right(concat('" & CADENA1 & "',scafac.codclien)," & LCad & ")) codmacta  ,0 cliabono,year(fecfactu) as anofaccl,"
    End If
    
    SQL = SQL & "scafac.dtoppago,scafac.dtognral,baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,imporiv1,imporiv2,imporiv3,"
    SQL = SQL & "totalfac,codigiv1,codigiv2,codigiv3,aportacion "
    
    'Cuando MIS facfuras llevan recargo equivalencia
    SQL = SQL & ",porciva1re,porciva2re,porciva3re,imporiv1re,imporiv2re,imporiv3re,0 tipoiva,"
    
    SQL = SQL & "scafac.nomclien,scafac.domclien,scafac.codpobla,scafac.pobclien,scafac.proclien,scafac.nifclien,scafac.codforpa "
    
    SQL = SQL & " FROM (" & "scafac inner join " & "stipom on scafac.codtipom=stipom.codtipom) "
    SQL = SQL & "INNER JOIN " & "sclien ON scafac.codclien=sclien.codclien "
    SQL = SQL & " WHERE " & cadWHERE
    
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
    
        vCF.NumeroFactura = DBLet(Rs!NumFactu)
        vCF.Anofac = Year(DBLet(Rs!FecFactu))
        vCF.Serie = DBLet(Rs!LetraSer)
    
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = Rs!DtoPPago
        DtoGnral = Rs!DtoGnral
        BaseImp = Rs!baseimp1 + CCur(DBLet(Rs!baseimp2, "N")) + CCur(DBLet(Rs!baseimp3, "N"))
        IvaImp = DBLet(Rs!imporiv1, "N") + DBLet(Rs!imporiv2, "N") + DBLet(Rs!imporiv3, "N")
        '---- Laura 10/10/2006:  a�adir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        DatosAportacion = ""
        If Rs!Aportacion > 0 Then
            'Deberia dar error si vparam.ctaaportacion=""
            DatosAportacion = Rs!codmacta & "|" & Rs!Aportacion & "|"
        Else
            
        End If
        '----
        conCtaAlt = Rs!cliabono
        
        
        'Guardamos los valores de la factura que estoy integrando
        If vCF.RealizarContabilizacion Then vCF.FijarNumeroFactura Rs!NumFactu, Year(Rs!FecFactu), Rs!LetraSer
        
        SQL = ""
        SQL = SQL & "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!codmacta, "T") & "," & Year(Rs!FecFactu) & ","
        
        'MAYO 2009
        'Si es una factura rectificativa, y hemos encontrado
        ' a k factura rectifica entonces meto esto, sino sigue como antes
        If FraRectifica = "" Then
            Select Case vParamAplic.ObsFactura
            Case 0
                'Vacio
                SQL = SQL & ValorNulo
            Case 1
                'N� Factura
                SQL = SQL & "'" & DevNombreSQL("N/Fra " & Rs!NumFactu) & "'"
            Case 2
                'Fecha integracion
                SQL = SQL & "'" & Format(Now, FormatoFecha) & "'"
            End Select
        Else
            SQL = SQL & "'" & FraRectifica & "'"
        End If
        
        '## LAURA (25/07/2008)
        Nulo2 = "N"
        Nulo3 = "N"
        If DBLet(Rs!codigiv2, "N") = 0 Then Nulo2 = "S"
        If DBLet(Rs!codigiv3, "N") = 0 Then Nulo3 = "S"
        
        If Not vParamAplic.ContabilidadNueva Then
            'Abril
            SQL = SQL & "," & DBSet(Rs!baseimp1, "N") & "," & DBSetDavid(Rs!baseimp2, "N", Nulo2) & "," & DBSetDavid(Rs!baseimp3, "N", Nulo3) & ","
            
            
            'SQL = SQL & DBSet(RS!porciva1, "N") & "," & DBSet(RS!porciva2, "N", Nulo2) & "," & DBSet(RS!porciva3, "N", Nulo3)
            SQL = SQL & DBSet(Rs!porciva1, "N") & "," & DBSetDavid(Rs!porciva2, "N", Nulo2) & "," & DBSetDavid(Rs!porciva3, "N", Nulo3)
            
            
            SQL = SQL & "," & DBSet(Rs!porciva1re, "N", "S") & "," & DBSet(Rs!porciva2re, "N", "S") & "," & DBSet(Rs!porciva3re, "N", "S")
            
            
            SQL = SQL & "," & DBSet(Rs!imporiv1, "N", "N") & "," & DBSetDavid(Rs!imporiv2, "N", Nulo2) & "," & DBSetDavid(Rs!imporiv3, "N", Nulo3)
            
            'ANTES
            'SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & "," & DBSet(Rs!imporiv1re, "N", "S") & "," & DBSet(Rs!imporiv2re, "N", "S") & "," & DBSet(Rs!imporiv3re, "N", "S") & ","
            
            
            SQL = SQL & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!codigiv1, "N") & "," & DBSet(Rs!codigiv2, "N", Nulo2) & "," & DBSet(Rs!codigiv3, "N", Nulo3) & ","
            
            'INTRACOM
            If Rs!TipoIVA = 3 Then
                'Tipo de iva intrcomunitatro
                SQL = SQL & "1"
            Else
                SQL = SQL & "0"
            End If
            
            SQL = SQL & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(Rs!FecFactu, "F")
            Cad = Cad & "(" & SQL & ")"
            
            
            'Insertar en la contabilidad
            SQL = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            SQL = SQL & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            SQL = SQL & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
            SQL = SQL & " VALUES " & Cad
            ConnConta.Execute SQL
            
        Else ' contabilidad nueva
            If InStr(1, cadWHERE, "'FRT'") > 0 Then
            'If FraRectifica <> "" Then
                SQL = SQL & ",'D',"
            Else
                SQL = SQL & ",'0',"
            End If
            
            SQL = SQL & "0," & DBSet(Rs!codforpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
            SQL = SQL & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!FecFactu, "F") & ","
            SQL = SQL & DBSet(Rs!nomclien, "T") & "," & DBSet(Rs!domclien, "T") & "," & DBSet(Rs!codpobla, "T") & "," & DBSet(Rs!pobclien, "T") & ","
            SQL = SQL & DBSet(Rs!proclien, "T") & "," & DBSet(Rs!nifClien, "T") & ",'ES',1"
            
            Cad = "(" & SQL & ")"
        
            SQL = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            SQL = SQL & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            SQL = SQL & "codpais,codagente)"
            SQL = SQL & " VALUES " & Cad
            ConnConta.Execute SQL
            
            CadenaInsertFaclin2 = ""

            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            SQL2 = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
            SQL2 = SQL2 & "1," & DBSet(Rs!baseimp1, "N") & "," & Rs!codigiv1 & "," & DBSet(Rs!porciva1, "N") & ","
            SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!imporiv1, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & SQL2 & ")"
        
            'para las lineas
            vTipoIva(0) = Rs!codigiv1
            vPorcIva(0) = Rs!porciva1
            vPorcRec(0) = 0
            vImpIva(0) = Rs!imporiv1
            vImpRec(0) = 0
            vBaseIva(0) = Rs!baseimp1
            
            vTipoIva(1) = 0: vTipoIva(2) = 0
            
            If Not IsNull(Rs!porciva2) Then
                SQL2 = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
                SQL2 = SQL2 & "2," & DBSet(Rs!baseimp2, "N") & "," & Rs!codigiv2 & "," & DBSet(Rs!porciva2, "N") & ","
                SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!imporiv2, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                vTipoIva(1) = Rs!codigiv2
                vPorcIva(1) = Rs!porciva2
                vPorcRec(1) = 0
                vImpIva(1) = Rs!imporiv2
                vImpRec(1) = 0
                vBaseIva(1) = Rs!baseimp2
            End If
            If Not IsNull(Rs!porciva3) Then
                SQL2 = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
                SQL2 = SQL2 & "3," & DBSet(Rs!baseimp3, "N") & "," & Rs!codigiv3 & "," & DBSet(Rs!porciva3, "N") & ","
                SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!imporiv3, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                vTipoIva(2) = Rs!codigiv3
                vPorcIva(2) = Rs!porciva3
                vPorcRec(2) = 0
                vImpIva(2) = Rs!imporiv3
                vImpRec(2) = 0
                vBaseIva(2) = Rs!baseimp3
            End If
    
            SQL = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            SQL = SQL & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnConta.Execute SQL
                
                
        End If
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactSOC = False
        cadErr = Err.Description
    Else
        InsertarCabFactSOC = True
    End If
End Function




Private Function InsertarCabFactCuota(cadWHERE As String, cadErr As String, ByRef vCF As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim SQL2 As String

Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim FraRectifica As String
Dim i As Integer
Dim CadenaInsertFaclin2 As String

    On Error GoTo EInsertar
    
    Set Rs = New ADODB.Recordset
    
    
    FraRectifica = ""
    If InStr(1, cadWHERE, "'FRC'") > 0 Then
        '�Voy a intentar sacar le numero de factura a la que rectifica. Sera de laobservacion
        SQL = Replace(cadWHERE, "scafac.", "scafac1.")
        Cad = "select observa1 from scafac1 where " & SQL
        Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            If Not IsNull(Rs!observa1) Then
                'Tiene valor, Vere si el valor es del tipo:
                ' RECTIFICA A FACTURA: A, 2600007, 30/1/2006
                Cad = CStr(Rs!observa1)
                i = InStr(1, Cad, "TURA:")  '.. FACTURA: A, 2600007 ....
                If i > 0 Then
                    Cad = Mid(Cad, i + 5)
                    i = InStr(1, Cad, ",")
                    If i > 0 Then
                        'La letra
                        SQL = Trim(Mid(Cad, 1, i - 1))
                        Cad = Mid(Cad, i + 1)
                        'Busco el NUMERO DE factura
                        i = InStr(1, Cad, ",")
                        If i > 0 Then
                            Cad = Mid(Cad, 1, i - 1)
                            If IsNumeric(Cad) Then
                                'Biennnnnnnnnnnnnnn
                                'Ya tengo el numero de factura
                                SQL = SQL & Cad
                            Else
                                SQL = ""
                            End If
                            FraRectifica = SQL
                        End If 'De buscando letra
                    End If 'De buscando n�fac
                End If 'RECTIFICA A FACTURA: A, 2600007, 30/1/2006
            End If
        End If
        Rs.Close
        Cad = ""
        
    End If
    
    Dim CADENA1 As String
    Dim LCad As Integer
    
    CADENA1 = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
    LCad = Len(CADENA1)
'    mCtaSocioLiq = vParamAplic.Raiz_Cta_Soc_Liqui & Format(mCodSocio, CADENA)

    SQL = " SELECT stipom.letraser,numfactu,fecfactu, concat('" & vParamAplic.Raiz_CtaClien_Soc & "',right(concat('" & CADENA1 & "',scafac.codclien)," & LCad & ")) codmacta,0 cliabono,year(fecfactu) as anofaccl,"
    SQL = SQL & "scafac.dtoppago,scafac.dtognral,baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,imporiv1,imporiv2,imporiv3,"
    SQL = SQL & "totalfac,codigiv1,codigiv2,codigiv3,aportacion "
    
    'Cuando MIS facfuras llevan recargo equivalencia
    SQL = SQL & ",porciva1re,porciva2re,porciva3re,imporiv1re,imporiv2re,imporiv3re, 0 tipoiva"
    
    SQL = SQL & ",scafac.nomclien,scafac.domclien,scafac.codpobla,scafac.pobclien,scafac.proclien,scafac.nifclien,scafac.codforpa "
    
    SQL = SQL & " FROM (" & "scafac inner join " & "stipom on scafac.codtipom=stipom.codtipom) "
    SQL = SQL & "INNER JOIN " & "sclien ON scafac.codclien=sclien.codclien "
    SQL = SQL & " WHERE " & cadWHERE
    
  
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        vCF.NumeroFactura = DBLet(Rs!NumFactu)
        vCF.Anofac = Year(DBLet(Rs!FecFactu))
        vCF.Serie = DBLet(Rs!LetraSer)
    
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = Rs!DtoPPago
        DtoGnral = Rs!DtoGnral
        BaseImp = Rs!baseimp1 + CCur(DBLet(Rs!baseimp2, "N")) + CCur(DBLet(Rs!baseimp3, "N"))
        IvaImp = DBLet(Rs!imporiv1, "N") + DBLet(Rs!imporiv2, "N") + DBLet(Rs!imporiv3, "N")
        
        '---- Laura 10/10/2006:  a�adir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        DatosAportacion = ""
        If Rs!Aportacion > 0 Then
            'Deberia dar error si vparam.ctaaportacion=""
            DatosAportacion = Rs!codmacta & "|" & Rs!Aportacion & "|"
        Else
            
        End If
        '----
        conCtaAlt = Rs!cliabono
        
        
        'Guardamos los valores de la factura que estoy integrando
        If vCF.RealizarContabilizacion Then vCF.FijarNumeroFactura Rs!NumFactu, Year(Rs!FecFactu), Rs!LetraSer
        
        SQL = ""
        SQL = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!codmacta, "T") & "," & Year(Rs!FecFactu) & ","
        
        'MAYO 2009
        'Si es una factura rectificativa, y hemos encontrado
        ' a k factura rectifica entonces meto esto, sino sigue como antes
        If FraRectifica = "" Then
            Select Case vParamAplic.ObsFactura
            Case 0
                'Vacio
                SQL = SQL & ValorNulo
            Case 1
                'N� Factura
                SQL = SQL & "'" & DevNombreSQL("N/Fra " & Rs!NumFactu) & "'"
            Case 2
                'Fecha integracion
                SQL = SQL & "'" & Format(Now, FormatoFecha) & "'"
            End Select
            
        Else
            SQL = SQL & "'" & FraRectifica & "'"
        End If
        
        '## LAURA (25/07/2008)
        Nulo2 = "N"
        Nulo3 = "N"
        If DBLet(Rs!codigiv2, "N") = 0 Then Nulo2 = "S"
        If DBLet(Rs!codigiv3, "N") = 0 Then Nulo3 = "S"
        
        'Abril
        If Not vParamAplic.ContabilidadNueva Then
            SQL = SQL & "," & DBSet(Rs!baseimp1, "N") & "," & DBSetDavid(Rs!baseimp2, "N", Nulo2) & "," & DBSetDavid(Rs!baseimp3, "N", Nulo3) & ","
            
            'SQL = SQL & DBSet(RS!porciva1, "N") & "," & DBSet(RS!porciva2, "N", Nulo2) & "," & DBSet(RS!porciva3, "N", Nulo3)
            SQL = SQL & DBSet(Rs!porciva1, "N") & "," & DBSetDavid(Rs!porciva2, "N", Nulo2) & "," & DBSetDavid(Rs!porciva3, "N", Nulo3)
            
            SQL = SQL & "," & DBSet(Rs!porciva1re, "N", "S") & "," & DBSet(Rs!porciva2re, "N", "S") & "," & DBSet(Rs!porciva3re, "N", "S")
            
            SQL = SQL & "," & DBSet(Rs!imporiv1, "N", "N") & "," & DBSetDavid(Rs!imporiv2, "N", Nulo2) & "," & DBSetDavid(Rs!imporiv3, "N", Nulo3)
            
            'ANTES
            'SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & "," & DBSet(Rs!imporiv1re, "N", "S") & "," & DBSet(Rs!imporiv2re, "N", "S") & "," & DBSet(Rs!imporiv3re, "N", "S") & ","
            
            SQL = SQL & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!codigiv1, "N") & "," & DBSet(Rs!codigiv2, "N", Nulo2) & "," & DBSet(Rs!codigiv3, "N", Nulo3) & ","
            
            'INTRACOM
            If Rs!TipoIVA = 3 Then
                'Tipo de iva intrcomunitatro
                SQL = SQL & "1"
            Else
                SQL = SQL & "0"
            End If
            
            SQL = SQL & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(Rs!FecFactu, "F")
            Cad = Cad & "(" & SQL & ")"
    '        RS.MoveNext
            
            'Insertar en la contabilidad
            SQL = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            SQL = SQL & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            SQL = SQL & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
            SQL = SQL & " VALUES " & Cad
            ConnConta.Execute SQL
        Else
            SQL = SQL & ","
'            If FraRectifica <> "" Then
            If InStr(1, cadWHERE, "'FRC'") > 0 Then
                SQL = SQL & "'D',"
            Else
                SQL = SQL & "'0',"
            End If
            
            SQL = SQL & "0," & DBSet(Rs!codforpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
            SQL = SQL & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!FecFactu, "F") & ","
            SQL = SQL & DBSet(Rs!nomclien, "T") & "," & DBSet(Rs!domclien, "T") & "," & DBSet(Rs!codpobla, "T") & "," & DBSet(Rs!pobclien, "T") & ","
            SQL = SQL & DBSet(Rs!proclien, "T") & "," & DBSet(Rs!nifClien, "T") & ",'ES',1"
            
            Cad = "(" & SQL & ")"
        
            SQL = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            SQL = SQL & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            SQL = SQL & "codpais,codagente)"
            SQL = SQL & " VALUES " & Cad
            ConnConta.Execute SQL
            
            CadenaInsertFaclin2 = ""

            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            SQL2 = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
            SQL2 = SQL2 & "1," & DBSet(Rs!baseimp1, "N") & "," & Rs!codigiv1 & "," & DBSet(Rs!porciva1, "N") & ","
            SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!imporiv1, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & SQL2 & ")"
        
            'para las lineas
            vTipoIva(0) = Rs!codigiv1
            vPorcIva(0) = Rs!porciva1
            vPorcRec(0) = 0
            vImpIva(0) = Rs!imporiv1
            vImpRec(0) = 0
            vBaseIva(0) = Rs!baseimp1
            
            vTipoIva(1) = 0: vTipoIva(2) = 0
            
            If Not IsNull(Rs!porciva2) Then
                SQL2 = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
                SQL2 = SQL2 & "2," & DBSet(Rs!baseimp2, "N") & "," & Rs!codigiv2 & "," & DBSet(Rs!porciva2, "N") & ","
                SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!imporiv2, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                vTipoIva(1) = Rs!codigiv2
                vPorcIva(1) = Rs!porciva2
                vPorcRec(1) = 0
                vImpIva(1) = Rs!imporiv2
                vImpRec(1) = 0
                vBaseIva(1) = Rs!baseimp2
            End If
            If Not IsNull(Rs!porciva3) Then
                SQL2 = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
                SQL2 = SQL2 & "3," & DBSet(Rs!baseimp3, "N") & "," & Rs!codigiv3 & "," & DBSet(Rs!porciva3, "N") & ","
                SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!imporiv3, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                vTipoIva(2) = Rs!codigiv3
                vPorcIva(2) = Rs!porciva3
                vPorcRec(2) = 0
                vImpIva(2) = Rs!imporiv3
                vImpRec(2) = 0
                vBaseIva(2) = Rs!baseimp3
            End If
    
            SQL = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            SQL = SQL & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnConta.Execute SQL
        
        End If
    
    
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactCuota = False
        cadErr = Err.Description
    Else
        InsertarCabFactCuota = True
    End If
End Function





'Private Function InsertarLinFact(cadTabla As String, cadWhere As String, cadErr As String, Optional numRegis As Long) As Boolean
''cadWHere: selecciona un registro de scafac
''codtipom=x and numfactu=y and fecfactu=z
'Dim SQL As String
'Dim SQLaux As String
'Dim SQL2 As String
'Dim RS As ADODB.Recordset
'Dim Cad As String, Aux As String
'Dim I As Byte
'Dim TotImp As Currency, ImpLinea As Currency
'
'    On Error GoTo EInLinea
'
'    If cadTabla = "scafac" Then
'        SQL = " SELECT stipom.letraser,slifac.codtipom,numfactu,fecfactu,sartic.codfamia,sfamia.ctaventa,sfamia.ctavent1,sfamia.aboventa,sfamia.abovent1,sum(importel) as importe "
'        SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
'        SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafac", "slifac")
'        SQL = SQL & " GROUP BY sfamia.codfamia "
'    Else
'        SQL = " SELECT slifpc.codprove,numfactu,fecfactu,sartic.codfamia,sfamia.ctacompr,sfamia.abocompr,sum(importel) as importe "
'        SQL = SQL & " FROM (slifpc  "
'        SQL = SQL & " inner join sartic on slifpc.codartic=sartic.codartic) "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "slifpc")
'        SQL = SQL & " GROUP BY sfamia.codfamia "
'    End If
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    Cad = ""
'    I = 1
'    TotImp = 0
'    SQLaux = ""
'    While Not RS.EOF
'        SQLaux = Cad
'        'calculamos la Base Imp del total del importe para cada cta cble ventas
'        '---- Laura: 10/10/2006
'        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
'        ImpLinea = RS!Importe - CalcularPorcentaje(RS!Importe, DtoPPago, 2)
'        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
'        ImpLinea = ImpLinea - CalcularPorcentaje(RS!Importe, DtoGnral, 2)
'        'ImpLinea = Round(ImpLinea, 2)
'        '----
'        TotImp = TotImp + ImpLinea
'
'        'concatenamos linea para insertar en la tabla de conta.linfact
'        SQL = ""
'        SQL2 = ""
'        If cadTabla = "scafac" Then
'            SQL = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & Year(RS!FecFactu) & "," & I & ","
'            If Not conCtaAlt Then 'cliente no tiene cuenta alternativa
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctaventa, "T")
'                Else
'                    SQL = SQL & DBSet(RS!aboventa, "T")
'                End If
'            Else
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctavent1, "T")
'                Else
'                    SQL = SQL & DBSet(RS!abovent1, "T")
'                End If
'            End If
'        Else
'            SQL = numRegis & "," & Year(RS!FecFactu) & "," & I & ","
'            If ImpLinea >= 0 Then
'                SQL = SQL & DBSet(RS!ctacompr, "T")
'            Else
'                SQL = SQL & DBSet(RS!abocompr, "T")
'            End If
'        End If
'        SQL2 = SQL & ","
'        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
'
'        If CCoste = "" Then
'            SQL = SQL & ValorNulo
'        Else
'            SQL = SQL & DBSet(CCoste, "T")
'        End If
'
'        Cad = Cad & "(" & SQL & ")" & ","
'
'        I = I + 1
'        RS.MoveNext
'    Wend
'    RS.Close
'    Set RS = Nothing
'
'    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
'    'de la factura
'    If TotImp <> BaseImp Then
''        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
'        'en SQL esta la ult linea introducida
'        TotImp = BaseImp - TotImp
'        TotImp = ImpLinea + TotImp '(+- diferencia)
'        SQL2 = SQL2 & DBSet(TotImp, "N") & ","
'        If CCoste = "" Then
'            SQL2 = SQL2 & ValorNulo
'        Else
'            SQL2 = SQL2 & DBSet(CCoste, "T")
'        End If
'        If SQLaux <> "" Then 'hay mas de una linea
'            Cad = SQLaux & "(" & SQL2 & ")" & ","
'        Else 'solo una linea
'            Cad = "(" & SQL2 & ")" & ","
'        End If
'
''        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
''        cad = Replace(cad, SQL, Aux)
'    End If
'
'
'    'Insertar en la contabilidad
'    If Cad <> "" Then
'        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
'        If cadTabla = "scafac" Then
'            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
'        Else
'            SQL = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
'        End If
'        SQL = SQL & " VALUES " & Cad
'        ConnConta.Execute SQL
'    End If
'
'EInLinea:
'    If Err.Number <> 0 Then
'        InsertarLinFact = False
'        cadErr = Err.Description
'    Else
'        InsertarLinFact = True
'    End If
'End Function
'


'
'Si lleva retencion(FRAPRO) se a�adiren dos lineas codprove contra ctareten

Private Function InsertarLinFact_NUEVO(cadTabla As String, cadWHERE As String, cadErr As String, vLlevaRetencion As Boolean, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SQLaux As String
Dim SQL2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Byte
Dim TotImp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim LineaCentroCoste As Boolean
    'Puede ser que teniendo analitica, la cuenta no sea del grupo 6 o 7 , con lo cual nodebe poner el CC
    'Por si acaso alguna linea no es del grupo venta o grupo compras, no

    On Error GoTo EInLinea
    

    '
    '   Habra que ver en funcion de CC que tenga si agrupo, o no, por  codtraba
    '
    Select Case cadTabla
        Case "scafaccli" ' ventas
             'comprobar si el cliente utiliza cuenta alternativa
            If conCtaAlt Then
                'utilizamos sfamia.ctavent1 o sfamia.abovent1
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctavent1"
                Else
                    cadCampo = "sfamia.abovent1" 'si es negativa es un abono
                End If
            Else
                'utilizamos sfamia.ctaventa o sfamia.aboventa
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctaventa"
                Else
                    cadCampo = "sfamia.aboventa"
                End If
            End If
            
            
            SQL = " SELECT stipom.letraser,slifaccli.codtipom,slifaccli.numfactu,slifaccli.fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
            
            'Tiene analitica. Luego el codtraba tiene que aparecer
            If vCCos > 0 Then SQL = SQL & ",slifaccli.codccost"
            
            SQL = SQL & " FROM ((slifaccli inner join stipom on slifaccli.codtipom=stipom.codtipom) "
            SQL = SQL & " inner join sartic on slifaccli.codartic=sartic.codartic) "
            SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            
            
            SQL = SQL & " WHERE "
            
            
            SQL = SQL & " " & Replace(cadWHERE, "scafaccli", "slifaccli")
            
            '[Monica]15/01/2018: no cogemos lo correspondiente a suplidos
            SQL = SQL & " and slifaccli.codartic <> " & DBSet(vParamAplic.ArtSuplidos, "T")
            
            
            
            SQL = SQL & " GROUP BY "
            
            'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
            If vCCos > 0 Then SQL = SQL & " codccost, "
                      
            'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
            SQL = SQL & cadCampo
    
    Case "scafac"
             'comprobar si el cliente utiliza cuenta alternativa
            If conCtaAlt Then
                'utilizamos sfamia.ctavent1 o sfamia.abovent1
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctavent1"
                Else
                    cadCampo = "sfamia.abovent1" 'si es negativa es un abono
                End If
            Else
                'utilizamos sfamia.ctaventa o sfamia.aboventa
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctaventa"
                Else
                    cadCampo = "sfamia.aboventa"
                End If
            End If
            
            
            SQL = " SELECT stipom.letraser,slifac.codtipom,slifac.numfactu,slifac.fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
            
            'Tiene analitica. Luego el codtraba tiene que aparecer
            If vCCos > 0 Then SQL = SQL & ",slifac.codccost"
            
            SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
            SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
            SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            
            
            SQL = SQL & " WHERE "
            
            
            SQL = SQL & " " & Replace(cadWHERE, "scafac", "slifac")
            SQL = SQL & " GROUP BY "
            
            'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
            If vCCos > 0 Then SQL = SQL & " codccost, "
                      
            'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
            SQL = SQL & cadCampo
    
    Case Else 'COMPRAS
        'utilizamos sfamia.ctaventa o sfamia.aboventa
        If TotalFac >= 0 Then
            cadCampo = "sfamia.ctacompr"
        Else
            cadCampo = "sfamia.abocompr"
        End If
        
        SQL = "SELECT slifpc.codprove,slifpc.numfactu,slifpc.fecfactu," & cadCampo & " as cuenta, sum(importel) as importe  "
        
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & ",slifpc.codccost"
                
        
        SQL = SQL & " FROM (slifpc  "
        SQL = SQL & " inner join sartic on slifpc.codartic=sartic.codartic) "
        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        
        If vCCos > 0 Then SQL = SQL & ",scafpa "
        
        SQL = SQL & " WHERE "
        
        'si tiene analitica, enlazo por con scafpa
        If vCCos > 0 Then SQL = SQL & " slifpc.NumFactu = scafpa.NumFactu And slifpc.FecFactu = scafpa.FecFactu and slifpc.codprove=scafpa.codprove AND slifpc.numalbar=scafpa.numalbar AND "
            
        SQL = SQL & Replace(cadWHERE, "scafpc", "slifpc")
        SQL = SQL & " GROUP BY "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos = 2 Then SQL = SQL & " codccost, "
                  
        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
        SQL = SQL & cadCampo
        
        
        
    End Select
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    TotImp = 0
    SQLaux = ""
    Aux = ""
    While Not Rs.EOF
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = Rs!Importe - CCur(CalcularPorcentaje(Rs!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(Rs!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        TotImp = TotImp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL2 = ""
        Select Case cadTabla
            Case "scafaccli" 'VENTAS a clientes
                'En aux guardaremos el trozo comun de las lineas (letra/numero/an�o
                If Aux = "" Then Aux = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & ","
                SQL = Aux & i & ","
                SQL = SQL & DBSet(Rs!cuenta, "T")
            Case "scafac"
                'En aux guardaremos el trozo comun de las lineas (letra/numero/an�o
                If Aux = "" Then Aux = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & ","
                SQL = Aux & i & ","
                SQL = SQL & DBSet(Rs!cuenta, "T")
            
            Case Else 'COMPRAS
                'Laura 24/10/2006
                'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
                SQL = numRegis & "," & AnyoFacPr & "," & i & ","
                
    '            If ImpLinea >= 0 Then
                    SQL = SQL & DBSet(Rs!cuenta, "T")
    '            Else
    '                SQL = SQL & DBSet(RS!abocompr, "T")
    '            End If
        End Select
        

        
        SQL2 = SQL & "," 'nos guardamos la linea sin el importe por si a la �ltima hay q descontarle para q coincida con total factura
        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
        
        
        'CENTRO DE COSTE
        LineaCentroCoste = False
        If vCCos = 0 Then
            'NO NECESTIA CENTRO DE COSTE.. seguro
           ' SQL = SQL & ValorNulo
        Else
            LineaCentroCoste = CuentaNecesitaCentroCoste(CStr(Rs!cuenta))
            
        End If
        If LineaCentroCoste Then
            CCoste2 = Rs!CodCCost
            SQL = SQL & DBSet(CCoste2, "T")
        Else
            SQL = SQL & ValorNulo
        End If
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close

    
    '[Monica]16/01/2018: los suplidos van a la cuenta del cliente
    If cadTabla = "scafaccli" Then
        SQL = " SELECT stipom.letraser,slifaccli.codtipom,slifaccli.numfactu,slifaccli.fecfactu,scliente.codmacta as cuenta,sum(importel) as importe"
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & ",slifaccli.codccost"
        
        SQL = SQL & ", sartic.codigiva "
        
        
        SQL = SQL & " FROM (((slifaccli inner join stipom on slifaccli.codtipom=stipom.codtipom) inner join scafaccli on slifaccli.codtipom = scafaccli.codtipom and slifaccli.numfactu = scafaccli.numfactu and slifaccli.fecfactu = scafaccli.fecfactu) "
        SQL = SQL & " inner join scliente on scafaccli.codclien = scliente.codclien) inner join sartic on slifaccli.codartic = sartic.codartic "
        SQL = SQL & " WHERE "
        SQL = SQL & " " & Replace(cadWHERE, "scafaccli", "slifaccli")
        
        '[Monica]15/01/2018: no cogemos lo correspondiente a suplidos
        SQL = SQL & " and slifaccli.codartic = " & DBSet(vParamAplic.ArtSuplidos, "T")
        
        SQL = SQL & " GROUP BY scliente.codmacta "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos > 0 Then SQL = SQL & ", codccost "
        
        Rs.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            ImpLinea = Rs!Importe
            TotImp = TotImp + ImpLinea
        
            'concatenamos linea para insertar en la tabla de conta.linfact
            SQL = ""
            SQL2 = ""
            
            'En aux guardaremos el trozo comun de las lineas (letra/numero/an�o
            If Aux = "" Then Aux = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & ","
            SQL = Aux & i & ","
            SQL = SQL & DBSet(Rs!cuenta, "T")


            SQL2 = SQL & "," 'nos guardamos la linea sin el importe por si a la �ltima hay q descontarle para q coincida con total factura
            SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
            
            
            'CENTRO DE COSTE
            LineaCentroCoste = False
            If vCCos = 0 Then
                'NO NECESTIA CENTRO DE COSTE.. seguro
               ' SQL = SQL & ValorNulo
            Else
                LineaCentroCoste = CuentaNecesitaCentroCoste(CStr(Rs!cuenta))
                
            End If
            If LineaCentroCoste Then
                CCoste2 = Rs!CodCCost
                SQL = SQL & DBSet(CCoste2, "T")
            Else
                SQL = SQL & ValorNulo
            End If
            
            Cad = Cad & "(" & SQL & ")" & ","
            
            i = i + 1
        End If
        Rs.Close
    End If
' hasta aqui
    
    
    
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If TotImp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        TotImp = BaseImp - TotImp
        TotImp = ImpLinea + TotImp '(+- diferencia)
        SQL2 = SQL2 & DBSet(TotImp, "N") & ","
        If CCoste2 = "" Then
            SQL2 = SQL2 & ValorNulo
        Else
            SQL2 = SQL2 & DBSet(CCoste2, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            Cad = SQLaux & "(" & SQL2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & SQL2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If



    'Si lleva retencion, solo sera en caso de facturas proveedores, entonces metere dos lineas mas
    '
    If vLlevaRetencion Then
        'Cojere los datos del proveedor
        'Reutilizo total fac
        TotalFac = CCur(RecuperaValor(DatosRetencion, 2))
        i = i + 1
        SQL = "(" & numRegis & "," & AnyoFacPr & "," & i & ",'" & RecuperaValor(DatosRetencion, 1) & "'," & TransformaComasPuntos(CStr(-TotalFac)) & ",NULL)"
        Cad = Cad & SQL
        i = i + 1
        SQL = ",(" & numRegis & "," & AnyoFacPr & "," & i & ",'" & vParamAplic.CtaReten & "'," & TransformaComasPuntos(CStr(TotalFac)) & ",NULL),"
        Cad = Cad & SQL
        
    End If

    
    
    'Facturas clientes. Ver si lleva aportacion al terminal
    If cadTabla = "scafac" Or cadTabla = "scafaccli" Then
        If DatosAportacion <> "" Then
            
            
            SQL = "(" & Aux & i & ",'" & RecuperaValor(DatosAportacion, 1) & "',"
            'Dejo en DatosAportacion solo el importe
            DatosAportacion = TransformaComasPuntos(RecuperaValor(DatosAportacion, 2))
            SQL = SQL & DatosAportacion & ",NULL),"
            Cad = Cad & SQL
            i = i + 1                                                                                   'Importe en negativo
            SQL = "(" & Aux & i & ",'" & vParamAplic.ctaAportacion & "',-" & DatosAportacion & ",NULL),"
            Cad = Cad & SQL
        
        
        
    
        End If
    End If

    Set Rs = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTabla = "scafaccli" Or cadTabla = "scafac" Then
            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            SQL = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        SQL = SQL & " VALUES " & Cad
        ConnConta.Execute SQL
    End If




EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_NUEVO = False
        cadErr = Err.Description
    Else
        InsertarLinFact_NUEVO = True
    End If
End Function



Private Function InsertarLinFact_NUEVOContaNueva(cadTabla As String, cadWHERE As String, cadErr As String, vLlevaRetencion As Boolean, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SQLaux As String
Dim SQL2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Byte
Dim TotImp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim LineaCentroCoste As Boolean

Dim NumeroIVA As Byte
Dim K As Integer
Dim HayQueAjustar As Boolean

Dim ImpImva As Currency
Dim ImpREC As Currency




    'Puede ser que teniendo analitica, la cuenta no sea del grupo 6 o 7 , con lo cual nodebe poner el CC
    'Por si acaso alguna linea no es del grupo venta o grupo compras, no

    On Error GoTo EInLinea
    '
    '   Habra que ver en funcion de CC que tenga si agrupo, o no, por  codtraba
    '
    Select Case cadTabla
        Case "scafaccli" ' ventas
             'comprobar si el cliente utiliza cuenta alternativa
            If conCtaAlt Then
                'utilizamos sfamia.ctavent1 o sfamia.abovent1
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctavent1"
                Else
                    cadCampo = "sfamia.abovent1" 'si es negativa es un abono
                End If
            Else
                'utilizamos sfamia.ctaventa o sfamia.aboventa
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctaventa"
                Else
                    cadCampo = "sfamia.aboventa"
                End If
            End If
            
            
            SQL = " SELECT stipom.letraser,slifaccli.codtipom,slifaccli.numfactu,slifaccli.fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
            
            'Tiene analitica. Luego el codtraba tiene que aparecer
            If vCCos > 0 Then SQL = SQL & ",slifaccli.codccost"
            
            
            SQL = SQL & ", sartic.codigiva "
            
            SQL = SQL & " FROM ((slifaccli inner join stipom on slifaccli.codtipom=stipom.codtipom) "
            SQL = SQL & " inner join sartic on slifaccli.codartic=sartic.codartic) "
            SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            
            
            SQL = SQL & " WHERE "
            
            
            SQL = SQL & " " & Replace(cadWHERE, "scafaccli", "slifaccli")
            
            '[Monica]15/01/2018: no cogemos lo correspondiente a suplidos
            SQL = SQL & " and slifaccli.codartic <> " & DBSet(vParamAplic.ArtSuplidos, "T")
            
            
            SQL = SQL & " GROUP BY "
            
            'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
            If vCCos > 0 Then SQL = SQL & " codccost, "
                      
            'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
            SQL = SQL & cadCampo & ", sartic.codigiva "
    
    Case "scafac"
             'comprobar si el cliente utiliza cuenta alternativa
            If conCtaAlt Then
                'utilizamos sfamia.ctavent1 o sfamia.abovent1
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctavent1"
                Else
                    cadCampo = "sfamia.abovent1" 'si es negativa es un abono
                End If
            Else
                'utilizamos sfamia.ctaventa o sfamia.aboventa
                If TotalFac >= 0 Then
                    cadCampo = "sfamia.ctaventa"
                Else
                    cadCampo = "sfamia.aboventa"
                End If
            End If
            
            
            SQL = " SELECT stipom.letraser,slifac.codtipom,slifac.numfactu,slifac.fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
            
            'Tiene analitica. Luego el codtraba tiene que aparecer
            If vCCos > 0 Then SQL = SQL & ",slifac.codccost"
            
            SQL = SQL & ", sartic.codigiva "
            
            SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
            SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
            SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            
            
            SQL = SQL & " WHERE "
            
            
            SQL = SQL & " " & Replace(cadWHERE, "scafac", "slifac")
            SQL = SQL & " GROUP BY "
            
            'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
            If vCCos > 0 Then SQL = SQL & " codccost, "
                      
            'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
            SQL = SQL & cadCampo & ", sartic.codigiva "
    
    Case Else 'COMPRAS
        'utilizamos sfamia.ctaventa o sfamia.aboventa
        If TotalFac >= 0 Then
            cadCampo = "sfamia.ctacompr"
        Else
            cadCampo = "sfamia.abocompr"
        End If
        
        SQL = "SELECT slifpc.codprove,slifpc.numfactu,slifpc.fecfactu," & cadCampo & " as cuenta, sum(importel) as importe  "
        
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & ",slifpc.codccost"
                
        SQL = SQL & ", sartic.codigiva "
        
        SQL = SQL & " FROM (slifpc  "
        SQL = SQL & " inner join sartic on slifpc.codartic=sartic.codartic) "
        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        
        If vCCos > 0 Then SQL = SQL & ",scafpa "
        
        SQL = SQL & " WHERE "
        
        'si tiene analitica, enlazo por con scafpa
        If vCCos > 0 Then SQL = SQL & " slifpc.NumFactu = scafpa.NumFactu And slifpc.FecFactu = scafpa.FecFactu and slifpc.codprove=scafpa.codprove AND slifpc.numalbar=scafpa.numalbar AND "
            
        SQL = SQL & Replace(cadWHERE, "scafpc", "slifpc")
        SQL = SQL & " GROUP BY "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos = 2 Then SQL = SQL & " codccost, "
                  
        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
        SQL = SQL & cadCampo & ", sartic.codigiva "
        
        
        
    End Select
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    TotImp = 0
    SQLaux = ""
    Aux = ""
    While Not Rs.EOF
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = Rs!Importe - CCur(CalcularPorcentaje(Rs!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(Rs!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        TotImp = TotImp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL2 = ""
        Select Case cadTabla
            Case "scafaccli" 'VENTAS a clientes
                'En aux guardaremos el trozo comun de las lineas (letra/numero/an�o
                If Aux = "" Then Aux = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & ","
                SQL = Aux & i & ","
                SQL = SQL & DBSet(Rs!cuenta, "T")
            Case "scafac"
                'En aux guardaremos el trozo comun de las lineas (letra/numero/an�o
                If Aux = "" Then Aux = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & ","
                SQL = Aux & i & ","
                SQL = SQL & DBSet(Rs!cuenta, "T")
            
            Case Else 'COMPRAS
                'Laura 24/10/2006
                'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
                SQL = DBSet(SerieFraPro, "T") & "," & numRegis & "," & DBSet(FechaRecepcion, "F") & "," & AnyoFacPr & "," & i & ","
                SQL = SQL & DBSet(Rs!cuenta, "T")
        
        End Select
        
        'Vemos que tipo de IVA es en el vector de importes
        NumeroIVA = 127
        For K = 0 To 2
            If Rs!codigiva = vTipoIva(K) Then
                NumeroIVA = K
                Exit For
            End If
        Next
        If NumeroIVA > 100 Then Err.Raise 513, "Error obteniendo IVA: " & Rs!codigiva

        
        'CENTRO DE COSTE
        LineaCentroCoste = False
        If vCCos = 0 Then
            'NO NECESTIA CENTRO DE COSTE.. seguro
           ' SQL = SQL & ValorNulo
        Else
            LineaCentroCoste = CuentaNecesitaCentroCoste(CStr(Rs!cuenta))
            
        End If
        SQL = SQL & ","
        
        If LineaCentroCoste Then
            CCoste2 = Rs!CodCCost
            SQL = SQL & DBSet(CCoste2, "T")
        Else
            SQL = SQL & ValorNulo
        End If
        
        If cadTabla = "scafac" Or cadTabla = "scafaccli" Then
            SQL = SQL & "," & DBSet(Rs!FecFactu, "F")
        End If
        
        vBaseIva(NumeroIVA) = vBaseIva(NumeroIVA) - ImpLinea   'Para ajustar el importe y que no haya descuadre
        
        'Caluclo el importe de IVA y el de recargo de equivalencia
        ImpImva = vPorcIva(NumeroIVA) / 100
        ImpImva = Round2(ImpLinea * ImpImva, 2)
        If vPorcRec(NumeroIVA) = 0 Then
            ImpREC = 0
        Else
            ImpREC = vPorcRec(NumeroIVA) / 100
            ImpREC = Round2(ImpLinea * ImpREC, 2)
        End If
        vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
        vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpREC
        
        
        HayQueAjustar = False
        If vBaseIva(NumeroIVA) <> 0 Or vImpIva(NumeroIVA) <> 0 Or vImpRec(NumeroIVA) <> 0 Then
            'falta importe.
            'Puede ser que hayan mas lineas, o haya descuadre. Como esta ordenado por tipo de iva
            Rs.MoveNext
            If Rs.EOF Then
                'No hay mas lineas
                'Hay que ajustar SI o SI
                HayQueAjustar = True
            Else
                'Si que hay mas lineas.
                'Son del mismo tipo de IVA
                If Rs!codigiva <> vTipoIva(NumeroIVA) Then
                    'NO es el mismo tipo de IVA
                    'Hay que ajustar
                    HayQueAjustar = True
                End If
            End If
            Rs.MovePrevious
        End If
        
        SQL = SQL & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","
        
        If HayQueAjustar Then
            
            If vBaseIva(NumeroIVA) <> 0 Then ImpLinea = ImpLinea + vBaseIva(NumeroIVA)
            If vImpIva(NumeroIVA) <> 0 Then ImpImva = ImpImva + vImpIva(NumeroIVA)
            If vImpRec(NumeroIVA) <> 0 Then ImpREC = ImpREC + vImpRec(NumeroIVA)
            
        End If

        
        ' baseimpo , impoiva, imporec, aplicret, CodCCost
        SQL = SQL & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S")
        
        ' si la linea lleva retencion
        If cadTabla = "scafac" Or cadTabla = "scafaccli" Then 'VENTAS a clientes
        Else
            SQL = SQL & ",0"
        End If
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close

    '[Monica]16/01/2018: los suplidos los metemos sobre la cuenta del cliente
    If cadTabla = "scafaccli" Then
        
        SQL = " SELECT stipom.letraser,slifaccli.codtipom,slifaccli.numfactu,slifaccli.fecfactu,scliente.codmacta as cuenta,sum(importel) as importe"
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & ",slifaccli.codccost"
        
        SQL = SQL & ", sartic.codigiva "
        
        
        SQL = SQL & " FROM (((slifaccli inner join stipom on slifaccli.codtipom=stipom.codtipom) inner join scafaccli on slifaccli.codtipom = scafaccli.codtipom and slifaccli.numfactu = scafaccli.numfactu and slifaccli.fecfactu = scafaccli.fecfactu) "
        SQL = SQL & " inner join scliente on scafaccli.codclien = scliente.codclien) inner join sartic on slifaccli.codartic = sartic.codartic "
        SQL = SQL & " WHERE "
        SQL = SQL & " " & Replace(cadWHERE, "scafaccli", "slifaccli")
        
        '[Monica]15/01/2018: no cogemos lo correspondiente a suplidos
        SQL = SQL & " and slifaccli.codartic = " & DBSet(vParamAplic.ArtSuplidos, "T")
        
        SQL = SQL & " GROUP BY scliente.codmacta "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos > 0 Then SQL = SQL & ", codccost "
        
        Rs.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
        
            ImpLinea = Rs!Importe
            
            If Aux = "" Then Aux = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & ","
            SQL = Aux & i & ","
            SQL = SQL & DBSet(Rs!cuenta, "T")
            
            
            'CENTRO DE COSTE
            LineaCentroCoste = False
            If vCCos = 0 Then
                'NO NECESTIA CENTRO DE COSTE.. seguro
               ' SQL = SQL & ValorNulo
            Else
                LineaCentroCoste = CuentaNecesitaCentroCoste(CStr(Rs!cuenta))
            End If
            SQL = SQL & ","
            
            If LineaCentroCoste Then
                CCoste2 = Rs!CodCCost
                SQL = SQL & DBSet(CCoste2, "T")
            Else
                SQL = SQL & ValorNulo
            End If
            
            SQL = SQL & "," & DBSet(Rs!FecFactu, "F")
        
            'Calculo el importe de IVA y el de recargo de equivalencia
            ImpImva = 0
            ImpREC = 0
        
            SQL = SQL & "," & DBSet(Rs!codigiva, "N") & ",0," & DBSet(0, "N", "S") & ","
    
            
            ' baseimpo , impoiva, imporec, aplicret, CodCCost
            SQL = SQL & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S")
            
            Cad = Cad & "(" & SQL & ")" & ","
        
        End If
    End If



    
    'Facturas clientes. Ver si lleva aportacion al terminal
    If cadTabla = "scafac" Or cadTabla = "scafaccli" Then
        If DatosAportacion <> "" Then
            
            SQL = "(" & Aux & i & ",'" & RecuperaValor(DatosAportacion, 1) & "',"
            'Dejo en DatosAportacion solo el importe
            DatosAportacion = TransformaComasPuntos(RecuperaValor(DatosAportacion, 2))
            SQL = SQL & DatosAportacion & ",NULL),"
            Cad = Cad & SQL
            i = i + 1                                                                                   'Importe en negativo
            SQL = "(" & Aux & i & ",'" & vParamAplic.ctaAportacion & "',-" & DatosAportacion & ",NULL),"
            Cad = Cad & SQL
    
        End If
    End If

    Set Rs = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTabla = "scafaccli" Or cadTabla = "scafac" Then
            SQL = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec) "
        Else
            SQL = "INSERT INTO factpro_lineas (numserie,numregis,fecharec,anofactu,numlinea,codmacta,codccost,codigiva,porciva,porcrec,baseimpo,impoiva,imporec,aplicret) "
        End If
        SQL = SQL & " VALUES " & Cad
        ConnConta.Execute SQL
    End If


EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_NUEVOContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFact_NUEVOContaNueva = True
    End If
End Function



Private Function ActualizarCabFact(cadTabla As String, cadWHERE As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    
    SQL = "UPDATE " & cadTabla & " SET intconta=1 "
    SQL = SQL & " WHERE " & cadWHERE

    conn.Execute SQL
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFact = False
        cadErr = Err.Description
    Else
        ActualizarCabFact = True
    End If
End Function



'----------------------------------------------------------------------
' FACTURAS PROVEEDOR
'----------------------------------------------------------------------
'Ccoste
'   0: No tendra analitica
'   1: Solo hay un CC que tratar. NO agruparemos por trabajador
'   2: Mas de un CC. Agruparemos por trabajador


'Ahora la retencion puede llevarla CUALQUIERA de las facturas.
'   0. Retencion NORMAL
'   1. Retencion SOCIOS

Public Function PasarFacturaProv(cadWHERE As String, CodCCost As Byte, FechaFin As String, ByRef vContaFra As cContabilizarFacturas) As Boolean

Dim b As Boolean
Dim cadMen As String
Dim SQL As String
Dim Mc As Contadores
Dim vLlevaRetencion As Boolean
Dim i As Integer

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    vLlevaRetencion = False 'Si llevara retencion me lo devolvera la fucion insertar
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactProv(cadWHERE, cadMen, Mc, FechaFin, vLlevaRetencion, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        
        'Veremos que opcion de CC es la que hay que pasar (agrupar o no agrupar)
        vCCos = CodCCost
        '---- Insertar lineas de Factura en la Conta
        If vParamAplic.ContabilidadNueva Then
            b = InsertarLinFact_NUEVOContaNueva("scafpc", cadWHERE, cadMen, vLlevaRetencion, Mc.Contador)
        Else
            b = InsertarLinFact_NUEVO("scafpc", cadWHERE, cadMen, vLlevaRetencion, Mc.Contador)
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            If vParamAplic.ContabilidadNueva Then
                vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
            End If
        End If
        
        If b Then
            '---- Poner intconta=1 en aritaxi.scafac
            b = ActualizarCabFact("scafpc", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
        

        
    End If
    
    
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaProv = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaProv = False

        InsertarTMPErrFac cadMen, cadWHERE
        
        'Si es correcto entonces creo una entrada en tmp para luego listar los resultados de
        'la contabilizacion
         If Mc.Contador > 0 Then
            SQL = "DELETE from tmpinformes where codusu = " & vUsu.Codigo & " AND codigo1= " & Mc.Contador
            conn.Execute SQL
        End If
    
    End If
End Function


Private Function InsertarCabFactProv(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, ByRef LlevaRetencion As Boolean, ByRef vCF As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el a�o de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim SQL As String
Dim SQL2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim TipoOpera As Byte
Dim Aux As String

Dim CadenaInsertFaclin2     As String
Dim ImporAux As Currency
Dim EsFacturaIntracom2 As String


    On Error GoTo EInsertar
       
    
    SQL = SQL & " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,sprove.codmacta,"
    SQL = SQL & "scafpc.dtoppago,scafpc.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    SQL = SQL & "totalfac,tipoiva1,tipoiva2,tipoiva3,tipprove,impret,scafpc.nomprove,scafpc.codprove,tiporet,PorRet,impret "   'Modificacion facturas socios
    SQL = SQL & ", scafpc.nomprove,scafpc.nifprove,scafpc.domprove,scafpc.codpobla,scafpc.pobprove,scafpc.proprove,scafpc.codforpa "
    SQL = SQL & " FROM " & "scafpc "
    SQL = SQL & "INNER JOIN " & "sprove ON scafpc.codprove=sprove.codprove "
    SQL = SQL & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (Rs!FecRecep <= CDate(FechaFin) - 365), True) = 0 Then
        
            vCF.NumeroFactura = Mc.Contador
            vCF.Anofac = Year(DBLet(Rs!FecFactu))
            
            FechaRecepcion = Rs!FecRecep
        
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = Rs!DtoPPago
            DtoGnral = Rs!DtoGnral
            BaseImp = Rs!BaseIVA1 + CCur(DBLet(Rs!BaseIVA2, "N")) + CCur(DBLet(Rs!BaseIVA3, "N"))
            IvaImp = DBLet(Rs!impoiva1, "N") + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
            
            TotalFac = Rs!TotalFac
            AnyoFacPr = Rs!anofacpr
            
            'Para que contabilice las facturas automaticamente
            If vCF.RealizarContabilizacion Then vCF.FijarNumeroFactura Mc.Contador, AnyoFacPr, ""
            
            'SI es facutra socio y tiene retencion
            If Rs!TipoRet = 1 Then 'FACTURA SOCIO, con retencion
                If DBLet(Rs!ImpRet, "N") <> 0 Then
                    'El total factura es totafac+ retencion
                    DatosRetencion = Rs!codmacta & "|" & Rs!ImpRet & "|"
                    TotalFac = TotalFac + Rs!ImpRet  'Luego en las lineas va la resta de este importe
                    LlevaRetencion = True
                Else
                    DatosRetencion = ""
                End If
            End If
            
            Nulo2 = "N"
            Nulo3 = "N"
            If DBLet(Rs!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            
            SQL = ""
            If vParamAplic.ContabilidadNueva Then SQL = SQL & DBSet(SerieFraPro, "T") & ","
            SQL = SQL & Mc.Contador & "," & DBSet(Rs!FecFactu, "F") & "," & Rs!anofacpr & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!NumFactu, "T") & "," & DBSet(Rs!codmacta, "T") & ","
            
            Select Case vParamAplic.ObsFactura
            Case 0
                'Vacio
                SQL = SQL & ValorNulo
            Case 1
                'N� Factura
                SQL = SQL & "'" & DevNombreSQL("S/Fra " & Rs!NumFactu) & "'"
            Case 2
                'Fecha integracion
                SQL = SQL & "'" & Format(Now, FormatoFecha) & "'"
            End Select
            
            If Not vParamAplic.ContabilidadNueva Then
                SQL = SQL & "," & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & ","
                SQL = SQL & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!impoiva2, "N", Nulo2) & "," & DBSet(Rs!impoiva3, "N", Nulo3) & ","
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                'ANTES era dbset de Rs!totalfac, ahora lo haremos de la variabele totalfac
                SQL = SQL & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", Nulo2) & "," & DBSet(Rs!TipoIVA3, "N", Nulo3) & ",0,"
                
                
                'RETENCION.   29 MAYO 2008
                ' retfacpr,trefacpr,cuereten              Las facturas pueden llevar retencion
                Nulo2 = ""
                If Rs!TipoRet = 0 Then
                    If Not IsNull(Rs!PorRet) And Not IsNull(Rs!ImpRet) Then Nulo2 = "O"
                End If
                If Nulo2 = "" Then
                    'NULOS
                    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Else
                    'TIene valor
                    SQL = SQL & DBSet(Rs!PorRet, "N") & "," & DBSet(Rs!ImpRet, "N") & ",'" & vParamAplic.CtaReten & "',"
                End If
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!FecRecep, "F") & ",0"
                
                Cad = Cad & "(" & SQL & ")"
                
                'Insertar en la contabilidad
                SQL = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                SQL = SQL & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                SQL = SQL & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
                SQL = SQL & " VALUES " & Cad
                ConnConta.Execute SQL
                
            Else
            
                SQL = SQL & "," & DBSet(Rs!nomprove, "T") & "," & DBSet(Rs!domprove, "T", "S") & ","
                SQL = SQL & DBSet(Rs!codpobla, "T", "S") & "," & DBSet(Rs!pobprove, "T", "S") & "," & DBSet(Rs!proprove, "T", "S") & ","
                SQL = SQL & DBSet(Rs!nifProve, "F", "S") & ",'ES',"
                SQL = SQL & DBSet(Rs!codforpa, "N") & ","
                
                TipoOpera = 0
                 'IVA ES CERO
                If Rs!tipprove = 1 Then
                    'intracomunitaria
                    TipoOpera = 1
                Else
                    'Exstranjero
                     If Rs!tipprove = 1 Then TipoOpera = 2
                End If
                
                Aux = "0"
                Select Case TipoOpera
                Case 0
                    If Rs!TotalFac < 0 Then
                        Aux = "D"
                    Else
                        If Not IsNull(Rs!TipoIVA2) Then Aux = "C"
                    End If
                
                Case 1
                    Aux = "P"
                
                Case 4
                    Aux = "I"
                End Select
                
                'codopera,codconce340,codintra
                SQL = SQL & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                
                'para las lineas
                'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Aux = "'" & SerieFraPro & "'," & Mc.Contador & "," & DBSet(Rs!FecRecep, "F") & "," & Rs!anofacpr & ","
                
                
                SQL2 = Aux & "1," & DBSet(Rs!BaseIVA1, "N") & "," & Rs!TipoIVA1 & "," & DBSet(Rs!porciva1, "N") & ","
                SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & SQL2 & ")"
                vTipoIva(0) = Rs!TipoIVA1
                vPorcIva(0) = Rs!porciva1
                vPorcRec(0) = 0
                vImpIva(0) = Rs!impoiva1
                vImpRec(0) = 0
                vBaseIva(0) = Rs!BaseIVA1
                
                vTipoIva(1) = 0: vTipoIva(2) = 0
                
                If Not IsNull(Rs!porciva2) Then
                    SQL2 = Aux & "2," & DBSet(Rs!BaseIVA2, "N") & "," & Rs!TipoIVA2 & "," & DBSet(Rs!porciva2, "N") & ","
                    SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                    vTipoIva(1) = Rs!TipoIVA2
                    vPorcIva(1) = Rs!porciva2
                    vPorcRec(1) = 0
                    vImpIva(1) = Rs!impoiva2
                    vImpRec(1) = 0
                    vBaseIva(1) = Rs!BaseIVA2
                
                End If
                If Not IsNull(Rs!porciva3) Then
                    SQL2 = Aux & "3," & DBSet(Rs!BaseIVA3, "N") & "," & Rs!TipoIVA3 & "," & DBSet(Rs!porciva3, "N") & ","
                    SQL2 = SQL2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & SQL2 & ")"
                    vTipoIva(2) = Rs!TipoIVA3
                    vPorcIva(2) = Rs!porciva3
                    vPorcRec(2) = 0
                    vImpIva(2) = Rs!impoiva3
                    vImpRec(2) = 0
                    vBaseIva(2) = Rs!BaseIVA3
                End If
                
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                ImporAux = Rs!BaseIVA1 + DBLet(Rs!BaseIVA2, "N") + DBLet(Rs!BaseIVA3, "N")
                SQL = SQL & DBSet(ImporAux, "N") & "," & ValorNulo & ","
                'totivas
                ImporAux = Rs!impoiva1 + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
                SQL = SQL & DBSet(ImporAux, "N") & "," & DBSet(Rs!TotalFac, "N") & ","
                        
                  
                EsFacturaIntracom2 = ""
                If DBLet(Rs!tipprove, "N") = 1 Then
                    'OK es intracomunitaria
                    EsFacturaIntracom2 = Rs!TipoIVA1
                End If
            
                Nulo2 = ""
                If Rs!TipoRet = 0 Then
                    If Not IsNull(Rs!PorRet) And Not IsNull(Rs!ImpRet) Then Nulo2 = "O"
                End If
                If Nulo2 = "" Then
                    'NULOS
                    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,"
                Else
                    'TIene valor
                    SQL = SQL & DBSet(Rs!PorRet, "N") & "," & DBSet(Rs!ImpRet, "N") & ",'" & vParamAplic.CtaReten & "',1,"
                End If
                
                SQL = SQL & DBSet(Rs!FecRecep, "F")
                
                Cad = Cad & "(" & SQL & ")"
            
                SQL = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,numfactu,codmacta,observa,nommacta,"
                SQL = SQL & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                SQL = SQL & "totbases,totbasesret,totivas,totfacpr,retfacpr , trefacpr, cuereten, tiporeten, fecliqpr)"
                SQL = SQL & " VALUES " & Cad
                ConnConta.Execute SQL
            
            
                'Las  lineas de IVA
                SQL = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                SQL = SQL & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute SQL
            
            End If
            
            
            'Para saber el numreo de registro que le asigna a la factrua
            SQL = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            SQL = SQL & ",'" & DevNombreSQL(Rs!NumFactu) & " @ " & Format(Rs!FecFactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomprove) & "'," & Rs!codProve & ")"
            conn.Execute SQL
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactProv = False
        cadErr = Err.Description
    Else
        InsertarCabFactProv = True
    End If
End Function



Public Sub FechasEjercicioConta(FIni As String, FFin As String)
'Dim RS As ADODB.Recordset
'
'    On Error GoTo EFechas
'
'    FIni = "Select fechaini,fechafin From parametros"
'    Set RS = New ADODB.Recordset
'    RS.Open FIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        FIni = DBLet(RS!FechaIni, "F")
'        FFin = DBLet(RS!FechaFin, "F")
'    End If
'    RS.Close
'    Set RS = Nothing
'
'EFechas:
'    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function InsertarLinFact_TicketsAgrupados(cadTabla As String, cadWHERE As String, cadErr As String, LlevaRetencion As Boolean, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SQLaux As String
Dim SQL2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Byte
Dim TotImp As Currency, ImpLinea As Currency
Dim cadCampo As String


    On Error GoTo EInLinea
    
        
    
            
            
         'comprobar si el cliente utiliza cuenta alternativa
        If conCtaAlt Then
            'utilizamos sfamia.ctavent1 o sfamia.abovent1
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctavent1"
            Else
                cadCampo = "sfamia.abovent1" 'si es negativa es un abono
            End If
        Else
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctaventa"
            Else
                cadCampo = "sfamia.aboventa"
            End If
        End If
        
        
        'Monto el WHERE buscando los tikets que estan asociados a este numfact FTG
        SQLaux = Replace(cadWHERE, "scafac.", "")
        SQLaux = Replace(SQLaux, "numfactu", "numfacftg")
        SQLaux = Replace(SQLaux, "fecfactu", "fecfacftg")
        SQLaux = "select sfactik.* from sfactik ,scafac where sfactik.numfacFTG=scafac.numfactu and sfactik.fecfacftg=scafac.fecfactu AND " & SQLaux
    
    
    
    
        Set Rs = New ADODB.Recordset
        Rs.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = Rs!NumFacftg & " as numfactu ,'" & Format(Rs!FecFacftg, FormatoFecha) & "' as fecfactu,"
        'En aux guardare el codtraba
        Aux = Rs!CodTraba
        SQLaux = ""
        Do
            SQLaux = SQLaux & "," & Rs!NumFactu
            Rs.MoveNext
        Loop Until Rs.EOF
        Rs.Close
        
        
        
        
        
        SQL = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", "FTG", "T")
        SQL = " SELECT '" & SQL & "' as LetraSer,slifac.codtipom," & Cad & cadCampo & " as cuenta,sum(importel) as importe"
        
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & "," & Aux & " as CodTraba"
        
        SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
        SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        
        'David.
        'Lleva anal. Necesitare el trabajador para obtener el CC
        If vCCos > 0 Then SQL = SQL & " ,scafac1 "
        
        SQL = SQL & " WHERE "
        
        'Si lleva analitica
        If vCCos > 0 Then
            'Linkamos la tabla
            SQL = SQL & " slifac.codTipoM = scafac1.codTipoM And slifac.NumFactu = scafac1.NumFactu And slifac.FecFactu = scafac1.FecFactu"
            SQL = SQL & " and slifac.codtipoa=scafac1.codtipoa and slifac.numalbar=scafac1.numalbar AND "
        End If
        

        
                
        
        
        
        SQLaux = Mid(SQLaux, 2)
        SQLaux = "   slifac.codtipom='FTI' AND slifac.numfactu IN (" & SQLaux & ")"
        SQL = SQL & SQLaux
        SQL = SQL & " GROUP BY "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos = 2 Then SQL = SQL & " codtraba, "
                  
        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
        SQL = SQL & cadCampo
        
    
    

    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    TotImp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = Rs!Importe - CCur(CalcularPorcentaje(Rs!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(Rs!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        TotImp = TotImp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL2 = ""
        

        SQL = "'" & Rs!LetraSer & "'," & Rs!NumFactu & "," & Year(Rs!FecFactu) & "," & i & ","
        SQL = SQL & DBSet(Rs!cuenta, "T")
        

        
        SQL2 = SQL & "," 'nos guardamos la linea sin el importe por si a la �ltima hay q descontarle para q coincida con total factura
        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
        
        If vCCos = 0 Then
            SQL = SQL & ValorNulo
        Else
            'Obtendremos el centro de coste a partir del trabajador
            CCoste2 = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", Rs!CodTraba)
            If CCoste2 = "" Then
                cadErr = "ERROR en el centro de coste del trabajador: " & Rs!CodTraba
                'CIerro el rs y salgo por patas
                Rs.Close
                Set Rs = Nothing
    
            End If
            SQL = SQL & DBSet(CCoste2, "T")
        End If
        
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close

    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If TotImp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        TotImp = BaseImp - TotImp
        TotImp = ImpLinea + TotImp '(+- diferencia)
        SQL2 = SQL2 & DBSet(TotImp, "N") & ","
        If CCoste2 = "" Then
            SQL2 = SQL2 & ValorNulo
        Else
            SQL2 = SQL2 & DBSet(CCoste2, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            Cad = SQLaux & "(" & SQL2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & SQL2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If



    'Si lleva retencion, solo sera en caso de facturas proveedores, entonces metere dos lineas mas
    '
    If LlevaRetencion Then
        'Cojere los datos del proveedor
        'Reutilizo total fac
        TotalFac = CCur(RecuperaValor(DatosRetencion, 2))
        i = i + 1
        SQL = "(" & numRegis & "," & AnyoFacPr & "," & i & ",'" & RecuperaValor(DatosRetencion, 1) & "'," & TransformaComasPuntos(CStr(-TotalFac)) & ",NULL)"
        Cad = Cad & SQL
        SQL = ",(" & numRegis & "," & AnyoFacPr & "," & i + 1 & ",'" & vParamAplic.CtaReten & "'," & TransformaComasPuntos(CStr(TotalFac)) & ",NULL),"
        Cad = Cad & SQL
    End If





    Set Rs = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        SQL = SQL & " VALUES " & Cad
        ConnConta.Execute SQL
    End If




EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_TicketsAgrupados = False
        cadErr = Err.Description
    Else
        InsertarLinFact_TicketsAgrupados = True
    End If
End Function







'=============================================================================
'==========     CENTROS DE COSTE
'=============================================================================
'LAURA
Public Function PonerNombreCCoste(ByRef txt As TextBox) As String
'Obtener el nombre de un centro de coste
Dim codCCoste As String
Dim Cad As String

    If txt.Text = "" Then
         PonerNombreCCoste = ""
         Exit Function
    End If
    
    codCCoste = Trim(txt.Text)
    
    Cad = DevuelveDesdeBDNew(conConta, "cabccost", "nomccost", "codccost", codCCoste, "T")
    If Cad = "" Then
        If Not txt.Locked Then MsgBox "No existe el Centro de coste : " & codCCoste, vbExclamation
        PonerNombreCCoste = ""
        txt.Text = ""
    Else
        txt.Text = codCCoste
        PonerNombreCCoste = Cad
    End If
    
End Function

'=============================================================================
'==========     CONCEPTOS
'=============================================================================
'LAURA
Public Function PonerNombreConcepto(ByRef txt As TextBox) As String
'Obtener el nombre de un concepto
Dim codConce As String
Dim Cad As String

     If txt.Text = "" Then
         PonerNombreConcepto = ""
         Exit Function
    End If
    codConce = txt.Text
    If ConceptoCorrecto(codConce, Cad) Then
        txt.Text = Format(codConce, "000")
        PonerNombreConcepto = Cad
    Else
        MsgBox Cad, vbExclamation
        txt.Text = ""
        PonerNombreConcepto = ""
        PonerFoco txt
    End If
End Function



'LAURA
Public Function ConceptoCorrecto(ByRef Concep As String, ByRef devuelve As String) As Boolean
    Dim SQL As String
    
    ConceptoCorrecto = False
 
    'BD 2: conexion a BD Conta
    SQL = DevuelveDesdeBDNew(conConta, "conceptos", "nomconce", "codconce", Concep, "N")
    If SQL = "" Then
        devuelve = "No existe el concepto : " & Concep
        Exit Function
    Else
        devuelve = SQL
        ConceptoCorrecto = True
    End If
End Function




Private Function CuentaNecesitaCentroCoste(cta As String) As Boolean
Dim i As Integer
Dim C As String
    
    CuentaNecesitaCentroCoste = False
    
    'vEmpresa.RaizAnalitica    lleva: gripo gasto |grupo vta| otros grupo
    For i = 1 To 3
        C = RecuperaValor(vEmpresa.RaizAnalitica, i)
        If i < 3 Then
            'UN DIGITO
            If Mid(cta, 1, 1) = C Then
                CuentaNecesitaCentroCoste = True
                Exit Function
            End If
        Else
            'Subgrupo a tres digitos
            If Mid(cta, 1, 3) = C Then
                CuentaNecesitaCentroCoste = True
                Exit Function
            End If
        End If
    Next i
End Function







































































'************************************************************************************
'************************************************************************************
'************************************************************************************
'************************************************************************************
'
'
'ANTES de cambiar el tema de los centros de coste
'
'
'
'
'************************************************************************************
'
'''''''''''Private Function InsertarLinFact_new(cadTabla As String, cadWhere As String, cadErr As String, vLlevaRetencion As Boolean, Optional numRegis As Long) As Boolean
''''''''''''cadWHere: selecciona un registro de scafac
''''''''''''codtipom=x and numfactu=y and fecfactu=z
'''''''''''Dim SQL As String
'''''''''''Dim SQLaux As String
'''''''''''Dim SQL2 As String
'''''''''''Dim RS As ADODB.Recordset
'''''''''''Dim Cad As String, Aux As String
'''''''''''Dim I As Byte
'''''''''''Dim TotImp As Currency, ImpLinea As Currency
'''''''''''Dim cadCampo As String
'''''''''''
'''''''''''
'''''''''''    On Error GoTo EInLinea
'''''''''''
'''''''''''
'''''''''''    '
'''''''''''    '   Habra que ver en funcion de CC que tenga si agrupo, o no, por  codtraba
'''''''''''    '
'''''''''''    If cadTabla = "scafac" Then 'VENTAS
'''''''''''         'comprobar si el cliente utiliza cuenta alternativa
'''''''''''        If conCtaAlt Then
'''''''''''            'utilizamos sfamia.ctavent1 o sfamia.abovent1
'''''''''''            If TotalFac >= 0 Then
'''''''''''                cadCampo = "sfamia.ctavent1"
'''''''''''            Else
'''''''''''                cadCampo = "sfamia.abovent1" 'si es negativa es un abono
'''''''''''            End If
'''''''''''        Else
'''''''''''            'utilizamos sfamia.ctaventa o sfamia.aboventa
'''''''''''            If TotalFac >= 0 Then
'''''''''''                cadCampo = "sfamia.ctaventa"
'''''''''''            Else
'''''''''''                cadCampo = "sfamia.aboventa"
'''''''''''            End If
'''''''''''        End If
'''''''''''
'''''''''''
'''''''''''        SQL = " SELECT stipom.letraser,slifac.codtipom,slifac.numfactu,slifac.fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
'''''''''''
'''''''''''        'Tiene analitica. Luego el codtraba tiene que aparecer
'''''''''''        If vCCos > 0 Then SQL = SQL & ",CodTraba"
'''''''''''
'''''''''''        SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
'''''''''''        SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
'''''''''''        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
'''''''''''
'''''''''''        'David.
'''''''''''        'Lleva anal. Necesitare el trabajador para obtener el CC
'''''''''''        If vCCos > 0 Then SQL = SQL & " ,scafac1 "
'''''''''''
'''''''''''        SQL = SQL & " WHERE "
'''''''''''
'''''''''''        'Si lleva analitica
'''''''''''        If vCCos > 0 Then
'''''''''''            'Linkamos la tabla
'''''''''''            SQL = SQL & " slifac.codTipoM = scafac1.codTipoM And slifac.NumFactu = scafac1.NumFactu And slifac.FecFactu = scafac1.FecFactu"
'''''''''''            SQL = SQL & " and slifac.codtipoa=scafac1.codtipoa and slifac.numalbar=scafac1.numalbar AND "
'''''''''''        End If
'''''''''''
'''''''''''        SQL = SQL & " " & Replace(cadWhere, "scafac", "slifac")
'''''''''''        SQL = SQL & " GROUP BY "
'''''''''''
'''''''''''        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
'''''''''''        If vCCos = 2 Then SQL = SQL & " codtraba, "
'''''''''''
'''''''''''        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
'''''''''''        SQL = SQL & cadCampo
'''''''''''
'''''''''''    Else 'COMPRAS
'''''''''''        'utilizamos sfamia.ctaventa o sfamia.aboventa
'''''''''''        If TotalFac >= 0 Then
'''''''''''            cadCampo = "sfamia.ctacompr"
'''''''''''        Else
'''''''''''            cadCampo = "sfamia.abocompr"
'''''''''''        End If
'''''''''''
'''''''''''        SQL = "SELECT slifpc.codprove,slifpc.numfactu,slifpc.fecfactu," & cadCampo & " as cuenta, sum(importel) as importe  "
'''''''''''
'''''''''''        'Tiene analitica. Luego el codtraba tiene que aparecer
'''''''''''        If vCCos > 0 Then SQL = SQL & ",CodTrab2 as codtraba"
'''''''''''
'''''''''''
'''''''''''        SQL = SQL & " FROM (slifpc  "
'''''''''''        SQL = SQL & " inner join sartic on slifpc.codartic=sartic.codartic) "
'''''''''''        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
'''''''''''
'''''''''''        If vCCos > 0 Then SQL = SQL & ",scafpa "
'''''''''''
'''''''''''        SQL = SQL & " WHERE "
'''''''''''
'''''''''''        'si tiene analitica, enlazo por con scafpa
'''''''''''        If vCCos > 0 Then SQL = SQL & " slifpc.NumFactu = scafpa.NumFactu And slifpc.FecFactu = scafpa.FecFactu and slifpc.codprove=scafpa.codprove AND slifpc.numalbar=scafpa.numalbar AND "
'''''''''''
'''''''''''        SQL = SQL & Replace(cadWhere, "scafpc", "slifpc")
'''''''''''        SQL = SQL & " GROUP BY "
'''''''''''
'''''''''''        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
'''''''''''        If vCCos = 2 Then SQL = SQL & " codtraba, "
'''''''''''
'''''''''''        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
'''''''''''        SQL = SQL & cadCampo
'''''''''''
'''''''''''
'''''''''''
'''''''''''    End If
'''''''''''
'''''''''''
'''''''''''    Set RS = New ADODB.Recordset
'''''''''''    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'''''''''''
'''''''''''    Cad = ""
'''''''''''    I = 1
'''''''''''    TotImp = 0
'''''''''''    SQLaux = ""
'''''''''''    Aux = ""
'''''''''''    While Not RS.EOF
'''''''''''        SQLaux = Cad
'''''''''''        'calculamos la Base Imp del total del importe para cada cta cble ventas
'''''''''''        '---- Laura: 10/10/2006
'''''''''''        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
'''''''''''        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
'''''''''''        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
'''''''''''        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
'''''''''''        'ImpLinea = Round(ImpLinea, 2)
'''''''''''        '----
'''''''''''        TotImp = TotImp + ImpLinea
'''''''''''
'''''''''''        'concatenamos linea para insertar en la tabla de conta.linfact
'''''''''''        SQL = ""
'''''''''''        SQL2 = ""
'''''''''''
'''''''''''        If cadTabla = "scafac" Then 'VENTAS a clientes
'''''''''''            'En aux guardaremos el trozo comun de las lineas (letra/numero/an�o
'''''''''''            If Aux = "" Then Aux = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & Year(RS!FecFactu) & ","
'''''''''''            SQL = Aux & I & ","
'''''''''''            SQL = SQL & DBSet(RS!Cuenta, "T")
'''''''''''
'''''''''''        Else 'COMPRAS
'''''''''''            'Laura 24/10/2006
'''''''''''            'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
'''''''''''            SQL = numRegis & "," & AnyoFacPr & "," & I & ","
'''''''''''
''''''''''''            If ImpLinea >= 0 Then
'''''''''''                SQL = SQL & DBSet(RS!Cuenta, "T")
''''''''''''            Else
''''''''''''                SQL = SQL & DBSet(RS!abocompr, "T")
''''''''''''            End If
'''''''''''        End If
'''''''''''
'''''''''''
'''''''''''
'''''''''''        SQL2 = SQL & "," 'nos guardamos la linea sin el importe por si a la �ltima hay q descontarle para q coincida con total factura
'''''''''''        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
'''''''''''
'''''''''''        If vCCos = 0 Then
'''''''''''            SQL = SQL & ValorNulo
'''''''''''        Else
'''''''''''            'Obtendremos el centro de coste a partir del trabajador
'''''''''''            CCoste2 = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", RS!CodTraba)
'''''''''''            If CCoste2 = "" Then
'''''''''''                cadErr = "ERROR en el centro de coste del trabajador: " & RS!CodTraba
'''''''''''                'CIerro el rs y salgo por patas
'''''''''''                RS.Close
'''''''''''                Set RS = Nothing
'''''''''''
'''''''''''            End If
'''''''''''            SQL = SQL & DBSet(CCoste2, "T")
'''''''''''        End If
'''''''''''
''''''''''''        If CCoste = "" Then
''''''''''''            SQL = SQL & ValorNulo
''''''''''''        Else
''''''''''''            SQL = SQL & DBSet(CCoste, "T")
''''''''''''        End If
'''''''''''
'''''''''''        Cad = Cad & "(" & SQL & ")" & ","
'''''''''''
'''''''''''        I = I + 1
'''''''''''        RS.MoveNext
'''''''''''    Wend
'''''''''''    RS.Close
'''''''''''
'''''''''''
'''''''''''    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
'''''''''''    'de la factura
'''''''''''    If TotImp <> BaseImp Then
''''''''''''        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
'''''''''''        'en SQL esta la ult linea introducida
'''''''''''        TotImp = BaseImp - TotImp
'''''''''''        TotImp = ImpLinea + TotImp '(+- diferencia)
'''''''''''        SQL2 = SQL2 & DBSet(TotImp, "N") & ","
'''''''''''        If CCoste2 = "" Then
'''''''''''            SQL2 = SQL2 & ValorNulo
'''''''''''        Else
'''''''''''            SQL2 = SQL2 & DBSet(CCoste2, "T")
'''''''''''        End If
'''''''''''        If SQLaux <> "" Then 'hay mas de una linea
'''''''''''            Cad = SQLaux & "(" & SQL2 & ")" & ","
'''''''''''        Else 'solo una linea
'''''''''''            Cad = "(" & SQL2 & ")" & ","
'''''''''''        End If
'''''''''''
''''''''''''        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
''''''''''''        cad = Replace(cad, SQL, Aux)
'''''''''''    End If
'''''''''''
'''''''''''
'''''''''''
'''''''''''    'Si lleva retencion, solo sera en caso de facturas proveedores, entonces metere dos lineas mas
'''''''''''    '
'''''''''''    If vLlevaRetencion Then
'''''''''''        'Cojere los datos del proveedor
'''''''''''        'Reutilizo total fac
'''''''''''        TotalFac = CCur(RecuperaValor(DatosRetencion, 2))
'''''''''''        I = I + 1
'''''''''''        SQL = "(" & numRegis & "," & AnyoFacPr & "," & I & ",'" & RecuperaValor(DatosRetencion, 1) & "'," & TransformaComasPuntos(CStr(-TotalFac)) & ",NULL)"
'''''''''''        Cad = Cad & SQL
'''''''''''        I = I + 1
'''''''''''        SQL = ",(" & numRegis & "," & AnyoFacPr & "," & I & ",'" & vParamAplic.CtaReten & "'," & TransformaComasPuntos(CStr(TotalFac)) & ",NULL),"
'''''''''''        Cad = Cad & SQL
'''''''''''
'''''''''''    End If
'''''''''''
'''''''''''
'''''''''''
'''''''''''
'''''''''''    'Facturas clientes. Ver si lleva aportacion al terminal
'''''''''''    If cadTabla = "scafac" Then
'''''''''''        If DatosAportacion <> "" Then
'''''''''''
'''''''''''
'''''''''''            SQL = "(" & Aux & I & ",'" & RecuperaValor(DatosAportacion, 1) & "',"
'''''''''''            'Dejo en DatosAportacion solo el importe
'''''''''''            DatosAportacion = TransformaComasPuntos(RecuperaValor(DatosAportacion, 2))
'''''''''''            SQL = SQL & DatosAportacion & ",NULL),"
'''''''''''            Cad = Cad & SQL
'''''''''''            I = I + 1                                                                                   'Importe en negativo
'''''''''''            SQL = "(" & Aux & I & ",'" & vParamAplic.ctaAportacion & "',-" & DatosAportacion & ",NULL),"
'''''''''''            Cad = Cad & SQL
'''''''''''
'''''''''''
'''''''''''
'''''''''''
'''''''''''        End If
'''''''''''    End If
'''''''''''
'''''''''''    Set RS = Nothing
'''''''''''
'''''''''''    'Insertar en la contabilidad
'''''''''''    If Cad <> "" Then
'''''''''''        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
'''''''''''        If cadTabla = "scafac" Then
'''''''''''            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
'''''''''''        Else
'''''''''''            SQL = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
'''''''''''        End If
'''''''''''        SQL = SQL & " VALUES " & Cad
'''''''''''        ConnConta.Execute SQL
'''''''''''    End If
'''''''''''
'''''''''''
'''''''''''
'''''''''''
'''''''''''EInLinea:
'''''''''''    If Err.Number <> 0 Then
'''''''''''        InsertarLinFact_new = False
'''''''''''        cadErr = Err.Description
'''''''''''    Else
'''''''''''        InsertarLinFact_new = True
'''''''''''    End If
'''''''''''End Function




