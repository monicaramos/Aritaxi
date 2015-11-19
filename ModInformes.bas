Attribute VB_Name = "ModInformes"
Option Explicit


'============================================================
'====== FUNCIONES GENERALES  ================================


Public Sub AbrirListado(numero As Integer)
    Screen.MousePointer = vbHourglass
    frmListado.OpcionListado = numero
    If numero = 223 Then
        frmListado.OptClientes.Value = True
        frmListado.FrameContab.Enabled = False
    End If
    If numero = 224 Then
        frmListado.FrameContab.Enabled = False
    End If
    frmListado.Show vbModal
    Screen.MousePointer = vbDefault
End Sub


Public Function AnyadirAFormula(ByRef cadFormula As String, arg As String) As Boolean
'Concatena los criterios del WHERE para pasarlos al Crystal como FormulaSelection
    If arg = "Error" Then
        AnyadirAFormula = False
        Exit Function
    ElseIf arg <> "" Then
        If cadFormula <> "" Then
            cadFormula = cadFormula & " AND (" & arg & ")"
        Else
            cadFormula = arg
        End If
    End If
    AnyadirAFormula = True
End Function

Public Function AnyadirAFormulaOr(ByRef cadFormula As String, arg As String) As Boolean
'Concatena los criterios del WHERE para pasarlos al Crystal como FormulaSelection
' La modificaci�n es que la concatenaci�n de criterios se hace con OR si se utiliza esta
' funci�n [SERVICIOS]
    If arg = "Error" Then
        AnyadirAFormulaOr = False
        Exit Function
    ElseIf arg <> "" Then
        If cadFormula <> "" Then
            cadFormula = cadFormula & " OR (" & arg & ")"
        Else
            cadFormula = arg
        End If
    End If
    AnyadirAFormulaOr = True
End Function

Public Function NumRegistros(vSQL As String, Optional vBD As Byte) As Integer
'Devuelve si hay registros con la seleccion
'realizada. Si no hay nada que mostrar devuelve 0
Dim RS As ADODB.Recordset

    On Error Resume Next

    Set RS = New ADODB.Recordset
    If vBD = conConta Then
        RS.Open vSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        RS.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    
    NumRegistros = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then
            NumRegistros = RS.Fields(0).Value
'            If RS.Fields(0).Value = 1 Then
'                RegistrosAListar = 1  'Solo es para saber que hay registros que mostrar
'            Else
'                RegistrosAListar = 2  'Solo es para saber que hay registros que mostrar
'            End If
        End If
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        NumRegistros = 0
        Err.Clear
    End If
End Function



Public Function RegistrosAListar(vSQL As String, Optional vBD As Byte) As Byte
'Devuelve si hay algun registro para mostrar en el Informe con la seleccion
'realizada. Si no hay nada que mostrar devuelve 0 y no abrir� el informe
Dim RS As ADODB.Recordset

    On Error GoTo ErrReg

    Set RS = New ADODB.Recordset
    If vBD = conConta Then
        RS.Open vSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        RS.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    
    RegistrosAListar = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then
            If RS.Fields(0).Value = 1 Then
                RegistrosAListar = 1  'Solo es para saber que hay registros que mostrar
            Else
                RegistrosAListar = 2  'Solo es para saber que hay registros que mostrar
            End If
        End If
    End If
    RS.Close
    Set RS = Nothing

    Exit Function
    
ErrReg:
    RegistrosAListar = 0
    MuestraError Err.Number, "Comprobar si hay registros seleccionados", Err.Description
End Function



'Para que no muestre el mensaje de NO hay datos
'   optional: por defecto FALSE
Public Function HayRegParaInforme(cTabla As String, cWhere As String, Optional OcultarMsg As Boolean) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        cWhere = Replace(cWhere, "[", "(")
        cWhere = Replace(cWhere, "]", ")")
        SQL = SQL & " WHERE " & cWhere
    End If
    If RegistrosAListar(SQL) = 0 Then
        'Por defecto SI que lo muestra
        If Not OcultarMsg Then MsgBox "No hay datos para mostrar.", vbInformation
        HayRegParaInforme = False
    Else
        HayRegParaInforme = True
    End If
End Function


Public Function CadenaDesdeHasta(cadDesde As String, cadHasta As String, campo As String, TipoCampo As String, Optional NomCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= cadDesde and campo<=cadHasta) "
'para Crystal Report
Dim cadAux As String
On Error GoTo ErrDH

    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            cadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    cadAux = campo & " >= " & Val(cadDesde)
                Case "T"
                    cadAux = campo & " >= """ & cadDesde & """"
                Case "F"
                    cadAux = campo & " >= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
                Case "FH"
                    cadAux = campo & " >= DateTime(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & "," & Hour(cadDesde) & "," & Minute(cadDesde) & "," & Second(cadDesde) & ")"
                    
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If cadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If cadDesde > cadHasta Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                        End If
                        
                    Case "FH"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                                   
                            cadAux = cadAux & " AND " & campo & " <= DateTime(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ","
                            If Len(cadHasta) = 10 Then
                                cadAux = cadAux & "23,59,59"
                            Else
                                cadAux = cadAux & Hour(cadHasta) & "," & Minute(cadHasta) & "," & Second(cadHasta)
                            End If
                            cadAux = cadAux & ")"
                        End If
                End Select
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        cadAux = campo & " <= " & Val(cadHasta)
                    Case "T"
                        cadAux = campo & " <= """ & cadHasta & """"
                    Case "F"
                        cadAux = campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                    Case "FH"
                            cadAux = campo & " <= DateTime(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ","
                            If Len(cadHasta) = 10 Then
                                cadAux = cadAux & "23,59,59"
                            Else
                                cadAux = cadAux & Hour(cadHasta) & "," & Minute(cadHasta) & "," & Second(cadHasta)
                            End If
                            cadAux = cadAux & ")"
                        
                End Select
            End If
        End If
    End If
    If cadAux <> "" And cadAux <> "Error" Then cadAux = "(" & cadAux & ")"
    CadenaDesdeHasta = cadAux
ErrDH:
    If Err.Number <> 0 Then CadenaDesdeHasta = "Error"
End Function


Public Function CadenaDesdeHastaBD(cadDesde As String, cadHasta As String, campo As String, TipoCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= valor1 and campo<=valor2) "
'Para MySQL
Dim cadAux As String

    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            cadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    cadAux = campo & " >= " & Val(cadDesde)
                Case "T"
                    cadAux = campo & " >= """ & cadDesde & """"
                Case "F"
                    cadAux = "(" & campo & " >= '" & Format(cadDesde, FormatoFecha) & "')"
                Case "FH"
                    If Len(cadDesde) = 10 Then cadDesde = cadDesde & " 00:00:00"
                    cadAux = "(" & campo & " >= '" & Format(cadDesde, FormatoFechaHora) & "')"
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If cadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and (" & campo & " <= '" & Format(cadHasta, FormatoFecha) & "')"
                        End If
                    Case "FH"
                        If Len(cadHasta) = 10 Then cadHasta = cadHasta & " 23:59:59"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " AND (" & campo & " <= '" & Format(cadHasta, FormatoFechaHora) & "')"
                        End If

                    

                End Select
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        cadAux = campo & " <= " & Val(cadHasta)
                    Case "T"
                        cadAux = campo & " <= """ & cadHasta & """"
                    Case "F"
                        cadAux = campo & " <= '" & Format(cadHasta, FormatoFecha) & "'"
                End Select
            End If
        End If
    End If
    If cadAux <> "" And cadAux <> "Error" Then cadAux = "(" & cadAux & ")"
    CadenaDesdeHastaBD = cadAux
End Function



Public Function QuitarCaracterACadena(cadForm As String, Caracter As String) As String
'IN: [cadForm] es la cadena en la que se eliminara todos los caractes iguales a la vble [Caracter]
'OUT: cadena sin los caracteres
Dim i As Long
Dim J As Long
Dim Aux As String

    Aux = cadForm
    i = InStr(1, Aux, Caracter, vbTextCompare)
    While i > 0
        i = InStr(1, Aux, Caracter, vbTextCompare)
        If i > 0 Then
            J = Len(Caracter)
            Aux = Mid(Aux, 1, i - 1) & Mid(Aux, i + J, Len(Aux) - 1)
        End If
    Wend
    QuitarCaracterACadena = Aux
End Function


Public Sub PonerFrameVisible(ByRef vFrame As Frame, visible As Boolean, H As Integer, W As Integer)
'Pone el Frame Visible y Ajustado al Formulario, y visualiza los controles

    vFrame.visible = visible
    If visible = True Then
        'Ajustar Tama�o del Frame para ajustar tama�o de Formulario al del Frame
        vFrame.Top = -90
        vFrame.Left = 0
        vFrame.Width = W
        vFrame.Height = H
    End If
End Sub



'Public Function SustituirCadenas(CADENA As String, cad1 As String, cad2 As String) As String
''IN: Cadena es la cadena de texto en la que se va a sustituir la cad1 por la cad2
''OUT: cadena con la sustitucion
'
''EJEMPLO: cadena = "scaalb.codtipom='ALV' AND scaalb.numalbar=1"
''         cad1 = "scaalb"
''         cad2 = "slialb"
'
''         Resultado = "slialb.codtipom='ALV' AND slialb.numalbar=1"
'
'Dim i As Integer
'Dim J As Integer
'Dim Aux As String
'
'    Aux = CADENA
'    Do
'        i = InStr(1, Aux, cad1, vbTextCompare)
'        If i > 0 Then
'            J = Len(cad1)
'            Aux = Mid(Aux, 1, i - 1) & cad2 & Mid(Aux, i + J, Len(Aux) - 1)
'        End If
'    Loop Until i <= 0
'    SustituirCadenas = Aux
'End Function



'============================================================
'====== FUNCIONES PARA ARITAXI  ==============================

Public Sub AbrirListadoOfer(numero As Integer)
'Abre el Form con los listados de Ofertas
    Screen.MousePointer = vbHourglass
    frmListadoOfer.OpcionListado = numero
    frmListadoOfer.Show vbModal
    Screen.MousePointer = vbDefault
End Sub


Public Sub AbrirListadoPed(numero As Integer)
'Abre el Form con los listados de Pedidos
    Screen.MousePointer = vbHourglass
    frmListadoPed.OpcionListado = numero
    frmListadoPed.Show vbModal
    Screen.MousePointer = vbDefault
End Sub



Public Function PonerParamEmpresa(cadParam As String, numParam As Byte) As Boolean
Dim DomiEmp As String
Dim WebEmp As String
Dim cad As String

        DomiEmp = vParam.DomicilioEmpresa & " - " & vParam.CPostal & " " & vParam.Poblacion
        If vParam.Provincia <> vParam.Poblacion Then DomiEmp = DomiEmp & " " & vParam.Provincia
        DomiEmp = DomiEmp & " - Telf. " & vParam.Telefono & " - Fax. " & vParam.Fax
        WebEmp = "Internet: " & vParam.WebEmpresa & " - E-mail: " & vParam.MailEmpresa
        'Resto parametros
        cad = ""
        cad = cad & "pNomEmpre=""" & vParam.NombreEmpresa & """|"
        cad = cad & "pDomEmpre=""" & DomiEmp & """|"
        cad = cad & "pWebEmpre=""" & WebEmp & """|"
        
        numParam = numParam + 3
        cadParam = cadParam & cad
        PonerParamEmpresa = True
End Function


Public Function PonerParamRPT(Indice As Byte, cadParam As String, numParam As Byte, nomDocu As String, ByRef ImpresionDirecta As Boolean, NomPDF As String) As Boolean
Dim vParamRpt As CParamRpt 'Tipos de Documentos
Dim cad As String

    Set vParamRpt = New CParamRpt
    
    NomPDF = ""  'Reestablezco
    ImpresionDirecta = False 'psi acaso
    
    If vParamRpt.Leer(Indice) = 1 Then
        cad = "No se han podido cargar los Par�metros de Tipos de Documentos." & vbCrLf
        MsgBox cad & "Debe configurar la aplicaci�n.", vbExclamation
        Set vParamRpt = Nothing
        PonerParamRPT = False
        Exit Function
    Else
        If cadParam = "" Then
            cad = "|"
        Else
            cad = ""
        End If
        cad = cad & "pCodigoISO=""" & vParamRpt.CodigoISO & """|"
        If vParamRpt.CodigoRevision = -1 Then
            cad = cad & "pCodigoRev=""" & "" & """|"
        Else
            cad = cad & "pCodigoRev=""" & Format(vParamRpt.CodigoRevision, "00") & """|"
        End If
        numParam = numParam + 2
        If vParamRpt.LineaPie1 <> "" Then
            cad = cad & "pLinea1=""" & vParamRpt.LineaPie1 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie2 <> "" Then
            cad = cad & "pLinea2=""" & vParamRpt.LineaPie2 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie3 <> "" Then
            cad = cad & "pLinea3=""" & vParamRpt.LineaPie3 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie4 <> "" Then
            cad = cad & "pLinea4=""" & vParamRpt.LineaPie4 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie5 <> "" Then
            cad = cad & "pLinea5=""" & vParamRpt.LineaPie5 & """|"
            numParam = numParam + 1
        End If
        cadParam = cadParam & cad
        nomDocu = vParamRpt.Documento
        NomPDF = vParamRpt.PDFrpt
        ImpresionDirecta = vParamRpt.ImprimeDirecto
        PonerParamRPT = True
        Set vParamRpt = Nothing
    End If
End Function


Public Sub PonerParamCadOferta(cadParam As String, numParam As Byte, cadSelect As String)
'Concatena los N� de Ofertas que se van a imprimir, y lo a�ade como parametro
' a los parametros que se pasaran al Report.
'A�ade el parametro: pCadOfertas= 0000001, 0000002, ...
'RPT que lo utiliza: AriOfertas.rpt
Dim cadOfertas As String
Dim SQL As String
Dim i As Byte
Dim RS As ADODB.Recordset

    On Error GoTo EPonParam
    
    cadOfertas = ""
    SQL = "scapre"

    i = InStr(1, cadSelect, "scapre")
    If Not (i > 0) Then SQL = "schpre"

    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")

    SQL = "SELECT distinct numofert from  " & SQL
    SQL = SQL & " WHERE " & cadSelect
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    While Not RS.EOF
        If Len(cadOfertas) > 75 Then
            If InStr(cadOfertas, "...") > 0 Then
                RS.MoveNext
            Else
                cadOfertas = cadOfertas & ", ..."
            End If
            
        Else
            If cadOfertas = "" Then
                cadOfertas = Format(RS.Fields(0).Value, "0000000")
            Else
                cadOfertas = cadOfertas & ", " & Format(RS.Fields(0).Value, "0000000")
            End If
            RS.MoveNext
        End If
    Wend
    RS.Close
    Set RS = Nothing

    cadParam = cadParam & "pCadOfertas=""" & cadOfertas & """|"
    numParam = numParam + 1
    
EPonParam:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo param�tros del informe.", Err.Description
End Sub





Public Function PonerNombreImpresora() As String
On Error Resume Next
    PonerNombreImpresora = Printer.DeviceName
    If Err.Number <> 0 Then
        PonerNombreImpresora = "No hay impresora instalada"
        Err.Clear
    End If
End Function


Public Sub EstablecerImpresora(Nombre As String)
Dim X As Printer
    For Each X In Printers
       If X.DeviceName = Nombre Then
          ' La define como predeterminada del sistema.
          Set Printer = X
          ' Sale del bucle.
          Exit For
       End If
    Next

End Sub
  


Public Function NombreImpresoraTicket(nTermi As Integer) As String
    On Error GoTo ErrNomImp
    
    If vParamTPV Is Nothing Then
    
        'Establecemos la impresora de ticket
        Set vParamTPV = New CParamTPV
        If vParamTPV.Leer2(CStr(nTermi)) = 0 Then
             NombreImpresoraTicket = vParamTPV.NomImpresora
    '        If vParamTPV.NomImpresora <> "" Then
    
    '            If Printer.DeviceName <> vParamTPV.NomImpresora Then
    '                NomImpre = Printer.DeviceName
    '                EstablecerImpresora vParamTPV.NomImpresora
    '            End If
    '        End If
        End If
        Set vParamTPV = Nothing
    Else
        NombreImpresoraTicket = vParamTPV.NomImpresora
    End If
    
    Exit Function
ErrNomImp:
    MuestraError Err.Number, "Obtener nombre impresora de Ticket", Err.Description
End Function



Public Function ObtenerTerminal() As Integer
Dim SQL As String
    
    On Error GoTo ErrTermi

    'Obtener que terminal es
    'Terminal con el que trabajaremos, leemos el nombre del ordenador
    SQL = ComputerName 'Nombre PC conectado por Terminal Server / local
    SQL = DevuelveDesdeBDNew(conAri, "spatpvt", "numtermi", "nombrepc", SQL, "T")
    If Not IsNumeric(SQL) Then
        MsgBox "No se ha podido establecer la impresora de ticket." & vbCrLf & "Debe configurar primero los par�metros del TPV.", vbExclamation
'    Else
'        bImpre = True
    End If
    ObtenerTerminal = CInt(SQL)
    Exit Function
    
ErrTermi:
    MuestraError Err.Number, "Obtener terminal.", Err.Description
    ObtenerTerminal = 0
End Function



Public Function SaltosDeLinea(ByVal CADENA As String) As String
    Dim Devu As String
    Dim i As Integer
    
    Devu = ""
    Do
        i = InStr(1, CADENA, vbCrLf)
        If i > 0 Then
            If Devu <> "" Then Devu = Devu & """ + chr(13) + """
            Devu = Devu & Mid(CADENA, 1, i - 1)
            CADENA = Mid(CADENA, i + 2)
            
       Else
            Devu = Devu & CADENA
       End If
    Loop While i > 0
    SaltosDeLinea = Devu
End Function


' posicionamos el combo cogiendo s�lo las 20 primeras posiciones

Public Sub PosicionarCombo2(ByRef Combo1 As ComboBox, Valor As String)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Mid(Combo1.List(J), 1, 20) = Trim(Valor) Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub

