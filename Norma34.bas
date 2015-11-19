Attribute VB_Name = "Norma34"
Option Explicit

Dim AuxD As String
Private NumeroTransferencia As Integer
'----------------------------------------------------------------------
'  Copia fichero generado bajo
Public Sub CopiarFicheroNorma43(Destino As String)

    
    'If Not CopiarEnDisquette(True, 3) Then
        AuxD = Destino
        CopiarEnDisquette False, 0  'A disco
    
        
End Sub

Private Function CopiarEnDisquette(A_disquetera As Boolean, Intentos As Byte) As Boolean
Dim i As Integer
Dim cad As String

On Error Resume Next

    CopiarEnDisquette = False
    
    If A_disquetera Then
        For i = 1 To Intentos
            cad = "Introduzca un disco vacio. (" & i & ")"
            MsgBox cad, vbInformation
            FileCopy App.Path & "\norma34.txt", "a:\norma34.txt"
            If Err.Number <> 0 Then
                MuestraError Err.Number, "Copiar En Disquette"
            Else
                CopiarEnDisquette = True
                Exit For
            End If
        Next i
    Else
        If AuxD = "" Then
            cad = Format(Now, "ddmmyyhhnn")
            cad = App.Path & "\" & cad & ".txt"
        Else
            cad = AuxD
        End If
        FileCopy App.Path & "\norma34.txt", cad
        If Err.Number <> 0 Then
            MsgBox "Error creando copia fichero. Consulte soporte t�cnico." & vbCrLf & Err.Description, vbCritical
        Else
            MsgBox "El fichero esta guardado como: " & cad, vbInformation
        End If
            
    End If
End Function



'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma34(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, Pagos As Boolean) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim RS As ADODB.Recordset
Dim Aux As String
Dim cad As String


    On Error GoTo EGen
    GeneraFicheroNorma34 = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
    'Cargamos la cuenta
    cad = "Select * from ctabancaria where codmacta='" & CuentaPropia & "'"
    Set RS = New ADODB.Recordset
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 10)
    If RS.EOF Then
        cad = ""
    Else
        If IsNull(RS!entidad) Then
            cad = ""
        Else
            cad = Format(RS!entidad, "0000") & "|" & Format(DBLet(RS!oficina, "T"), "0000") & "|" & DBLet(RS!Control, "T") & "|" & Format(DBLet(RS!CtaBanco, "T"), "0000000000") & "|"
            CuentaPropia = cad
        End If
        
        
        'Identificador norma bancaria
        If Not IsNull(RS!idnorma34) Then Aux = RS!idnorma34
    End If
    RS.Close
    Set RS = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NFich
    
    
    
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 10)  'CIF EMPRESA
    
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, cad
    Cabecera2 NFich, CodigoOrdenante, cad
    Cabecera3 NFich, CodigoOrdenante, cad
    Cabecera4 NFich, CodigoOrdenante, cad
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set RS = New ADODB.Recordset
    If Pagos Then
        Aux = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,despobla from spagop,cuentas"
        Aux = Aux & " where codmacta=ctaprove and transfer =" & NumeroTransferencia
    Else
        'ABONOS
         '
        Aux = "Select scobro.codbanco as entidad,scobro.codsucur as oficina,scobro.cuentaba,scobro.digcontr as CC"
        Aux = Aux & ",nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta from scobro,cuentas"
        Aux = Aux & " where cuentas.codmacta=scobro.codmacta and transfer =" & NumeroTransferencia
    End If
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If RS.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not RS.EOF
            If Pagos Then
                Im = DBLet(RS!imppagad, "N")
                Im = RS!impefect - Im
                Aux = RellenaAceros(RS!CtaProve, False, 12)
            
            Else
                Im = Abs(RS!ImpVenci)
                Aux = RellenaAceros(RS!Codmacta, False, 12)
            End If
            
            'Cad = "06"
            'Cad = Cad & "56"
            'Cad = Cad & " "
            'Aux = "06" & "56" & " " & CodigoOrdenante & Aux  'Ordenante y socio juntos
        
            Aux = "06" & "56" & CodigoOrdenante & Aux   'Ordenante y socio juntos
        
            Linea1 NFich, Aux, RS, Im, cad, ConceptoTransferencia
            Linea2 NFich, Aux, RS, cad
            Linea3 NFich, Aux, RS, cad
            Linea4 NFich, Aux, RS, cad
            Linea5 NFich, Aux, RS, cad
            Linea6 NFich, Aux, RS, cad, ConceptoTr, Pagos
            If Pagos Then Linea7 NFich, Aux, RS, cad
        
        
        
        
            Importe = Importe + Im
            Regs = Regs + 1
            RS.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, cad, Pagos
    End If
    RS.Close
    Set RS = Nothing
    Close (NFich)
    If Regs > 0 Then GeneraFicheroNorma34 = True
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function


'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma34New(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, Pagos As Boolean) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim RS As ADODB.Recordset
Dim Aux As String
Dim cad As String


    On Error GoTo EGen
    GeneraFicheroNorma34New = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
'    'Cargamos la cuenta
'    cad = "Select * from ctabancaria where codmacta='" & CuentaPropia & "'"
'    Set Rs = New ADODB.Recordset
'    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 10)
'    If Rs.EOF Then
'        cad = ""
'    Else
'        If IsNull(Rs!entidad) Then
'            cad = ""
'        Else
'            cad = Format(Rs!entidad, "0000") & "|" & Format(DBLet(Rs!oficina, "T"), "0000") & "|" & DBLet(Rs!Control, "T") & "|" & Format(DBLet(Rs!CtaBanco, "T"), "0000000000") & "|"
'            CuentaPropia = cad
'        End If
'
'
'        'Identificador norma bancaria
'        If Not IsNull(Rs!idnorma34) Then Aux = Rs!idnorma34
'    End If
'    Rs.Close
'    Set Rs = Nothing
'    If cad = "" Then
'        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
'        Exit Function
'    End If

    NFich = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NFich
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 10)  'CIF EMPRESA
    
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, cad
    Cabecera2 NFich, CodigoOrdenante, cad
    Cabecera3 NFich, CodigoOrdenante, cad
    Cabecera4 NFich, CodigoOrdenante, cad
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set RS = New ADODB.Recordset
    Aux = "Select codtraba,sum(impnomi), sum(impgasto)"
    
'    Aux = "select tmpimpor.*, straba.codbanco as entidad, straba.codsucur as oficina, straba.digcontr as CC, straba.cuentaba as cuentaba, "
'    Aux = Aux & " straba.nomtraba as nommacta, straba.domtraba as dirdatos, straba.codpobla as codposta, straba.pobtraba as despobla "
'    Aux = Aux & " from tmpimpor, straba where tmpimpor.codtraba = straba.codtraba "
    
''    If Pagos Then
''        Aux = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,despobla from spagop,cuentas"
''        Aux = Aux & " where codmacta=ctaprove and transfer =" & NumeroTransferencia
''    Else
''        'ABONOS
''         '
''        Aux = "Select scobro.codbanco as entidad,scobro.codsucur as oficina,scobro.cuentaba,scobro.digcontr as CC"
''        Aux = Aux & ",nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta from scobro,cuentas"
''        Aux = Aux & " where cuentas.codmacta=scobro.codmacta and transfer =" & NumeroTransferencia
''    End If



    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If RS.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not RS.EOF
'            If Pagos Then
'                Im = DBLet(Rs!imppagad, "N")
'                Im = Rs!impefect - Im
'                Aux = RellenaAceros(Rs!CtaProve, False, 12)
'
'            Else
'                Im = Abs(Rs!ImpVenci)
'                Aux = RellenaAceros(Rs!Codmacta, False, 12)
'            End If

            Im = DBLet(RS!Importe, "N")
            Aux = RellenaAceros("0", False, 12) 'Rs!Codmacta, False, 12)
            
            'Cad = "06"
            'Cad = Cad & "56"
            'Cad = Cad & " "
            'Aux = "06" & "56" & " " & CodigoOrdenante & Aux  'Ordenante y socio juntos
        
            Aux = "06" & "56" & CodigoOrdenante & Aux   'Ordenante y socio juntos
        
            Linea1 NFich, Aux, RS, Im, cad, ConceptoTransferencia
            Linea2 NFich, Aux, RS, cad
            Linea3 NFich, Aux, RS, cad
            Linea4 NFich, Aux, RS, cad
            Linea5 NFich, Aux, RS, cad
            Linea6 NFich, Aux, RS, cad, ConceptoTr, Pagos
            If Pagos Then Linea7 NFich, Aux, RS, cad
        
            Importe = Importe + Im
            Regs = Regs + 1
            RS.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, cad, Pagos
    End If
    RS.Close
    Set RS = Nothing
    Close (NFich)
    If Regs > 0 Then GeneraFicheroNorma34New = True
    Exit Function
    
EGen:
    MuestraError Err.Number, Err.Description
End Function

Private Function RellenaABlancos(CADENA As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Space(longitud)
    If PorLaDerecha Then
        cad = CADENA & cad
        RellenaABlancos = Left(cad, longitud)
    Else
        cad = cad & CADENA
        RellenaABlancos = Right(cad, longitud)
    End If
    
End Function



Private Function RellenaAceros(CADENA As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Mid("00000000000000000000", 1, longitud)
    If PorLaDerecha Then
        cad = CADENA & cad
        RellenaAceros = Left(cad, longitud)
    Else
        cad = cad & CADENA
        RellenaAceros = Right(cad, longitud)
    End If
    
End Function




Private Sub Cabecera1(NF As Integer, ByRef CodOrde As String, Fecha As Date, Cta As String, ByRef cad As String)

    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "001"
    cad = cad & Format(Now, "ddmmyy")
    cad = cad & Format(Fecha, "ddmmyy")
    'Cuenta bancaria
    cad = cad & RecuperaValor(Cta, 1)
    cad = cad & RecuperaValor(Cta, 2)
    cad = cad & RecuperaValor(Cta, 4)
    cad = cad & "0"  'Sin relacion
    cad = cad & "   " & RecuperaValor(Cta, 3)  'Digito de control bancario
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



Private Sub Cabecera2(NF As Integer, ByRef CodOrde As String, ByRef cad As String)
    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "002"
    
    cad = cad & RellenaABlancos(vParam.NombreEmpresa, True, 30)   'Nombre empresa
  
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Cabecera3(NF As Integer, ByRef CodOrde As String, ByRef cad As String)
    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "003"
    
    
'    AuxD = DevuelveDesdeBD("direccion", "empresa2", "codigo", 1, "N")
    cad = cad & RellenaABlancos(vParam.DomicilioEmpresa, True, 30) 'AuxD, True, 30)   'Nombre empresa
    cad = cad & RellenaABlancos("", True, 30)   'Nombre empresa
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



Private Sub Cabecera4(NF As Integer, ByRef CodOrde As String, ByRef cad As String)

    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "004"
    
'    AuxD = DevuelveDesdeBD("codpos", "empresa2", "codigo", 1, "N")
    cad = cad & RellenaABlancos(vParam.CPostal, False, 5) '   AuxD, False, 5)
    cad = cad & " "
'    AuxD = DevuelveDesdeBD("provincia", "empresa2", "codigo", 1, "N")
    cad = cad & RellenaABlancos(vParam.Provincia, True, 30) 'AuxD, True, 30)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



'ConceptoTransferencia
'1.- Abono nomina
'9.- Transferencia ordinaria
Private Sub Linea1(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef Importe1 As Currency, ByRef cad As String, vConceptoTransferencia As String)


   
    '
    cad = CodOrde   'llevara tb la ID del socio
    cad = cad & "010"
    cad = cad & RellenaAceros(CStr(Round(Importe1, 2) * 100), False, 12)
    
    cad = cad & RellenaAceros(CStr(rs1!entidad), False, 4)     'Entidad
    cad = cad & RellenaAceros(CStr(rs1!oficina), False, 4)   'Sucur
    cad = cad & RellenaAceros(CStr(rs1!cuentaba), False, 10)  'Cta
    cad = cad & "1" & vConceptoTransferencia
    cad = cad & "  "
    cad = cad & RellenaAceros(CStr(rs1!CC), False, 2)  'CC
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea2(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "011"
    cad = cad & RellenaABlancos(rs1!nommacta, False, 36)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea3(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "012"
    cad = cad & RellenaABlancos(DBLet(rs1!dirdatos, "T"), False, 36)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea4(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "013"
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea5(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "014"
    cad = cad & RellenaABlancos(DBLet(rs1!codposta, "T"), False, 5) & " "
    cad = cad & RellenaABlancos(DBLet(rs1!despobla, "T"), False, 30)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea6(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef cad As String, ByRef ConceptoT As String, Pagos As Boolean)
Dim Aux As String

    Aux = ConceptoT
'    If Pagos Then
'        'Tiene dos campos para las descripcion. Si no tiene nada pondre la descripcion de la transferencia
'        Aux = Trim(DBLet(RS1!text1csb, "T"))
'        If Aux = "" Then Aux = ConceptoT
'    End If

    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "016"
    cad = cad & RellenaABlancos(Aux, False, 35)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea7(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef cad As String)


    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "017"
    cad = cad & RellenaABlancos(DBLet(rs1!text2csb, "T"), False, 35)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub




Private Sub Totales(NF As Integer, ByRef CodOrde As String, total As Currency, Registros As Integer, ByRef cad As String, Pagos As Boolean)
    cad = "08" & "56"
    cad = cad & CodOrde    'llevara tb la ID del socio
    cad = cad & Space(15)
    cad = cad & RellenaAceros(CStr(Int(Round(total * 100, 2))), False, 12)
    cad = cad & RellenaAceros(CStr(Registros), False, 8)
    If Pagos Then
        cad = cad & RellenaAceros(CStr((Registros * 7) + 4 + 1), False, 10)
    Else
        cad = cad & RellenaAceros(CStr((Registros * 6) + 4 + 1), False, 10)
    End If
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub
