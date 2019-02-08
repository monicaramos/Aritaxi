Attribute VB_Name = "ModFunciones"
Option Explicit

Public Const ValorNulo = "Null"


Public NombreCheck As String

Public Function CompForm(ByRef formulario As Form, Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is CommonDialog Then
        'ElseIf TypeOf Control Is MSComm Then
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or (Opcion = 2 And Control.Name = "Text3") Or (Opcion = 3 And Control.Name = "txtAux") Then
                Carga = mTag.Cargar(Control)
                If Carga = True Then
                    Correcto = mTag.Comprobar(Control)
                    If Not Correcto Then Exit Function
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function

                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                            MsgBox "Seleccione una dato para: " & mTag.NOMBRE, vbExclamation
                            Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function



Public Sub limpiar(ByRef formulario As Form)
    Dim Control As Object

    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
End Sub

'-----------------------------------
Public Function ValorParaSQL(Valor, ByRef vtag As CTag) As String
Dim Dev As String
Dim D As Single
Dim i As Integer
Dim V
    Dev = ""
    If Valor <> "" Then
        Select Case vtag.TipoDato
        Case "N"
            V = Valor
            If InStr(1, Valor, ",") Then
                If InStr(1, Valor, ".") Then
                    'ABRIL 2004

                    'Ademas de la coma lleva puntos
                    V = ImporteFormateado(CStr(Valor))
                    Valor = V
                Else

                    V = CSng(Valor)
                    Valor = V
                End If
            Else

            End If
            Dev = TransformaComasPuntos(CStr(Valor))

        Case "F"
            Dev = "'" & Format(Valor, FormatoFecha) & "'"
        Case "H"
            Dev = "'" & Format(Valor, FormatoFecha & " hh:mm:ss") & "'"
        Case "T"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & Valor & "'"
        End Select

    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vtag.Vacio = "S" Then
            Dev = ValorNulo
        Else
            'Modifica Laura: 04/10/05
            If vtag.TipoDato = "N" Then
                Dev = "0"
            Else
                Dev = "''"
            End If
        End If
    End If
    ValorParaSQL = Dev
End Function



Public Function InsertarDesdeForm(ByRef formulario As Form, Optional Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or Opcion = 0 Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & ""
                        
                            'Parte VALUES
                            Cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                    Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Pr�cticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    conn.Execute Cad, , adCmdText
    
    InsertarDesdeForm = True
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function



Public Function CadenaInsertarDesdeForm(ByRef formulario As Form) As String
'Equivale a InsertarDesdeForm, excepto que devuelve la candena SQL y hace el execute fuera de la funci�n.
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    CadenaInsertarDesdeForm = ""
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        Cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                    Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Pr�cticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
'    Conn.Execute cad, , adCmdText
    
    CadenaInsertarDesdeForm = Cad
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function PonerCamposForma(ByRef formulario As Form, ByRef vData As Adodc) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim i As Integer


    On Error GoTo EPonerCamposForma


    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In formulario.Controls
        'TEXTO
        If TypeOf Control Is CommonDialog Then
        
        ElseIf (TypeOf Control Is TextBox) And (Control.visible = True) And (Control.Name = "Text1") Then
'                If TypeOf control Is TextBox Then

            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    
                    If mTag.columna <> "" Then
                        'Debug.Print mTag.columna
                        'If mTag.columna = "porciva3re" Then Stop
                        
                        campo = mTag.columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(campo))
                        Else
                            Valor = vData.Recordset.Fields(campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            If mTag.TipoDato = "N" Then
                                'Mariela 25/06/2010
                                'If Val(Valor) = 0 Then
                                    'Control.Text = ""
                                'Else
                                    Control.Text = Valor
                                'End If
                            Else
                                Control.Text = Valor
                            End If
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) And (Control.visible = True) Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    campo = mTag.columna
                    Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                End If
                If IsNull(Valor) Then Valor = 0
                Control.Value = Valor
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    campo = mTag.columna
                    Valor = DBLet(vData.Recordset.Fields(campo))
                    i = 0
                    For i = 0 To Control.ListCount - 1
                        If Control.ItemData(i) = Val(Valor) Then
                            Control.ListIndex = i
                            Exit For
                        End If
                    Next i
                    If i = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    Cad = Err.Description
    Cad = "Poner campos formulario. " & vbCrLf & campo & vbCrLf & Cad & vbCrLf
    MsgBox Cad, vbExclamation
End Function



Public Function PonerCamposFormaFrame(ByRef formulario As Form, NomTxtBox As String, ByRef vData As Adodc, Optional NomCheck As String, Optional NomCombo As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim i As Integer

    Set mTag = New CTag
    PonerCamposFormaFrame = False


        For Each Control In formulario.Controls
        If TypeOf Control Is TextBox And Control.visible = True And Control.Name = NomTxtBox Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.columna <> "" Then
                        campo = mTag.columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(campo))
                        Else
                            Valor = vData.Recordset.Fields(campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = Valor
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True And Control.Name = NomCheck Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    campo = mTag.columna
                    Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                End If
                Control.Value = Valor
            End If

         'COMBOBOX
         ElseIf TypeOf Control Is ComboBox And Control.visible = True And Control.Name = NomCombo Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    campo = mTag.columna
                    Valor = vData.Recordset.Fields(campo)
                    i = 0
                    For i = 0 To Control.ListCount - 1
                        If Control.ItemData(i) = Val(Valor) Then
                            Control.ListIndex = i
                            Exit For
                        End If
                    Next i
                    If i = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If

    Next Control

    'Veremos que tal
    PonerCamposFormaFrame = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function


Private Function ObtenerMaximoMinimo(ByRef vSQL As String) As String
Dim Rs As Recordset
    ObtenerMaximoMinimo = ""
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.EOF) Then
            ObtenerMaximoMinimo = CStr(Rs.Fields(0))
        End If
    End If
    Rs.Close
    Set Rs = Nothing
End Function

'====DAVID
'Public Function ObtenerBusqueda(ByRef formulario As Form) As String
'    Dim Control As Object
'    Dim Carga As Boolean
'    Dim mTag As CTag
'    Dim Aux As String
'    Dim cad As String
'    Dim SQL As String
'    Dim tabla As String
'    Dim RC As Byte
'
'    On Error GoTo EObtenerBusqueda
'
'    'Exit Function
'    Set mTag = New CTag
'    ObtenerBusqueda = ""
'    SQL = ""
'
'    'Recorremos los text en busca de ">>" o "<<"
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If Aux = ">>" Or Aux = "<<" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'                    If Aux = ">>" Then
'                        cad = " MAX(" & mTag.Columna & ")"
'                    Else
'                        cad = " MIN(" & mTag.Columna & ")"
'                    End If
'                    SQL = "Select " & cad & " from " & mTag.tabla
'                    SQL = ObtenerMaximoMinimo(SQL)
'                    Select Case mTag.TipoDato
'                    Case "N"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
'                    Case "F"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
'                    Case Else
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & SQL & "'"
'                    End Select
'                    SQL = "(" & SQL & ")"
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los text en busca del NULL
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If UCase(Aux) = "NULL" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'
'                    SQL = mTag.tabla & "." & mTag.Columna & " is NULL"
'                    SQL = "(" & SQL & ")"
'                    Control.Text = ""
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los textbox
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            'Cargamos el tag
'            Carga = mTag.Cargar(Control)
'            If Carga Then
'                If mTag.Cargado Then
'                    Aux = Trim(Control.Text)
'                    If Aux <> "" Then
'                        If mTag.tabla <> "" Then
'                            tabla = mTag.tabla & "."
'                        Else
'                            tabla = ""
'                        End If
'                    RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, cad)
'                    If RC = 0 Then
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'            Else
'                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
'                Exit Function
'            End If
'
'        'COMBO BOX
'        ElseIf TypeOf Control Is ComboBox Then
'            mTag.Cargar Control
'            If mTag.Cargado Then
'                If Control.ListIndex > -1 Then
'                    If mTag.TipoDato <> "T" Then
'                        cad = Control.ItemData(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = " & cad
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    Else
'                        cad = Control.List(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = '" & cad & "'"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'
'
'        'CHECK
'        ElseIf TypeOf Control Is CheckBox Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    If Control.Value = 1 Then
'                        cad = mTag.tabla & "." & mTag.Columna & " = 1"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'        End If
'
'
'    Next Control
'    ObtenerBusqueda = SQL
'Exit Function
'EObtenerBusqueda:
'    ObtenerBusqueda = ""
'    MuestraError Err.Number, "Obtener b�squeda. "
'End Function

'A�ado Optional CHECK As String. Para poder realizar las busquedas con los checks
Public Function ObtenerBusqueda(ByRef formulario As Form, paraRPT As Boolean, Optional CHECK As String) As String
Dim Control As Object
Dim Carga As Boolean
Dim mTag As CTag
Dim Aux As String
Dim Cad As String
Dim Sql As String
Dim Tabla As String, columna As String
Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) Then
            If Control.visible Then
                Aux = Trim(Control.Text)
                If Aux = ">>" Or Aux = "<<" Then
                    Carga = mTag.Cargar(Control)
                    If Carga Then
                        If Aux = ">>" Then
                            If Not paraRPT Then
                                Cad = " MAX(" & mTag.columna & ")"
                            Else
                                Cad = " MAX({" & mTag.Tabla & "." & mTag.columna & "})"
                            End If
                        Else
                            If Not paraRPT Then
                                Cad = " MIN(" & mTag.columna & ")"
                            Else
                                Cad = " MIN({" & mTag.Tabla & "." & mTag.columna & "})"
                            End If
                        End If
                        If Not paraRPT Then
                            Sql = "Select " & Cad & " from " & mTag.Tabla
                        Else
                            Sql = "Select " & Cad & " from {" & mTag.Tabla & "}"
                        End If
                        Sql = ObtenerMaximoMinimo(Replace(Replace(Sql, "}", ""), "{", ""))
                        
                        Select Case mTag.TipoDato
                        Case "N"
                            If Not paraRPT Then
                                Sql = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                            Else
                                Sql = "{" & mTag.Tabla & "." & mTag.columna & "} = " & TransformaComasPuntos(Sql)
                            End If
                        Case "F"
                            If Not paraRPT Then
                                Sql = mTag.Tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                            Else
                                Sql = "{" & mTag.Tabla & "." & mTag.columna & "} = date(""" & Format(Sql, "dd/mm/yyyy") & """)"
                            End If
                        Case Else
                            If Not paraRPT Then
                                Sql = mTag.Tabla & "." & mTag.columna & " = '" & Sql & "'"
                            Else
                                Sql = "{" & mTag.Tabla & "." & mTag.columna & "} = '" & Sql & "'"
                            End If
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next

    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) Then
            If Control.visible Then
                Aux = Trim(Control.Text)
                If UCase(Aux) = "NULL" Then
                    Carga = mTag.Cargar(Control)
                    If Carga Then
                        If Not paraRPT Then
                            Sql = mTag.Tabla & "." & mTag.columna & " is NULL"
                        Else
                            ' cambiado para rpt [Monica]
    '                        SQL = "{" & mTag.tabla & "." & mTag.columna & "} is NULL"
                            Sql = "isnull({" & mTag.Tabla & "." & mTag.columna & "})"
                        End If
                        Sql = "(" & Sql & ")"
                        '[Monica] lo he quitado
    '                    Control.Text = ""
                    End If
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) Then
            If Control.visible Then
                'Cargamos el tag
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If mTag.Cargado Then
                        Aux = Trim(Control.Text)
                        Aux = QuitarCaracterEnter(Aux) 'Si es multilinea quitar ENTER
                        If Aux <> "" Then
                            If mTag.Tabla <> "" Then
                                If Not paraRPT Then
                                    Tabla = mTag.Tabla & "."
                                Else
                                    Tabla = "{" & mTag.Tabla & "."
                                End If
                            Else
                                Tabla = ""
                            End If
                            If Not paraRPT Then
                                columna = mTag.columna
                            Else
                                columna = mTag.columna & "}"
                            End If
                        Rc = SeparaCampoBusqueda(mTag.TipoDato, Tabla & columna, Aux, Cad, paraRPT)
                        If Rc = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            If Not paraRPT Then
                                Sql = Sql & "(" & Cad & ")"
                            Else
                                Sql = Sql & "(" & Cad & ")"
                            End If
                        End If
                    End If
                End If
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.visible Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex > -1 Then
                        If mTag.TipoDato <> "T" Then
                            Cad = Control.ItemData(Control.ListIndex)
                            If Not paraRPT Then
                                Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                            Else
                                Cad = "{" & mTag.Tabla & "." & mTag.columna & "} = " & Cad
                            End If
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & Cad & ")"
                        Else
                            Cad = Control.List(Control.ListIndex)
                            If Not paraRPT Then
                                Cad = mTag.Tabla & "." & mTag.columna & " = '" & Cad & "'"
                            Else
                                Cad = "{" & mTag.Tabla & "." & mTag.columna & "} = '" & Cad & "'"
                            End If
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & Cad & ")"
                        End If
                    End If
                End If
            End If

        'CHECK
                'CHECK
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    
                    Aux = ""
                    If Control.Value = 1 Then
                        Aux = "1"
                    Else
                        If CHECK <> "" Then
                            CheckBusqueda Control
                            Tabla = NombreCheck & "|"
                            If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                        End If
                    End If
                    If Aux <> "" Then
                        If Not paraRPT Then
                            Cad = mTag.Tabla & "." & mTag.columna
                        Else
                            Cad = "{" & mTag.Tabla & "." & mTag.columna & "} "
                        End If
                        
                        Cad = Cad & " = " & Aux
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & Cad & ")"
                    End If 'cargado
                End If '<>""
            End If
        End If
    
    Next Control
    ObtenerBusqueda = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener b�squeda. "
End Function



Public Function ObtenerBusqueda2(ByRef formulario As Form, Optional CHECK As String, Optional opcio As Integer, Optional nom_frame As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim Cad As String
    Dim Sql As String
    Dim Tabla As String
    Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda2 = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Aux = ">>" Then
                            Cad = " MAX(" & mTag.columna & ")"
                        Else
                            Cad = " MIN(" & mTag.columna & ")"
                        End If
                        Sql = "Select " & Cad & " from " & mTag.Tabla
                        Sql = ObtenerMaximoMinimo(Sql)
                        Select Case mTag.TipoDato
                        Case "N"
                            Sql = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Case "F"
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case Else
                            '[Monica]04/03/2013: quito las comillas y pongo el dbset
                            Sql = mTag.Tabla & "." & mTag.columna & " = " & DBSet(Sql, "T") ' & "'"
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next

'++monica: lo he a�adido del anterior obtenerbusqueda
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga And mTag.columna <> "" Then

                    Sql = mTag.Tabla & "." & mTag.columna & " is NULL"
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
          If Control.Tag <> "" Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.Tabla <> "" Then
                            Tabla = mTag.Tabla & "."
                            Else
                            Tabla = ""
                        End If
                        Rc = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.columna, Aux, Cad, mTag.EsClave)
                        If Rc = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & Cad & ")"
                        End If
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then ' +-+- 12/05/05: canvi de C�sar, no te sentit passar-li un control que no t� TAG +-+-
                mTag.Cargar Control
                If mTag.Cargado Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Control.ListIndex > -1 Then
                            Cad = Control.ItemData(Control.ListIndex)
                            Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & Cad & ")"
                        End If
                    End If
                End If
            End If
            
         ElseIf TypeOf Control Is CheckBox Then
            '=============== A�ade: Laura, 27/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    ' a�adido 12022007
                    Aux = ""
                    If CHECK <> "" Then
                        Tabla = DBLet(Control.Index, "T")
                        If Tabla <> "" Then Tabla = "(" & Tabla & ")"
                        Tabla = Control.Name & Tabla & "|"
                        If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        Cad = Control.Value
                        Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & Cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda2 = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda2 = ""
    MuestraError Err.Number, "Obtener b�squeda. " & vbCrLf & Err.Description
End Function





Public Function ModificaDesdeFormulario(ByRef formulario As Form, Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUpdate As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or (Opcion = 3 And Control.Name = "txtAux") Then
            If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                                 cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUpdate <> "" Then cadUpdate = cadUpdate & " , "
                                cadUpdate = cadUpdate & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUpdate <> "" Then cadUpdate = cadUpdate & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUpdate = cadUpdate & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUpdate <> "" Then cadUpdate = cadUpdate & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUpdate = cadUpdate & "" & mTag.columna & " = " & Aux
                End If
            End If
        ElseIf TypeOf Control Is OptionButton And Control.visible Then
            If Control.Enabled Then
                If Control.Value = True And Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        Aux = Control.Index
                        If cadUpdate <> "" Then cadUpdate = cadUpdate & " , "
                        cadUpdate = cadUpdate & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE Pa�sDestinatario = 'M�xico';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUpdate & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function


Public Sub FormateaCampo(vTex As TextBox)
'devuelve el valor del control vText.text formateado: 12 -> "0012"
    Dim mTag As CTag
    Dim Cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                Cad = TransformaPuntosComas(vTex.Text)
                Cad = Format(Cad, mTag.Formato)
                vTex.Text = Cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub

Public Function FormatoCampo(ByRef vTex As TextBox) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim Cad As String
On Error GoTo EFormatoCampo

    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        FormatoCampo = mTag.Formato
    End If
EFormatoCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef CADENA As String, Orden As Integer, Optional Separador As String) As String
Dim i As Integer
Dim J As Integer
Dim cont As Integer
Dim Cad As String

    If Separador = "" Then Separador = "|"

    i = 0
    cont = 1
    Cad = ""
    Do
        J = i + 1
        i = InStr(J, CADENA, Separador)
        If i > 0 Then
            If cont = Orden Then
                Cad = Mid(CADENA, J, i - J)
                i = Len(CADENA) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until i = 0
    RecuperaValor = Cad
End Function



'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim i As Integer
Dim J As Integer

On Error GoTo EPonerOpcionesMenuGeneral

'A�adir, modificar y borrar deshabilitados si no nivel
With formulario

    'LA TOOLBAR  .--> Requisito, k se llame toolbar1
    For i = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(i).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(i).Tag)
            If J < vUsu.Nivel Then
                .Toolbar1.Buttons(i).Enabled = False
            End If
        End If
    Next i

    'Esto es un poco salvaje. Por si acaso , no existe en este trozo pondremos los errores on resume next

    On Error Resume Next

    'Los MENUS
    'K sean mnAlgo
    J = Val(.mnnuevo.HelpContextID)
    If J < vUsu.Nivel Then .mnnuevo.Enabled = False

    J = Val(.mnModificar.HelpContextID)
    If J < vUsu.Nivel Then .mnModificar.Enabled = False

    J = Val(.mnEliminar.HelpContextID)
    If J < vUsu.Nivel Then .mnEliminar.Enabled = False
    
    J = Val(.mnLineas.HelpContextID)
    If J < vUsu.Nivel Then .mnLineas.Enabled = False
    
    On Error GoTo 0
End With

Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



Public Function BLOQUEADesdeFormulario(ByRef formulario As Form, Optional Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or Opcion <> 1 Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next Control

    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.Tabla
        Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
    
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function BloqueaRegistro(cadTabla As String, cadWHERE As String) As Boolean
Dim Aux As String
On Error GoTo EBloqueaRegistro

        BloqueaRegistro = False
        
        Aux = "SELECT * FROM " & cadTabla
        Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BloqueaRegistro = True
        
EBloqueaRegistro:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
End Function


Public Function BloqueaRegistroForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.Name = "Text1" Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    Next Control

    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "Insert into zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.Tabla
        Aux = Aux & "',""" & ComprobarComillas(AuxDef) & """)"
        conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '�Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Private Function ComprobarComillas(Cad As String) As String
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, Cad, """")
        If i > 0 Then
            Aux = Mid(Cad, 1, i - 1) & "\"
            Cad = Aux & Mid(Cad, i)
            J = i + 2
        End If
    Loop Until i = 0
    ComprobarComillas = Cad
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim Sql As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
        Sql = "DELETE from zbloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.Tabla & "'"
        conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function



Public Function BloqueoManual(cadTabla As String, cadWHERE As String, Optional OcultarMsg As Boolean) As Boolean
Dim Aux As String

On Error GoTo EBLOQ
    BloqueoManual = False
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "INSERT INTO zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & cadTabla
        Aux = Aux & "',""" & cadWHERE & """)"
        conn.Execute Aux
        BloqueoManual = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '�Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            If Not OcultarMsg Then MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
'    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueoManual(cadTabla As String) As Boolean
Dim Sql As String

'Solo me interesa la tabla
On Error Resume Next

        Sql = "DELETE FROM zbloqueos WHERE codusu=" & vUsu.Codigo & " and tabla='" & cadTabla & "'"
        conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function


'====================== LAURA

Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function


Public Function QuitarCero(Valor As String) As String
    On Error Resume Next
    
    If Valor <> "" Then
        If CSng(Valor) = 0 Then
            QuitarCero = ""
        Else
            QuitarCero = Valor
        End If
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function CalcularImporte(Cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    Cantidad = ComprobarCero(Cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    vImp = CCur(Cantidad) * CCur(vPre)
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    '// Enero 2009.  Hacia mal el redondeo pq ahora cantidad lleva decimales
    '   Ponemos Round2
    vImp = Round2(vImp, 2)
    CalcularImporte = CStr(vImp)
End Function

'Redondeo a 4 digitos
Public Function CalcularImporte4(Cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    Cantidad = ComprobarCero(Cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    vImp = CCur(Cantidad) * CCur(vPre)
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    vImp = Round(vImp, 4)
    CalcularImporte4 = CStr(vImp)
End Function



Public Function CalcularImporteSng(Cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran,
'donde PRECIO es sng                                          *********************** MAYO 2009
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Single
Dim vDto1 As Single, vDto2 As Single
Dim vPre As Single
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    Cantidad = ComprobarCero(Cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    vImp = CSng(Cantidad) * CSng(vPre)
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    '// Enero 2009.  Hacia mal el redondeo pq ahora cantidad lleva decimales
    '   Ponemos Round2
    vImp = Round2(vImp, 2)
    CalcularImporteSng = CStr(vImp)
End Function





Public Function CalcularDto(Importe As String, Dto As String) As String
'devuelve el Dto% del Importe
'Ej el 16% de 120 = 19.2
Dim vImp As Currency
Dim vDto As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    Dto = ComprobarCero(Dto)
    
    vImp = CCur(Importe)
    vDto = CCur(Dto)
    
    vImp = ((vImp * vDto) / 100)
    'vImp = Round(vImp, 2)
    
    CalcularDto = CStr(vImp)
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function CalcularNumBultos(Cantidad As Currency, UdsCaja As Integer) As Integer
Dim numUds As Integer 'unidades
    
    If UdsCaja > 0 Then
        '- calcular los bultos q necesitamos para la cantidad
        numUds = Int(Cantidad / UdsCaja)
        If Cantidad Mod UdsCaja > 0 Then
            numUds = numUds + 1
        ElseIf Cantidad > Int(UdsCaja * numUds) Then
             numUds = numUds + 1
        End If
        
        
        If numUds = 0 And Cantidad <> 0 Then numUds = numUds + 1
    End If
    
    CalcularNumBultos = numUds
End Function


'Si pone algo en DevuelveImporte en lugar del msg metera en esa cadena el importe
Public Sub ComprobarCobrosCliente2(CodClien As String, FechaDoc As String, Optional DevuelveImporte As String)
'Comprueba en la tabla de Cobros Pendientes (scobro) de la Base de datos de Contabilidad
'si el cliente tiene alguna factura pendiente de cobro que ha vendido
'con fecha de vencimiento anterior a la fecha del documento: Oferta, Pedido, ALbaran,...
Dim Sql As String, vWhere As String
Dim codmacta As String
Dim Rs As ADODB.Recordset
Dim cadMen As String
Dim ImporteCred As Currency
Dim Importe As Currency
Dim Impaux As Currency

    Set Rs = New ADODB.Recordset
    ImporteCred = 0
    'Obtener la cuenta del cliente de la tabla sclien en Aritaxi
    Sql = "Select nomclien,codmacta,limcredi from sclien where codclien=" & CodClien
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        Sql = ""
    Else
        'CodClien = CodClien & " - " & sql
        CodClien = CodClien & " - " & Rs!nomclien
        ImporteCred = DBLet(Rs!limcredi, "N")
        If ImporteCred > 0 Then CodClien = CodClien & "   L�mite credito: " & Format(ImporteCred, FormatoImporte)
        codmacta = Rs!codmacta
    End If
    Rs.Close
    If Sql = "" Then Exit Sub
    
    If vParamAplic.ContabilidadNueva Then
        'AHORA FEBRERO 2010
        Sql = "SELECT cobros.* FROM cobros INNER JOIN formapago ON cobros.codforpa=formapago.codforpa "
        vWhere = " WHERE cobros.codmacta = '" & codmacta & "'"
        vWhere = vWhere & " AND fecvenci <= ' " & Format(FechaDoc, FormatoFecha) & "' "
        'Antes mayo 2010
        'vWhere = vWhere & " AND (sforpa.tipforpa between 0 and 3)"
        vWhere = vWhere & " AND recedocu=0 ORDER BY fecfactu, numfactu"
        Sql = Sql & vWhere
    Else
        'AHORA FEBRERO 2010
        Sql = "SELECT scobro.* FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
        vWhere = " WHERE scobro.codmacta = '" & codmacta & "'"
        vWhere = vWhere & " AND fecvenci <= ' " & Format(FechaDoc, FormatoFecha) & "' "
        'Antes mayo 2010
        'vWhere = vWhere & " AND (sforpa.tipforpa between 0 and 3)"
        vWhere = vWhere & " AND recedocu=0 ORDER BY fecfaccl, codfaccl"
        Sql = Sql & vWhere
    End If
    'Lee de la Base de Datos de CONTABILIDAD
    Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    While Not Rs.EOF
    
        'QUITO LO DE DEVUELTO. MAYO 2010
        'If Val(RS!Devuelto) = 1 Then
        '    'SALE SEGURO (si no esta girado otra vez �no?
        '    'Si esta girado otra vez tendra impcobro, con lo cual NO tendra diferencia de importes
        '    Impaux = RS!ImpVenci + DBLet(RS!gastos, "N") - DBLet(RS!impcobro, "N")
            
        'Else
            'Si esta recibido NO lo saco
            If Val(Rs!recedocu) = 1 Then
                Impaux = 0
            Else
                'NO esta recibido. Si tiene diferencia
                Impaux = Rs!ImpVenci + DBLet(Rs!Gastos, "N") - DBLet(Rs!impcobro, "N")
        
            End If
    '    End If
        If Impaux <> 0 Then Importe = Importe + Impaux
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        If Importe > 0 Then
        
            If DevuelveImporte <> "" Then
                'Meto aqui el importer
                DevuelveImporte = CStr(Importe)
            Else
                cadMen = "El Socio tiene facturas vencidas con valor de: " & Format(Importe, FormatoImporte) & " �."
                If ImporteCred > 0 Then cadMen = cadMen & vbCrLf & "L�mite cr�dito: " & Format(ImporteCred, FormatoImporte) & " �."
                cadMen = cadMen & vbCrLf & "�Desea Ver Detalle?"
                If MsgBox(cadMen, vbYesNo, "Cobros Pendientes") = vbYes Then
                    'Mostrar los detalles de los cobros pendientes
                    frmMensajes.cadWHERE = vWhere
                    frmMensajes.vCampos = CodClien
                    frmMensajes.OpcionMensaje = 1
                    frmMensajes.Show vbModal
                End If
            End If
        End If
    
    
End Sub


Public Function EsArticuloVarios(codArtic As String) As Boolean
Dim devuelve As String

    EsArticuloVarios = False
    devuelve = DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", codArtic, "T")
    
    If devuelve = "1" Or devuelve = "2" Then 'Es Articulo de Varios y podemos modificar la Denominaci�n del Articulo
        EsArticuloVarios = True
    Else
        EsArticuloVarios = False
    End If
End Function


Public Function EsClienteVarios(vCodClien As String) As Boolean
'Devuelve true si es un cliente de varios
Dim devuelve As String

    EsClienteVarios = False
    devuelve = DevuelveDesdeBD(conAri, "clivario", "scliente", "codclien", vCodClien, "N")
    If devuelve <> "" Then EsClienteVarios = CBool(devuelve)
    'Es cliente de varios Y podemos recuperar de sclvar los datos
    'del cliente por el NIF
End Function


Public Function EsProveedorVarios(codProve As String) As Boolean
Dim devuelve As String

    EsProveedorVarios = False
    devuelve = DevuelveDesdeBD(conAri, "provario", "sprove", "codprove", codProve, "N")
    If devuelve <> "" Then EsProveedorVarios = CBool(devuelve)
    'Es proveedor de varios Y podemos recuperar de ????
End Function


Public Function ObtenerNSerieSiguiente(cadNSerie As String) As String
'IN -> cadNSerie: cadena con el N� Serie de Tipo: "0000-12-0011"
'OUT -> RETURN: cadena con el sig. N�Serie : "0000-12-0012"
Dim NumAux As String, numAnt As String
Dim NumAux2 As String
Dim i As Integer

    On Error Resume Next
    
    NumAux = cadNSerie
    numAnt = ""
    'Quitar los cararacter '-' y quedarse con la parte dcha
    i = InStr(1, NumAux, "-")
    While Not i = 0
        numAnt = numAnt & Mid(NumAux, 1, i)
        NumAux = Mid(NumAux, i + 1, Len(NumAux) - i)
        i = InStr(1, NumAux, "-")
    Wend
    
    If NumAux <> "" Then 'Hay q coger la parte derecha del - : 0011
        i = Len(NumAux)
        If IsNumeric(NumAux) Then
            NumAux = CStr(NumAux + 1)
            While Len(NumAux) < i
                NumAux = "0" & NumAux
            Wend
        Else
        'Coger el n� mas a la derecha, incrementarlo y concatenarlo con el principio
            NumAux2 = Mid(NumAux, i, Len(NumAux))
            While IsNumeric(NumAux2)
                i = i - 1
                NumAux2 = Mid(NumAux, i, Len(NumAux))
            Wend
            NumAux2 = Right(NumAux2, Len(NumAux2) - 1)
            numAnt = numAnt & Mid(NumAux, 1, i)
            NumAux = CStr(NumAux2 + 1)
            While Len(NumAux) < Len(NumAux2)
                NumAux = "0" & NumAux
            Wend
        End If
        
        If numAnt <> "" Then
            ObtenerNSerieSiguiente = numAnt & NumAux
        Else
            ObtenerNSerieSiguiente = NumAux
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function PonerTrabajadorConectado(NomTraba As String) As String
'Pone en el campo del Form "Realizada Por" el trabajador que esta conectado en ese momento
'OUT: codTraba, NomTraba
Dim devuelve As String

    On Error Resume Next

    NomTraba = "nomtraba"
    devuelve = DevuelveDesdeBDNew(conAri, "straba", "codtraba", "login", vUsu.Login, "T", NomTraba)
    If devuelve <> "" Then
        PonerTrabajadorConectado = Format(devuelve, "0000") 'Cod. Trabajador
    Else
        PonerTrabajadorConectado = ""
        NomTraba = ""
    End If
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function PonerAlmacen(codAlm As String) As String
'Comprueba si existe el Almacen y lo pone en el Text
Dim devuelve As String
    
    On Error Resume Next

    If codAlm = "" Then
        MsgBox "Debe introducir el Almacen.", vbInformation
    Else
        devuelve = DevuelveDesdeBDNew(conAri, "salmpr", "codalmac", "codalmac", codAlm, "N")
        If devuelve = "" Then
            MsgBox "No existe el Almacen: " & Format(codAlm, "000"), vbInformation
            PonerAlmacen = ""
        Else
            PonerAlmacen = Format(codAlm, "000")
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


'=============================================================================
'==================== REPARACIONES ===========================================

Public Sub ComprobarReparaciones(Modo As Byte, numSerie As String, codArtic As String)
Dim numRep As Integer

    'Comprobar si ya esta en Reparacion
    If Modo = 3 Then ComprobarSiReparandose numSerie, codArtic
    'Comprobar cuantas veces se ha reparado ya el articulo(ver historico Reparaciones)
    numRep = ComprobarNumRepHco(numSerie, codArtic)
    If numRep > 0 Then
        MsgBox "Este aparato ya ha sido reparado " & numRep & " veces.", vbInformation
    End If
End Sub



Public Function ComprobarSiReparandose(numSerie As String, codArtic As String) As Boolean
'Comprueba si ya el Articulo se esta reparando, es decir si existe un registro
' en la tabla scarep
'IN -> numSerie, codArtic
Dim devuelve As String

    devuelve = DevuelveDesdeBDNew(conAri, "scarep", "numrepar", "numserie", numSerie, "T", , "codartic", codArtic, "T")
    If devuelve <> "" Then
        MsgBox "Este aparato ya esta en Reparaci�n.", vbInformation
        ComprobarSiReparandose = True
    Else
        ComprobarSiReparandose = False
    End If
End Function


Public Function ComprobarNumRepHco(numSerie As String, codArtic As String) As Integer
'Comprueba cuantas veces se ha reparado ya el articulo
'Ver cuantos registros existen en la tabla de historico Reparaciones (schrep)
'IN -> numserie, codartic
'RETURN -> N� Reparaciones
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error GoTo ENumRep

    Sql = " SELECT count(numrepar) FROM schrep "
    Sql = Sql & " WHERE numserie=" & DBSet(numSerie, "T") & " and "
    Sql = Sql & " codartic=" & DBSet(codArtic, "T")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ComprobarNumRepHco = Rs.Fields(0).Value
    Else
        ComprobarNumRepHco = 0
    End If
    
    Rs.Close
    Set Rs = Nothing
    
ENumRep:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Public Function ObtenerLetraSerie(tipMov As String) As String
'Devuelve la letra de serie asociada al tipo de movimiento
Dim LEtra As String

    On Error Resume Next
    
    LEtra = DevuelveDesdeBDNew(conAri, "stipom", "letraser", "codtipom", tipMov, "T")
    If LEtra = "" Then MsgBox "Las factura de venta no tienen asignada una letra de serie", vbInformation
    ObtenerLetraSerie = LEtra
End Function


Public Function ObtenerPoblacion(CPostal As String, ByRef provin As String) As String
'IN: "cpostal"
'OUT: en "provin" devolvemos la provincia
'     en ObtenerPoblacion devolvemos la poblacion
Dim devuelve As String

    On Error GoTo EPoblacion

    If CPostal <> "" Then
        devuelve = DevuelveDesdeBDNew(conAri, "scpostal", "provincia", "cpostal", CPostal, "T")
        ObtenerPoblacion = devuelve 'Nombre Poblacion
        If devuelve <> "" Then 'Nombre Provincia
            provin = DevuelveDesdeBDNew(conAri, "scpostal", "provincia", "cpostal", Mid(CPostal, 1, 2), "T")
        Else
            provin = ""
            MsgBox "No existe el CPostal " & CPostal, vbInformation
        End If
    Else
        ObtenerPoblacion = ""
        provin = ""
    End If
    
EPoblacion:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener Poblaci�n", Err.Description
End Function


Public Sub ObtenerCtasBancoPropio2(banPr As String, ctaBan As String, ctaCble As String)
'obtener la cuenta bancaria y la cuenta contable del banco propio
'(IN) banPr: cod. banco propio
'(OUT) ctaBan: cuenta bancaria
'(OUT) ctaCble: cuenta contable
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Aux As String

    ctaBan = ""
    ctaCble = ""

    Sql = "SELECT codbanco,codsucur,digcontr,cuentaba,codmacta"
    Sql = Sql & " from sbanpr where codbanpr=" & banPr

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        Aux = Right("0000" & DBLet(Rs!codbanco, "T"), 4)
        ctaBan = Aux & "-"
        Aux = Right("0000" & DBLet(Rs!codsucur, "T"), 4) & "-"
        ctaBan = ctaBan & Aux
        ctaBan = ctaBan & DBLet(Rs!digcontr, "T") & "-" & DBLet(Rs!cuentaba, "T")
        ctaCble = DBLet(Rs!codmacta, "T")
        'obtener el nombre de la cuenta contable
        Sql = ""
        Sql = DevuelveDesdeBD(conConta, "nommacta", "cuentas", "codmacta", ctaCble, "T")
        If Sql <> "" Then ctaCble = ctaCble & "-" & Sql
    End If
    Set Rs = Nothing
End Sub



Public Function ObtenerSQLcomponentes(cadWHERE As String) As String
'Obtiene la consulta SQL que selecciona los articulos con n� de serie
'agrupados por tipo de articulo
Dim Sql As String

    Sql = "Select distinct sserie.codtipar, nomtipar, count(numserie) as cantidad "
    Sql = Sql & "FROM sserie INNER JOIN stipar ON sserie.codtipar=stipar.codtipar "
    Sql = Sql & cadWHERE
    Sql = Sql & " GROUP by codtipar "
    
    ObtenerSQLcomponentes = Sql
End Function



Public Function ComprobarStock(codArtic As String, codAlmac As String, cant As String, CodTipMov As String) As Boolean
'Comprueba si el Articulo existe en el Almacen Origen y si hay
'stock suficiente para poder realizar el traspaso
Dim vStock As String
Dim vArtic As CArticulo
Dim b As Boolean

    Set vArtic = New CArticulo
    b = vArtic.Existe(codArtic)
    If b Then
        b = vArtic.ExisteEnAlmacen(codAlmac, vStock)
        If b Then
            b = ComprobarHayStock(CSng(vStock), CSng(cant), codArtic, vArtic.NOMBRE, CodTipMov)
'            If Not ComprobarHayStock(CSng(vStock), CSng(cant), codArtic, vArtic.Nombre, CodTipMov) Then
'                b = False
'            Else
'                b = True
'            End If
        End If
    End If
    Set vArtic = Nothing
    ComprobarStock = b
End Function



Public Function ObtenerPrecioSinIVAvarios(codArtic As String, Precio As String) As Currency
Dim vArtic As CArticulo
Dim PreuSinIVA  As Currency

'    On Error GoTo ErrTotal
'
''    If sPorce <> "" Then curPorce = ImporteFormateado(sPorce)
'    If Precio <> "" Then PreuConIVA = ImporteFormateado(Precio) 'precio con iva

    Set vArtic = New CArticulo
    If vArtic.LeerDatos(codArtic) Then
        'precio con iva del articulo
        PreuSinIVA = vArtic.ObtenerPrecioSinIVA(Precio)
    Else
        PreuSinIVA = CCur(ComprobarCero(Precio))
    End If

'
'
'    curPorce = curPorce / 100
'    curImporte = curImporte / (1 + curPorce) 'importe sin iva
'    curCuota = Round((curPorce * curImporte), 2)
'    curImporte = Round(curImporte, 2)
'
'    'valores que devuelve: Importe sin iva, cuota de iva
'    ImporteSinIVA = Format(curImporte, FormatoImporte)
'    sCuota = Format(curCuota, FormatoImporte)
'
'    Exit Function


'    Set vArtic = New CArticulo
'    If vArtic.LeerDatos(codArtic) Then
'        'precio con iva del articulo
'        PreuIVA = vArtic.ObtenerPrecioConIVA
'    End If
'
'
'    'El precio con IVA calculado a partir del importe del articulo no coincide con el
'    'precio con IVA introducido en la linea.
'    'recalculamos el importe del articulo SIN iva (se modifica precio original del artic)
'    If Round(PreuIVA, 2) <> Round(CCur(Precio), 2) Then
'        If PreuIVA <> 0 Then
'            PreuIVA = Round((vArtic.PrecioVenta * CCur(Precio)) / PreuIVA, 4)
'        Else
'            PreuIVA = Round((CCur(Precio) * 100) / (100 + vArtic.ObtenerPorceIVA), 4)
'        End If
'    Else
'        PreuIVA = vArtic.PrecioVenta
'    End If
    Set vArtic = Nothing
    ObtenerPrecioSinIVAvarios = PreuSinIVA
End Function




 



Public Function TipoCamp(ByRef objec As Object) As String
Dim mTag As CTag
Dim Cad As String

    On Error GoTo ETipoCamp

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        TipoCamp = mTag.TipoDato
    End If

ETipoCamp:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


Public Function CApos(Texto As String) As String
'-- (RAFA/ALZIRA) 07092006
'-- Esta funci�n procesa caracteres extra�os y de control para sentencias SQL

    Dim i As Integer
    Dim i2 As Integer
    i2 = 1
    i = InStr(i2, Texto, "'")
    While i <> 0
        Texto = Mid(Texto, 1, i) & "'" & Mid(Texto, i + 1, Len(Texto) - i)
        i2 = i + 2
        i = InStr(i2, Texto, "'")
    Wend
    i2 = 1
    i = InStr(i2, Texto, "\")
    While i <> 0
        Texto = Mid(Texto, 1, i) & "\" & Mid(Texto, i + 1, Len(Texto) - i)
        i2 = i + 2
        i = InStr(i2, Texto, "\")
    Wend

End Function



Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim Ent As Integer
Dim Cad As String
  
  ' Comprobaciones
  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un n�mero."
    Exit Function
  End If
  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If
  
  ' Redondeo.
  Cad = "0"
  If NumDigitsAfterDecimals <> 0 Then Cad = Cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Format(Number, Cad)
  
End Function



Public Function CalcularPorcentaje(Importe As Currency, Porce As Currency, NumDecimales As Long) As Variant
'devuelve el valor del Porcentaje aplicado al Importe
'Ej el 16% de 120 = 19.2
'Dim vImp As Currency
'Dim vDto As Currency
    
    On Error Resume Next
'
'    Importe = ComprobarCero(Importe)
'    Dto = ComprobarCero(Dto)
'
'    vImp = CCur(Importe)
'    vDto = CCur(Dto)
    
    
    'vImp = Round(vImp, 2)
    
    CalcularPorcentaje = Round2((Importe * Porce) / 100, NumDecimales)
    
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function CalcularPorcentaje2(Importe As Currency, Porce As Currency, NumDecimales As Long) As Variant
'devuelve el valor del Porcentaje incluido en el Importe
'Ej el 16% de 120 = 19.2
'Dim vImp As Currency
'Dim vDto As Currency
    
    On Error Resume Next
'
'    Importe = ComprobarCero(Importe)
'    Dto = ComprobarCero(Dto)
'
'    vImp = CCur(Importe)
'    vDto = CCur(Dto)
    
    
    'vImp = Round(vImp, 2)
    
    CalcularPorcentaje2 = Round2(Importe / (1 + (Porce / 100)), NumDecimales)
    
    If Err.Number <> 0 Then Err.Clear
End Function






Public Function ArticuloTieneMargen(codArt As String) As Boolean
Dim Cad As String

    'Comprobar que el art�culo tiene margen comercial
    Cad = DevuelveDesdeBDNew(conAri, "sartic", "margecom", "codartic", codArt, "T")
    If Cad = "" Then
        Cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
        Cad = Cad & "El art�culo no tiene margen comercial para calcular nuevos precios."
        MsgBox Cad, vbExclamation
        ArticuloTieneMargen = False
        Exit Function
    End If
    
    
'    'comprobar que las tarifas del articulo tienen margen comercial
'    cad = "SELECT count(*)"
'    cad = cad & " FROM slista INNER JOIN starif ON slista.codlista = starif.codlista "
'    cad = cad & " WHERE slista.codartic=" & DBSet(codArt, "T") & " AND  isnull(margecom)"
'    If RegistrosAListar(cad) > 0 Then
'        cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
'        cad = cad & "El art�culo tiene tarifas sin %PVP necesario para calcular nuevos precios."
'        MsgBox cad, vbExclamation
'        ArticuloTieneMargen = False
'        Exit Function
'    End If
    
    ArticuloTieneMargen = True
    
End Function






Public Function TotalRegistros(vSQL As String, Optional vBD As Byte) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    If vBD = conConta Then 'Accede a BD de contabilidad
        Rs.Open vSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    
    TotalRegistros = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalRegistros = Rs.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function

'---------------------------------------------------------------------------------
'
'       Para buscar en los checks con las dos opciones de true y false
'
'A partir de un check cualquiera devolvera nombre e indice, si tiene. Si no sera ()
Public Sub CheckBusqueda(ByRef CH As CheckBox)
    NombreCheck = ""
    NombreCheck = CH.Name & "("
    On Error Resume Next
    NombreCheck = NombreCheck & CH.Index
    If Err.Number <> 0 Then Err.Clear
    NombreCheck = NombreCheck & ")"
End Sub



Public Sub CheckCadenaBusqueda(ByRef CH As CheckBox, ByRef CadenaCHECKs As String)
        CheckBusqueda CH
        If InStr(1, CadenaCHECKs, NombreCheck) = 0 Then CadenaCHECKs = CadenaCHECKs & NombreCheck & "|"
End Sub




'---------------------------------------------------------------------------------
'
'       Las tabla reparaciones esta relacionada, sin FOREING KEY con
'       SAT, tipoave,trabajorealizado
'       Para saber si se puede eliminar alguno de estos
'       mantenimientos entonces trendrmos esta funcion
'
'       Opcion
'           1:  sat
'           2:  tipoave
'           3:  trabajaorealizado
Public Function SePuedeEliminarRelReparacione(Opcion As Byte, Codigo As String) As Boolean
Dim CA As String
Dim C2 As String

    SePuedeEliminarRelReparacione = False
    If Opcion = 1 Then
        'SAT
        CA = "codman"
    Else
        If Opcion = 2 Then
            CA = "codavi" 'Deberia haber sido AVE de averia, no avi
        Else
            CA = "codtrabajo"
        End If
    End If
    'Miramos primero en scarep
    C2 = DevuelveDesdeBDNew(conAri, "scarep", "numrepar", CA, Codigo, "N")
    If C2 <> "" Then Exit Function
        
        
    'Ahora miraremos en hco reparaciones
    C2 = DevuelveDesdeBDNew(conAri, "schrep", "numrepar", CA, Codigo, "N")
    If C2 <> "" Then Exit Function

    
    SePuedeEliminarRelReparacione = True
End Function

Public Function SugerirCodAutomatico(marca As String, categoria As String, modelo As String, Formato As String) As String
    '-- SugerirCodAtomatico:
    '   Esta funci�n se utiliza en el marco del par�metro descriptores y sirve, al igual que se montaba un descriptor
    '   autom�tico a partir de las descripciones de los campos de marca, categoria, modelo y formato; hacer lo propio
    '   pero con el c�digo. Con el siguiente formato
    '   MMMMCCCCmmffXXXX -> M=marca, C=categoria, m=modelo, f=formato, x=un ordinal para el c�digo
    Dim inferior As String
    Dim superior As String
    Dim comun As String
    Dim Codigo As String
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim Valor As Integer
    '-- Primero trimeamos los valores por si acaso.
    marca = Left(Trim(marca) & "0000", 4)
    categoria = Left(Trim(categoria) & "0000", 4)
    modelo = Left(Trim(modelo) & "00", 2)
    Formato = Left(Trim(Formato) & "00", 2)
    '--
    comun = marca & categoria & modelo & Formato
'    inferior = comun & "0000"
'    superior = comun & "9999"
'
'    SQL = "select max(codartic) from sartic where" & _
'            " codartic >= '" & inferior & "'" & _
'            " and codartic <= '" & superior & "'"
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly
'    '-- por defecto el c�digo es:
'    codigo = comun & "0001"
'    If Not RS.EOF Then
'        If Not IsNull(RS.Fields(0)) Then
'            If Not IsNumeric(Right(RS.Fields(0), 4)) Then
'                MsgBox "La cola de c�digo: " & RS.Fields(0) & " no es num�rica. No puedo sugerir el c�digo siguiente", vbExclamation
'                codigo = ""
'            Else
'                Valor = Val(Right(RS.Fields(0), 4)) + 1
'                codigo = comun & Format(Valor, "0000")
'            End If
'        End If
'    End If
'    SugerirCodAutomatico = codigo
    SugerirCodAutomatico = comun
End Function

Public Function CambiaTagDescriptores(ByRef txt As TextBox, descriptor As String) As String
    '-- Cambia el comienzo del tag del descriptor en el tag, para que cuando diga xxx no exista, aparezca
    '   la etiqueta correcta.
    Dim pos As Integer
    Dim ntag As String
    ntag = txt.Tag
    pos = InStr(1, ntag, "|")
    If pos Then
        ntag = descriptor & Mid(ntag, pos, (Len(ntag) - pos) + 1)
    End If
    txt.Tag = ntag
    CambiaTagDescriptores = ntag
End Function


'                                                                       CINCO DECIMALES
Public Function ArticuloConTasaReciclado(ArticuloLinea As String, ByRef ImporteSng As Single) As Boolean
Dim RT As ADODB.Recordset
Dim Sql As String
        On Error GoTo EArticuloConTasaReciclado
        ArticuloConTasaReciclado = False
        Sql = "select tasareciclado from sunida,sartic where sunida.codunida =sartic.codunida and sartic.codartic=" & DBSet(ArticuloLinea, "T")
        Set RT = New ADODB.Recordset
        RT.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RT.EOF Then
            If Not IsNull(RT!tasareciclado) Then
                ImporteSng = RT!tasareciclado
                ArticuloConTasaReciclado = True
            End If
        End If
        RT.Close
        Set RT = Nothing
        Exit Function
EArticuloConTasaReciclado:
    MuestraError Err.Number, Err.Description, "Calculando tasa reciclado."
    Set RT = Nothing
End Function



Public Function DevuelveUltimoAlmacen(Tabla As String, Where As String) As Integer
Dim C As String
Dim Rs As ADODB.Recordset

    DevuelveUltimoAlmacen = -1
    C = "Select codalmac FROM " & Tabla & Where & " ORDER BY numlinea DESC"
    Set Rs = New ADODB.Recordset
    Rs.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then DevuelveUltimoAlmacen = CInt(Rs.Fields(0))
    End If
    Rs.Close
    Set Rs = Nothing
End Function




Public Function TotalRegistrosConsulta(cadSQL) As Long
Dim Cad As String
Dim Rs As ADODB.Recordset

    On Error GoTo ErrTotReg
    Cad = "SELECT count(*) FROM (" & cadSQL & ") x"
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not Rs.EOF Then
        TotalRegistrosConsulta = DBLet(Rs.Fields(0).Value, "N")
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
ErrTotReg:
    MuestraError Err.Number, "", Err.Description
End Function


Public Function DevuelveValor(vSQL As String) As Variant
'Devuelve el valor de la SQL
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    DevuelveValor = 0
    If Not Rs.EOF Then
        ' antes RS.Fields(0).Value > 0
        If Not IsNull(Rs.Fields(0).Value) Then DevuelveValor = Rs.Fields(0).Value   'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        DevuelveValor = 0
        Err.Clear
    End If
End Function

Public Sub RecalcularImportes(TImporte As TextBox, CalculoImportes As Boolean, Optional TCantidad As TextBox, Optional TPrecio As TextBox, Optional TDto1 As TextBox, Optional TDto2 As TextBox)
'CalculoImportes = true cuando vengo de albaranes con precio, cantidad y dtos o fras normales donde modifico el precio
'                = false cuando vengo de facturas de importe de llamadas
Dim Aux As String

    If CalculoImportes Then
        If Not PonerFormatoDecimal(TCantidad, 1) Then TCantidad.Text = ""
        If Not PonerFormatoDecimal(TPrecio, 2) Then TPrecio.Text = ""
        If Not PonerFormatoDecimal(TDto1, 4) Then TDto1.Text = ""
        If Not PonerFormatoDecimal(TDto2, 4) Then TDto2.Text = ""
    
        Aux = CalcularImporte(TCantidad.Text, TPrecio.Text, TDto1.Text, TDto2.Text, vParamAplic.TipoDtos)
        Aux = Format(Aux, FormatoImporte)
        If Aux <> TImporte.Text Then TImporte.Text = Aux
    Else
        If Not PonerFormatoDecimal(TImporte, 1) Then TImporte.Text = ""
    End If
End Sub

Public Sub PonerContRegIndicador(ByRef lblIndicador As Label, ByRef vData As Adodc, cadBuscar As String)
'cuando esta en el MODO 2 pone el label de contador de registros a�adiendo
'la palabra 'Busqueda' si es el resultado de una busqueda
'devolvera: "1 de 20" o "BUSQUEDA: 1 de 20"
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    cadReg = PonerContRegistros(vData) 'devuelve: "1 de 20"
    
    If cadBuscar = "" Or cadReg = "" Then
        lblIndicador.Caption = cadReg
    Else
        lblIndicador.Caption = "BUSQUEDA: " & cadReg
    End If
End Sub


Public Function ObtenerBusquedaNew(ByRef formulario As Form, Optional CHECK As String, Optional vBD As Byte, Optional cadWHERE As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim Cad As String
    Dim Sql As String
    Dim Tabla As String
    Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusquedaNew = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                If Control.Tag <> "" Then
                    Carga = mTag.Cargar(Control)
                    If Carga Then
                        If Aux = ">>" Then
                            Cad = " MAX("
                        Else
                            Cad = " MIN("
                        End If
                        'monica
                        Select Case mTag.TipoDato
                            Case "FHF"
                                Cad = Cad & "date(" & mTag.columna & "))"
                            Case "FHH"
                                Cad = Cad & "time(" & mTag.columna & "))"
                            Case Else
                                Cad = Cad & mTag.columna & ")"
                        End Select
                        
                        Sql = "Select " & Cad & " from " & mTag.Tabla
                        If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
                        Sql = ObtenerMaximoMinimoNew(Sql, vBD)
                        Select Case mTag.TipoDato
                        Case "N"
                            Sql = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Case "F"
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case "FHF"
                            Sql = "date(" & mTag.Tabla & "." & mTag.columna & ") = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case "FHH"
                            Sql = "time(" & mTag.Tabla & "." & mTag.columna & ") = '" & Format(Sql, "hh:mm:ss") & "'"
                        Case Else
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Sql & "'"
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next


'++monica: lo he a�adido del anterior obtenerbusqueda
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    Sql = mTag.Tabla & "." & mTag.columna & " is NULL"
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next
 

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                'Cargamos el tag
                Carga = mTag.Cargar(Control)
                If Carga Then
'                    Debug.Print Control.Tag
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.Tabla <> "" Then
                            Tabla = mTag.Tabla & "."
                            Else
                            Tabla = ""
                        End If
                        Rc = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.columna, Aux, Cad)
                        If Rc = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & Cad & ")"
                        End If
                    End If
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex > -1 Then
                        If mTag.TipoDato = "N" Then
                            Cad = Control.ItemData(Control.ListIndex)
                        Else
                            Cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & Cad & ")"
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is CheckBox Then
            '=============== A�ade: Laura, 15/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Aux = ""
                    If CHECK <> "" Then
                        Tabla = DBLet(Control.Index, "T")
                        If Tabla <> "" Then Tabla = "(" & Tabla & ")"
                        Tabla = Control.Name & Tabla & "|"
                        If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        Cad = Control.Value
                        Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & Cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusquedaNew = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusquedaNew = ""
    MuestraError Err.Number, "Obtener b�squeda. " & vbCrLf & Err.Description
End Function

Private Function ObtenerMaximoMinimoNew(vSQL As String, Optional vBD As Byte) As String
Dim Rs As Recordset
    ObtenerMaximoMinimoNew = ""
    Set Rs = New ADODB.Recordset
    If vBD = conConta Then
        Rs.Open vSQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else
        Rs.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            ObtenerMaximoMinimoNew = CStr(Rs.Fields(0))
        End If
    End If
    Rs.Close
    Set Rs = Nothing
End Function


Public Function TieneChofer(Socio As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset


    TieneChofer = False
    Set Rs = New ADODB.Recordset

    Sql = "select * from sclien_chofer where codsocio=" & DBSet(Socio, "N") & " and (fechabaj is null or fechabaj <= '" & Format(Date, FormatoFecha)
    Sql = Sql & "')"
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        TieneChofer = True
    End If
    Rs.Close
    Set Rs = Nothing
End Function


Public Function MiFormat(Valor As String, Formato As String) As String
    If Trim(Valor) = "" Then
       MiFormat = ""
    Else
        If Formato = "" Then
            MiFormat = Valor
        Else
            MiFormat = Format(Valor, Formato)
        End If
    End If
End Function



Public Sub AyudaTiposDocumentos(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 1405 + 3000 + 2595  hay que quitarle al width 0
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|1405|;S|txtAux(1)|T|Descripci�n|3000|;S|txtAux(2)|T|Fichero|2595|;"
    frmBas.CadenaConsulta = "SELECT scryst.codcryst, scryst.nomcryst, scryst.documrpt "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM scryst "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo Documento|N|N|||scryst|codcryst|0000|S|"
    frmBas.Tag2 = "Descripci�n|T|N|||scryst|nomcryst|||"
    frmBas.Tag3 = "Fichero rpt|T|N|||scryst|documrpt|||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 100
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "scryst"
    frmBas.CampoCP = "codcryst"
    frmBas.Caption = "Tipos de Documentos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
End Sub


Public Sub AyudaFamiliasArticulos(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|905|;S|txtAux(1)|T|Descripci�n|4595|;"
    frmBas.CadenaConsulta = "SELECT sfamia.codfamia, sfamia.nomfamia "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM sfamia "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo|N|N|0|999|sfamia|codfamia|000|S|"
    frmBas.Tag2 = "Descripci�n|T|N|||sfamia|nomfamia|||"
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 130
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "sfamia"
    frmBas.CampoCP = "codfamia"
    frmBas.Caption = "Familias de Art�culos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub

Private Sub Redimensiona(frmBas As frmBasico2, cant As Integer)
    frmBas.Width = frmBas.Width + cant
    frmBas.DataGrid1.Width = frmBas.DataGrid1.Width + cant
    frmBas.cmdAceptar.Left = frmBas.cmdAceptar.Left + cant
    frmBas.cmdCancelar.Left = frmBas.cmdCancelar.Left + cant
    frmBas.cmdRegresar.Left = frmBas.cmdRegresar.Left + cant

End Sub


Public Sub AyudaFormasPago(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|905|;S|txtAux(1)|T|Descripci�n|4595|;"
    frmBas.CadenaConsulta = "SELECT sforpa.codforpa, sforpa.nomforpa "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM sforpa "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo|N|N|0|999|sforpa|codforpa|000|S|"
    frmBas.Tag2 = "Descripci�n|T|N|||sforpa|nomforpa|||"
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 130
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "sforpa"
    frmBas.CampoCP = "codforpa"
    frmBas.Caption = "Formas de Pago"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub


Public Sub AyudaAgentesComerciales(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|905|;S|txtAux(1)|T|Descripci�n|4595|;"
    frmBas.CadenaConsulta = "SELECT sagent.codagent, sagent.nomagent "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM sagent "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo Agente Comercial|N|N|0|9999|sagent|codagent|0000|S|"
    frmBas.Tag2 = "Nombre del Agente Comercial|T|N|||sagent|nomagent||N|"
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 130
    frmBas.Maxlen3 = 0
    
    
    frmBas.pConn = conAri
    
    
    frmBas.Tabla = "sagent"
    frmBas.CampoCP = "codagent"
    frmBas.Caption = "Agentes Comerciales"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub




Public Sub AyudaArticulos(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 1405 + 3000 + 2595  hay que quitarle al width 0
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|1405|;S|txtAux(1)|T|Descripci�n|3500|;S|txtAux(2)|T|Cod.Asociaci�n|2095|;"
    frmBas.CadenaConsulta = "SELECT sartic.codartic, sartic.nomartic, sartic.codtelem "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM sartic "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo Art�culo|T|N|||sartic|codartic||S|"
    frmBas.Tag2 = "Denominaci�n Art�culo|T|N|||sartic|nomartic||N|"
    frmBas.Tag3 = "C�digo Asociaci�n|T|S|||sartic|codtelem||N|"
    
    frmBas.Maxlen1 = 16
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 18
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "sartic"
    frmBas.CampoCP = "codartic"
    frmBas.Caption = "Art�culos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
End Sub


Public Sub AyudaArticulosInventario(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 1405 + 3000 + 2595  hay que quitarle al width 0
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|1405|;S|txtAux(1)|T|Descripci�n|4595|;"
    frmBas.CadenaConsulta = "SELECT distinct shinve.codartic, sartic.nomartic "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM shinve, sartic "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE shinve.codartic = sartic.codartic "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo Art�culo|T|N|||shinve|codartic||S|"
    frmBas.Tag2 = "Denominaci�n Art�culo|T|N|||sartic|nomartic||N|"
    
    frmBas.Maxlen1 = 16
    frmBas.Maxlen2 = 40
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "shinve"
    frmBas.CampoCP = "codartic"
    frmBas.Caption = "Historico Inventario"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1000
    
    frmBas.Show vbModal
End Sub


Public Sub AyudaBancosPropios(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|905|;S|txtAux(1)|T|Descripci�n|4595|;"
    frmBas.CadenaConsulta = "SELECT sbanpr.codbanpr, sbanpr.nombanpr "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM sbanpr "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo Banco Propio|N|N|0|9999|sbanpr|codbanpr|0000|S|"
    frmBas.Tag2 = "Denominaci�n Banco Propio|T|N|||sbanpr|nombanpr||N|"
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 130
    frmBas.Maxlen3 = 0
    
    
    frmBas.pConn = conAri
    
    
    frmBas.Tabla = "sbanpr"
    frmBas.CampoCP = "codbanpr"
    frmBas.Caption = "Bancos Propios"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub


Public Sub AyudaCartasOferta(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|Codigo|1405|;S|txtAux(1)|T|Nombre|4095|;"
    frmBas.CadenaConsulta = "SELECT scartas.codcarta, scartas.descarta "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM scartas "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Cod. Carta|N|N|0|999|scartas|codcarta|000|S|"
    frmBas.Tag2 = "Descripci�n|T|S|||scartas|descarta||N|"
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 130
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "scartas"
    frmBas.CampoCP = "codcarta"
    frmBas.Caption = "Cartas"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub



Public Sub AyudaChoferes(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|Codigo|1405|;S|txtAux(1)|T|Nombre|4095|;"
    frmBas.CadenaConsulta = "SELECT schofe.codchofe, schofe.nomchofe "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM schofe "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Codigo chofer|N|N|||schofe|codchofe|0000|S|"
    frmBas.Tag2 = "Nombre chofer|T|N|||schofe|nomchofe|||"
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 130
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "schofe"
    frmBas.CampoCP = "codchofe"
    frmBas.Caption = "Choferes"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub


Public Sub AyudaCRMTipos(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|Codigo|1405|;S|txtAux(1)|T|Nombre|4095|;"
    frmBas.CadenaConsulta = "SELECT scrmtipo.codigo, scrmtipo.denominacion "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM scrmtipo "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo|N|N|0||scrmtipo|codigo|0000|S|"
    frmBas.Tag2 = "Denominacion|T|N|||scrmtipo|denominacion||N|"
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 130
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "scrmtipo"
    frmBas.CampoCP = "codigo"
    frmBas.Caption = "Tipos de CRM"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub








Public Sub AyudaCuentasContables(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 1500 + 4000 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|1500|;S|txtAux(1)|T|Descripci�n|4000|;"
    frmBas.CadenaConsulta = "SELECT cuentas.codmacta, cuentas.nommacta "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM cuentas "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo de Cuenta|T|N|||cuentas|codmacta||S|"
    frmBas.Tag2 = "Descripci�n de la Cuenta|T|N|||cuentas|nommacta|||"
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 130
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conConta
    
    frmBas.Tabla = "cuentas"
    frmBas.CampoCP = "codmacta"
    frmBas.Caption = "Cuentas Contables"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub


Public Sub AyudaTiposIva(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 1500 + 4000 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|1500|;S|txtAux(1)|T|Descripci�n|3500|;S|txtAux(2)|T|Porcentaje|2095|;"
    frmBas.CadenaConsulta = "SELECT tiposiva.codigiva,tiposiva.nombriva,tiposiva.porceiva "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM tiposiva "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo Iva|N|N|||tiposiva|codigiva|###0|S|"
    frmBas.Tag2 = "Descripci�n|T|N|||tiposiva|nombriva|||"
    frmBas.Tag3 = "Porcentaje|N|N|||tiposiva|porceiva|##0.00||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 120
    frmBas.Maxlen3 = 10
    
    frmBas.pConn = conConta
    
    frmBas.Tabla = "tiposiva"
    frmBas.CampoCP = "codigiva"
    frmBas.Caption = "Tipos de Iva"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    
    frmBas.Show vbModal
End Sub



Public Sub AyudaMovimientosArticulos(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|1805|;S|txtAux(1)|T|Descripci�n|3495|;"
    frmBas.CadenaConsulta = "SELECT smoval.codartic, sartic.nomartic "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM smoval left join sartic on smoval.codartic = sartic.codartic "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " group by 1 "
    
    frmBas.Tag1 = "C�digo Art�culo|T|N|||smoval|codartic||S|"
    frmBas.Tag2 = "Denominaci�n|T|N|||sartic|nomartic||N|"
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 16
    frmBas.Maxlen2 = 114
    frmBas.Maxlen3 = 0
    
    
    frmBas.pConn = conAri
    
    
    frmBas.Tabla = "smoval"
    frmBas.CampoCP = "codartic"
    frmBas.Caption = "Movimientos Art�culos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub

Public Sub AyudaDirecciones(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 1405 + 3000 + 2595  hay que quitarle al width 0
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|1405|;S|txtAux(1)|T|Tipo|2595|;S|txtAux(2)|T|Nombre|3000|;"
    frmBas.CadenaConsulta = "SELECT coddirec,if(tipodire=0,'Albaran','Factura') tipo,domdirec "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM sdirpr "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Cod. Direcci�n|N|N|0|999|sdirpr|coddirec|000|S|"
    frmBas.Tag2 = "Tipo|T|N|||sdirpr|tipo|||"
    frmBas.Tag3 = "Nombre Direcci�n|T|N|||sdirpr|nomdirec||N|"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 100
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "sdirpr"
    frmBas.CampoCP = "coddirec"
    frmBas.Caption = "Direccciones"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
End Sub

Public Sub AyudaProveedoresV(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|Nif|1405|;S|txtAux(1)|T|Nombre|4095|;"
    frmBas.CadenaConsulta = "SELECT sprvar.nifprove, sprvar.nomprove "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM sprvar "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo|T|N|||sprvar|nifprove|000|S|"
    frmBas.Tag2 = "Nombre|T|N|||sprvar|nomprove|||"
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 130
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "sprvar"
    frmBas.CampoCP = "nifprove"
    frmBas.Caption = "Proveedores Varios"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub



Public Sub AyudaProveedores(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|1000|;S|txtAux(1)|T|Nombre|4000|;S|txtAux(2)|T|Nombre Comercial|4000|;"
    frmBas.CadenaConsulta = "SELECT sprove.codprove, sprove.nomprove, sprove.nomcomer "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM sprove "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo Proveedor|N|N|0|999999|sprove|codprove|000000|S|"
    frmBas.Tag2 = "Nombre Proveedor|T|N|||sprove|nomprove||N|"
    frmBas.Tag3 = "Nombre Comercial|T|N|||sprove|nomcomer||N|"
    
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 75
    frmBas.Maxlen3 = 75
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "sprove"
    frmBas.CampoCP = "codprove"
    frmBas.Caption = "Proveedores"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 2010
    
    
    frmBas.Show vbModal
End Sub



Public Sub AyudaTrabajadores(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 1000+4000+2000 no hay que quitarle al width nada
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|1000|;S|txtAux(1)|T|Nombre|4000|;S|txtAux(2)|T|Nif|2000|;"
    frmBas.CadenaConsulta = "SELECT straba.codtraba, straba.nomtraba, straba.niftraba "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM straba "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "C�digo|N|N|||straba|codtraba|0000|S|"
    frmBas.Tag2 = "Nombre|T|N|||straba|nomtraba||N|"
    frmBas.Tag3 = "Nif|T|N|||straba|niftraba||N|"
    
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 15
    
    
    frmBas.pConn = conAri
    
    
    frmBas.Tabla = "straba"
    frmBas.CampoCP = "codtraba"
    frmBas.Caption = "Trabajadores"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    frmBas.Show vbModal
End Sub



Public Sub AyudaCentroCoste(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|905|;S|txtAux(1)|T|Descripci�n|4595|;"
    If vParamAplic.ContabilidadNueva Then
        frmBas.CadenaConsulta = "SELECT ccoste.codccost, ccoste.nomccost "
        frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM ccoste "
        frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
        If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
        frmBas.Tag1 = "C�digo|T|N|||ccoste|codccost||S|"
        frmBas.Tag2 = "Descripci�n|T|N|||ccoste|nomccost|||"
    Else
        frmBas.CadenaConsulta = "SELECT cabccost.codccost, cabccost.nomccost "
        frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM cabccost "
        frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
        If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
        frmBas.Tag1 = "C�digo|T|N|||cabccost|codccost||S|"
        frmBas.Tag2 = "Descripci�n|T|N|||cabccost|nomccost|||"
    
    End If
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 130
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conConta
    If vParamAplic.ContabilidadNueva Then
        frmBas.Tabla = "ccoste"
    Else
        frmBas.Tabla = "cabccost"
    End If
    frmBas.CampoCP = "codccost"
    frmBas.Caption = "Centros de Coste"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub


Public Sub AyudaDiarios(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|905|;S|txtAux(1)|T|Descripci�n|4595|;"
    frmBas.CadenaConsulta = "SELECT tiposdiario.numdiari, tiposdiario.desdiari "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM tiposdiario "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo|N|N|||tiposdiario|numdiari|000|S|"
    frmBas.Tag2 = "Descripci�n|T|N|||tiposdiario|desdiari|||"
    frmBas.Tag3 = ""
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 130
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conConta
    frmBas.Tabla = "tiposdiario"
    frmBas.CampoCP = "numdiari"
    frmBas.Caption = "Tipos de Diario"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
End Sub


Public Sub AyudaClientes(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 1405 + 3000 + 2595  hay que quitarle al width 0
    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|1005|;S|txtAux(1)|T|Nombre|3400|;S|txtAux(2)|T|Nombre Comercial|2595|;"
    frmBas.CadenaConsulta = "SELECT scliente.codclien, scliente.nomclien, scliente.nomcomer "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM scliente "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "C�digo|N|N|||scliente|codclien|000000|S|"
    frmBas.Tag2 = "Nombre|T|N|||scliente|nomclien|||"
    frmBas.Tag3 = "Nombre Comercial|T|N|||scliente|nomcomer||N|"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 100
    
    frmBas.pConn = conAri
    
    frmBas.Tabla = "scliente"
    frmBas.CampoCP = "codclien"
    frmBas.Caption = "Clientes"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
End Sub


Public Function EsTeletaxi() As Boolean
Dim Sql As String

    EsTeletaxi = ((vParamAplic.Cooperativa = 0) And (Trim(vParam.CifEmpresa) <> "B98877806"))
    
End Function


Public Function Jason_GET(CADENA As String) As String
Dim httpURL As WinHttp.WinHttpRequest
Dim vJson As String

    Jason_GET = ""
 
    Set httpURL = New WinHttp.WinHttpRequest
    vJson = ""
    httpURL.Open "GET", CADENA
    httpURL.send vJson
    Jason_GET = httpURL.responseText
    
End Function



Public Function Jason_POST(CADENA As String) As String

Dim httpURL As WinHttp.WinHttpRequest
Dim vJson As String

    Jason_POST = ""
 
    Set httpURL = New WinHttp.WinHttpRequest
    
    vJson = ""
    httpURL.Open "POST", CADENA, False
    'vJson = Text3.Text
    httpURL.setRequestHeader "Content-Type", "application/json"
    httpURL.send vJson

    Jason_POST = httpURL.responseText
    
End Function





