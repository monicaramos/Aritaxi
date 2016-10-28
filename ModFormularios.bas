Attribute VB_Name = "ModFormularios"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'============================================================
'====== FUNCIONES GENERALES  ================================


'======== A�ade: Laura

'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Public Function SugerirCodigoSiguienteStr(NomTabla As String, NomCodigo As String, Optional CondLineas As String) As String
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo ESugerirCodigo

    'SQL = "Select Max(codtipar) from stipar"
    SQL = "Select Max(" & NomCodigo & ") from " & NomTabla
    If CondLineas <> "" Then
        SQL = SQL & " WHERE " & CondLineas
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, , , adCmdText
    SQL = "1"
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If IsNumeric(RS.Fields(0)) Then
                SQL = CStr(RS.Fields(0) + 1)
            Else
                If Asc(Left(RS.Fields(0), 1)) <> 122 Then 'Z
                SQL = Left(RS.Fields(0), 1) & CStr(Asc(Right(RS.Fields(0), 1)) + 1)
                End If
            End If
        End If
    End If
    RS.Close
    Set RS = Nothing
    SugerirCodigoSiguienteStr = SQL
ESugerirCodigo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Public Function EsCodigoCero(Cod As String, Formato As String) As Boolean
'comprueba que para algunas tablas en las que el codigo 0000 se reserva para
'un valor gen�rico no se modifique ni se borre

    EsCodigoCero = False
    If Cod <> "" Then
        If Val(Cod) = Val(0) Then
            EsCodigoCero = True
            MsgBox "El c�digo " & Formato & " no se puede modificar ni eliminar.", vbExclamation
            Screen.MousePointer = vbDefault
        End If
    End If
End Function


Public Sub BloquearText1(ByRef formulario As Form, Modo As Byte)
'Bloquea controles q se llamen TEXT1 si no estamos en Modo: 3.-Insertar, 4.-Modificar
'si estamos en modo modificar bloquea solo los campos que son clave primaria
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualizaci�n
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)
Dim i As Byte
Dim b As Boolean
Dim vtag As CTag
On Error Resume Next

    With formulario
        b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'And ModoLineas = 1))
        
        For i = 0 To .Text1.Count - 1 'En principio todos los TExt1 tiene TAG
            Set vtag = New CTag
            vtag.Cargar .Text1(i)
            If vtag.Cargado Then
                If vtag.EsClave And (Modo = 2 Or Modo = 4 Or Modo = 5) Then
                    .Text1(i).Locked = True
                    .Text1(i).BackColor = &H80000018 'amarillo claro
                Else
                    .Text1(i).Locked = Not b  '((Not b) And (Modo <> 1))
                    If b Then
                        .Text1(i).BackColor = vbWhite
                    Else
                        .Text1(i).BackColor = &H80000018 'amarillo claro
                    End If
                    If Modo = 3 Then .Text1(i).Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            Else
                .Text1(i).Locked = Not b  '((Not b) And (Modo <> 1))
                If b Then
                    .Text1(i).BackColor = vbWhite
                Else
                    .Text1(i).BackColor = &H80000018 'amarillo claro
                End If
            End If
        Set vtag = Nothing
        Next i
        
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearTxt(ByRef Text As TextBox, b As Boolean, Optional EsContador As Boolean)
'Bloquea un control de tipo TextBox
'Si lo bloquea lo pone de color amarillo claro sino lo pone en color blanco (sino es contador)
'pero si es contador lo pone color azul claro
On Error Resume Next

    Text.Locked = b
    If Not b And Text.Enabled = False Then Text.Enabled = True
    If b Then
        If EsContador Then
            'Si Es un campo que se obtiene de un contador poner color azul
'            Text.BackColor = &H80000013 'Azul Claro
            Text.BackColor = &HFFFFC0   'Azul claro con vista
        Else
            Text.BackColor = &H80000018 'Amarillo Claro
        End If
    Else
        Text.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearImg(ByRef imgF As Image, b As Boolean)
On Error Resume Next

    imgF.Enabled = Not b
    imgF.visible = Not b
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearCmb(ByRef cmb As ComboBox, b As Boolean, Optional EsContador As Boolean)
'Bloqueja un control de tipo ComboBox
'Si el bloqueja el posa de color gris claro, sino el posa de color blanc (sino es contador)
'pero si es contador el posa color blau clar
    On Error Resume Next

    cmb.Locked = b
    cmb.Enabled = True
    
    'cmb.Enabled = Not b
    
    'If Not b And Cmb.Enabled = False Then Cmb.Enabled = True
    If b Then
        If EsContador Then
            'Si Es un campo que se obtiene de un contador poner color azul
            cmb.BackColor = &H80000013 'Azul Claro
        Else
            cmb.BackColor = &H80000018 'Amarillo Claro
        End If
    Else
        cmb.BackColor = vbWhite
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub BloquearChecks(ByRef formulario As Form, Modo As Byte)
'Bloquea controles  CheckBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualizaci�n
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)
Dim b As Boolean
Dim Control As Control
On Error Resume Next

    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
    With formulario
        For Each Control In formulario.Controls
            If TypeOf Control Is CheckBox Then
                If Control.Name <> "chkVistaPrevia" Then
                    'modo Insertar o modificar
                    If Modo = 3 Or Modo = 4 Then
                        If Control.Value = 2 Then Control.Value = 1
                    End If
                    'modo consulta
                    If Modo = 0 Or Modo = 2 Then
                        If Control.Value = 1 Then Control.Value = 2
                    End If
                    'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                    If (Modo = 3 Or Modo = 1) Then Control.ListIndex = -1
                End If
            End If
        Next Control
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BloquearChk(ByRef ChK As CheckBox, b As Boolean)
'Bloquea un control de tipo CheckBox
'(IN) b : sera true o false segun si bloquea o no
    On Error Resume Next

    ChK.Enabled = Not b
   
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub PonerLongCamposGnral(ByRef formulario As Form, Modo As Byte, Opcion As Byte)
    Dim i As Integer
    
    On Error Resume Next

    With formulario
        If Modo = 1 Then 'BUSQUEDA
            Select Case Opcion
                Case 1 'Para los TEXT1
                    For i = 0 To .Text1.Count - 1
                        With .Text1(i)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = (.HelpContextID * 2) + 1 'el doble + 1
                            End If
                        End With
                    Next i
                
                Case 3 'para los TXTAUX
                    For i = 0 To .txtAux.Count - 1
                        With .txtAux(i)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = (.HelpContextID * 2) + 1 'el doble + 1
                            End If
                        End With
                    Next i
            End Select
            
        Else 'resto de modos
            Select Case Opcion
                Case 1
                    For i = 0 To .Text1.Count - 1
                        With .Text1(i)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next i
                Case 3
                    For i = 0 To .txtAux.Count - 1
                        With .txtAux(i)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next i
            End Select
        End If
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub
 

Public Sub CargarICO(btn As CommandButton, Nombre As String)
    On Error Resume Next
    btn.Picture = LoadPicture(App.Path & "\iconos\" & Nombre)
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub DesplazamientoData(ByRef vData As Adodc, Index As Integer)
'Para desplazarse por los registros de control Data
    If vData.Recordset.EOF Then Exit Sub
    Select Case Index
        Case 0 'Primer Registro
            If Not vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 1 'Anterior
            vData.Recordset.MovePrevious
            If vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 2 'Siguiente
            vData.Recordset.MoveNext
            If vData.Recordset.EOF Then vData.Recordset.MoveLast
        Case 3 'Ultimo
            vData.Recordset.MoveLast
    End Select
End Sub




'===========================
Public Function SituarData(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String) As Boolean
'Situa un DataControl en el registo que cumple vwhere
'para cuando la clave primaria esta formada por 1 campo
On Error GoTo ESituarData
        'Actualizamos el recordset
        vData.Refresh
        vData.Recordset.MoveFirst
        'El sql para que se situe en el registro en especial es el siguiente
        vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then GoTo ESituarData
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarData = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        If vData.Recordset.RecordCount > 0 Then vData.Recordset.MoveFirst
        SituarData = False
End Function


'===========================
Public Function SituarDataPosicion(ByRef vData As Adodc, NumPos As Long, ByRef Indicador As String) As Boolean
'Situa un DataControl en el registro que ocupa la posicion NumPos
Dim TotalReg As Long
On Error GoTo ESituarDataPosicion
        'Actualizamos el recordset
'        vData.Refresh  'Refresh al cargar el grid

        TotalReg = vData.Recordset.RecordCount
        If vData.Recordset.EOF Then GoTo ESituarDataPosicion
        If NumPos <= TotalReg Then
            vData.Recordset.Move NumPos - 1
        Else
'            vData.Recordset.Move NumPos
            vData.Recordset.MoveLast
        End If
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataPosicion = True
        Exit Function
ESituarDataPosicion:
        If Err.Number <> 0 Then Err.Clear
        SituarDataPosicion = False
End Function


'===========================
Public Function SituarDataMULTI(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String) As Boolean
'Situa un DataControl en el registo que cumple vwhere
On Error GoTo ESituarData
        'Actualizamos el recordset
        vData.Refresh
        'El sql para que se situe en el registro en especial es el siguiente
        Multi_Find vData.Recordset, vWhere
        'vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then GoTo ESituarData
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataMULTI = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarDataMULTI = False
End Function



'===========================
Public Function SituarRSetMULTI(ByRef vData As ADODB.Recordset, vWhere As String) As Boolean
'Situa un ADODB.Recordset en el registo que cumple vwhere
On Error GoTo ESituarData
    
        'El sql para que se situe en el registro en especial es el siguiente
        Multi_Find2 vData, vWhere
        If vData.EOF Or vData.BOF Then GoTo ESituarData
        
        SituarRSetMULTI = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarRSetMULTI = False
End Function



Public Sub Multi_Find(ByRef oRs As ADODB.Recordset, sCriteria As String)
'para el situarDataMULTI
On Error Resume Next
    Dim clone_rs As ADODB.Recordset
    Set clone_rs = oRs.Clone
    
    clone_rs.Filter = sCriteria
    
    If clone_rs.EOF Or clone_rs.BOF Then
     oRs.MoveLast
     oRs.MoveNext
    Else
     oRs.Bookmark = clone_rs.Bookmark
    End If
    
    clone_rs.Close
    Set clone_rs = Nothing
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub Multi_Find2(ByRef oRs As ADODB.Recordset, sCriteria As String)
'para el situarDataMULTI
On Error Resume Next

    oRs.Filter = ""
    oRs.MoveFirst
    oRs.Filter = sCriteria
    
    If oRs.EOF Or oRs.BOF Then
     oRs.MoveLast
     oRs.MoveNext
    Else
'        x = oRs.AbsolutePosition
'     oRs.Bookmark = clone_rs.Bookmark
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function SituarDataTrasEliminar(ByRef vData As Adodc, NumReg, Optional NoActualiza As Boolean) As Boolean
'NumReg: numero de registro que acabo de eliminar
'NoActualiza: si se hace el refresh o no, por defecto siempre se hace el refresh
'             pero si hemos eliminado de un Grid ya se hizo en el cargaGrid y
'             no lo volvemos a hacer para mantener las columnas bien.

    On Error GoTo ESituarDataElim
    
        If NoActualiza = False Then vData.Refresh
        
        If Not vData.Recordset.EOF Then    'Solo habia un registro
            If NumReg > vData.Recordset.RecordCount Then
                vData.Recordset.MoveLast
            Else
                vData.Recordset.MoveFirst
                vData.Recordset.Move NumReg - 1
            End If
            SituarDataTrasEliminar = True
        Else
            SituarDataTrasEliminar = False
        End If
        
ESituarDataElim:
    If Err.Number <> 0 Then
        Err.Clear
        SituarDataTrasEliminar = False
    End If
End Function


Public Sub PonerFoco(ByRef Text As TextBox)
On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoBtn(ByRef btn As CommandButton)
On Error Resume Next
    If btn.visible Then btn.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoChk(ByRef ChK As CheckBox)
On Error Resume Next
    ChK.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoCbo(ByRef Cbo As ComboBox)
On Error Resume Next
    Cbo.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub PonerFocoGrid(ByRef DGrid As DataGrid)
    On Error Resume Next
    DGrid.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub ConseguirFoco(ByRef Text As TextBox, Modo As Byte, Optional cadkey As Integer)
'Acciones que se realizan en el evento:GotFocus de los TextBox:Text1
'en los formularios de Mantenimiento
On Error Resume Next

    If Modo = 5 Then Exit Sub
    
    If (Modo <> 0 And Modo <> 2) Then
        If Modo = 1 Then
            Text.BackColor = vbYellow  'Modo 1: Busqueda
        Else
            If Text.Locked Then 'si el control esta bloqueado pasamos el foco al sig. campo
                Text.BackColor = &H80000018 'amarillo claro
                 If cadkey = 0 Then cadkey = 40
                 KEYdown cadkey
                 Exit Sub
            Else
                Text.BackColor = vbWhite
            End If
        End If
        Text.SelStart = 0
        Text.SelLength = Len(Text.Text)
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub ConseguirFocoLin(ByRef Text As TextBox, Optional cadkey As Integer)
'Acciones que se realizan en el evento:GotFocus de los TextBox:TxtAux para LINEAS
'en los formularios de Mantenimiento
On Error Resume Next

    With Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    'si el control esta bloqueado pasamos el foco al sig. campo
    If Text.Locked Then
        Text.BackColor = &H80000018 'amarillo claro
        If cadkey = 0 Then cadkey = 40
        KEYdown cadkey
'        Exit Sub

    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Function ObtenerCadKey(actCampo As Integer, sigCampo As Integer) As Integer
    Dim cadkey As Integer

    On Error Resume Next
    
    If actCampo > sigCampo Then
        cadkey = 38 'flecha superior
    Else
        cadkey = 40 'flecha inferior
    End If
    If sigCampo = 0 Then cadkey = 0
    
    ObtenerCadKey = cadkey
    
    If Err.Number <> 0 Then Err.Clear
End Function


Public Sub ConseguirfocoChk(Modo As Byte)
     If Modo = 0 Or Modo = 2 Then
        KEYpressGnral 13, Modo, False
    End If
End Sub


Public Function PerderFocoGnral(ByRef Text As TextBox, Modo As Byte) As Boolean
Dim Comprobar As Boolean
On Error Resume Next

    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
        PerderFocoGnral = False
        Exit Function
    End If
    
    With Text
        'Quitamos blancos por los lados
        .Text = Trim(.Text)
        
        If .BackColor = vbYellow Then .BackColor = vbWhite
        
        'Si no estamos en modo: 3=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion
        If (Modo <> 3 And Modo <> 4 And Modo <> 1) Then
            PerderFocoGnral = False
            Exit Function
        End If
        
        If Modo = 1 Then
            'Si estamos en modo busqueda y contiene un caracter especial no realizar
            'las comprobaciones
            Comprobar = ContieneCaracterBusqueda(.Text)
            If Comprobar Then
                PerderFocoGnral = False
                Exit Function
            End If
        End If
        PerderFocoGnral = True
    End With
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function PerderFocoGnralLineas(ByRef txt As TextBox, ModoLineas As Byte) As Boolean
'Para el LostFocus de los txtAux de Mto de lineas
Dim Comprobar As Boolean

    On Error Resume Next

    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
        PerderFocoGnralLineas = False
        Exit Function
    End If

    With txt
        'Quitamos blancos por los lados
        .Text = Trim(.Text)
        
        If .BackColor = vbYellow Then .BackColor = vbWhite
        
        'Si no estamos en modo: 1=Insertar o 2=Modificar , no hacer ninguna comprobacion
        If (ModoLineas <> 1 And ModoLineas <> 2) Then
            PerderFocoGnralLineas = False
            Exit Function
        End If
        
        If ModoLineas = 4 Then 'Busqueda
            'Si estamos en modo busqueda y contiene un caracter especial no realizar
            'las comprobaciones
            Comprobar = ContieneCaracterBusqueda(.Text)
            If Comprobar Then
                PerderFocoGnralLineas = False
                Exit Function
            End If
        End If
        
        'si el campo esta bloqueado no actualizar campos
        If .Locked Then
            PerderFocoGnralLineas = False
            Exit Function
        End If
        
        PerderFocoGnralLineas = True
    End With
    If Err.Number <> 0 Then Err.Clear
End Function


Public Sub AnyadirLinea(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
On Error Resume Next
    vDataGrid.AllowAddNew = True
    If vData.Recordset.RecordCount > 0 Then
        vDataGrid.HoldFields
        vData.Recordset.MoveLast
        vDataGrid.Row = vDataGrid.Row + 1
    End If
    
    vDataGrid.Enabled = False
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub DeseleccionaGrid(ByRef vDataGrid As DataGrid)
    On Error GoTo EDeseleccionaGrid

    While vDataGrid.SelBookmarks.Count > 0
        vDataGrid.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
    Err.Clear
End Sub


Public Sub CargaGridGnral(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, SQL As String, PrimeraVez As Boolean)
On Error GoTo ECargaGrid

    vDataGrid.Enabled = True
    '    vdata.Recordset.Cancel
    vData.ConnectionString = conn
    vData.RecordSource = SQL
    vData.CursorType = adOpenDynamic
    vData.LockType = adLockPessimistic
    vDataGrid.ScrollBars = dbgNone
    vData.Refresh
    
    Set vDataGrid.DataSource = vData
    vDataGrid.AllowRowSizing = False
    vDataGrid.AllowUpdate = False
    vDataGrid.RowHeight = 290

    If PrimeraVez Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "CargaGrid", Err.Description
End Sub


Public Sub CargarCombo_SiNo(ByRef Cbo As ComboBox)
'Carga un combo con los valores de opcion SI/NO
    On Error GoTo ErrCarga
    
    Cbo.Clear
    
    Cbo.AddItem "NO"
    Cbo.ItemData(Cbo.NewIndex) = 0
    
    Cbo.AddItem "SI"
    Cbo.ItemData(Cbo.NewIndex) = 1
    
    Exit Sub
    
ErrCarga:
    MuestraError Err.Number, "Cargar combo.", Err.Description
End Sub


Public Sub CargarCombo_Tabla(ByRef Cbo As ComboBox, NomTabla As String, NomCodigo As String, nomDescrip As String, Optional strWhere As String, Optional ItemNulo As Boolean)
'Carga un objeto ComboBox con los registros de una Tabla
'(IN) cbo: ComboBox en el q se van a cargar los datos
'(IN) nomTabla: nombre de la tabla de la q leeremos los datos a cargar
'(IN) nomCodigo: nombre del campo codigo de la tabla q queremos cargar
'(IN) nomDescrip: nombre del campo descripcion de la tabla a cargar
'(IN) strWhere: para filtrar los registros de la tabla q queremos cargar
'(IN) ItemNulo: si es true se a�ade el primer item con linea en blanco
Dim SQL As String
Dim RS As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrCombo
    
    Cbo.Clear
    
    SQL = "SELECT " & NomCodigo & "," & nomDescrip & " FROM " & NomTabla
    If strWhere <> "" Then SQL = SQL & " WHERE " & strWhere
    SQL = SQL & " ORDER BY " & nomDescrip
    
'    If AbrirRecordset(SQL, RS) Then
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    '- si valor del parametro ItemNulo=true hay que a�adir linea en blanco
    If Not RS.EOF And ItemNulo Then
        Cbo.AddItem "  "
        Cbo.ItemData(Cbo.NewIndex) = 0
    End If
    
    If Not RS.EOF Then
        If IsNumeric(RS.Fields(0).Value) Then
            '- si el codigo NomCodigo es numerico en el ItemData se carga el campo clave primaria
            '- y en List la descripcion NomDescrip
            While Not RS.EOF
              Cbo.AddItem RS.Fields(1).Value 'descrip
              Cbo.ItemData(Cbo.NewIndex) = RS.Fields(0).Value 'codigo
              RS.MoveNext
            Wend
        Else
            '- si el codigo NomCodigo en alfanumerico no se puede cargar
            '- el codigo en ItemData y cargamos un indice ficticio
            '- y en el List el campo codigo NomCodigo
            i = 1
            While Not RS.EOF
              Cbo.AddItem RS.Fields(0).Value 'campo del codigo
              Cbo.ItemData(Cbo.NewIndex) = i
              i = i + 1
              RS.MoveNext
            Wend
        End If
    End If
'    End If
    
'    CerrarRecordset RS
    RS.Close
    Set RS = Nothing
    Exit Sub
    
ErrCombo:
    MuestraError Err.Number, "Cargar combo." & NomTabla, Err.Description
End Sub




Public Sub CargarCombo_TipMov(ByRef Cbo As ComboBox, NomTabla As String, NomCodigo As String, nomDescrip As String, Optional strWhere As String, Optional ItemNulo As Boolean)
'Carga un objeto ComboBox con los registros de una Tabla
'(IN) cbo: ComboBox en el q se van a cargar los datos
'(IN) nomTabla: nombre de la tabla de la q leeremos los datos a cargar
'(IN) nomCodigo: nombre del campo codigo de la tabla q queremos cargar
'(IN) nomDescrip: nombre del campo descripcion de la tabla q queremos cargar
'(IN) strWhere: para filtrar los registros de la tabla q queremos cargar
'(IN) ItemNulo: si es true se a�ade el primer item con linea en blanco
Dim SQL As String
Dim RS As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrCombo
    
    Cbo.Clear
    
    SQL = "SELECT " & NomCodigo & "," & nomDescrip & " FROM " & NomTabla
    If strWhere <> "" Then SQL = SQL & " WHERE " & strWhere
    SQL = SQL & " ORDER BY " & NomCodigo
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    '- si valor del parametro ItemNulo=true hay que a�adir linea en blanco
    If Not RS.EOF And ItemNulo Then
        Cbo.AddItem "  "
        Cbo.ItemData(Cbo.NewIndex) = 0
    End If
       
    i = 1
    While Not RS.EOF
        SQL = Replace(RS.Fields(1).Value, "Factura", "Fac.")
        SQL = RS.Fields(0).Value & " - " & SQL
        Cbo.AddItem SQL 'campo del codigo
        Cbo.ItemData(Cbo.NewIndex) = i
        i = i + 1
        RS.MoveNext
    Wend

    RS.Close
    Set RS = Nothing
    Exit Sub
    
ErrCombo:
    MuestraError Err.Number, "Cargar combo." & NomTabla, Err.Description
End Sub



Public Sub CancelaADODC(ByRef vData As Adodc)
On Error Resume Next
    vData.Recordset.Cancel
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function EsVacio(ByRef campo As TextBox) As Boolean
    If (campo.Text = "" Or campo.Text = "0") Then
        EsVacio = True
    Else
        EsVacio = False
    End If
End Function


Public Sub DesplazamientoVisible(ByRef toolb As Toolbar, iniBoton As Byte, bol As Boolean, nreg As Byte)
'Oculta o Muestra las botones de  flechas de desplazamiento de la toolbar
Dim i As Byte

    Select Case nreg
        Case 0, 1 '0 o 1 registro no mostrar los botones despl.
            For i = iniBoton To iniBoton + 3
                toolb.Buttons(i).visible = False
            Next i
        Case Else '>1 reg, mostrar si bol
            For i = iniBoton To iniBoton + 3
                toolb.Buttons(i).visible = bol
            Next i
    End Select
End Sub


Public Sub PonerIndicador(ByRef lblIndicador As Label, Modo As Byte)
'Pone el titulo del label lblIndicador
    lblIndicador.FontBold = True
    Select Case Modo
        Case 0    'Modo Inicial
            lblIndicador.Caption = ""
        Case 1 'Modo Buscar
            lblIndicador.Caption = "BUSQUEDA"
        Case 2    'Preparamos para que pueda Modificar
        
        Case 3 'Modo Insertar
            lblIndicador.Caption = "INSERTAR"
        Case 4 'MODIFICAR
            lblIndicador.Caption = "MODIFICAR"
        Case Else
            lblIndicador.Caption = ""
    End Select
End Sub


Public Sub ActualizarToolbarGnral(ByRef Toolbar1 As Toolbar, Modo As Byte, Kmodo As Byte, posic As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner
Dim b As Boolean
    
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
    b = (Modo = 5 Or Modo = 6 Or Modo = 7 Or Modo = 8)
    
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
    If (b) And (Kmodo <> 5 And Kmodo <> 6 And Kmodo <> 7 And Kmodo <> 8) Then 'Cabecera
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(posic).Image = 3
        Toolbar1.Buttons(posic).ToolTipText = "Nuevo"
        '-- Modificar
        Toolbar1.Buttons(posic + 1).Image = 4
        Toolbar1.Buttons(posic + 1).ToolTipText = "Modificar"
        '-- eliminar
        Toolbar1.Buttons(posic + 2).Image = 5
        Toolbar1.Buttons(posic + 2).ToolTipText = "Eliminar"
    End If
    
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
    If (Kmodo = 5 Or Kmodo = 6 Or Kmodo = 7 Or Kmodo = 8) Then 'Lineas
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(posic).Image = 12
        Toolbar1.Buttons(posic).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(posic + 1).Image = 13
        Toolbar1.Buttons(posic + 1).ToolTipText = "Modificar linea"
        '-- eliminar
        Toolbar1.Buttons(posic + 2).Image = 14
        Toolbar1.Buttons(posic + 2).ToolTipText = "Eliminar linea"
    End If
End Sub


Public Sub KEYpressGnral(KeyAscii As Integer, Modo As Byte, cerrar As Boolean)
'IN: codigo keyascii tecleado, y modo en que esta el formulario
'OUT: si se tiene que cerrar el formulario o no
    cerrar = False
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
'        SendKeys "{tab}"
        CreateObject("WScript.Shell").SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then cerrar = True
    End If
End Sub


Public Sub KEYdown(KeyCode As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
On Error Resume Next
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            SendKeys "+{tab}"
        Case 40 'Desplazamiento Flecha Hacia Abajo
            SendKeys "{tab}"
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub KEYdownLineas(KeyCode As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
On Error Resume Next
    Select Case KeyCode
        Case 37 'Desplazamiento Flecha Izquierda
            SendKeys "+{tab}"
        Case 38 'Desplazamieto Flecha Hacia Arriba
            SendKeys "+{tab}"
        Case 39 'Desplaz. Flecha Derecha
            SendKeys "{tab}"
        Case 40 'Desplazamiento Flecha Hacia Abajo
            SendKeys "{tab}"
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub SituarMultiTextFinal(ByRef txt As TextBox)
    On Error GoTo ErrMulti
    
    'situa el cursor del text multilinea al final para poder empezar a escribir
    If txt.Text <> "" And txt.MultiLine And txt.Enabled And txt.Locked = False Then SendKeys "^{END}"
    Exit Sub
     
ErrMulti:
    MuestraError Err.Number, "", Err.Description
End Sub




Public Sub SituarCombo(ByRef Cbo As ComboBox, Valor As Byte)
Dim i As Byte

    On Error Resume Next

        For i = 0 To Cbo.ListCount - 1
            If Cbo.ItemData(i) = Val(Valor) Then
                Cbo.ListIndex = i
                Exit For
            End If
        Next i
        If i = Cbo.ListCount Then Cbo.ListIndex = -1
    
    If Err.Number <> 0 Then
        Cbo.ListIndex = -1
        Err.Clear
    End If
End Sub


Public Function ObtenerAlto(ByRef vDataGrid As DataGrid, Optional alto As Integer) As Single
Dim anc As Single
    anc = vDataGrid.Top + alto
    If vDataGrid.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + vDataGrid.RowTop(vDataGrid.Row)
    End If
    ObtenerAlto = anc
End Function


'*********** LAURA : 13/09/2005
Public Function EsEnteroNew(Texto As String) As Boolean
Dim i As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEnteroNew = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            i = InStr(L, Texto, ".")
            If i > 0 Then
                L = i + 1
                C = C + 1
            End If
        Loop Until i = 0
        If C > 0 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                i = InStr(L, Texto, ",")
                If i > 0 Then
                    L = i + 1
                    C = C + 1
                End If
            Loop Until i = 0
            If C > 0 Then res = False
        End If
    End If
    EsEnteroNew = res
End Function




'=================================
'******** DAVID (NO LA USO)
Public Function EsEntero(Texto As String) As Boolean
Dim i As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEntero = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            i = InStr(L, Texto, ".")
            If i > 0 Then
                L = i + 1
                C = C + 1
            End If
        Loop Until i = 0
        If C > 1 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                i = InStr(L, Texto, ",")
                If i > 0 Then
                    L = i + 1
                    C = C + 1
                End If
            Loop Until i = 0
            If C > 1 Then res = False
        End If
        
    End If
        EsEntero = res
End Function



Public Function PonerFormatoEntero(ByRef T As TextBox) As Boolean
'Comprueba que el valor del textbox es un entero y le pone el formato
Dim mTag As CTag
Dim cad As String
Dim Formato As String
On Error GoTo EPonerFormato

    If T.Text = "" Then Exit Function
    PonerFormatoEntero = True
    
    Set mTag = New CTag
    mTag.Cargar T
    If mTag.Cargado Then
       cad = mTag.Nombre 'descripcion del campo
       Formato = mTag.Formato
    End If
    Set mTag = Nothing

    If Not EsEnteroNew(T.Text) Then
        PonerFormatoEntero = False
        MsgBox "El campo " & cad & " tiene que ser un n�mero entero.", vbExclamation
        PonerFoco T
    Else
         T.Text = Format(T.Text, Formato)
    End If
    
EPonerFormato:
    If Err.Number <> 0 Then Err.Clear
End Function


'******* IMPORTANTE
' El tipo de datos CURRENCY solo admite 4 decimales
Public Function PonerFormatoDecimal_Single(ByRef T As TextBox, tipoF As Single) As Boolean
Dim Valor2 As Single
Dim PEntera As Currency
Dim NoOK As Boolean
Dim Tg As CTag
Dim FormatoTag As String
'tipoF: tipo de Formato a aplicar
'  1 -> Decimal(12,2)
'  2 -> Decimal(10,4)
'  3 -> Decimal(10,2)    '�FORMATO CANTIDAD
'  4 -> Decimal(4,2)
'  5 -> Decimal(8,4)
'  6 -> Decimal(8,2)
'  7 -> Decimal(5,2)
'
'
'  8 -> Lo que ponga en su TAG
'  9 ->  Formato precio2. Para cuando podamos parametrizarlo


    PonerFormatoDecimal_Single = False
    If T.Text = "" Then Exit Function
    NoOK = False
    With T
        If Not EsNumerico(.Text) Then
'            .Text = ""
            PonerFoco T
        Else
            If InStr(1, .Text, ",") > 0 Then
                Valor = ImporteFormateadoSingle(.Text)
            Else
                Valor = CSng(TransformaPuntosComas(.Text))
            End If

            'Comprobar la longitud de la Parte Entera
            PEntera = Int(Valor)
            Select Case tipoF 'Comprobar longitud
                Case 1 'Decimal(12,2)
                    If Len(PEntera) > 10 Then
                        MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 2, 9 'Decimal(10,4)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,9999", vbExclamation
                        NoOK = True
                    End If
                Case 3 'Decimal(10,2)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 4 'Decimal(4,2)
                    If Len(CStr(PEntera)) > 2 Then
                        MsgBox "El valor no puede ser mayor de 99,99", vbExclamation
                        NoOK = True
                    End If
                Case 5 'Decimal(8,4)
                    If Len(CStr(PEntera)) > 4 Then
                        MsgBox "El valor no puede ser mayor de 9999,9999", vbExclamation
                        NoOK = True
                    End If
                Case 6 'Decimal(8,2)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 7 'Decimal(5,2)
                    '---- Laura: 05/10/2006
                    '# ANTES:   If Len(CStr(PEntera)) > 3 Then
                    If Len(CStr(Abs(PEntera))) > 3 Then
                    '----
                        MsgBox "El valor no puede ser mayor de 100,00", vbExclamation
                        NoOK = True
                    End If
                    
                Case 8
                    'David 12 Feb 07
                    'Lo que ponga en su tag
                    Set Tg = New CTag
                    If Not Tg.Cargar(T) Then NoOK = True
                    FormatoTag = Tg.Formato
                    Set Tg = Nothing
            End Select
            
            If NoOK Then
                .Text = ""
                T.SetFocus
                PonerFormatoDecimal_Single = False
                Exit Function
            Else
                PonerFormatoDecimal_Single = True
            End If

            'Poner el Formato
            Select Case tipoF
                Case 1 'Formato Decimal(12,2)
                    .Text = Format(Valor, FormatoImporte)
                Case 2 'Formato Decimal(10,4)
                    .Text = Format(Valor, FormatoPrecio)
                Case 3 'Formato Decimal(10,2)
                    .Text = Format(Valor, FormatoCantidad)
                Case 4 'Formato Decimal(4,2)
                    .Text = Format(Valor, FormatoDescuento)
                Case 5 'Formato Decimal(8,4)
                    .Text = Format(Valor, FormatoKms)
                Case 6 'Formato Decimal(8,2)
                    .Text = Format(Valor, FormatoCantidad2)
                Case 7 'Formato Decimal(5,2)
                    .Text = Format(Valor, "##0.00")
                Case 8
                    .Text = Format(Valor, FormatoTag)
                Case 9 'Formato Decimal(10,5). Intentaremos parametrizarlo
                    .Text = Format(Valor, FormatoPrecio2)
                    
            End Select
        End If
    End With

End Function




'******* IMPORTANTE
'Leer el procedimiento de arriba.   IMPORTANTE:   PonerFormatoDecimal_Single
'---------------------------------------------------------------------------------
Public Function PonerFormatoDecimal(ByRef T As TextBox, tipoF As Single) As Boolean
'tipoF: tipo de Formato a aplicar
'  1 -> Decimal(12,2)
'  2 -> Decimal(10,4)
'  3 -> Decimal(10,2)    '�FORMATO CANTIDAD
'  4 -> Decimal(4,2)
'  5 -> Decimal(8,4)
'  6 -> Decimal(8,2)
'  7 -> Decimal(5,2)
'
'
'  8 -> Lo que ponga en su TAG
Dim Valor As Currency
Dim PEntera As Currency
Dim NoOK As Boolean
Dim Tg As CTag
Dim FormatoTag As String

    PonerFormatoDecimal = False
    If T.Text = "" Then Exit Function
    NoOK = False
    With T
        If Not EsNumerico(.Text) Then
'            .Text = ""
            PonerFoco T
        Else
            If InStr(1, .Text, ",") > 0 Then
                Valor = ImporteFormateado(.Text)
            Else
                Valor = CCur(TransformaPuntosComas(.Text))
            End If

            'Comprobar la longitud de la Parte Entera
            PEntera = Int(Valor)
            Select Case tipoF 'Comprobar longitud
                Case 1 'Decimal(12,2)
                    If Len(PEntera) > 10 Then
                        MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 2 'Decimal(10,4)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,9999", vbExclamation
                        NoOK = True
                    End If
                Case 3 'Decimal(10,2)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 4 'Decimal(4,2)
                    If Len(CStr(PEntera)) > 2 Then
                        MsgBox "El valor no puede ser mayor de 99,99", vbExclamation
                        NoOK = True
                    End If
                Case 5 'Decimal(8,4)
                    If Len(CStr(PEntera)) > 4 Then
                        MsgBox "El valor no puede ser mayor de 9999,9999", vbExclamation
                        NoOK = True
                    End If
                Case 6 'Decimal(8,2)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 7 'Decimal(5,2)
                    '---- Laura: 05/10/2006
                    '# ANTES:   If Len(CStr(PEntera)) > 3 Then
                    If Len(CStr(Abs(PEntera))) > 3 Then
                    '----
                        MsgBox "El valor no puede ser mayor de 100,00", vbExclamation
                        NoOK = True
                    End If
                    
                Case 8
                    'David 12 Feb 07
                    'Lo que ponga en su tag
                    Set Tg = New CTag
                    If Not Tg.Cargar(T) Then NoOK = True
                    FormatoTag = Tg.Formato
                    Set Tg = Nothing
            End Select
            
            If NoOK Then
                .Text = ""
                T.SetFocus
                PonerFormatoDecimal = False
                Exit Function
            Else
                PonerFormatoDecimal = True
            End If

            'Poner el Formato
            Select Case tipoF
                Case 1 'Formato Decimal(12,2)
                    .Text = Format(Valor, FormatoImporte)
                Case 2 'Formato Decimal(10,4)
                    .Text = Format(Valor, FormatoPrecio)
                Case 3 'Formato Decimal(10,2)
                    .Text = Format(Valor, FormatoCantidad)
                Case 4 'Formato Decimal(4,2)
                    .Text = Format(Valor, FormatoDescuento)
                Case 5 'Formato Decimal(8,4)
                    .Text = Format(Valor, FormatoKms)
                Case 6 'Formato Decimal(8,2)
                    .Text = Format(Valor, FormatoCantidad2)
                Case 7 'Formato Decimal(5,2)
                    .Text = Format(Valor, "##0.00")
                Case 8
                    .Text = Format(Valor, FormatoTag)
            End Select
        End If
    End With
End Function


Public Function PonerNombreDeCod(ByRef txt As TextBox, bd As Byte, Tabla As String, campo As String, Optional Codigo As String, Optional Texto As String, Optional Tipo As String) As String
'Devuelve el nombre/Descripci�n asociado al C�digo correspondiente
'Adem�s pone formato al campo txt del c�digo a partir del Tag
Dim SQL As String
Dim devuelve As String
Dim vtag As CTag
Dim ValorCodigo As String
On Error GoTo EPonerNombresDeCod

    ValorCodigo = txt.Text
    If ValorCodigo <> "" Then
        Set vtag = New CTag
        If vtag.Cargar(txt) Then
            If Codigo = "" Then Codigo = vtag.columna
            If Tipo = "" Then Tipo = vtag.TipoDato
            SQL = DevuelveDesdeBD(bd, campo, Tabla, Codigo, ValorCodigo, Tipo)
            If vtag.TipoDato = "N" Then ValorCodigo = Format(ValorCodigo, vtag.Formato)
            If SQL = "" Then
                If Texto = "" Then
                    devuelve = "No existe " & vtag.Nombre & ": " & ValorCodigo
                Else
                    devuelve = "No existe " & Texto & ": " & ValorCodigo
                End If
                MsgBox devuelve, vbExclamation
'                Txt.Text = ""
                'si ponemos foco bucle
'                PonerFoco Txt
'                Txt.SetFocus
            Else
                PonerNombreDeCod = SQL 'Descripcion del codigo
                'Poner valor codigo formateado
                txt.Text = ValorCodigo 'Valor codigo formateado
            End If
        End If
        Set vtag = Nothing
    Else
        PonerNombreDeCod = ""
    End If
    Exit Function
EPonerNombresDeCod:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Nombre asociado a c�digo: " & Codigo, Err.Description
End Function


Public Function ExisteCP(T As TextBox) As Boolean
'comprueba para un campo de texto que sea clave primaria, si ya existe un
'registro con ese valor
Dim vtag As CTag
Dim devuelve As String
On Error GoTo EExiste

    ExisteCP = False
    If T.Text <> "" Then
        If T.Tag <> "" Then
            Set vtag = New CTag
            If vtag.Cargar(T) Then
                devuelve = DevuelveDesdeBDNew(conAri, vtag.Tabla, vtag.columna, vtag.columna, T.Text, vtag.TipoDato)
                If devuelve <> "" Then
                    MsgBox "Ya existe un registro para " & vtag.Nombre & ": " & T.Text, vbExclamation
                    ExisteCP = True
                End If
            End If
            Set vtag = Nothing
        End If
    End If
EExiste:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar c�digo", Err.Description
End Function



Public Sub SubirItemList(ByRef LView As ListView)
'Subir el item seleccionado del listview una posicion
Dim i As Byte, item As Byte
Dim Aux As String
On Error Resume Next
   
    For i = 2 To LView.ListItems.Count
        If LView.ListItems(i).Selected Then
            item = i
            Aux = LView.ListItems(i).Text
            LView.ListItems(i).Text = LView.ListItems(i - 1).Text
            LView.ListItems(i - 1).Text = Aux
        End If
    Next i
    If item <> 0 Then
        LView.ListItems(item).Selected = False
        LView.ListItems(item - 1).Selected = True
    End If
    LView.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BajarItemList(ByRef LView As ListView)
'Bajar el item seleccionado del listview una posicion
Dim i As Byte, item As Byte
Dim Aux As String
On Error Resume Next

    For i = 1 To LView.ListItems.Count - 1
        If LView.ListItems(i).Selected Then
            item = i
            Aux = LView.ListItems(i).Text
            LView.ListItems(i).Text = LView.ListItems(i + 1).Text
            LView.ListItems(i + 1).Text = Aux
        End If
    Next i
    If item <> 0 Then
        LView.ListItems(item).Selected = False
        LView.ListItems(item + 1).Selected = True
    End If
    LView.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub CargarProgres(ByRef PBar As ProgressBar, Valor As Integer)
On Error Resume Next
    PBar.Max = 100
    PBar.Value = 0
    PBar.Tag = Valor
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub IncrementarProgres(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub CargarProgresNew(ByRef PBar As ProgressBar, Valor As Integer)
On Error Resume Next
    PBar.Max = Valor
    PBar.Value = 0
'    PBar.Tag = Valor
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub IncrementarProgresNew(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
'    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    PBar.Value = PBar.Value + Veces
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub PosicionarComboDes(ByRef Cbo As ComboBox, Valor As String)
Dim i As Byte

    For i = 0 To Cbo.ListCount - 1
        If Trim(Cbo.List(i)) = Trim(Valor) Then
            Cbo.ListIndex = i
            Exit For
        End If
    Next i
    If i = Cbo.ListCount Then Cbo.ListIndex = -1
    
End Sub



Public Sub PosicionarCombo(ByRef Combo1 As ComboBox, Valor As Integer)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(J) = Valor Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub





'============================================================
'====== FUNCIONES PARA ARITAXI ==============================
'============================================================

Public Function PonerNombreCuenta(ByRef txt As TextBox, Modo As Byte, Optional clien As String) As String
Dim DevfrmCCtas As String
Dim SQL As String

     If txt.Text = "" Then
         PonerNombreCuenta = ""
         Exit Function
    End If
    DevfrmCCtas = txt.Text
    If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
        If InStr(SQL, "No existe la cuenta") > 0 Then
            txt.Text = DevfrmCCtas
            
            If (Modo = 3 Or Modo = 4) Then  'si insertar o modificar
                SQL = SQL & "  �Desea crearla?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                    'SI MODO es insetar NO me sirve el metodo anterior. Pq? Pq aun no he creado el cli/prov
                    'De momento pondre una marca en el texto de descripcion para que la cree
                    If Modo = 3 Then
                        PonerNombreCuenta = vbCrearNuevaCta
                                                
                    Else
                        If InStr(1, txt.Tag, "scliente") > 0 Then
                            InsertarCuentaCble DevfrmCCtas, clien
                        ElseIf InStr(1, txt.Tag, "sclien") Then
                            ModificarCuentaCble DevfrmCCtas, clien, "sclien"
                        ElseIf InStr(1, txt.Tag, "sprove") > 0 Then
                            InsertarCuentaCble DevfrmCCtas, "", clien
                        ' ---- [02/10/2009] (LAURA): crear cuenta en familias articulos
                        ElseIf InStr(1, txt.Tag, "sfamia") > 0 Then
                            InsertarCuentaCble DevfrmCCtas, "", , clien
                        ' ----
                        End If
                        PonerNombreCuenta = clien
                    End If
                Else
                    'DAVID
                    'Si me dice que no quiere crearla, pongo el txt a blanco
                    txt.Text = ""
                End If
            Else
                MsgBox SQL, vbExclamation
            End If
        Else
            txt.Text = DevfrmCCtas
            
            If InStr(1, txt.Tag, "scliente") > 0 Then
                ModificarCuentaCble DevfrmCCtas, clien, "scliente"
            ElseIf InStr(1, txt.Tag, "sclien") Then
                ModificarCuentaCble DevfrmCCtas, clien, "sclien"
            ElseIf InStr(1, txt.Tag, "sprove") > 0 Then
                ModificarCuentaCble DevfrmCCtas, clien, "sprove"
            End If
            
            PonerNombreCuenta = clien
        End If
    Else
        If Modo = 3 Or Modo = 4 Or Modo = 1 Then 'si insertar o modificar
            MsgBox SQL, vbExclamation
'            PonerNombreCuenta = ""
        End If
'        Txt.Text = ""
        PonerNombreCuenta = ""
'        ConseguirFoco Txt, Modo
        PonerFoco txt
    End If
    DevfrmCCtas = ""
End Function

'He cambiado el metodo a public
Public Function InsertarCuentaCble(cuenta As String, Optional cadClien As String, Optional cadProve As String, Optional cadFamia As String, Optional cadSocio As String, Optional Desmarcar347 As Byte) As Boolean
Dim SQL As String
Dim vClien As CCliente
Dim vProve As CProveedor
Dim vSocio As CSocio
Dim b As Boolean

    On Error GoTo EInsCta
    
    SQL = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,maidatos,webdatos,obsdatos,pais,forpa,ctabanco,entidad,oficina,CC,cuentaba" ') "
    '[Monica]22/11/2013: iban
    If vEmpresa.HayNorma19_34Nueva = 1 Then
        SQL = SQL & ", iban)"
    Else
        SQL = SQL & ")"
    End If
    SQL = SQL & " VALUES (" & DBSet(cuenta, "T") & ","
    
    If cadClien <> "" Then
        Set vClien = New CCliente
        If vClien.LeerDatos(cadClien) Then
            If Desmarcar347 Then
                SQL = SQL & DBSet(vClien.Nombre, "T") & ",'S',0," & DBSet(vClien.Nombre, "T") & "," & DBSet(vClien.Domicilio, "T") & ","
            Else
                SQL = SQL & DBSet(vClien.Nombre, "T") & ",'S',1," & DBSet(vClien.Nombre, "T") & "," & DBSet(vClien.Domicilio, "T") & ","
            End If
            
            SQL = SQL & DBSet(vClien.CPostal, "T") & "," & DBSet(vClien.Poblacion, "T") & "," & DBSet(vClien.Provincia, "T") & "," & DBSet(vClien.NIF, "T") & "," & DBSet(vClien.EMailAdm, "T") & "," & DBSet(vClien.WebClien, "T") & "," & ValorNulo & "," & ValorNulo
            'Forma pago y cuenta banco por defecto
            SQL = SQL & "," & DBSet(vClien.ForPago, "N", "S") & "," & ValorNulo & "," & DBSet(Format(vClien.Banco, "0000"), "T") & "," & DBSet(Format(vClien.Sucursal, "0000"), "T") & "," & DBSet(vClien.DigControl, "T") & "," & DBSet(vClien.CuentaBan, "T") '& ")"
            
            '[Monica]22/11/2013: iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                SQL = SQL & "," & DBSet(vClien.Iban, "T") & ")"
            Else
                SQL = SQL & ")"
            End If
            
            ConnConta.Execute SQL
            cadClien = vClien.Nombre
            b = True
        Else
            b = False
        End If
        Set vClien = Nothing
    End If
    
    If cadSocio <> "" Then
        Set vSocio = New CSocio
        If vSocio.LeerDatos(cadSocio) Then
            If Desmarcar347 Then
                SQL = SQL & DBSet(vSocio.Nombre, "T") & ",'S',0," & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Domicilio, "T") & ","
            Else
                SQL = SQL & DBSet(vSocio.Nombre, "T") & ",'S',1," & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Domicilio, "T") & ","
            End If
            SQL = SQL & DBSet(vSocio.CPostal, "T") & "," & DBSet(vSocio.Poblacion, "T") & "," & DBSet(vSocio.Provincia, "T") & "," & DBSet(vSocio.NIF, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
            'Forma pago y cuenta banco por defecto
            SQL = SQL & "," & DBSet(vSocio.ForPago, "N", "S") & "," & ValorNulo & "," & DBSet(Format(vSocio.Banco, "0000"), "T") & "," & DBSet(Format(vSocio.Sucursal, "0000"), "T") & "," & DBSet(vSocio.DigControl, "T") & "," & DBSet(vSocio.CuentaBan, "T") '& ")"
            
            '[Monica]22/11/2013: iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                SQL = SQL & "," & DBSet(vSocio.Iban, "T") & ")"
            Else
                SQL = SQL & ")"
            End If
            
            ConnConta.Execute SQL
            cadClien = vSocio.Nombre
            b = True
        Else
            b = False
        End If
        Set vSocio = Nothing
    End If
    
    If cadProve <> "" Then
        Set vProve = New CProveedor
        If vProve.LeerDatos(cadProve) Then
            If Desmarcar347 Then
                SQL = SQL & DBSet(vProve.Nombre, "T") & ",'S',0," & DBSet(vProve.Nombre, "T") & "," & DBSet(vProve.Domicilio, "T") & ","
            Else
                SQL = SQL & DBSet(vProve.Nombre, "T") & ",'S',1," & DBSet(vProve.Nombre, "T") & "," & DBSet(vProve.Domicilio, "T") & ","
            End If
            SQL = SQL & DBSet(vProve.CPostal, "T") & "," & DBSet(vProve.Poblacion, "T") & "," & DBSet(vProve.Provincia, "T") & "," & DBSet(vProve.NIF, "T") & ","
            SQL = SQL & DBSet(vProve.EMailAdmon, "T") & "," & DBSet(vProve.WebProve, "T") & "," & ValorNulo & "," & ValorNulo
            'Forma pago y cuenta banco por defecto
            cadProve = DevuelveDesdeBD(conAri, "codmacta", "sbanpr", "codbanpr", vProve.BancoPropio)
            SQL = SQL & "," & DBSet(vProve.ForPago, "N", "S") & "," & DBSet(cadProve, "N", "S") '& ")"
            
            '[Monica]22/11/2013: iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                SQL = SQL & "," & DBSet(vProve.Iban, "T") & ")"
            Else
                SQL = SQL & ")"
            End If
            
            cadProve = ""
            ConnConta.Execute SQL
            cadProve = vProve.Nombre
            b = True
        Else
            b = False
        End If
        Set vProve = Nothing
    
    ' ---- [02/10/2009] (LAURA): crear cuenta en familias articulos
    ElseIf cadFamia <> "" Then 'cuentas familias articulos
        SQL = SQL & DBSet(cadFamia, "T") & ",'S',0," & DBSet(cadFamia, "T") & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
        'Forma pago y cuenta banco por defecto
        SQL = SQL & "," & ValorNulo & "," & ValorNulo & ")"
        
        ConnConta.Execute SQL
        b = True
    ' ----
    End If
    
EInsCta:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Description, "Insertando cuenta contable", Err.Description
    End If
    InsertarCuentaCble = b
End Function


Public Function ModificarCuentaCble(cuenta As String, cadClien As String, Tabla As String) As Boolean
Dim SQL As String
Dim vClien As CCliente
Dim vProve As CProveedor
Dim vSocio As CSocio
Dim b As Boolean

    On Error GoTo EInsCta
    
    
    If cadClien <> "" And Tabla = "scliente" Then
        Set vClien = New CCliente
        If vClien.LeerDatos(cadClien) Then
            
            SQL = "UPDATE cuentas SET nommacta = " & DBSet(vClien.Nombre, "T")
            SQL = SQL & ", razosoci = " & DBSet(vClien.Nombre, "T")
            SQL = SQL & ", dirdatos = " & DBSet(vClien.Domicilio, "T")
            SQL = SQL & ", codposta = " & DBSet(vClien.CPostal, "T")
            SQL = SQL & ", despobla = " & DBSet(vClien.Poblacion, "T")
            SQL = SQL & ", desprovi = " & DBSet(vClien.Provincia, "T")
            SQL = SQL & ", nifdatos = " & DBSet(vClien.NIF, "T")
            SQL = SQL & ", maidatos = " & DBSet(vClien.EMailAdm, "T")
            SQL = SQL & ", forpa = " & DBSet(vClien.ForPago, "N", "S")
            SQL = SQL & ", entidad = " & DBSet(Format(vClien.Banco, "0000"), "T")
            SQL = SQL & ", oficina = " & DBSet(Format(vClien.Sucursal, "0000"), "T")
            SQL = SQL & ", CC = " & DBSet(vClien.DigControl, "T")
            SQL = SQL & ", cuentaba = " & DBSet(vClien.CuentaBan, "T")
            '[Monica]22/11/2013: iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                SQL = SQL & ", iban = " & DBSet(vClien.Iban, "T")
            End If
            
            SQL = SQL & " where codmacta = " & DBSet(cuenta, "T")
            
            ConnConta.Execute SQL
            cadClien = vClien.Nombre
            b = True
        Else
            b = False
        End If
        Set vClien = Nothing
    End If
    
    If cadClien <> "" And Tabla = "sclien" Then
        Set vSocio = New CSocio
        If vSocio.LeerDatos(cadClien) Then
            
            SQL = "UPDATE cuentas SET nommacta = " & DBSet(vSocio.Nombre, "T")
            SQL = SQL & ", razosoci = " & DBSet(vSocio.Nombre, "T")
            SQL = SQL & ", dirdatos = " & DBSet(vSocio.Domicilio, "T")
            SQL = SQL & ", codposta = " & DBSet(vSocio.CPostal, "T")
            SQL = SQL & ", despobla = " & DBSet(vSocio.Poblacion, "T")
            SQL = SQL & ", desprovi = " & DBSet(vSocio.Provincia, "T")
            SQL = SQL & ", nifdatos = " & DBSet(vSocio.NIF, "T")
            SQL = SQL & ", forpa = " & DBSet(vSocio.ForPago, "N", "S")
            SQL = SQL & ", maidatos = " & DBSet(vSocio.EMailSocio, "T")
            SQL = SQL & ", entidad = " & DBSet(Format(vSocio.Banco, "0000"), "T")
            SQL = SQL & ", oficina = " & DBSet(Format(vSocio.Sucursal, "0000"), "T")
            SQL = SQL & ", CC = " & DBSet(vSocio.DigControl, "T")
            SQL = SQL & ", cuentaba = " & DBSet(vSocio.CuentaBan, "T")
            '[Monica]22/11/2013: iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                SQL = SQL & ", iban = " & DBSet(vSocio.Iban, "T")
            End If
            SQL = SQL & " where codmacta = " & DBSet(cuenta, "T")
            
            ConnConta.Execute SQL
            cadClien = vSocio.Nombre
            b = True
        Else
            b = False
        End If
        Set vSocio = Nothing
    End If
    
    
    If cadClien <> "" And Tabla = "sprove" Then
        Set vProve = New CProveedor
        If vProve.LeerDatos(cadClien) Then
            SQL = "UPDATE cuentas SET nommacta = " & DBSet(vProve.Nombre, "T")
            SQL = SQL & ", razosoci = " & DBSet(vProve.Nombre, "T")
            SQL = SQL & ", dirdatos = " & DBSet(vProve.Domicilio, "T")
            SQL = SQL & ", codposta = " & DBSet(vProve.CPostal, "T")
            SQL = SQL & ", despobla = " & DBSet(vProve.Poblacion, "T")
            SQL = SQL & ", desprovi = " & DBSet(vProve.Provincia, "T")
            SQL = SQL & ", nifdatos = " & DBSet(vProve.NIF, "T")
            SQL = SQL & ", forpa = " & DBSet(vProve.ForPago, "N", "S")
            SQL = SQL & ", entidad = " & DBSet(Format(vProve.Banco, "0000"), "T")
            SQL = SQL & ", oficina = " & DBSet(Format(vProve.Sucursal, "0000"), "T")
            SQL = SQL & ", CC = " & DBSet(vProve.DigControl, "T")
            SQL = SQL & ", cuentaba = " & DBSet(vProve.CuentaBan, "T")
            '[Monica]22/11/2013: iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                SQL = SQL & ", iban = " & DBSet(vProve.Iban, "T")
            End If
            SQL = SQL & " where codmacta = " & DBSet(cuenta, "T")
            
            ConnConta.Execute SQL
            cadClien = vProve.Nombre
            b = True
            
        Else
            b = False
        End If
        Set vProve = Nothing
    End If
    
EInsCta:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Description, "Insertando cuenta contable", Err.Description
    End If
    ModificarCuentaCble = b
End Function








'He cambiado el metodo a public
Public Function InsertarCuentaCbleDescripcion(cuenta As String, Descripcion As String) As Boolean
Dim SQL As String


   
    
    SQL = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci) "
    SQL = SQL & " VALUES ('" & cuenta & "','" & DevNombreSQL(Descripcion) & "','S',0,'" & DevNombreSQL(Descripcion) & "')"
    ConnConta.Execute SQL

End Function






Public Function ComprobarHayStock(stockOrig As Single, stockTras As Single, codArtic As String, NomArtic As String, tipoMov As String)
'IN: stockOrig: stock existente en almacen Origen
'    stockTras: stock a traspasar del origen a otro almacen
Dim b As Boolean
Dim devuelve As String

    ComprobarHayStock = False
    If stockOrig >= CSng(stockTras) Then
    'Si cantidad en stock > cantidad a traspasar entonces
        b = True
    Else    'No hay suficiente stock en almacen origen
        devuelve = "Control de Stock : " & vbCrLf
        devuelve = devuelve & "---------------------- " & vbCrLf & vbCrLf
        devuelve = devuelve & " No hay suficiente Stock en el Almacen del Art�culo:"
        devuelve = devuelve & vbCrLf & " C�digo:   " & codArtic & vbCrLf
        devuelve = devuelve & " Desc.: " & NomArtic & vbCrLf & vbCrLf
        devuelve = devuelve & "(Stock=" & stockOrig & ")"

        If tipoMov = "OFE" Then
            MsgBox devuelve, vbInformation
        Else
            If vParamAplic.ControlStock Then
            'Si hay control Stock no permitir traspaso
                b = False
                Select Case tipoMov
                    Case "REG"
                        devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Movimiento de Almacen. "
                    Case "TRA"
                        devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Traspaso de Almacen. "
                End Select
                MsgBox devuelve, vbExclamation
            Else
                Select Case tipoMov
                Case "REG"
                    devuelve = devuelve & vbCrLf & vbCrLf & " �Desea realizar el Movimiento de Almacen? "
                Case "TRA"
                    devuelve = devuelve & vbCrLf & vbCrLf & " �Desea realizar el Traspaso de Almacen? "
                End Select
                If MsgBox(devuelve, vbQuestion + vbYesNo) = vbYes Then
                    b = True
                Else
                    b = False
                End If
            End If
        End If
    End If
    ComprobarHayStock = b
End Function


Public Function LanzaHomeGnral(nomWeb As String) As Boolean
On Error GoTo ELanzaHome

    LanzaHomeGnral = False
    'Obtenemos la pagina web de los parametros
'    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "spara1", Opcion, "codigo", "1", "N")
'    If CadenaDesdeOtroForm = "" Then
'        MsgBox "Falta configurar los datos en Par�metros de la Aplicaci�n.", vbExclamation
'        Exit Function
'    End If
    If nomWeb = "" Then
        MsgBox "No hay una direcci�n Web para mostrar.", vbInformation
        Exit Function
    End If

    'Lanzamos
'    If cad <> "" Then Shell cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
    If vConfig.Explorador <> "" Then
        Shell vConfig.Explorador & " " & nomWeb, vbMaximizedFocus
        LanzaHomeGnral = True
    End If
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, nomWeb & vbCrLf & Err.Description
End Function


Public Function LanzaMailGnral(dirMail As String) As Boolean
'LLama al Programa de Correo (Outlook,...)
On Error GoTo ELanzaHome

    LanzaMailGnral = False
    If dirMail = "" Then
        MsgBox "No hay direcci�n e-mail a la que enviar.", vbExclamation
        Exit Function
    End If

    Call ShellExecute(hwnd, "Open", "mailto: " & dirMail, "", "", vbNormalFocus)
    LanzaMailGnral = True
    
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, vbCrLf & Err.Description
End Function


Public Function PonerArticuloEan(ByRef txtCod As TextBox, ByRef txtNom As TextBox, codAlm As String, tipoMov As String, Optional Modo As Byte, Optional AntCodArtic As String, Optional sConLotes As Boolean, Optional ByRef TxTProv As String) As Boolean
Dim C As String
    PonerArticuloEan = False
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN y quitar de la cabecera
'    C = DevuelveDesdeBD(conAri, "codartic", "sartic", "codigoea", txtCod.Text, "T")
    C = DevuelveDesdeBD(conAri, "codartic", "sarti3", "codigoea", txtCod.Text, "T")
    '----
    
    If C = "" Then
        MsgBox "El codigo EAN no corresponde con ningun articulo", vbExclamation
    Else
        txtCod.Text = C
        PonerArticuloEan = PonerArticulo(txtCod, txtNom, codAlm, tipoMov, Modo, AntCodArtic, sConLotes, TxTProv)
    End If
End Function


Public Function PonerArticulo(ByRef txtCod As TextBox, ByRef txtNom As TextBox, codAlm As String, tipoMov As String, Optional Modo As Byte, Optional AntCodArtic As String, Optional sConLotes As Boolean, Optional ByRef TxTProv As String) As Boolean
'Poner el codigo y nombre correcto de un Articulo
'IN: txtCod: codigo del articulo
'    txtNom: nombre del articulo
'    codAlm: codigo del almacen en el que comprobamos si se esta inventariando (almacen en el que se va a realizar el movimiento)
Dim vArtic As CArticulo
Dim Bloquea As Boolean

    PonerArticulo = False
    sConLotes = False
    
    Set vArtic = New CArticulo
        
    If vArtic.Existe(txtCod.Text) Then
        If vArtic.LeerDatos(txtCod.Text) Then
            'comprobar que existe el articulo en el almacen del movimiento
            If vArtic.ExisteEnAlmacen(codAlm) Then
            
                'comprobar si el articulo esta inventariandose
                If vArtic.EnInventario(codAlm) Then
                    If Modo = 1 Then 'Insertar lineas
                        txtCod.Text = ""
                        txtNom.Text = ""
                    End If
                    PonerFoco txtCod
                Else
                    'comprobar si el articulo esta bloqueado
                    vArtic.MostrarStatusArtic Bloquea
                    
                    If Bloquea Then 'El articulo esta bloqueado
                        If Modo = 1 Then
                            txtCod.Text = ""
                            txtNom.Text = ""
                        End If
                        PonerFoco txtCod
                    Else 'Articulo OK
                        PonerArticulo = True
                        
                        'Si es articulo DE VARIOS podemos modificar la descripci�n del articulo, sino bloqueamos.
                        If Not EsArticuloVarios(txtCod.Text) Then
                            BloquearTxt txtNom, True
                            'si insertando lineas
                            'If Modo = 1 Then txtNom.Text = vArtic.Nombre
                            txtNom.Text = vArtic.Nombre
                        Else
                            'si insertando lineas
                            If Modo = 1 Then
                                txtNom.Text = vArtic.Nombre
                            ElseIf Modo = 2 And AntCodArtic <> "" Then
                                If txtCod.Text <> AntCodArtic Then txtNom.Text = vArtic.Nombre
                            End If
                            BloquearTxt txtNom, False
'                            PonerFoco txtNom
                        End If

                        Select Case tipoMov
                            Case "OFE", "PEV", "ALV", "ALR", "FAV", "FTI": If vArtic.TextoVentas <> "" Then vArtic.MostrarTextoVen
                            Case "PEC", "ALC", "FAC": If vArtic.TextoCompras <> "" Then vArtic.MostrarTextoCom
                        End Select
                        txtCod.Text = UCase(txtCod.Text)
                        
                       
                        'Si me ha indicado el text donde va el codprove, entonces le pongo
                        If vEmpresa.TieneAnalitica Then
                            If vParamAplic.ModoAnalitica = 0 Then 'ccoste trabajador
                            
                            ElseIf vParamAplic.ModoAnalitica = 1 Then 'ccoste familia
                                'centro de coste
                                TxTProv = DevuelveDesdeBDNew(conAri, "sfamia", "codccost", "codfamia", vArtic.Familia, "N")
                                
                            Else
                                TxTProv = ""
                            End If
                           
                        Else
                            TxTProv = vArtic.codProve
                        End If
                    End If
                End If
            Else
                txtNom.Text = vArtic.Nombre
            End If
        End If
    End If
    
    Set vArtic = Nothing
End Function


'Lineas de ofertas, pedido y albaranes
'Para decir que hace el F2
Public Sub LabelAyudatxtAux(indice As Integer, ByRef Lbl As Label)
    Select Case indice
    Case 3
        'Ver referencia
        Lbl.Caption = "F2  Ver articulo"
    Case 4
        'Consultar precios del articulo
        Lbl.Caption = "F2  Ver precios"
    Case Else
        Lbl.Caption = ""
    End Select
End Sub


Public Sub AbrirConsultaPrecio(Cliente As String, Articulo As String)
    frmFacConsultaPrecios.ConsultaDesdeFrm = Cliente & "|" & Articulo & "|"
    frmFacConsultaPrecios.Show vbModal
End Sub

' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
'Puesto en el modulo en Abril 2010
Public Sub AbrirForm_Articulos(Articulo As String)
Dim frmArt As frmAlmArticulos

    If Trim(Articulo) = "" Then Exit Sub
    
    Set frmArt = New frmAlmArticulos
    frmArt.DeConsulta = True
    frmArt.DatosADevolverBusqueda2 = "::" & Trim(Articulo)
    frmArt.parNumTAb = 6
    frmArt.Show vbModal
    Set frmArt = Nothing
End Sub
' -----


'********************************************************************
'********************************************************************
'********************************************************************
'
' COMBOS para el CRM. Asi no lo coje de una tabla
'
'********************************************************************
Public Sub CargaComboMediosCRM(ByRef Co As ComboBox)
    Co.Clear

    Co.AddItem "Tel�fono"
    Co.ItemData(Co.NewIndex) = 0
    Co.AddItem "eMail"
    Co.ItemData(Co.NewIndex) = 1
    Co.AddItem "Fax"
    Co.ItemData(Co.NewIndex) = 2
    Co.AddItem "Carta"
    Co.ItemData(Co.NewIndex) = 3
    Co.AddItem "Otros"
    Co.ItemData(Co.NewIndex) = 4
End Sub

Public Sub CargaComboEstadoCRM(ByRef Co As ComboBox)
    Co.Clear
    
    Co.AddItem "Pendiente"
    Co.ItemData(Co.NewIndex) = 0
    Co.AddItem "En curso"
    Co.ItemData(Co.NewIndex) = 1
    Co.AddItem "Finalizada"
    Co.ItemData(Co.NewIndex) = 2
    
End Sub

Public Function PonerContRegistros(ByRef vData As Adodc) As String
'indicador del registro donde nos encontramos: "1 de 20"
    On Error GoTo EPonerReg
    
    If Not vData.Recordset.EOF Then
        PonerContRegistros = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
    Else
        PonerContRegistros = ""
    End If
    
EPonerReg:
    If Err.Number <> 0 Then
        Err.Clear
        PonerContRegistros = ""
    End If
End Function

