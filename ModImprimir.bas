Attribute VB_Name = "ModImprimir"
'convertix una posici� d'un Adodc en una Selection Formula SF
Public Function POS2SF(ByRef ado As Adodc, ByRef formu As Form, Optional opcio As Integer, Optional nom_frame As String) As String
'si opcio = 1 OR opcio = 1 => funcionament normal
'si opcio = 2 => funcionament per a ll�nies (NOTA: el manteniment de llinies ha d'estar dins d'un frameAux)
    Dim cadSQL2 As String
    Dim nom_camp As String
    Dim Control As Object
    Dim mTag As CTag
    Dim i As Integer
    
    Set mTag = New CTag
    cadSQL2 = ""

    For Each Control In formu.Controls
        If Control.Tag <> "" Then
            mTag.Cargar Control
            
            'If (mTag.Cargado) And (mTag.EsClave) And (InStr(1, Control.Container.Name, "FrameAux")) = 0 Then 'el control es clau primaria i no forma part de les ll�nies
            If (mTag.Cargado) And (mTag.EsClave) Then
                If (((opcio = 0) Or (opcio = 1)) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    For i = 0 To ado.Recordset.Fields.Count - 1
                        If mTag.columna = ado.Recordset.Fields(i).Name Then
                        
                            If cadSQL2 = "" Then
                                cadSQL2 = "{" & mTag.tabla & "." & mTag.columna & "} = "
                            Else
                                cadSQL2 = cadSQL2 & " AND {" & mTag.tabla & "." & mTag.columna & "} = "
                            End If
                            
                            If mTag.TipoDato = "T" Then 'text
                                cadSQL2 = cadSQL2 & "'" & ado.Recordset.Fields(i).Value & "'"
                            ElseIf mTag.TipoDato = "N" Then 'integer i decimal
                                cadSQL2 = cadSQL2 & ado.Recordset.Fields(i).Value
                            ElseIf mTag.TipoDato = "F" Then 'fecha
                                'cadSQL2 = cadSQL2 & "'" & ado.Recordset.Fields(i).Value & "'"
                                cadSQL2 = cadSQL2 & Date2SF("'" & ado.Recordset.Fields(i).Value & "'")
                            ElseIf mTag.TipoDato = "H" Then 'hora [Monica] A�adido
                                'cadSQL2 = cadSQL2 & "'" & ado.Recordset.Fields(i).Value & "'"
                                cadSQL2 = cadSQL2 & "time('" & Format(ado.Recordset.Fields(i).Value, "hh:mm:ss") & "')"
                            End If
                            
                            Exit For
                        End If
                    Next i
                End If
            End If
        End If
    Next Control
    
    POS2SF = cadSQL2

End Function

'convertix un SQL a una Selection Formula SF
Public Function SQL2SF(cadSQL As String) As String
    Dim cadSQL2 As String
    Dim posP As Integer 'posici� del Punt
    
    cadSQL2 = cadSQL
    
    cadSQL2 = Replace(cadSQL2, "AND (1=1)", "") 'lleva el AND (1=1)  NOTA: a�� ha d'estar abans de traure els parentesi
    cadSQL2 = Replace(cadSQL2, "%", "*") 'canvia el % per un *
    'cadSQL2 = Replace(cadSQL2, "_", "?") 'canvia el _ per un ?
    cadSQL2 = Replace(cadSQL2, "(", "{") 'canvia el ( per una {
    cadSQL2 = Replace(cadSQL2, ")", "") 'lleva )
    
    '+-+- per a posar el } +-+-
    posP = 0
    Do
        If posP = 0 Then
            posP = InStr(cadSQL2, ".")
        Else
            posP = InStr(posP + 1, cadSQL2, ".")
        End If
        If posP > 0 Then cadSQL2 = Left(cadSQL2, posP - 1) & Replace(cadSQL2, " ", "} ", posP, 1)
    Loop Until (InStr(posP + 1, cadSQL2, ".") = 0)
    '+-+-+-+-+-+-+-+-+-+-+-+-+
    
    'per a canviar el format de les dates
    cadSQL2 = Date2SF(CStr(cadSQL2))
    
    'per a canviar els _ per ?
    cadSQL2 = Like2SF(CStr(cadSQL2))
    
    SQL2SF = cadSQL2
End Function

'convertix una data per a pasar-li-la a una Selection Formula SF
' funciona tant en format 2005-01-17 com en 17/01/2005
Public Function Date2SF(cadData As String) As String
' DAVIDV [08/11/2006]: Cambios a causa del mal funcionamiento ocasionado por los criterios
' de b�squeda que contienen los car�cteres - y /, y q no son fechas.
    Dim data, n_data As String
    Dim pos As Integer, pos_desde, pos_hasta As Integer

    pos = InStr(1, cadData, "-")
    While pos <> 0
        pos_desde = InStrRev(cadData, "'", pos) + 1
        pos_hasta = InStr(pos, cadData, "'")
        data = Mid(cadData, pos_desde, pos_hasta - pos_desde)
        If IsDate(data) Then
          n_data = "Date(" & Year(data) & "," & Month(data) & "," & Day(data) & ")"
          cadData = Replace(cadData, "'" & data & "'", n_data)
          '-- LAURA: 27/04/2007
          pos = InStr(pos + 1, cadData, "-")
          '--
        Else
          pos = InStr(pos + 1, cadData, "-")
        End If
    Wend
    
    pos = InStr(1, cadData, "/")
    While pos <> 0
        pos_desde = InStrRev(cadData, "'", pos) + 1
        pos_hasta = InStr(pos, cadData, "'")
        data = Mid(cadData, pos_desde, pos_hasta - pos_desde)
        If IsDate(data) Then
          n_data = "Date(" & Year(data) & "," & Month(data) & "," & Day(data) & ")"
          cadData = Replace(cadData, "'" & data & "'", n_data)
          '-- LAURA: 27/04/2007
          pos = InStr(pos + 1, cadData, "/")
          '--
        Else
          pos = InStr(pos + 1, cadData, "/")
        End If
    Wend

'    While InStr(cadData, "-") <> 0
''        data = Mid(cadData, InStr(cadData, "-") - 5, 12) 'pa llevar les ' '
''        n_data = "Date(" & Mid(data, InStr(data, "-") - 4, 4) & "," & Mid(data, InStr(data, "-") + 1, 2)
''        n_data = n_data & "," & Mid(data, InStr(data, "-") + 4, 2) & ")"
''        cadData = Replace(cadData, data, n_data)
'        data = Mid(cadData, InStr(cadData, "-") - 5, 12) 'pa llevar les ' '
'        'data = cadData
'        n_data = "Date(" & Mid(data, InStr(data, "-") - 4, 4) & "," & Mid(data, InStr(data, "-") + 1, 2)
'        n_data = n_data & "," & Mid(data, InStr(data, "-") + 4, 2) & ")"
'        cadData = Replace(cadData, data, n_data)
'    Wend
    
'    While InStr(cadData, "/") <> 0
''        data = Mid(cadData, InStr(cadData, "/") - 2, 10) 'pa llevar les ' '
''        n_data = "Date(" & Mid(data, InStr(data, "/") + 4, 4) & "," & Mid(data, InStr(data, "/") + 1, 2)
''        n_data = n_data & "," & Mid(data, InStr(data, "/") - 2, 2) & ")"
''        cadData = Replace(cadData, data, n_data)
'
'        data = Mid(cadData, InStr(cadData, "/") - 2, 10) 'pa llevar les ' '
'        'data = cadData
'        n_data = "Date(" & Mid(data, InStr(data, "/") + 4, 4) & "," & Mid(data, InStr(data, "/") + 1, 2)
'        n_data = n_data & "," & Mid(data, InStr(data, "/") - 2, 2) & ")"
'        cadData = Replace(cadData, data, n_data)
'    Wend
    
    Date2SF = cadData
    
End Function

' funci� per a llevar els _ de la cadena
' nom�s els lleva de lo que hi haja entre ' ' i despr�s de LIKE
'per a que no canvie el _ dels noms del camps
Public Function Like2SF(CADENA As String) As String
    Dim cadLike As String
    Dim cadTemp As String
    
    cadLike = CADENA
    
    While InStr(cadLike, "LIKE") <> 0
        cadLike = Mid(cadLike, InStr(cadLike, "LIKE") + 5, Len(cadLike) - 1)
        cadTemp = Mid(cadLike, 1, InStr(2, cadLike, "'"))
        CADENA = Replace(CADENA, cadTemp, Replace(cadTemp, "_", "?"))
    Wend
    
    Like2SF = CADENA
    
End Function
