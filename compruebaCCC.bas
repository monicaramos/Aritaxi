Attribute VB_Name = "CompruebaCCC"
'-- Esta librer�a contiene un conjunto de funciones de utilidad general
Public Function Comprueba_CC(CC As String) As Boolean
    Dim Ent As String ' Entidad
    Dim Suc As String ' Oficina
    Dim DC As String ' Digitos de control
    Dim i, i2, i3, i4 As Integer
    Dim NumCC As String ' N�mero de cuenta propiamente dicho
    '-- Esta funci�n comprueba la correcci�n de un n�mero de cuenta pasado en CC
    If Len(CC) <> 20 Then Exit Function '-- Las cuentas deben contener 20 d�gitos en total
    
    
    '-- Calculamos el primer d�gito de control
    i = Val(Mid(CC, 1, 1)) * 4
    i = i + Val(Mid(CC, 2, 1)) * 8
    i = i + Val(Mid(CC, 3, 1)) * 5
    i = i + Val(Mid(CC, 4, 1)) * 10
    i = i + Val(Mid(CC, 5, 1)) * 9
    i = i + Val(Mid(CC, 6, 1)) * 7
    i = i + Val(Mid(CC, 7, 1)) * 3
    i = i + Val(Mid(CC, 8, 1)) * 6
    i2 = Int(i / 11)
    i3 = i - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    If i4 <> Val(Mid(CC, 9, 1)) Then Exit Function '-- El primer d�gito de control no coincide
    '-- Calculamos el segundo d�gito de control
    i = Val(Mid(CC, 11, 1)) * 1
    i = i + Val(Mid(CC, 12, 1)) * 2
    i = i + Val(Mid(CC, 13, 1)) * 4
    i = i + Val(Mid(CC, 14, 1)) * 8
    i = i + Val(Mid(CC, 15, 1)) * 5
    i = i + Val(Mid(CC, 16, 1)) * 10
    i = i + Val(Mid(CC, 17, 1)) * 9
    i = i + Val(Mid(CC, 18, 1)) * 7
    i = i + Val(Mid(CC, 19, 1)) * 3
    i = i + Val(Mid(CC, 20, 1)) * 6
    i2 = Int(i / 11)
    i3 = i - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    If i4 <> Val(Mid(CC, 10, 1)) Then Exit Function '-- El segundo d�gito de control no coincide
    '-- Si llega aqu� ambos figitos de control son correctos
    Comprueba_CC = True
End Function


'---- A�ade Laura: 04/10/05
Public Function Comprueba_CuentaBan(CC As String) As Boolean
    'Validar que la cuenta bancaria es correcta
    Comprueba_CuentaBan = False
    If Trim(CC) <> "" Then
        If Not Comprueba_CC(CC) Then
            MsgBox "La cuenta bancaria no es correcta", vbInformation
            Exit Function
        Else
            Comprueba_CuentaBan = True
        End If
    End If
End Function
'------------------------------
