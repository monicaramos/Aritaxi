Attribute VB_Name = "modCharMultibase"



Public Function RevisaCaracterMultibase(CADENA As String) As String
Dim i As Integer
Dim J As Integer
Dim L As String
Dim C As String

    L = ""
    For i = 1 To Len(CADENA)
        C = Mid(CADENA, i, 1)
        J = Asc(C)
        If J > 125 Then
            Select Case J
            Case 128
                C = "�"
            Case 164  '� minuscula
                C = "�"
            Case 165
                'Es la �
                C = "�"
            Case 166
                C = "�"
            Case 167, 186
                C = "�"
            Case 194
                C = ""
            Case 209
            
            Case Else
                
            End Select
        End If
        L = L & C
    Next i
    
    
    
' CAMBIOS EN MySQL a MySQL by MASL 08092009


       If InStr(CADENA, "Ñ") Then L = Replace(CADENA, "Ñ", "�")
       If InStr(CADENA, "ª") Then L = Replace(CADENA, "ª", "�")
       If InStr(CADENA, "º") Then L = Replace(CADENA, "º", "�")
       If InStr(CADENA, "Ñ") Then L = Replace(CADENA, "Ñ", "�")
       If InStr(CADENA, "§") Then L = Replace(CADENA, "§", "�")
       If InStr(CADENA, "�") Then L = Replace(CADENA, "�", "�")

RevisaCaracterMultibase = L

End Function
