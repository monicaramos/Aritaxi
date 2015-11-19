Attribute VB_Name = "modBackup"
Option Explicit


Public Sub BACKUP_TablaIzquierda(ByRef Rs As ADODB.Recordset, ByRef CADENA As String)
Dim I As Integer
Dim nexo As String

    CADENA = ""
    nexo = ""
    For I = 0 To Rs.Fields.Count - 1
        CADENA = CADENA & nexo & Rs.Fields(I).Name
        nexo = ","
    Next I
    CADENA = "(" & CADENA & ")"
End Sub



Public Sub BACKUP_Tabla(ByRef Rs As ADODB.Recordset, ByRef Derecha As String, Optional canvi_nom As String, Optional canvi_valor As String)
Dim I As Integer
Dim nexo As String
Dim Valor As String
Dim Tipo As Integer

    Derecha = ""
    nexo = ""
    For I = 0 To Rs.Fields.Count - 1
        Tipo = Rs.Fields(I).Type
        
        If (canvi_nom <> "" And Rs.Fields(I).Name = canvi_nom) Then
            Valor = canvi_valor
            If Tipo = 133 Then
                Valor = "'" & Format(Valor, "yyyy-mm-dd") & "'"
            End If
        Else
            If Tipo = 201 Then 'MEMO
                Valor = DBLetMemo(Rs.Fields(I).Value)
                If Valor <> "" Then
                    NombreSQL Valor
                    Valor = "'" & Valor & "'"
                Else
                    Valor = "NULL"
                End If
            
            Else
                If IsNull(Rs.Fields(I)) Then
                    Valor = "NULL"
                Else
                    'pruebas
                    Select Case Tipo
                    'TEXTO
                    Case 129, 200
                        Valor = Rs.Fields(I)
                        NombreSQL Valor
                        Valor = "'" & Valor & "'"
                    'Fecha
                    Case 133
                        Valor = CStr(Rs.Fields(I))
                        Valor = "'" & Format(Valor, "yyyy-mm-dd") & "'"
                        
                    Case 134 'HORA
                        Valor = DBSet(Valor, "H")
                        
                    Case 135 'Fecha/Hora
                        Valor = DBSet(Rs.Fields(I), "FH", "S")
                    'Numero normal, sin decimales
                    Case 2, 3, 16 To 19
                        Valor = Rs.Fields(I)
                    
                    'Numero con decimales
                    Case 131, 6
                        Valor = CStr(Rs.Fields(I))
                        Valor = TransformaComasPuntos(Valor)
                    Case Else
                        Valor = "Error grave. Tipo de datos no tratado." & vbCrLf
                        Valor = Valor & vbCrLf & "SQL: " & Rs.Source
                        Valor = Valor & vbCrLf & "Pos: " & I
                        Valor = Valor & vbCrLf & "Campo: " & Rs.Fields(I).Name
                        Valor = Valor & vbCrLf & "Valor: " & Rs.Fields(I)
                        MsgBox Valor, vbExclamation
                        MsgBox "El programa finalizara. Avise al soporte t�cnico.", vbCritical
                        End
                    End Select
                End If
            End If
        End If
        Derecha = Derecha & nexo & Valor
        nexo = ","
    Next I
    Derecha = "(" & Derecha & ")"
End Sub

