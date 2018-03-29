Attribute VB_Name = "ModHistorico"
Option Explicit
'Modulo para el traspaso de registros de cabecera y lineas de las tablas
'de OFERTAS,PEDIDOS,ALBARANES
'A las tablas del HISTORICO de Ofertas,Pedidos,Albaranes
'OFERTAS:
' scapre --> schpre
' slipre --> slhpre
'PEDIDOS:
' scaped --> schped
' sliped --> slhped


Dim CodTipoMov As String
Dim NomTabla As String 'nombre de la tabla
Dim NomTablaH As String 'nombre de la tabla del historico al que movemos
Dim NomTablaLin As String 'nombre tabla de lineas
Dim NomTablaLinH As String 'nombre tabla de lineas del historico


Public Function ActualizarElTraspaso(ByRef ADonde As String, cadWHERE As String, codMovim As String, Optional cadL As String) As Boolean
'codMovim: tipo de movimiento que estamos hacienco: OFE,PEV,ALV,PEC,ALC,....
    
    ActualizarElTraspaso = False
    CodTipoMov = codMovim
    
    '[Monica]++20/12/2010 Añadido todo este punto borrar de elementos eliminados
    ADonde = "Borrar cabeceras y lineas de elementos eliminados."
    If Not EliminarPreexistente(False, cadWHERE) Then Exit Function
    
    
    
    'Insertamos en cabeceras Historico
    ADonde = "Insertando datos en histórico cabeceras "
    If Not InsertarCabeceraHistorico(cadWHERE, cadL) Then Exit Function
'    IncrementarProgres 2
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Histórico lineas "
    If Not InsertarLineasHistorico(cadWHERE) Then Exit Function
'    IncrementarProgres 2
    
    'Borramos cabeceras y lineas
    ADonde = "Borrar cabeceras y lineas"
    If Not BorrarTraspaso(False, cadWHERE) Then Exit Function
'    IncrementarProgres 2

    ActualizarElTraspaso = True
End Function


Private Function InsertarCabeceraHistorico(cadWHERE As String, Optional cadL As String) As Boolean
Dim Sql As String
On Error Resume Next

    Select Case CodTipoMov
      Case "PEV" 'pedidos de venta a clientes
        NomTabla = "scaped"
        NomTablaH = "schped"
        Sql = " SELECT numpedcl,fecpedcl," & vUsu.Codigo Mod 1000 & " as codigusu," & cadL & ","
        Sql = Sql & "fecentre,sementre,visadore,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        Sql = Sql & "coddirec,nomdirec,referenc,codtraba,codagent,codforpa,dtoppago,dtognral,"
        Sql = Sql & "tipofact,observa01,observa02,observa03,observa04,observa05,servcomp,restoped,numofert,fecofert,observap1,observap2,recogecl"
        
      Case "ALV", "ALM", "ALR", "ALS", "ART", "ARC"  '[1.3.1] 'Albaran de venta a clientes
        NomTabla = "scaalb"
        NomTablaH = "schalb"
        Sql = " SELECT codtipom,numalbar,fechaalb," & vUsu.Codigo Mod 1000 & " as codigusu," & cadL & ","
        Sql = Sql & "factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        Sql = Sql & "coddirec,nomdirec,referenc,facturkm,cantidkm,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,"
        Sql = Sql & "tipofact,observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,esticket,numtermi,numventa "
        
        
      Case "OFE" 'Ofertas a Clientes
        NomTabla = "scapre"
        NomTablaH = "schpre"
        Sql = " SELECT numofert, fecofert," & "'" & Format(Now, FormatoFecha) & "' as fechamov, fecentre, aceptado, codclien, nomclien, domclien, codpobla, "
        Sql = Sql & "pobclien, proclien, nifclien, telclien, coddirec, nomdirec, referenc, codtraba, codagent, codforpa, dtoppago, dtognral, tipofact, "
        Sql = Sql & "plazos01, plazos02, plazos03, asunto01, asunto02, asunto03, asunto04, asunto05, observa01, observa02, observa03, observa04, observa05, "
        Sql = Sql & "concepto, seguiofe "
        
      Case "ALC" 'Albaranes a Proveedores (Compras)
        NomTabla = "scaalp"
        NomTablaH = "schalp"
        Sql = " (numalbar,fechaalb,codprove,codigusu,fechelim,trabelim,codincid,nomprove,domprove,"
        Sql = Sql & "codpobla,pobprove,proprove,nifprove,telprove,codforpa,codtraba,codtrab1,dtoppago,dtognral,"
        Sql = Sql & "observa1,observa2,observa3,observa4,observa5,numpedpr,fecpedpr) "
        Sql = Sql & " SELECT numalbar,fechaalb,codprove," & vUsu.Codigo Mod 1000 & " as codigusu," & cadL & ","
        Sql = Sql & "nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,"
        Sql = Sql & "codforpa,codtraba,codtrab1,dtoppago,dtognral,"
        Sql = Sql & "observa1,observa2,observa3,observa4,observa5,numpedpr,fecpedpr"
      
      Case "PEC" 'Pedidos a Proveedores (Compras)
        NomTabla = "scappr"
        NomTablaH = "schppr"
        Sql = " SELECT numpedpr,fecpedpr," & vUsu.Codigo Mod 1000 & " as codigusu," & cadL & ","
        Sql = Sql & "codprove,nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,"
        Sql = Sql & "coddirea,coddiref,codforpa,codtraba,codtrab1,dtognral,dtoppago,"
        Sql = Sql & "restoped,codclien,observa1,observa2,observa3,observa4,observa5,tipoporte"
        
      Case "ARN", "ARP"
        NomTabla = "scaalbcli"
        NomTablaH = "schalbcli"
        Sql = " SELECT codtipom,numalbar,fechaalb," & vUsu.Codigo Mod 1000 & " as codigusu," & cadL & ","
        Sql = Sql & "factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        Sql = Sql & "coddirec,nomdirec,referenc,facturkm,cantidkm,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,"
        Sql = Sql & "tipofact,observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,esticket,numtermi,numventa "
      
    End Select
    
    Sql = Sql & " FROM " & NomTabla & " WHERE " & cadWHERE
    Sql = "INSERT INTO " & NomTablaH & Sql
    
    conn.Execute Sql
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabeceraHistorico = False
    Else
        InsertarCabeceraHistorico = True
    End If
End Function


Private Function InsertarLineasHistorico(cadWHERE As String) As Boolean
Dim Sql As String
On Error Resume Next

    Select Case CodTipoMov
      Case "PEV" 'pedidos ventas a clientes
        NomTablaLin = "sliped"
        NomTablaLinH = "slhped"
        Sql = " SELECT scaped.numpedcl,scaped.fecpedcl,sliped.numlinea,sliped.codalmac,sliped.codartic,sliped.nomartic,sliped.ampliaci,sliped.cantidad,servidas,numbultos,precioar,dtoline1,dtoline2,importel,origpre,numlote,codccost "
        Sql = Sql & " FROM scaped INNER JOIN sliped on scaped.numpedcl=sliped.numpedcl "
        Sql = Sql & " WHERE " & cadWHERE
        
      Case "ALV", "ALM", "ALR", "ALS", "ART", "ARC" '[1.3.1] 'Albaranes ventas a clientes, Mantenimientos y Reparaciones
        NomTablaLin = "slialb"
        NomTablaLinH = "slhalb"
        Sql = " SELECT scaalb.codtipom,scaalb.numalbar,scaalb.fechaalb,slialb.numlinea,slialb.codalmac,slialb.codartic,slialb.nomartic,slialb.ampliaci,slialb.cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre ,codproveX,numlote,codccost"
        Sql = Sql & " FROM scaalb INNER JOIN slialb on scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar "
        Sql = Sql & " WHERE " & cadWHERE
        
      Case "OFE" 'Ofertas a clientes
        NomTablaLin = "slipre"
        NomTablaLinH = "slhpre"
        Sql = " SELECT scapre.numofert,scapre.fecofert,slipre.numlinea,slipre.codalmac,slipre.codartic,slipre.nomartic,slipre.ampliaci,slipre.cantidad,precioar,dtoline1,dtoline2,importel,origpre,codprovex "
        Sql = Sql & " FROM scapre INNER JOIN slipre on scapre.numofert=slipre.numofert"
        Sql = Sql & " WHERE " & cadWHERE
        
      Case "ALC" 'Albaranes compras a proveedores
        NomTablaLin = "slialp"
        NomTablaLinH = "slhalp"
        Sql = "(numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,numlotes,codccost) "
        Sql = Sql & " SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codprove,slialp.numlinea,slialp.codartic,slialp.codalmac,slialp.nomartic,slialp.ampliaci,slialp.cantidad,precioar,dtoline1,dtoline2,importel,numlotes,codccost "
        Sql = Sql & " FROM scaalp INNER JOIN slialp on scaalp.numalbar=slialp.numalbar AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        Sql = Sql & " WHERE " & cadWHERE
      
      Case "PEC" 'Pedidos compras a proveedores
        NomTablaLin = "slippr"
        NomTablaLinH = "slhppr"
        Sql = " SELECT scappr.numpedpr,scappr.fecpedpr,slippr.numlinea,slippr.codartic,slippr.codalmac,slippr.nomartic,slippr.ampliaci,slippr.cantidad,slippr.recibida,precioar,dtoline1,dtoline2,importel,slippr.codccost "
        Sql = Sql & " FROM scappr INNER JOIN slippr on scappr.numpedpr=slippr.numpedpr "
        Sql = Sql & " WHERE " & cadWHERE
      
      Case "ARN", "ARP"
        NomTablaLin = "slialbcli"
        NomTablaLinH = "slhalbcli"
        Sql = " SELECT scaalbcli.codtipom,scaalbcli.numalbar,scaalbcli.fechaalb,slialbcli.numlinea,slialbcli.codalmac,slialbcli.codartic,slialbcli.nomartic,slialbcli.ampliaci,slialbcli.cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre ,codproveX,numlote,codccost"
        Sql = Sql & " FROM scaalbcli INNER JOIN slialbcli on scaalbcli.codtipom=slialbcli.codtipom AND scaalbcli.numalbar=slialbcli.numalbar "
        Sql = Sql & " WHERE " & cadWHERE
    
    
    End Select
    
    Sql = "INSERT INTO " & NomTablaLinH & Sql
    
    conn.Execute Sql
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasHistorico = False
    Else
        InsertarLineasHistorico = True
    End If
End Function


Private Function BorrarTraspaso(EnHistorico As Boolean, cadWHERE As String) As Boolean
'Si EnHistorico=true borra de las tablas de historico: "schtra" y "slhtra"
'Si EnHistorico=false borra de las tablas de traspaso: "scatra" y "slitra"
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String, cadAux As String

    BorrarTraspaso = False
    On Error GoTo EBorrar
    
    
    'Eliminamos las lineas
    Select Case CodTipoMov
      Case "PEV" 'pedidos ventas  a clientes
        Sql = "Select numpedcl from scaped WHERE " & cadWHERE
        cadAux = " numpedcl IN "
      Case "ALV", "ALM", "ALR", "ALS", "ART", "ARC" '[1.3.1] 'albaranes ventas a clientes,Mantenimientos y Reparaciones
        Sql = "Select numalbar from scaalb WHERE " & cadWHERE
        cadAux = "codtipom=" & DBSet(CodTipoMov, "T") & " AND numalbar IN "
      Case "OFE" 'Ofertas a clientes
        Sql = "Select numofert from scapre WHERE " & cadWHERE
        cadAux = " numofert IN "
      Case "ALC" 'Albaranes compras a proveedores
'        SQL = "Select numalbar,fechaalb,codprove from scaalp WHERE " & cadWHERE
'        cadAux = " numalbar IN "
      Case "ARN", "ARP"
        Sql = "Select numalbar from scaalbcli WHERE " & cadWHERE
        cadAux = "codtipom=" & DBSet(CodTipoMov, "T") & " AND numalbar IN "
    
    End Select
    
    If CodTipoMov <> "ALC" And CodTipoMov <> "PEC" Then
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rs.EOF
            If CodTipoMov <> "ALC" Then
                Cad = Cad & Rs.Fields(0).Value & ","
            Else
                Cad = Cad & "numalbar="
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        'Quitar la ultima coma de la cadena
        Cad = Mid(Cad, 1, Len(Cad) - 1)
        
        cadAux = cadAux & "(" & Cad & ")"
    Else
        cadAux = Replace(cadWHERE, NomTabla, NomTablaLin)
    End If
    
    Sql = "DELETE FROM " & NomTablaLin & " WHERE " & cadAux

    conn.Execute Sql
    
    'La cabecera
    Sql = "Delete from " & NomTabla
    Sql = Sql & " WHERE " & cadWHERE
    conn.Execute Sql
    BorrarTraspaso = True
    
EBorrar:
    If Err.Number <> 0 Then
        BorrarTraspaso = False
    Else
        BorrarTraspaso = True
    End If
End Function

Private Function EliminarPreexistente(EnHistorico As Boolean, cadWHERE As String) As Boolean
'Si EnHistorico=true borra de las tablas de historico: "schtra" y "slhtra"
'Si EnHistorico=false borra de las tablas de traspaso: "scatra" y "slitra"
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String, cadAux As String

    EliminarPreexistente = False
    On Error GoTo EBorrar
    
    
    'Eliminamos las lineas
    Select Case CodTipoMov
      Case "PEV" 'pedidos ventas  a clientes
        Sql = "Select numpedcl from schped WHERE " & Replace(cadWHERE, "scaped", "schped")
        cadAux = " numpedcl IN "
      Case "ALV", "ALM", "ALR", "ALS", "ART", "ARC" '[1.3.1] 'albaranes ventas a clientes,Mantenimientos y Reparaciones
        Sql = "Select numalbar from schalb WHERE " & Replace(cadWHERE, "scaalb", "schalb")
        cadAux = "codtipom=" & DBSet(CodTipoMov, "T") & " AND numalbar IN "
      Case "OFE" 'Ofertas a clientes
        Sql = "Select numofert from schpre WHERE " & Replace(cadWHERE, "scapre", "schpre")
        cadAux = " numofert IN "
      Case "ALC" 'Albaranes compras a proveedores
'        SQL = "Select numalbar,fechaalb,codprove from scaalp WHERE " & cadWHERE
'        cadAux = " numalbar IN "
      Case "ARN", "ARP" ' rectificativas de servicios de cliente
        Sql = "Select numalbar from schalbcli WHERE " & Replace(cadWHERE, "scaalbcli", "schalbcli")
        cadAux = "codtipom=" & DBSet(CodTipoMov, "T") & " AND numalbar IN "
    

    End Select
    
    If CodTipoMov <> "ALC" And CodTipoMov <> "PEC" Then
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        
        If Not Rs.EOF Then
            Select Case CodTipoMov
              Case "PEV" 'pedidos ventas  a clientes
                    Sql = "DELETE FROM slhped WHERE " & Replace(cadWHERE, "scaped", "slhped")
                    conn.Execute Sql
    
                    'La cabecera
                    Sql = "Delete from schped WHERE " & Replace(cadWHERE, "scaped", "schped")
                    conn.Execute Sql
        
              Case "ALV", "ALM", "ALR", "ALS", "ART", "ARC" '[1.3.1] 'albaranes ventas a clientes,Mantenimientos y Reparaciones
                    Sql = "DELETE FROM slhalb WHERE " & Replace(cadWHERE, "scaalb", "slhalb")
                    conn.Execute Sql
    
                    'La cabecera
                    Sql = "Delete from schalb WHERE " & Replace(cadWHERE, "scaalb", "schalb")
                    conn.Execute Sql
              
              
              Case "OFE" 'Ofertas a clientes
                    Sql = "DELETE FROM slhpre WHERE " & Replace(cadWHERE, "scapre", "slhpre")
                    conn.Execute Sql
    
                    'La cabecera
                    Sql = "Delete from schpre WHERE " & Replace(cadWHERE, "scapre", "schpre")
                    conn.Execute Sql
              
              Case "ALC" 'Albaranes compras a proveedores
        '        SQL = "Select numalbar,fechaalb,codprove from scaalp WHERE " & cadWHERE
        '        cadAux = " numalbar IN "
        
              Case "ARN", "ARP"
                    Sql = "DELETE FROM slhalbcli WHERE " & Replace(cadWHERE, "scaalbcli", "slhalbcli")
                    conn.Execute Sql
    
                    'La cabecera
                    Sql = "Delete from schalbcli WHERE " & Replace(cadWHERE, "scaalbcli", "schalbcli")
                    conn.Execute Sql
              
            
            End Select
        End If
        Set Rs = Nothing
    End If
    EliminarPreexistente = True
    
EBorrar:
    If Err.Number <> 0 Then
        EliminarPreexistente = False
    Else
        EliminarPreexistente = True
    End If
End Function




'========================================================

Public Sub CargarTagsHco(ByRef F As Form, vTabla As String, vTablaHco As String)
'Sustituye en los tags del formulario la tabla de Reparaciones (scarep)
'por la del historico de Reparaciones (schrep)
Dim Control As Object
Dim vtag As String

    For Each Control In F.Controls
        If Control.Tag <> "" Then
            vtag = Control.Tag
'            vtag = SustituirCadenas(vtag, vTabla, vTablaHco)
            vtag = Replace(vtag, vTabla, vTablaHco)
            Control.Tag = vtag
        End If
    Next Control
End Sub
