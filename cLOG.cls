VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ACCIONES DISPONIBLES
'
'       1.- Eliminar facturas
'       2.- Facturacion
'       3.- Contabilizar facturas
'       4.- Correccion precios
'       5.- Facturacion desde hasta COMPRAS
'       6.- Facturacion TICKETS agrupados
'
'       7.- Eliminar articulo
'       8.- Modificar FACTURA
'       9.- Modificacion de llamada de SHILLA
'       10.- Eliminacion de llamada de SHILLA
'
Public Function Insertar(Accion As Byte, ByRef ElUsuario As USUARIO, Descripcion As String) As Boolean
Dim Sql As String

    On Error GoTo EI
    Insertar = False
    
    Sql = "insert into `slog` (`fecha`,`accion`,`usuario`,`pc`,`descripcion`) values ( "
    Sql = Sql & " now()," & Accion & " ,'" & DevNombreSQL(ElUsuario.Login) & "','"
    Sql = Sql & DevNombreSQL(ElUsuario.PC) & "',"
    If Descripcion = "" Then
        Sql = Sql & "NULL)"
    Else
        Sql = Sql & "'" & DevNombreSQL(Descripcion) & "')"
    End If
    conn.Execute Sql
    
    
    Insertar = True
    Exit Function
EI:
    MuestraError Err.Number, Err.Description

End Function




Public Function DevuelveAcciones(ByRef ListaAcciones As Collection) As Boolean
 On Error GoTo EDevuelveAcciones
    ListaAcciones.Add "1|Eliminar facturas|"
    ListaAcciones.Add "2|Facturación|"
    ListaAcciones.Add "3|Contabilizar facturas|"
    ListaAcciones.Add "4|Corrección precios|"
    ListaAcciones.Add "5|Facturacion COMPRAS|"
    ListaAcciones.Add "6|Facturacion TICKETS agrupados|"
    ListaAcciones.Add "7|Eliminar articulo|"
    ListaAcciones.Add "8|Modificar Factura|"
    ListaAcciones.Add "9|Modifica SHILLA|"
    ListaAcciones.Add "10|Eliminar SHILLA|"
    DevuelveAcciones = True
    Exit Function
EDevuelveAcciones:
    MuestraError Err.Number
End Function

