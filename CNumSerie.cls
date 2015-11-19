VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNumSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TABLA: sserie
'Nº SERIE DE EQUIPOS

'Variables locales que contienen valores de propiedad
Private mNumserie As String 'num. serie
Private mCodartic As String 'cod. articulo
Private mCodtipar As String 'cod. tipo articulo

'Private mCodClien As Integer 'cod. cliente
Private mCodClien As Long 'cod. cliente  FAVID. En la bD es un INT

Private mCoddirec As String 'cod. direc./dpto
Private mTieneman As Boolean 'tiene manteniento S/N
Private mNummante As String 'num. mantenimiento
Private mUltrepar As String 'ult. repar
Private mFingaran As String 'fin garantia

Private mCodTipoM As String 'tipo movim.
Private mFechavta As String 'Fecha de venta
Private mNumalbar As String 'Numero de albaran venta
Private mNumFactu As String 'Numero de factura venta
Private mNumLine1 As String 'Num. linea del albaran

'Private mCodProve As Integer 'cod. proveedor
Private mCodProve As Long 'cod. proveedor DAVID

Private mNumalbpr As String 'mum. Albaran prove
Private mFechaCom As String 'Fecha de compra
Private mNumline2 As String 'Num. linea Albaran prove



'------------------------------------------------
'Propiedades del modulo CSerie
'------------------------------------------------

'**** num. Serie

Public Property Let numSerie(ByVal vData As String)
     mNumserie = vData
End Property

Public Property Get numSerie() As String
     numSerie = mNumserie
End Property


'**** cod. Artic

Public Property Let Articulo(ByVal vData As String)
     mCodartic = vData
End Property

Public Property Get Articulo() As String
     Articulo = mCodartic
End Property


'**** Tipo Articulo

Public Property Let TipoArt(ByVal vData As String)
     mCodtipar = vData
End Property

Public Property Get TipoArt() As String
     TipoArt = mCodtipar
End Property


'**** Fecha fin garantia
Public Property Let FinGarantia(ByVal vData As String)
     mFingaran = vData
End Property

Public Property Get FinGarantia() As String
     FinGarantia = mFingaran
End Property



'**** Cod. Cliente

Public Property Let Cliente(ByVal vData As Long)
     mCodClien = vData
End Property

Public Property Get Cliente() As Long
     Cliente = mCodClien
End Property


'**** Cod. Direc/dpto

Public Property Let DirDpto(ByVal vData As String)
     mCoddirec = vData
End Property

Public Property Get DirDpto() As String
     DirDpto = mCoddirec
End Property


'**** Tiene mantenimiento

Public Property Let conMante(ByVal vData As Boolean)
     mTieneman = vData
End Property

Public Property Get conMante() As Boolean
     conMante = mTieneman
End Property


'**** Num. MAntenimiento

Public Property Let nummante(ByVal vData As String)
     mNummante = vData
End Property

Public Property Get nummante() As String
     nummante = mNummante
End Property


'**** Cod. tipo movim.

Public Property Let tipoMov(ByVal vData As String)
     mCodTipoM = vData
End Property

Public Property Get tipoMov() As String
     tipoMov = mCodTipoM
End Property


'**** Fecha Venta

Public Property Let FechaVta(ByVal vData As String)
     mFechavta = vData
End Property

Public Property Get FechaVta() As String
     FechaVta = mFechavta
End Property


'**** Num. Albaran

Public Property Let NumAlbaran(ByVal vData As String)
     mNumalbar = vData
End Property

Public Property Get NumAlbaran() As String
     NumAlbaran = mNumalbar
End Property


'**** Num. Linea Albaran

Public Property Let NumLinAlb(ByVal vData As String)
     mNumLine1 = vData
End Property

Public Property Get NumLinAlb() As String
     NumLinAlb = mNumLine1
End Property



'**** Num. Factura

Public Property Let NumFactura(ByVal vData As String)
     mNumFactu = vData
End Property

Public Property Get NumFactura() As String
     NumFactura = mNumFactu
End Property



'**** Cod. Proveedor

Public Property Let Proveedor(ByVal vData As Long)
     mCodProve = vData
End Property

Public Property Get Proveedor() As Long
     Proveedor = mCodProve
End Property

'**** Num. Albaran Proveedor (compras)

Public Property Let NumAlbProve(ByVal vData As String)
     mNumalbpr = vData
End Property

Public Property Get NumAlbProve() As String
     NumAlbProve = mNumalbpr
End Property


'**** Fecha Compra

Public Property Let fechaCom(ByVal vData As String)
     mFechaCom = vData
End Property

Public Property Get fechaCom() As String
     fechaCom = mFechaCom
End Property



'**** Num. Linea Albaran Proveedor

Public Property Let NumLinAlbPr(ByVal vData As String)
     mNumline2 = vData
End Property

Public Property Get NumLinAlbPr() As String
     NumLinAlbPr = mNumline2
End Property


'------------------------------------------------
'Procedimientos del modulo CNumSerie
'------------------------------------------------


Public Function LeerDatos(NumSer As String, codArt As String) As Boolean
'Leer los datos de un numero de serie
'Lee de la BD: Aritaxi, Tabla: sserie
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error GoTo ELeer
    LeerDatos = False
    
    Sql = "SELECT *"
    Sql = Sql & " FROM sserie "
    Sql = Sql & " WHERE numserie=" & DBSet(NumSer, "T") & " AND codartic=" & DBSet(codArt, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Rs.EOF Then
        LeerDatos = False
    Else
        Me.numSerie = DBLet(Rs!numSerie, "T")
        Me.Articulo = Rs!codArtic
        
        Me.TipoArt = Rs!codTipar
        Me.Cliente = DBLet(Rs!CodClien, "N")
        Me.DirDpto = DBLet(Rs!CodDirec, "T")
        
        Me.conMante = IIf(DBLet(Rs!TieneMan, "N") = 1, True, False)
        Me.nummante = DBLet(Rs!nummante, "T")
        Me.FinGarantia = DBLet(Rs!fingaran, "F")
        
        Me.tipoMov = DBLet(Rs!codtipom, "T")
        Me.NumFactura = DBLet(Rs!NumFactu, "T")
        Me.FechaVta = DBLet(Rs!FechaVta, "F")
        Me.NumAlbaran = DBLet(Rs!NumAlbar, "T")
        Me.NumLinAlb = DBLet(Rs!numline1, "T")
        
        'FALTA leer datos proveedor

        LeerDatos = True
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Set Rs = Nothing
    LeerDatos = False
    MuestraError Err.Number, "Leer datos nº serie.", Err.Description
End Function




Public Function InsertarNumSerie() As Boolean
Dim Sql As String

    On Error GoTo EInsertar

    If TipoArt = "" Then
        TipoArt = DevuelveDesdeBDNew(conAri, "sartic", "codtipar", "codartic", Articulo, "T")
    End If

    'Insertar en la tabla sserie
    Sql = "INSERT INTO sserie (numserie, codartic, codtipar, codclien, coddirec,tieneman, nummante, ultrepar, fingaran, "
    Sql = Sql & " codtipom, numfactu, fechavta, numalbar, numline1, codprove, numalbpr, fechacom, numline2) "
    
    Sql = Sql & " VALUES (" & DBSet(numSerie, "T") & ", " & DBSet(Articulo, "T") & ", " & DBSet(TipoArt, "T") & ","
    Sql = Sql & DBSet(Cliente, "N", "S") & "," & DBSet(DirDpto, "N", "S") & ","
    Sql = Sql & DBSet(conMante, "B") & "," & DBSet(nummante, "T") & ","
    Sql = Sql & ValorNulo & "," & DBSet(FinGarantia, "F") & "," 'ult. repar, fin garantia
    Sql = Sql & DBSet(tipoMov, "T") & "," & ValorNulo & "," & DBSet(FechaVta, "F", "S") & "," 'codtipom,numfactu, fechavta
    Sql = Sql & DBSet(NumAlbaran, "N", "S") & "," & DBSet(NumLinAlb, "N", "S") & "," 'numalbar, numline1
    Sql = Sql & DBSet(Proveedor, "N", "S") & ","
    Sql = Sql & DBSet(NumAlbProve, "T") & "," & DBSet(fechaCom, "F") & "," & DBSet(NumLinAlbPr, "N", "S") & ")"
    
    conn.Execute Sql
    InsertarNumSerie = True
    Exit Function
    
EInsertar:
    InsertarNumSerie = False
    MsgBox "Error al insertar el Nº de Serie: " & numSerie & " del artículo: " & Articulo, vbExclamation
End Function



Public Function ActualizarNumSerie(DeVentas As Boolean) As Boolean
'UPDATE en la tabla sserie de los campos de venta o compra
Dim Sql As String

    On Error GoTo EActualizar

    'Insertar en la tabla sserie
    If DeVentas Then 'Se llama desde Ventas a Clientes
        Sql = "UPDATE sserie SET codclien=" & DBSet(Cliente, "N", "S") & ","
        Sql = Sql & " coddirec=" & DBSet(DirDpto, "N", "S") & ","
        Sql = Sql & " tieneman=" & DBSet(conMante, "B") & ","
        Sql = Sql & " nummante=" & DBSet(nummante, "T", "S") & ","
        Sql = Sql & " fingaran=" & DBSet(FinGarantia, "F", "S") & ","
        Sql = Sql & " codtipom=" & DBSet(tipoMov, "T", "S") & ","
        Sql = Sql & " fechavta=" & DBSet(FechaVta, "F", "S") & ","
        Sql = Sql & " numalbar=" & DBSet(NumAlbaran, "N", "S") & ","
        Sql = Sql & " numfactu=" & DBSet(NumFactura, "N", "S") & ","
        Sql = Sql & " numline1=" & DBSet(NumLinAlb, "N", "S")
        
        
    Else 'se llama desde Compras a Proveedores
        Sql = "UPDATE sserie SET codprove=" & DBSet(Proveedor, "N", "S") & ","
        Sql = Sql & "numalbpr=" & DBSet(NumAlbProve, "T") & ","
        Sql = Sql & "fechacom=" & DBSet(fechaCom, "F", "S") & ","
        Sql = Sql & "numline2=" & DBSet(NumLinAlbPr, "N", "S")
    End If
    
    Sql = Sql & " WHERE numserie=" & DBSet(numSerie, "T") & " AND codartic=" & DBSet(Articulo, "T")
      
    conn.Execute Sql
    ActualizarNumSerie = True
    Exit Function
    
EActualizar:
    ActualizarNumSerie = False
End Function



Public Function EliminarNumSerie() As Boolean
'Elimina un nº serie de un articulo
Dim Sql As String
    
    On Error GoTo EEliminar

    Sql = "DELETE FROM sserie "
    Sql = Sql & " WHERE numserie=" & DBSet(numSerie, "T") & " AND codartic=" & DBSet(Articulo, "T")

    conn.Execute Sql
    EliminarNumSerie = True
    Exit Function

EEliminar:
    EliminarNumSerie = False
End Function


Public Function BorrarNumSeriesAlbVta() As Boolean
'Actualiza los campos cliente,codtipom,numfactu,numalbar,fehcavta,numline1
'de todos los nº serie de un articulo asignados a una determinada linea de un albaran de venta
Dim Sql As String
    
    On Error GoTo EBorrar

    Sql = "UPDATE sserie SET  codclien=" & ValorNulo & ", codtipom=" & ValorNulo & ","
    Sql = Sql & "numfactu=" & ValorNulo & ", fechavta=" & ValorNulo & ", numalbar= " & ValorNulo & ",numline1=" & ValorNulo
    Sql = Sql & " WHERE codtipom=" & DBSet(tipoMov, "T") & " AND numalbar=" & NumAlbaran & " AND numline1=" & NumLinAlb & " and year(fechavta)=" & Year(Me.FechaVta)

    conn.Execute Sql

EBorrar:
    If Err.Number <> 0 Then
        BorrarNumSeriesAlbVta = False
    Else
        BorrarNumSeriesAlbVta = True
    End If
End Function




Public Function TieneAsignadaVenta() As Boolean
'Comprueba si el nº de serie de un articulo ha sido asignado a algun albaran de venta
'o factura de venta en cuyo caso no se podrá eliminar el nº de serie
Dim Sql As String
    
    On Error GoTo EAsignada
    
    If Me.NumAlbaran <> "" Or Me.NumFactura <> "" Then
        TieneAsignadaVenta = True
    Else
        TieneAsignadaVenta = False
    End If
    
    Exit Function
    
EAsignada:
    TieneAsignadaVenta = False
    MuestraError Err.Number, "Comprovando si nº de serie asignado a una venta", Err.Description
End Function






Public Sub ObtenFechaFinGarantia(codArt As String, Fecha As String)
'calcula la fecha fin de garantia a partir de los dias de garantia del articulo
Dim cad As String
    On Error GoTo EFechaFin

    'obtenemos los dias de garantia del articulo
    cad = DevuelveDesdeBDNew(conAri, "sartic", "garantia", "codartic", codArt, "T")
    'fin garantia= fecha albaran + dias de garantia
    If Fecha <> "" Then FinGarantia = CStr(CDate(Fecha) + CInt(ComprobarCero(cad)))
EFechaFin:
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Function RecuperarParaVenta(FechaRec As String, Optional DesdeCompra As Boolean) As Boolean
Dim Sql As String
Dim NumL As String

    On Error GoTo ErrRecup
    
    '[Monica]24/10/2012: lo queremos en algunos casos fuera de transaccion
    '                    por defecto siempre dentro de transaccion
    If Not DesdeCompra Then conn.BeginTrans
    
    '-- obtener el siguiente num linea
    Sql = "numserie=" & DBSet(Me.numSerie, "T") & " and codartic=" & DBSet(Me.Articulo, "T")
    NumL = SugerirCodigoSiguienteStr("sserlin", "numlinea", Sql)
    
    '-- pasar datos la las lineas del nº serie
    Sql = "INSERT INTO sserlin (numserie,codartic,numlinea,fecharec,codclien,coddirec,tieneman,nummante,codtipom,numfactu,fechavta,numalbar,numline1"
    '[Monica]14/02/2014: insertamos los datos de quien lo compra para actualizar la parte de compra y dejarlo en hco
    Sql = Sql & ",codprove, numalbpr, fechacom, numline2)"
    '
    Sql = Sql & " VALUES (" & DBSet(Me.numSerie, "T") & "," & DBSet(Me.Articulo, "T") & "," & NumL & ","
    Sql = Sql & DBSet(FechaRec, "F", "N") & "," & Me.Cliente & "," & DBSet(Me.DirDpto, "N", "S") & ","
    Sql = Sql & DBSet(Me.conMante, "B") & "," & DBSet(Me.nummante, "T", "S") & "," & DBSet(Me.tipoMov, "T", "S") & ","
    Sql = Sql & DBSet(Me.NumFactura, "N", "S") & "," & DBSet(Me.FechaVta, "F", "S") & ","
    Sql = Sql & DBSet(Me.NumAlbaran, "N", "S") & "," & DBSet(Me.NumLinAlb, "N", "S")
    
    '[Monica]14/02/2014: antiguo comprador
    Dim sqlSerie As String
    Dim RsSerie As ADODB.Recordset
    sqlSerie = "select codprove, numalbpr, fechacom, numline2 from sserie where numserie = " & DBSet(numSerie, "T") & " AND codartic=" & DBSet(Articulo, "T")
    Set RsSerie = New ADODB.Recordset
    RsSerie.Open sqlSerie, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RsSerie.EOF Then
        Sql = Sql & "," & DBSet(RsSerie.Fields(0), "N") & "," & DBSet(RsSerie.Fields(1), "T") & "," & DBSet(RsSerie.Fields(2), "F") & "," & DBLet(RsSerie.Fields(3), "N") & ")"
    Else
        Sql = Sql & ")"
    End If
    
    
    conn.Execute Sql
    
    '-- actualizar la cabecera limpiando los datos venta
    Sql = "UPDATE sserie SET codclien=" & ValorNulo & ","
    Sql = Sql & " coddirec=" & ValorNulo & ","
    Sql = Sql & " tieneman=" & ValorNulo & ","
    Sql = Sql & " nummante=" & ValorNulo & ","
'    SQL = SQL & " fingaran=" & DBSet(FinGarantia, "F", "S") & ","
    Sql = Sql & " codtipom=" & ValorNulo & ","
    Sql = Sql & " numfactu=" & ValorNulo & ","
    Sql = Sql & " fechavta=" & ValorNulo & ","
    Sql = Sql & " numalbar=" & ValorNulo & ","
    Sql = Sql & " numline1=" & ValorNulo
    
    '[Monica]14/02/2014: falta insertar los datos de la nueva compra
    Sql = Sql & ", codprove=" & DBSet(Proveedor, "N")
    Sql = Sql & ", numalbpr=" & DBSet(NumAlbProve, "T")
    Sql = Sql & ", fechacom=" & DBSet(fechaCom, "F")
    Sql = Sql & ", numline2=" & DBSet(NumLinAlbPr, "N")
    
    
    Sql = Sql & " WHERE numserie=" & DBSet(numSerie, "T") & " AND codartic=" & DBSet(Articulo, "T")
    conn.Execute Sql
    
    
    If Not DesdeCompra Then conn.CommitTrans
    RecuperarParaVenta = True
    Exit Function
    
ErrRecup:
    RecuperarParaVenta = False
    If Not DesdeCompra Then conn.RollbackTrans
    MuestraError Err.Number, "Recuperar nº serie para venta.", Err.Description
End Function



