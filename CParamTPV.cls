VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamTPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros SPATPVG
Private mvarCodClien As String
Private mvarNomClien As String
Private mvarCtaBanc1 As String 'cta prevista de cobro

Private mvarCodForpa As String 'forma de pago por defecto
Private mvarNomForPa As String
Private mvarTipForPa As Byte 'Tipo de pago

Private mvarTipCierre As Byte 'tipo cierre: por terminal o conjunto.
'Private mvarModoCalculo As Byte 'calcular importes de la linea con iva o no
                                '1=con IVA 0=sin IVA

'Variables asociadas a cada campo de la tabla de parametros SPATPVT
'Del terminal conectado
Private mvarContador As Long
Private mvarNomImpre As String 'impresora impresion tickets
Private mvarHayVisor As Boolean
Private mvarNumPuerto As Byte
Private mvarVelocPue As String

'Introduccion de ticket rapida
Private mvarRapida As Boolean
Private mvarFormaPagoUnica As Boolean   'Este no lo lee de parametros, si no cuando abre el terminal lo comrpueba
Private mvarImprimiDirecto As Boolean

'Desglosa IVAS al pie del TICKET
Private mvarDesglosaIVATicket As Boolean
Private mvarAbreCajon As Boolean

'Solo tendra get
Private mvarSecuenciaCajon As String


'Enero 10
'Redondeo a dos decimales
Private mvarRedondea2 As Boolean
Private mvarCalculaIVAsobrePVP As Boolean


'Control stock al finalizar la venta
Private mvarCtrstockVenta As Boolean


Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes


'------------------------------------------------
'Propiedades del modulo CParamTPV
'------------------------------------------------

Public Property Let Cliente(ByVal vData As String)
     mvarCodClien = vData
End Property

Public Property Get Cliente() As String
     Cliente = mvarCodClien
End Property


Public Property Let NomCliente(ByVal vData As String)
     mvarNomClien = vData
End Property

Public Property Get NomCliente() As String
     NomCliente = mvarNomClien
End Property



'cuenta prevista de cobro por si pasamos el ticket a una factura de venta FAV
Public Property Let CtaPrevCobro(ByVal vData As String)
     mvarCtaBanc1 = vData
End Property

Public Property Get CtaPrevCobro() As String
     CtaPrevCobro = mvarCtaBanc1
End Property


Public Property Let ForPago(ByVal vData As String)
     mvarCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mvarCodForpa
End Property



Public Property Let NomForPago(ByVal vData As String)
     mvarNomForPa = vData
End Property

Public Property Get NomForPago() As String
     NomForPago = mvarNomForPa
End Property


Public Property Let TipoForPago(ByVal vData As Byte)
     mvarTipForPa = vData
End Property

Public Property Get TipoForPago() As Byte
     TipoForPago = mvarTipForPa
End Property


Public Property Let TipoCierre(ByVal vData As Byte)
     mvarTipCierre = vData
End Property

Public Property Get TipoCierre() As Byte
     TipoCierre = mvarTipCierre
End Property


'Public Property Let ModoCalculo(ByVal vData As Byte)
'     mvarModoCalculo = vData
'End Property
'
'Public Property Get ModoCalculo() As Byte
'     ModoCalculo = mvarModoCalculo
'End Property



Public Property Let Contador(ByVal vData As Long)
    mvarContador = vData
End Property


Public Property Let NomImpresora(ByVal vData As String)
     mvarNomImpre = vData
End Property

Public Property Get NomImpresora() As String
     NomImpresora = mvarNomImpre
End Property


Public Property Get Contador() As Long
    Contador = mvarContador
End Property


Public Property Let HayVisor(ByVal vData As Boolean)
    mvarHayVisor = vData
End Property


Public Property Get HayVisor() As Boolean
    HayVisor = mvarHayVisor
End Property



Public Property Let NumPuerto(ByVal vData As Byte)
    mvarNumPuerto = vData
End Property


Public Property Get NumPuerto() As Byte
    NumPuerto = mvarNumPuerto
End Property



Public Property Let VelociPuerto(ByVal vData As String)
    mvarVelocPue = vData
End Property


Public Property Get VelociPuerto() As String
    VelociPuerto = mvarVelocPue
End Property



Public Property Let Rapida(ByVal vData As Boolean)
     mvarRapida = vData
End Property

Public Property Get Rapida() As Boolean
     Rapida = mvarRapida
End Property



Public Property Let FormaPagoUnica(ByVal vData As Boolean)
     mvarFormaPagoUnica = vData
End Property

Public Property Get FormaPagoUnica() As Boolean
     FormaPagoUnica = mvarFormaPagoUnica
End Property



Public Property Let ImprimiDirecto(ByVal vData As Boolean)
     mvarImprimiDirecto = vData
End Property

Public Property Get ImprimiDirecto() As Boolean
     ImprimiDirecto = mvarImprimiDirecto
End Property







Public Property Let DesglosaIVATicket(ByVal vData As Boolean)
     mvarDesglosaIVATicket = vData
End Property

Public Property Get DesglosaIVATicket() As Boolean
     DesglosaIVATicket = mvarDesglosaIVATicket
End Property



Public Property Let AbreCajon(ByVal vData As Boolean)
     mvarAbreCajon = vData
End Property

Public Property Get AbreCajon() As Boolean
     AbreCajon = mvarAbreCajon
End Property

Public Property Get SecuenciaCajon() As String
    SecuenciaCajon = mvarSecuenciaCajon
End Property


'mvarRedondea2
Public Property Let Redondea2(ByVal vData As Boolean)
     mvarRedondea2 = vData
End Property

Public Property Get Redondea2() As Boolean
     Redondea2 = mvarRedondea2
End Property



Public Property Let CalculaIVAsobrePVP(ByVal vData As Boolean)
     mvarCalculaIVAsobrePVP = vData
End Property

Public Property Get CalculaIVAsobrePVP() As Boolean
     CalculaIVAsobrePVP = mvarCalculaIVAsobrePVP
End Property



Public Property Let CtrstockVenta(ByVal vData As Boolean)
     mvarCtrstockVenta = vData
End Property

Public Property Get CtrstockVenta() As Boolean
     CtrstockVenta = mvarCtrstockVenta
End Property


'------------------------------------------------
'FUNCIONES del modulo CParamTPV
'------------------------------------------------

'****   LEER  ***
Public Function Leer() As Byte
'Lee los parametros generales
Dim cad As String
Dim RS As ADODB.Recordset
On Error GoTo Err3
    
    SQL = "SELECT * FROM spatpvg WHERE codigo=1"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic

    If RS.EOF Then
        Leer = 1
    Else
        'cliente por defecto
        mvarCodClien = RS!codClien
        SQL = "codforpa"
        mvarNomClien = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", RS!codClien, "N", SQL)
        mvarCtaBanc1 = DBLet(RS!ctabanc1, "T")
        
        mvarCodForpa = SQL 'forma de pago del cliente
        
        cad = "tipforpa"
        SQL = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", mvarCodForpa, "N", cad)
        mvarNomForPa = SQL
        mvarTipForPa = CByte(cad)
        
        mvarTipCierre = CByte(RS!tipcierre)
        
        mvarRapida = Val(DBLet(RS!rapido, "N")) = 1
        mvarFormaPagoUnica = False  'Por defecto lo dejo asi
        mvarImprimiDirecto = Val(DBLet(RS!imprtick, "N")) = 1
        
        
        
        mvarDesglosaIVATicket = Val(DBLet(RS!BasesImp, "N")) = 1

        
        mvarRedondea2 = (DBLet(RS!Redondea2, "N") = 1)
        
        
        mvarCalculaIVAsobrePVP = (DBLet(RS!IVAsobrePVP, "N") = 1)
        
        mvarCtrstockVenta = (DBLet(RS!CtrstockVenta, "N")) = 1
        
'        mvarModoCalculo = 0 'sin IVA
        
'        mvarHayVisor = CBool(RS!HayVisor)
'        mvarNumPuerto = DBLet(RS!NumPuerto, "N")
'        mvarVelocPue = DBLet(RS!velocpue, "T")
        Leer = 0
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Set RS = Nothing
    Leer = 1
End Function




Public Function Leer2(NT As String) As Byte
'Lee los parametros del terminal
Dim cad As String
Dim RS As ADODB.Recordset
    
    On Error GoTo Err3
    
    SQL = "SELECT * FROM spatpvt WHERE numtermi=" & NT
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic

    If RS.EOF Then
        Leer2 = 1
    Else
        mvarNomImpre = DBLet(RS!NomImpre, "T")
        mvarHayVisor = CBool(RS!HayVisor)
        mvarNumPuerto = DBLet(RS!NumPuerto, "N")
        mvarVelocPue = DBLet(RS!velocpue, "T")
        
        
        mvarAbreCajon = Val(DBLet(RS!AbreCajon, "N")) = 1
        mvarSecuenciaCajon = DBLet(RS!SecuenciaCajon)
        
        Leer2 = 0
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Set RS = Nothing
    Leer2 = 1
End Function

Public Function ConseguirContador(vCodigo As Byte) As Long
Dim Ok As Boolean
Dim C1 As Long 'contador
Dim RS As ADODB.Recordset

    On Error GoTo Err1
    
    ConseguirContador = 1
    
    'Abrimos bloqueando
    SQL = "Select * from spatpvt WHERE numtermi=" & vCodigo & " FOR UPDATE"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        mvarContador = RS!Contador
        C1 = RS!Contador
        Ok = True
    End If
    RS.Close
    Set RS = Nothing
    
    If Ok Then ConseguirContador = C1 + 1
    Exit Function
    
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
End Function


Public Function IncrementarContador(vCodigo As Byte) As Boolean
Dim C1 As Long 'contador1: recibo de entrada

    On Error GoTo Err1
    
      
    'Actualizamos el contador
    C1 = Contador
    SQL = "UPDATE spatpvt set "
    C1 = C1 + 1
    mvarContador = C1
    SQL = SQL & " contador=" & C1
    
    SQL = SQL & " WHERE numtermi=" & vCodigo
    conn.Execute SQL
    IncrementarContador = True
    Exit Function
    
Err1:
    If Err.Number <> 0 Then
        IncrementarContador = False
    Else
        IncrementarContador = True
    End If
End Function



Public Function DevolverContador(vCodigo As Byte, Contador As Long) As Byte
Dim Ok As Boolean
Dim C1 As Long
Dim RS As ADODB.Recordset

    On Error GoTo Err1
    
    'Abrimos bloqueando
    SQL = "Select * from spatpvt WHERE numtermi= " & vCodigo & " FOR UPDATE " ' "' FOR UPDATE"
    'Esto esba asi antes, sin comentar
'    Conn.Execute "Set autocommit = 0"
    DevolverContador = 1
    Ok = False
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        C1 = RS!Contador
        Ok = True
    End If
    RS.Close
    Set RS = Nothing
    If Not Ok Then
        conn.Execute "commit"
        Exit Function
    End If

    Ok = False

    Ok = (C1 = Contador)
    If Ok Then
        'Actualizamos el contador
        SQL = "UPDATE spatpvt set "

        C1 = C1 - 1
        SQL = SQL & " contador=" & C1
        SQL = SQL & " WHERE numtermi = " & vCodigo
        conn.Execute SQL
    End If

    'Desbloqueamos
    conn.Execute "commit"
    DevolverContador = 1
    
Err1:
    If Err.Number <> 0 Then
        Mens = "Error: " & Err.Number & " : " & Err.Description
        MsgBox Mens, vbExclamation
        DevolverContador = 0
    End If
    conn.Execute "Set autocommit = 1"
End Function
