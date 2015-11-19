VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TABLA: sclien
'   SOCIOS
'Variables locales que contienen valores de propiedad
Private mCodSocio As String 'Cod. socio
Private mNomSocio As String 'Nombre socio
Private mDomSocio As String 'Domicilio socio
Private mCPostalSocio As String 'cpostal socio
Private mPobSocio As String 'poblacion socio
Private mProSocio As String 'provincia socio
Private mNIFSocio As String 'nif socio
Private mTfnoSocio As String 'telefono socio
Private mObservac As String 'observaciones socio

Private mMatricul As String 'matricula
Private mMaiSocio As String 'Mail administracion

Private mCodForpa As String 'cod. forma de pago
Private mCtaSocioPub As String ' cuenta contable de Publicidad
Private mCtaSocioLiq As String ' cuenta contable de liquidacion
Private mCodIva As String

Private mContador As String
Private mLicencia As String

'[Monica]22/11/2013: tema iban
Private mIban As String ' iban

Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria

Private mNumerUve As String 'numero de uve

'------------------------------------------------
'------------------------------------------------
'Propiedades del modulo CSocio
'------------------------------------------------
'------------------------------------------------


'**** Codigo del socio

Public Property Let Codigo(ByVal vData As String)
     mCodSocio = vData
End Property

Public Property Get Codigo() As String
     Codigo = mCodSocio
End Property


'**** Nombre del socio

Public Property Let Nombre(ByVal vData As String)
     mNomSocio = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomSocio
End Property


'**** Domicilio del socio

Public Property Let Domicilio(ByVal vData As String)
     mDomSocio = vData
End Property

Public Property Get Domicilio() As String
     Domicilio = mDomSocio
End Property

'**** CPostal del socio

Public Property Let CPostal(ByVal vData As String)
     mCPostalSocio = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCPostalSocio
End Property


'**** Poblacion  del socio

Public Property Let Poblacion(ByVal vData As String)
     mPobSocio = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobSocio
End Property


'**** Provincia  del socio

Public Property Let Provincia(ByVal vData As String)
     mProSocio = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProSocio
End Property

'**** NIF  del socio

Public Property Let NIF(ByVal vData As String)
     mNIFSocio = vData
End Property

Public Property Get NIF() As String
     NIF = mNIFSocio
End Property


'**** Tfno  del socio

Public Property Let TfnoSocio(ByVal vData As String)
     mTfnoSocio = vData
End Property

Public Property Get TfnoSocio() As String
     TfnoSocio = mTfnoSocio
End Property


'**** Observaciones del socio

Public Property Let Observaciones(ByVal vData As String)
     mObservac = vData
End Property

Public Property Get Observaciones() As String
     Observaciones = mObservac
End Property



'**** Matricula del socio

Public Property Let Matricula(ByVal vData As String)
     mMatricul = vData
End Property

Public Property Get Matricula() As String
     Matricula = mMatricul
End Property


'**** Forma de pago del socio

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property


'**** Cuenta Contable socio Liquidacion

Public Property Let CtaSocioLiq(ByVal vData As String)
     mCtaSocioLiq = vData
End Property

Public Property Get CtaSocioLiq() As String
     CtaSocioLiq = mCtaSocioLiq
End Property


'**** Cuenta Contable socio Publicidad

Public Property Let CtaSocioPub(ByVal vData As String)
     mCtaSocioPub = vData
End Property

Public Property Get CtaSocioPub() As String
     CtaSocioPub = mCtaSocioPub
End Property




'**** Código de IVA

Public Property Let CodIva(ByVal vData As String)
     mCodIva = vData
End Property

Public Property Get CodIva() As String
     CodIva = mCodIva
End Property


'**** Contador de Factura

Public Property Let Contador(ByVal vData As String)
     mContador = vData
End Property

Public Property Get Contador() As String
     Contador = mContador
End Property

'**** licencia

Public Property Let Licencia(ByVal vData As String)
     mLicencia = vData
End Property

Public Property Get Licencia() As String
     Licencia = mLicencia
End Property


'[Monica]22/11/2013: tema Iban
'**** IBAN
Public Property Let Iban(ByVal vData As String)
     mIban = vData
End Property

Public Property Get Iban() As String
     Iban = mIban
End Property


'**** Banco del socio

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal del socio

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property

'**** Digito Control de la cuenta del socio
Public Property Let DigControl(ByVal vData As String)
     mDigContr = vData
End Property

Public Property Get DigControl() As String
     DigControl = mDigContr
End Property

'**** Cuenta Bancaria del socio

Public Property Let CuentaBan(ByVal vData As String)
     mCuentaba = vData
End Property

Public Property Get CuentaBan() As String
     CuentaBan = mCuentaba
End Property


'**** e-MAil administracion del Socio

Public Property Let EMailSocio(ByVal vData As String)
     mMaiSocio = vData
End Property

Public Property Get EMailSocio() As String
     EMailSocio = mMaiSocio
End Property


'**** Numero de uve del Socio

Public Property Let UveSocio(ByVal vData As String)
     mNumerUve = vData
End Property

Public Property Get UveSocio() As String
     UveSocio = mNumerUve
End Property


'------------------------------------------------
'------------------------------------------------
'Procedimientos del modulo Csocio
'------------------------------------------------
'------------------------------------------------

Public Function Existe(vCodsocio As String) As Boolean
'Comprueba si existe el socio en la BD
Dim devuelve As String

    On Error GoTo EExiste
    
    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codclien", "codclien", vCodsocio, "N")
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe el Socio " & vCodsocio & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function


Public Function LeerDatos(vCodsocio As String) As Boolean
'Leer los datos de un tranportista dado
'Lee de la BD: Ariges, Tabla: agencias
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim Mens As String
Dim Sql As String
Dim CADENA As String

    On Error GoTo ELeer
    LeerDatos = False
    
    Sql = "SELECT codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,codbanco,codsucur,digcontr,cuentaba, telclie1, observac, licencia, maiclie1, numeruve, iban "
    Sql = Sql & " FROM sclien "
    Sql = Sql & " WHERE codclien=" & DBSet(vCodsocio, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If RS.EOF Then
        LeerDatos = False
    Else
        Mens = "Datos Socio"
        mCodSocio = CStr(RS!CodClien)
        mNomSocio = RS!nomclien
        mDomSocio = CStr(RS!domclien)
        mCPostalSocio = CStr(RS!codpobla)
        mPobSocio = CStr(RS!pobclien)
        mProSocio = CStr(RS!proclien)
        mNIFSocio = CStr(RS!nifClien)
        
        '[Monica]22/11/2013: tema iban
        mIban = CStr(DBLet(RS!Iban, "T"))
        
        mCodbanco = CStr(DBLet(RS!codbanco, "N"))
        mCodsucur = CStr(DBLet(RS!codsucur, "N"))
        mDigContr = DBLet(RS!digcontr, "T")
        mCuentaba = DBLet(RS!cuentaba, "T")
        mLicencia = CStr(DBLet(RS!Licencia, "N"))
        mTfnoSocio = DBLet(RS!telclie1, "T")
        mObservac = DBLet(RS!observac, "T")
        mMaiSocio = DBLet(RS!maiclie1, "T")
        mNumerUve = DBLet(RS!NumerUve, "N")
        
        CADENA = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
        
        mCtaSocioLiq = vParamAplic.Raiz_Cta_Soc_Liqui & Format(mCodSocio, CADENA)
        mCtaSocioPub = vParamAplic.Raiz_Cta_Soc_publi & Format(mCodSocio, CADENA)
        
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function


Public Function ConseguirContador(vTipom As String) As Long
Dim Ok As Boolean
Dim C1 As Long 'contador
Dim RS As ADODB.Recordset
Dim Sql As String
Dim SQL2 As String
Dim Mens As String

On Error GoTo Err1
    
    ConseguirContador = 1
    
    Mens = ""
    'Abrimos bloqueando
    Sql = "Select contador from sclien_contadores WHERE codsocio=" & Trim(mCodSocio)
    Sql = Sql & " and codtipom = " & DBSet(vTipom, "T") & " FOR UPDATE"
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        mContador = RS!Contador
        C1 = RS!Contador
        Ok = True
    Else
        'MsgBox "No existe contador de tipo " & vTipom & " para ese socio.", vbExclamation
        Mens = "Insertando contador socio " & mCodSocio & " de tipo " & vTipom & "."
        
        SQL2 = "insert into sclien_contadores (codsocio, codtipom, contador) values ("
        SQL2 = SQL2 & DBSet(mCodSocio, "N") & "," & DBSet(vTipom, "T") & ",0) "
        
        conn.Execute SQL2
        
        mContador = 0
        C1 = 0
        
        Ok = True
    End If
    RS.Close
    Set RS = Nothing
     
    If Ok Then ConseguirContador = C1 + 1
    Exit Function
Err1:
    Mens = Mens & "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    ConseguirContador = -1
End Function


Public Function IncrementarContador(vTipom As String) As Boolean
Dim C1 As Long 'contador1: recibo de entrada
Dim Sql As String
On Error GoTo Err1
    
      
    'Actualizamos el contador
    C1 = Contador
    Sql = "UPDATE sclien_contadores set "
    C1 = C1 + 1
    mContador = C1
    Sql = Sql & " contador=" & C1
    
    Sql = Sql & " WHERE codsocio = " & mCodSocio & " and codtipom = " & DBSet(vTipom, "T")
    conn.Execute Sql
    IncrementarContador = True
    Exit Function
Err1:
    If Err.Number <> 0 Then
        IncrementarContador = False
    Else
        IncrementarContador = True
    End If
End Function

Public Function ActualizarContador(vTipom As String) As Boolean
Dim C1 As Long 'contador1: recibo de entrada
Dim Sql As String
On Error GoTo Err1
    
      
    'Actualizamos el contador
    C1 = Contador
    Sql = "UPDATE sclien_contadores set "
    mContador = C1
    Sql = Sql & " contador=" & C1
    
    Sql = Sql & " WHERE codsocio = " & mCodSocio & " and codtipom = " & DBSet(vTipom, "T")
    conn.Execute Sql
    ActualizarContador = True
    Exit Function
Err1:
    If Err.Number <> 0 Then
        ActualizarContador = False
    Else
        ActualizarContador = True
    End If
End Function

Public Function DevolverContador(vCodigo As String, Contador As Long, vTipom As String) As Byte
Dim Ok As Boolean
Dim C1 As Long
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Mens As String

On Error GoTo Err1
    
    'Abrimos bloqueando
    Sql = "Select contador from sclien_contadores WHERE codsocio = " & vCodigo & " and codtipom = " & DBSet(vTipom, "T") & " FOR UPDATE " ' "' FOR UPDATE"
    'Esto esba asi antes, sin comentar
    'Conn.Execute "Set autocommit = 0"
    DevolverContador = 1
    Ok = False
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
        Sql = "UPDATE sclien_contadores set "

        C1 = C1 - 1
        Sql = Sql & " contador=" & C1
        Sql = Sql & " WHERE codsocio = " & vCodigo & " and codtipom = " & DBSet(vTipom, "T")
        conn.Execute Sql
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

