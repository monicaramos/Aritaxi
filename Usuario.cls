VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Soporta entidades financieras usadas en los ingresos"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales que contienen valores de propiedad
Private mvarCodigo As Long 'copia local
Private mvarNombre As String 'copia local
Private mvarPasswd As String 'copia local
Private mvarPasswdPROPIO As String 'copia local
Private mvarlogin As String 'copia local
Private mvarDirFich As String 'copia local
Private mvarNivel As Long 'copia local
Private mvarCadenaConexion As String
Private mvarPC As String

Private Rs As ADODB.Recordset
Private SQL As String
Private Mens As String

Private mvarId As Long 'copia local

Private mvarSkin As String


Private mvarTabPorDefecto As Integer   'Guadara el TAB que tiene seleccionado por defecto


Public Property Let Nivel(ByVal vData As Long)
    mvarNivel = vData
End Property


Public Property Get Nivel() As Long
    Nivel = mvarNivel
End Property

Public Property Let Passwd(ByVal vData As String)
Attribute Passwd.VB_Description = "Password de usuario"
    mvarPasswd = vData
End Property


Public Property Get Passwd() As String
    Passwd = mvarPasswd
End Property



Public Property Let CadenaConexion(ByVal vData As String)
    mvarCadenaConexion = vData
End Property


Public Property Get CadenaConexion() As String
    CadenaConexion = mvarCadenaConexion
End Property



Public Property Let PasswdPROPIO(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Passwd = 5
    mvarPasswdPROPIO = vData
End Property


Public Property Get PasswdPROPIO() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Passwd
    PasswdPROPIO = mvarPasswdPROPIO
End Property


Public Property Let Nombre(ByVal vData As String)
Attribute Nombre.VB_Description = "Nombre del grupo de formas de cobro."
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarNombre = vData
End Property


Public Property Get Nombre() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    Nombre = mvarNombre
End Property

'---------------------------
'Dirfich
Public Property Let Dirfich(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarDirFich = vData
End Property


Public Property Get Dirfich() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    Dirfich = mvarDirFich
End Property

'Login
Public Property Let Login(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarlogin = vData
End Property


Public Property Get Login() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    Login = mvarlogin
End Property



Public Property Let Codigo(ByVal vData As Long)
Attribute Codigo.VB_Description = "Código del grupo de formas de cobro"
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Codigo= 5
    mvarCodigo = vData
End Property


Public Property Get Codigo() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Codigo
    Codigo = mvarCodigo
End Property

Public Property Get PC() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Codigo
    PC = mvarPC
End Property



Public Property Let TabPorDefecto(ByVal vData As Integer)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nivel = 5
    mvarTabPorDefecto = vData
End Property


Public Property Get TabPorDefecto() As Integer
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nivel
    TabPorDefecto = mvarTabPorDefecto
End Property

Public Property Let Skin(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarSkin = vData
End Property


Public Property Get Skin() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    Skin = mvarSkin
End Property

Public Property Let id(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Codigo= 5
    mvarId = vData
End Property


Public Property Get id() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Codigo
    id = mvarId
End Property

Public Function Leer(ByVal vlogin As String) As Byte
Attribute Leer.VB_Description = "Lee un grupo de formas de cobro en función del código que se le ha pasado. Si devuelve 0 se ha conseguido leer el dato."
Dim i As Long
    On Error GoTo Err1
    SQL = "SELECT * FROM usuarios.usuarios WHERE login = '" & vlogin & "'"
    SQL = SQL & " AND NivelAritaxi >=0"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Rs.EOF Then
        Leer = 1
    Else
        i = DevuelveAumentoPC
        mvarCodigo = Val(DBLet(Rs!CodUsu)) + i
        mvarNombre = DBLet(Rs!nomusu)
        mvarNivel = Val(DBLet(Rs!nivelaritaxi))
        mvarPasswdPROPIO = Rs!passwordpropio
        mvarlogin = vlogin
        mvarCadenaConexion = ""  'Se le asigna en tiempo de diseño
        mvarSkin = DBLet(Rs!Skin)
        mvarId = Val(DBLet(Rs!CodUsu))
        Leer = 0
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Leer = 1
End Function


Private Function DevuelveAumentoPC() As Long

DevuelveAumentoPC = 32000
'mvarPC = ComputerName
mvarPC = ComputerName
SQL = DevuelveDesdeBD(conAri, "codpc", "usuarios.pcs", "nompc", mvarPC, "T")
If SQL <> "" Then DevuelveAumentoPC = Val(SQL) * 1000

End Function


Public Sub LeerTabPorDefecto()
    'Las tabs estan en mysql,menus, desde el 1 hasta el 14
    If mvarTabPorDefecto = 0 Then mvarTabPorDefecto = 2
    
    SQL = " codigo<=14 and aplicacion = 'aritaxi' and filtro =1 and codusu "
    SQL = DevuelveDesdeBD(conAri, "codigo", "menus_usuarios", SQL, CStr(mvarId))
    If SQL <> "" Then
        If Val(SQL) > 0 Then mvarTabPorDefecto = CInt(SQL)
    End If
    
End Sub

Public Sub GuardarTabPorDefecto()
    
    'En todas las contabilidades
    SQL = "SELECT codempre conta FROM usuarios.empresasariconta "
    SQL = SQL & " where mid(conta,1,8) = 'aritaxi'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        'SOLO en esta. Deberia dar error
        SQL = "UPDATE menus_usuarios SET filtro=if(codigo=" & vUsu.TabPorDefecto & ",1,0) WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " AND aplicacion = 'aritaxi' AND codigo <=14"
        conn.Execute SQL
    Else
        While Not Rs.EOF
            SQL = "UPDATE aritaxi" & Rs!conta & ".menus_usuarios "
            SQL = SQL & " SET filtro=if(codigo=" & vUsu.TabPorDefecto & ",1,0) WHERE codusu = " & vUsu.Codigo
            SQL = SQL & " AND aplicacion = 'aritaxi' AND codigo <=14"
            conn.Execute SQL
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
End Sub



Public Sub ActualizarSkin()
    SQL = "UPDATE usuarios.usuarios SET skin=" & mvarSkin & " WHERE codusu = " & vUsu.Codigo
    
    EjecutarSQL SQL
End Sub



