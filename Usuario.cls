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
Private mvarcodigo As Long 'copia local
Private mvarNombre As String 'copia local
Private mvarPasswd As String 'copia local
Private mvarPasswdPROPIO As String 'copia local
Private mvarlogin As String 'copia local
Private mvarDirFich As String 'copia local
Private mvarNivel As Long 'copia local
Private mvarCadenaConexion As String
Private mvarPC As String

Private RS As ADODB.Recordset
Private Sql As String
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
    mvarcodigo = vData
End Property


Public Property Get Codigo() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Codigo
    Codigo = mvarcodigo
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

Public Property Let Id(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Codigo= 5
    mvarId = vData
End Property


Public Property Get Id() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Codigo
    Id = mvarId
End Property

Public Function Leer(ByVal vlogin As String) As Byte
Attribute Leer.VB_Description = "Lee un grupo de formas de cobro en función del código que se le ha pasado. Si devuelve 0 se ha conseguido leer el dato."
Dim I As Long
    On Error GoTo Err1
    Sql = "SELECT * FROM usuarios.usuarios WHERE login = '" & vlogin & "'"
    Sql = Sql & " AND NivelAritaxi >=0"
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If RS.EOF Then
        Leer = 1
    Else
        I = DevuelveAumentoPC
        mvarcodigo = Val(DBLet(RS!codusu)) + I
        mvarNombre = DBLet(RS!nomusu)
        mvarNivel = Val(DBLet(RS!nivelaritaxi))
        mvarPasswdPROPIO = RS!passwordpropio
        mvarlogin = vlogin
        mvarCadenaConexion = ""  'Se le asigna en tiempo de diseño
        mvarSkin = DBLet(RS!Skin)
        mvarId = Val(DBLet(RS!codusu))
        
        
        Leer = 0
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Leer = 1
End Function


Private Function DevuelveAumentoPC() As Long

DevuelveAumentoPC = 32000
'mvarPC = ComputerName
mvarPC = ComputerName
Sql = DevuelveDesdeBD(conAri, "codpc", "usuarios.pcs", "nompc", mvarPC, "T")
If Sql <> "" Then DevuelveAumentoPC = Val(Sql) * 1000

End Function


Public Sub LeerTabPorDefecto()
    'Las tabs estan en mysql,menus, desde el 1 hasta el 14
    If mvarTabPorDefecto = 0 Then mvarTabPorDefecto = 2
    
    Sql = " codigo<=14 and aplicacion = 'aritaxi' and filtro =1 and codusu "
    Sql = DevuelveDesdeBD(conAri, "codigo", "menus_usuarios", Sql, CStr(mvarId))
    If Sql <> "" Then
        If Val(Sql) > 0 Then mvarTabPorDefecto = CInt(Sql)
    End If
    
End Sub

Public Sub GuardarTabPorDefecto()
    
    'En todas las contabilidades
    Sql = "SELECT codempre conta FROM usuarios.empresasaritaxi "
    Sql = Sql & " where mid(aritaxi,1,7) = 'aritaxi'"
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        'SOLO en esta. Deberia dar error
        Sql = "UPDATE menus_usuarios SET filtro=if(codigo=" & vUsu.TabPorDefecto & ",1,0) WHERE codusu = " & vUsu.Codigo
        Sql = Sql & " AND aplicacion = 'aritaxi' AND codigo <=14"
        conn.Execute Sql
    Else
        While Not RS.EOF
            Sql = "UPDATE aritaxi" & RS!conta & ".menus_usuarios "
            Sql = Sql & " SET filtro=if(codigo=" & vUsu.TabPorDefecto & ",1,0) WHERE codusu = " & vUsu.Id
            Sql = Sql & " AND aplicacion = 'aritaxi' AND codigo <=14"
            conn.Execute Sql
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
End Sub



Public Sub ActualizarSkin()
    Sql = "UPDATE usuarios.usuarios SET skin=" & mvarSkin & " WHERE codusu = " & vUsu.Codigo
    
    EjecutarSQL Sql
End Sub



