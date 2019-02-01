VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Acceso via WS"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ejemplo para  consumir ws por GET
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Autor del proyecto: Yvan Acosta (YAcosta)
' Este proyecto no hubiera sido posible sin la colaboracion
' de mis amigos Leandro Ascierto y Alberto Miñano
' Visite: leandroascierto.com
Dim CadenaJSON       As String


Private Sub btnAceptar_Click()
Dim p             As Object
Dim Texto         As String
Dim sInputJson    As String
Dim cab           As Integer

Set httpURL = New WinHttp.WinHttpRequest

usua = Trim(txtUsuario)
Pass = Trim(txtPassword)

cadena = "http://queryfull.com/yacosta/prueba/login.php?USUARIO=" & usua & "&PASSWORD=" & Pass
'httpURL.Open "GET", Cadena
'httpURL.send
'
cadena = "http://api.icndb.com/jokes/count"
httpURL.Open "POST", cadena
httpURL.send

Texto = httpURL.responseText
If Texto = "[]" Then
   MsgBox ("No se obtuvo resultados")
   Exit Sub
End If

sInputJson = "{items:" & Texto & "}"

Set p = JSON.parse(sInputJson)

NOMBRE = p.Item("items").Item(1).Item("NOMBRE")

MsgBox ("Bienvenido " & NOMBRE)

End Sub

Private Sub Command1_Click()
Call CreandoJson("JUAN PEREZ", "JPEREZ")

Debug.Print CadenaJSON
End Sub

Public Sub CreandoJson(NOMBRE As String, USUARIO As String)
'Creando string
CadenaJSON = ""
Dim sBuffer As String

AddParamJSON CadenaJSON, "NOMBRE", NOMBRE
AddParamJSON CadenaJSON, "USUARIO", USUARIO, True
'enviando al ws
End Sub
