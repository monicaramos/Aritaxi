VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "POST"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15720
   LinkTopic       =   "Form2"
   ScaleHeight     =   5880
   ScaleWidth      =   15720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Json"
      Height          =   375
      Left            =   14040
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   9120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1680
      Width           =   6255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "POST"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "{""userId"":13,""email"":""sm232Xxx3e@prueba.com"",""password"":""""}"
      Top             =   1080
      Width           =   6975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1680
      Width           =   8535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "http://192.168.1.40:8089/users"
      Top             =   360
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "GET"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "POST"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim vJson As String

If Text1.Text = "" Then Exit Sub

Set httpURL = New WinHttp.WinHttpRequest


'cadena = "http://api.icndb.com/jokes/count"
 cadena = Text1.Text
 Text2.Text = "Enviando: " & cadena
 
 vJson = ""
 
 If Me.Check1.Value = 0 Then
    httpURL.Open "GET", cadena
Else
    'vJson = "{""userId"":13,""email"":""u131313@prueba.com"",""password"":""13""}"

    


    
    httpURL.Open "POST", cadena, False
    vJson = Text3.Text
    httpURL.setRequestHeader "Content-Type", "application/json"
    

    
    
End If
httpURL.send vJson

Texto = httpURL.responseText

Text2.Text = Texto
    
End Sub

Private Sub Command2_Click()
Dim i As Integer
Dim J As Integer

    Dim D As Dictionary
    Dim p As Object
    Dim P1 As Object
    Dim key As Variant
    Set p = JSON.parse(Text2.Text)
   
    If p Is Nothing Then
        Text4.Text = "No es formato JSON"
        
    Else
   
         If p.Count = 0 Then
            Text4.Text = "Ningun dato devuelto"
            
        Else
            'Encabezado
            

            
            For i = 1 To p.Count
                Set P1 = p.Item(i)
                For J = 1 To P1.Count
                    Debug.Print "userid"
            
                Next J
            Next
    
        End If
    End If
End Sub

Private Sub Form_Load()
    
    'Text1.Text ="http://api.icndb.com/jokes/1"  ""
End Sub
