VERSION 5.00
Begin VB.Form frmFacClienteObser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmFacClienteObser.frx":0000
      Top             =   120
      Width           =   7935
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "frmFacClienteObser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Modificar As Boolean

Private Sub Command1_Click(Index As Integer)
    If Index = 1 Then
        CadenaDesdeOtroForm = "0|"
    Else
        'Desde el 3 en adelante
        CadenaDesdeOtroForm = "1|" & Text1.Text
    End If
    Unload Me
End Sub


Private Sub Form_Load()
    Me.Caption = "Observaciones"
    Me.Icon = frmPpal.Icon
    Text1.Locked = Not Modificar
    Me.Command1(0).Enabled = Modificar
    Screen.MousePointer = vbDefault
End Sub

'Private Sub Command1_GotFocus(Index As Integer)
'    If True Then
'        MsgBox "hola", vbExclamation
'    End If
'End Sub
'
'Private Sub Text1_Change()
'    If True Then
'        MsgBox "hola", vbExclamation
'    End If
'End Sub
'
'Private Sub Text1_Click()
'    If True Then
'        MsgBox "hola", vbExclamation
'    End If
'End Sub
'
'Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If True Then
'        MsgBox "hola", vbExclamation
'    End If
'End Sub
'
'Private Sub Text1_KeyPress(KeyAscii As Integer)
'    If True Then
''        MsgBox "hola", vbExclamation
'    End If
'End Sub
'
'Private Sub Text1_LostFocus()
'    If True Then
''        MsgBox "hola", vbExclamation
'    End If
'End Sub
'
'Private Sub Text1_Validate(Cancel As Boolean)
'    If True Then
''        MsgBox "hola", vbExclamation
'    End If
'
'End Sub
