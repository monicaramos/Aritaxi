VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación de datos"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
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
      Left            =   5520
      TabIndex        =   3
      Top             =   930
      Width           =   1135
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   6810
      TabIndex        =   2
      Top             =   930
      Width           =   1135
   End
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
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgFich 
      Height          =   240
      Index           =   0
      Left            =   2190
      Picture         =   "frmImportar.frx":0000
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero a importar:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Dim NomFich As String
    Dim NF As Integer
    Dim linea As String
    Dim Campos(20) As String
    Dim Mc As String
    Dim campo As String
    Dim I As Integer
    Dim i2 As Integer
    '-- CAMPOS
    Dim CodClien As Long
    Dim codusuar As Long
    Dim numlinea As Long
    Dim NUMTARJE As String
    Dim TEXTOA1 As String
    Dim TEXTOA2 As String
    Dim TEXTOA3 As String
    Dim FECHAEMI As Date
    Dim FECHACAD As Date
    Dim PISTAGR1 As String
    Dim PISTAGR2 As String
    Dim PISTAGR3 As String
    Dim NOMFICHE As String
    Dim NOMUSUAR As String
    Dim Sql As String
    Dim RS As ADODB.Recordset
    If Not DatosOk Then Exit Sub
    NomFich = Text1
    NF = FreeFile
    i2 = 0
    Open NomFich For Input As #NF
    Do While Not EOF(NF)
        i2 = 0
        Line Input #NF, linea
        For I = 1 To Len(linea)
            Mc = Mid(linea, I, 1)
            If Mc = ";" Then
                i2 = i2 + 1
                
                campo = Replace(campo, """", "")
                
                Campos(i2) = campo
                campo = ""
            Else
                campo = campo & Mc
            End If
        Next I
        i2 = i2 + 1
        Campos(i2) = Replace(campo, """", "")
        campo = ""
        '-- Asignar los campos?
        I = InStr(1, Campos(1), ".")
        CodClien = Val(Mid(Campos(1), 1, I - 1))
        codusuar = Val(Mid(Campos(1), I + 1, Len(Campos(1)) - I))
        FECHAEMI = Date
        FECHACAD = CDate("01/" & Mid(Campos(2), 1, 3) & Mid(Campos(2), 4, 2))
        NOMUSUAR = Campos(4)
        TEXTOA1 = Campos(3)
        TEXTOA2 = Campos(4)
        PISTAGR2 = Campos(6)
        I = InStr(1, Campos(6), "=")
        NUMTARJE = Val(Mid(Campos(6), 1, I - 1))
        NOMFICHE = App.Path & "\Informes\" & vConfig.Formato
        
        Sql = "select * from scliente where codclien = " & CodClien
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, , , adCmdText
        If RS.EOF Then
            MsgBox "El cliente " & CStr(CodClien) & " no existe, se cancela la importación.", vbCritical
            Exit Sub
        End If
        Set RS = Nothing
        
        Sql = "select * from scatar where codclien = " & CodClien & " and codusuar = " & codusuar
        
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
        If Not RS.EOF Then
            RS!NOMUSUAR = NOMUSUAR
            
            conn.Execute "Update scatar set nomusuar = " & DBSet(NOMUSUAR, "T") & " where codclien = " & CodClien & " and codusuar = " & codusuar
        Else
            'RS.AddNew
'            RS!CodClien = CodClien
'            RS!codusuar = codusuar
'            RS!NOMUSUAR = NOMUSUAR
            
            conn.Execute "insert into scatar (codclien,codusuar,nomusuar) values (" & DBSet(CodClien, "N") & "," & DBSet(codusuar, "N") & "," & DBSet(NOMUSUAR, "T") & ")"
            
        End If
'        Rs.Update
        Sql = "select * from slitar where numtarje ='" & NUMTARJE & "'"
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
        If RS.EOF Then
'            Rs.AddNew
'            Rs!codclien = codclien
'            Rs!codusuar = codusuar
'            Rs!numlinea = SugerirCodigoSiguiente(codclien, codusuar)
'            Rs!NUMTARJE = NUMTARJE

            Sql = "insert into slitar (codclien,codusuar,numlinea,numtarje,textoa1,textoa2,textoa3,"
            Sql = Sql & "fechaemi,fechacad,pistagr1,pistagr2,pistagr3,nomfiche) values ("
            Sql = Sql & DBSet(CodClien, "N") & "," & DBSet(codusuar, "N") & ","
            Sql = Sql & SugerirCodigoSiguiente(CodClien, codusuar) & ","
            Sql = Sql & DBSet(NUMTARJE, "T") & ","
            Sql = Sql & DBSet(TEXTOA1, "T") & "," & DBSet(TEXTOA2, "T") & "," & DBSet(TEXTOA3, "T") & ","
            Sql = Sql & DBSet(FECHAEMI, "F") & ","
            Sql = Sql & DBSet(FECHACAD, "F") & ","
            Sql = Sql & DBSet(PISTAGR1, "T") & "," & DBSet(PISTAGR2, "T") & "," & DBSet(PISTAGR3, "T") & ","
            Sql = Sql & DBSet(NOMFICHE, "T") & ")"
            
            conn.Execute Sql
            
        Else
            Sql = "update slitar set textoa1 = " & DBSet(TEXTOA1, "T")
            Sql = Sql & ",textoa2 = " & DBSet(TEXTOA2, "T")
            Sql = Sql & ",textoa3 = " & DBSet(TEXTOA3, "T")
            Sql = Sql & ",fechaemi = " & DBSet(FECHAEMI, "F")
            Sql = Sql & ",fechacad = " & DBSet(FECHACAD, "F")
            Sql = Sql & ",pistagr1 = " & DBSet(PISTAGR1, "T")
            Sql = Sql & ",pistagr2 = " & DBSet(PISTAGR2, "T")
            Sql = Sql & ",pistagr3 = " & DBSet(PISTAGR3, "T")
            Sql = Sql & ",nomfiche = " & DBSet(NOMFICHE, "T")
            Sql = Sql & "where numtarje ='" & NUMTARJE & "'"

            conn.Execute Sql

        End If
'        Rs!TEXTOA1 = TEXTOA1
'        Rs!TEXTOA2 = TEXTOA2
'        Rs!TEXTOA3 = TEXTOA3
'        Rs!FECHAEMI = FECHAEMI
'        Rs!FECHACAD = FECHACAD
'        Rs!PISTAGR1 = PISTAGR1
'        Rs!PISTAGR2 = PISTAGR2
'        Rs!PISTAGR3 = PISTAGR3
'        Rs!NOMFICHE = NOMFICHE
'        Rs.Update
        RS.Close
    Loop
    Close #NF
    MsgBox "Importación finalizada", vbInformation
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
End Sub

Private Sub imgFich_Click(Index As Integer)
    CommonDialog1.InitDir = "A:"
    CommonDialog1.Filter = "*.txt|*.*"
    CommonDialog1.ShowOpen
    Text1 = CommonDialog1.FileName
End Sub
Private Function DatosOk()
    If Text1 = "" Then
        MsgBox "Introduzca la ruta del fichero a importar", vbInformation
        Exit Function
    End If
    DatosOk = True
End Function

Private Function SugerirCodigoSiguiente(CodClien As Long, codusuar As Long) As Long

    Dim Sql As String
    Dim RS As ADODB.Recordset

    Sql = "Select Max(numlinea) from slitar"
    Sql = Sql & " where codclien = " & CStr(CodClien)
    Sql = Sql & " and codusuar = " & CStr(codusuar)
    
    SugerirCodigoSiguiente = 1
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, , , adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            SugerirCodigoSiguiente = RS.Fields(0) + 1
        End If
    End If
    RS.Close
End Function

