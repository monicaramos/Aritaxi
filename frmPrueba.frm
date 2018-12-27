VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmPrueba 
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2805
      Left            =   315
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   450
      Width           =   4920
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   465
      Left            =   3555
      TabIndex        =   0
      Top             =   3600
      Width           =   1635
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   225
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim Direccion As String

'        Direccion = "http://www.afilnet.com/http/sms/?email=" & Trim(vParamAplic.SMSemail) & "&pass=" & Trim(vParamAplic.SMSclave)
'        Direccion = Direccion & "&mobile=" & Trim(cad1) & "&id=" & Trim(vParamAplic.SMSremitente)
'        Direccion = Direccion & "&country=0034" & "&sms=" & txtcodigo(2).Text & "&now=" & Format(Check1.Value, "0")
'        Direccion = Direccion & "&date=" & Format(txtcodigo(0).Text, "yyyy/mm/dd") & " " & Format(txtcodigo(1).Text, "hh:mm")
'        Direccion = Direccion & "&type=" & Format(Check2.Value, "0")
        
        Screen.MousePointer = vbHourglass
        
        Direccion = Me.Text1.Text
        
       
        'Cargamos en el fichero el resultado de enviar un mensaje
        GetFileFromUrl Direccion, App.Path & "\RESULT.TXT"

        NF = FreeFile
        Open App.Path & "\RESULT.TXT" For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
        Cad = ""
        Line Input #NF, Cad
        Close NF

        Select Case Mid(Cad, 1, 2)
            Case "OK"
                Espera 2
            
                Me.Refresh
                DoEvents
                Espera 0.4
                cont = cont + 1
        
            Case "-1"
                MsgBox "Error en el Login, usuario o clave incorrectas", vbExclamation
                EstaOk = False
            Case Else
                If Mid(Cad, 1, 12) = "Sin Creditos" Then
                    MsgBox "No tiene créditos. Revise", vbExclamation
                    b = False
                Else
                    If MsgBox("Error en el envio de mensaje al socio " & ". ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then b = False
                End If
                EstaOk = False
        End Select
        
        Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Dim Direccion As String

    Direccion = "http://192.168.1.40:8089/version"
'    & Trim(vParamAplic.SMSemail) & "&pass=" & Trim(vParamAplic.SMSclave)
'    Direccion = Direccion & "&mobile=" & Trim(cad1) & "&id=" & Trim(vParamAplic.SMSremitente)
'    Direccion = Direccion & "&country=0034" & "&sms=" & txtcodigo(2).Text & "&now=" & Format(Check1.Value, "0")
'    Direccion = Direccion & "&date=" & Format(txtcodigo(0).Text, "yyyy/mm/dd") & " " & Format(txtcodigo(1).Text, "hh:mm")
'    Direccion = Direccion & "&type=" & Format(Check2.Value, "0")

    Text1.Text = Direccion


End Sub



Private Sub GetFileFromUrl(ByRef url As String, ByRef file As String)
    Dim fileBytes() As Byte
    Dim fileNum As Integer
    
    Dim Cad As String
    
    On Error GoTo DownloadError
    DoEvents
    
  '  fileBytes() = Inet1.OpenURL(url, icByteArray)
    
    Cad = Inet1.OpenURL(url)
    
'    fileNum = FreeFile
'    Open file For Binary Access Write As #fileNum
'    Put #fileNum, , fileBytes()
'    Close #fileNum
    
    Exit Sub
    
DownloadError:
    MsgBox Err.Description
End Sub

