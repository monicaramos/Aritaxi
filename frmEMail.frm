VERSION 5.00
Begin VB.Form frmEMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviar E-MAIL"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmEMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCartaConf 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   3120
      TabIndex        =   33
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame FrameASuntoMsg 
      Height          =   4815
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   6015
      Begin VB.TextBox Text1 
         Height          =   3135
         Index           =   5
         Left            =   1140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Text            =   "frmEMail.frx":0442
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   1140
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   31
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Width           =   555
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6075
      Begin VB.Frame Frame3 
         Caption         =   "E-mail"
         ForeColor       =   &H00972E0B&
         Height          =   1040
         Left            =   3360
         TabIndex        =   23
         Top             =   80
         Width           =   2175
         Begin VB.OptionButton OptMail 
            Caption         =   "Administraci�n"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptMail 
            Caption         =   "Comercial/Compras"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Para"
         ForeColor       =   &H00972E0B&
         Height          =   1040
         Left            =   1080
         TabIndex        =   20
         Top             =   80
         Width           =   2055
         Begin VB.OptionButton OptPara 
            Caption         =   "Trabajadores"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   26
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton OptPara 
            Caption         =   "Proveedores"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton OptPara 
            Caption         =   "Clientes"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1620
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   1080
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   2055
         Index           =   3
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmEMail.frx":0448
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Para"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   10
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   9
         Top             =   2100
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   720
         Picture         =   "frmEMail.frx":044E
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   8
         Top             =   2520
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5715
      Begin VB.TextBox Text3 
         Height          =   1695
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmEMail.frx":0550
         Top             =   1800
         Width           =   5355
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Otro"
         Height          =   255
         Index           =   2
         Left            =   2460
         TabIndex        =   15
         Top             =   1140
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Error"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   1140
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sugerencia"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   1140
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Mensaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   1500
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Asunto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enviar e-Mail Ariadna Software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   4305
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "frmEMail.frx":0556
      Top             =   4860
      Width           =   480
   End
End
Attribute VB_Name = "frmEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte
    '0 - Envio del PDF
    '1 - Envio Mail desde menu soporte
    '2 - Envio Mail masivo a varios clientes o proveedores
    '3 - Mail a un destinatario que se detalla en datosenvio
    '       y ademas envia copia a remitente
        
    '4 - haremos un multienvo. Es el envio de facturas por mail
    '5 -  Multienvio cartas renovacion
    '6 -  Confirmacion de pedido
    
Public DatosEnvio As String
    'Nombre para|email para|Asunto|Mensaje|    y para envio tipo3 el mail de otro persona mail|nombre|

'Private WithEvents frmC As frmFacClientes
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1


Dim cad As String
Dim PrimeraVez As Boolean




'Modificacion MUY IMPORTANTE
'Programaita de envio: arigesmail.exe
'Si la opcion es esa hara OOOtras cosas, si no lo dejamos como esta
Private Sub EnvioNuevo(ListaArch As Collection)

    If vParamAplic.ExeEnvioMail <> "" Then
        'Utliza el programa que lanza desde el outlook
        EnvioDesdeExeNuestro ListaArch
        If Opcion = 0 And DatosEnvio <> "" Then Me.DatosEnvio = "OK"
    Else
    
    
        'El que habia
        Enviar2 ListaArch
    End If

End Sub



'Modificacion: 10 Abril 2007
' Enviar siempre envia el documento llamado docum.pdf
' Ahora necesito enviar varios documentos por mail
' Para ello mandare si en la lista hay algo
' seran los path de los archivos, si no sera docum.pdf
Private Sub Enviar2(ListaArchivos As Collection)
    Dim success
    Dim mailman As ChilkatMailMan
    Dim Valores As String
    Dim J As Integer
    
    On Error GoTo GotException
    Set mailman = New ChilkatMailMan
    
    'Esta cadena es constante de la lincencia comprada a CHILKAT
    mailman.UnlockComponent "1AriadnaMAIL_BOVuuRWYpC9f"
'    mailman.LogMailSentFilename = App.Path & "\mailSent.log"
    
    
    'Servidor smtp
    If vParamAplic.EnvioDesdeOutlook Then
        Valores = "||||"
    Else
        Valores = ObtenerValoresEnvioMail  'Empipado: smtphost,smtpuser, pass, diremail
        If Valores = "" Then
            MsgBox "Falta configurar en paremtros la opcion de envio mail(servidor, usuario, clave)"
            Exit Sub
        End If
        mailman.SMTPhost = RecuperaValor(Valores, 1) ' vParam.SmtpHOST
        mailman.SmtpUsername = RecuperaValor(Valores, 2) 'vParam.SmtpUser
        mailman.SmtpPassword = RecuperaValor(Valores, 3) 'vParam.SmtpPass
        
        'David 2 Mayo 2007
        mailman.SmtpAuthMethod = "LOGIN"
        
    End If
    
    ' Create the email, add content, address it, and sent it.
    Dim email As ChilkatEmail
    Set email = New ChilkatEmail
    
    'Si es de SOPORTE
    
    
    If Opcion = 1 Then
         'Obtenemos la pagina web de los parametros
        '====David
'        Cad = DevuelveDesdeBD("mailsoporte", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
        '====
        cad = DevuelveDesdeBDNew(conAri, "sparam", "maiempre", "codempre", 1, "N")
        If cad = "" Then
            MsgBox "Falta configurar en parametros el mail de soporte", vbExclamation
            Exit Sub
        End If
    
        If cad = "" Then GoTo GotException
        email.AddTo "Soporte Gesti�n", cad
        cad = "Soporte AriTaxi. "
        If Option1(0).Value Then cad = cad & Option1(0).Caption
        If Option1(1).Value Then cad = cad & Option1(1).Caption
        If Option1(2).Value Then cad = cad & "Otro: " & Text2.Text
        email.Subject = cad
        
        'Ahora en text1(3).text generaremos nuestro mensaje
        cad = "Fecha: " & Format(Now, "dd/mm/yyyy") & vbCrLf
        cad = cad & "Hora: " & Format(Now, "hh:mm") & vbCrLf
        cad = cad & "Usuario: " & vUsu.Nombre & vbCrLf
        cad = cad & "Nivel USU: " & vUsu.Nivel & vbCrLf
        cad = cad & "Empresa: " & vEmpresa.nomempre & vbCrLf
        cad = cad & "&nbsp;<hr>"
        cad = cad & Text3.Text & vbCrLf & vbCrLf
        Text1(3).Text = cad
    Else
        'Opcion=0 or opcion= 3 or envio=4
        'Envio de mensajes normal
        ' ---- [04/11/2009] [LAURA] : concatenar al final del asunto [ARI] para poder crear regla correo
        
        
        If Opcion <> 6 Then
            email.Subject = Text1(2).Text & " [ARI]"
        Else
            email.Subject = Text1(2).Text
        End If
        ' ----
        email.AddTo Text1(0).Text, Text1(1).Text
        
        '### A�ade: Laura 11/10/05
        '### Modifica david.     Lo que hare sera para c
        If Opcion < 4 Then
            cad = RecuperaValor(Valores, 4)
            email.AddBcc RecuperaValor(Valores, 2), cad    'vParam.SmtpPass
            
        Else
            'Para el multienvio de facturacion y renovacion
            cad = RecuperaValor(DatosEnvio, 3)
            If cad = "1" Then
                cad = RecuperaValor(Valores, 4)
                email.AddBcc RecuperaValor(Valores, 2), cad    'vParam.SmtpPass
            End If
        End If
        'Si la opcion es 3   Envio del mail con tooodos los datos en datosenvio
        If Opcion = 3 Then
            CadenaDesdeOtroForm = RecuperaValor(DatosEnvio, 5)
            If CadenaDesdeOtroForm <> "" Then
                If CadenaDesdeOtroForm <> cad Then
                    'El usuario con el que envia el mail NO es el usuario que le indico con el datosenvio
                    'Por lo cual lo a�ado
                    cad = RecuperaValor(DatosEnvio, 6)
                    email.AddBcc "Aviso tomado", CadenaDesdeOtroForm
                End If
            End If
        End If
    End If
    
    'El resto lo hacemos comun
    'La imagen
    'imageContentID = email.AddRelatedContent(App.Path & "\minilogo.bmp")
    
    
    cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
    cad = cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
    cad = cad & "<TABLE BORDER=""0"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
    'Cuerpo del mensaje
    cad = cad & "<TR><TD VALIGN=""TOP""><P>"
    FijarTextoMensaje
    cad = cad & "</P></TD></TR>"
    cad = cad & "<TR><TD VALIGN=""TOP""><P><hr></P>"
    cad = cad & "<FONT SIZE=2>"
    cad = cad & "<P><P><P><P align=""justify"">Este correo electr�nico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a "
    cad = cad & " los destinatarios especificados. La informaci�n contenida puesde ser CONFIDENCIAL"
    cad = cad & " y/o estar LEGALMENTE PROTEGIDA.</P>"
    cad = cad & "<P align=""justify"">Si usted recibe este mensaje por ERROR, por favor comun�queselo inmediatamente al"
    
    cad = cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelaci�n, distribuci�n"
    cad = cad & " impresi�n o copia de toda o alguna parte de la informaci�n contenida, Gracias "
    cad = cad & ".</FONT></P><P><HR ALIGN=""LEFT"" SIZE=1></TD>"
    cad = cad & "</TR></TABLE></BODY></HTML>"
    
    email.SetHtmlBody (cad)
    
    'Texto alternativo
    cad = ""
    cad = cad & "Este correo electronico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a " & vbCrLf
    cad = cad & " los destinatarios especificados. La informacion contenida puesde ser CONFIDENCIAL" & vbCrLf
    cad = cad & " y/o estar LEGALMENTE PROTEGIDA." & vbCrLf & vbCrLf
    cad = cad & "Si usted recibe este mensaje por ERROR, por favor comuniqueselo inmediatamente al" & vbCrLf
    cad = cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelacion, distribucion" & vbCrLf
    cad = cad & " impresion o copia de toda o alguna parte de la informacion contenida, Gracias " & vbCrLf

    
    'Por si no acepta HTML
    cad = UCase(cad)
    email.AddPlainTextAlternativeBody Text1(3).Text & vbCrLf & vbCrLf & vbCrLf & cad
    email.From = RecuperaValor(Valores, 4) 'vParam.diremail
    
    
    If Opcion <> 1 Then   'Solo la opcion 1 NO lleva attachment
        'ADjunatmos el PDF
        If ListaArchivos Is Nothing Then
            email.AddFileAttachment App.Path & "\docum.pdf"
        Else
            
            For J = 1 To ListaArchivos.Count
                   email.AddFileAttachment ListaArchivos.item(J)
            Next J
        End If
    End If
        
    
    'email.SendEncrypted = 1
    
        'sI ENVIA POR OUTLOOK O NO
     If vParamAplic.EnvioDesdeOutlook Then
        'Si envia por outlook
         mailman.SendViaOutlook email
         success = 1
        
    Else
        success = mailman.SendEmail(email)
    End If
    If (success = 1) Then
        If Opcion <> 2 And Opcion <> 4 And Opcion <> 6 Then
            If vParamAplic.EnvioDesdeOutlook Then
                cad = "Enviado al outlook"
            Else
                cad = "Mensaje enviado correctamente."
            End If
            MsgBox cad, vbInformation
            Command2(0).SetFocus
        End If
        
        ' ---- [04/11/2009] [LAURA] : para saber q se ha enviado con exito y actualizar check de enviado
        If Opcion = 0 And DatosEnvio <> "" Then
            Me.DatosEnvio = "OK"
            Command2_Click (0)
        End If
        ' ---
    Else
        cad = "Han ocurrido errores durante el envio.Compruebe el archivo log.xml para mas informacion"
        mailman.SaveXmlLog App.Path & "\log.xml"
        MsgBox cad, vbExclamation
    End If
    
    
GotException:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set email = Nothing
    Set mailman = Nothing

End Sub

Private Sub cmdCartaConf_Click()
    'Carta confirmacion
    If Text1(4).Text = "" Then
        MsgBox "ponga el asunto", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    HacerMultiEnvioConfirPed
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Command1_Click()
Dim Col As Collection

    Screen.MousePointer = vbHourglass
    Image2.visible = True
    Me.Refresh
    
    
    'Opcion cero. Confirmacion entrega pedido
    If Opcion = 0 Then
        cad = RecuperaValor(Me.DatosEnvio, 5)
        If cad <> "" Then
            Set Col = New Collection
            Col.Add cad
        End If
    
    End If
                  
    EnvioNuevo Col
    Image2.visible = False
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Activate()
     If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 2 Or Opcion = 4 Or Opcion = 5 Then
            If Opcion = 2 Then
                HacerMultiEnvio
            Else
                'Opcion 4 y 5
                    Me.Command1.visible = False
                    Command2(0).visible = False
                    DoEvents
                    HacerMultiEnvioFacturacion
                    
                    
                    
                    Me.Command1.visible = True
                    Command2(0).visible = True
                    DoEvents


            End If
            Unload Me
            
            
        ' ---- [04/11/2009] [LAURA] : A�adir bot�n para enviar informe confirmacion entrega del Pedido
        ' ----                        para ello aqui a�ado opcion=0
        ElseIf (Opcion = 3) Or (Opcion = 0) Then
            If DatosEnvio <> "" Then
                'Fuerzo el envio de mail
    
                Text1(0).Text = RecuperaValor(DatosEnvio, 1)
                Text1(1).Text = RecuperaValor(DatosEnvio, 2)
                Text1(2).Text = RecuperaValor(DatosEnvio, 3)
                Text1(3).Text = RecuperaValor(DatosEnvio, 4)
                Me.Refresh
                DoEvents
                
                If Opcion = 3 Then
                    Command1_Click
                    Unload Me
                End If
            End If
        End If
        ' ----
    End If
End Sub


Private Sub Form_Load()
    PrimeraVez = True
    Image2.visible = False
    limpiar Me
    Frame1(0).visible = (Opcion = 0) Or (Opcion = 2)
    Frame1(1).visible = (Opcion = 1)
    FrameASuntoMsg.visible = (Opcion = 6)
    If Opcion = 1 Then HabilitarText

'    cad = DevuelveDesdeBD("smtpHost", "spara1", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
    cad = ""
    If vParamAplic.ExeEnvioMail <> "" Then
        cad = "OK"
    Else
        If vParamAplic.EnvioDesdeOutlook Then
            cad = "OK"
        Else
            cad = DevuelveDesdeBDNew(conAri, "spara1", "smtphost", "codigo", "1", "N")
        End If
    End If
    Me.Command1.Enabled = (cad <> "")
    
    
    Label13(0).Caption = ""
    If Opcion = 6 Then Label13(0).Caption = " Cartas Confirmaci�n de Pedidos"
    Me.cmdCartaConf.Enabled = Me.Command1.Enabled
    
    Me.cmdCartaConf.visible = Opcion = 6
    Me.Command1.visible = Opcion <> 6
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Opcion = 0
    
    ' ---- [04/11/2009] [LAURA]
'    DatosEnvio = ""
    If DatosEnvio <> "OK" Then DatosEnvio = ""
    ' ----
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)

'    Screen.MousePointer = vbHourglass
'    Text1(0).Tag = RecuperaValor(CadenaSeleccion, 1)
'    Text1(0).Text = RecuperaValor(CadenaSeleccion, 3)
'    'Si regresa con datos tengo k devolveer desde la bd el campo e-mail
'    Text1(1).Text = RecuperaValor(CadenaSeleccion, 4)
''    cad = DevuelveDesdeBDNew(conAri, "sclien", "maiclie1", "codclien", Text1(0).Tag, "T")
''    Text1(1).Text = cad
'    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    If CadenaDevuelta <> "" Then
'        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If Me.OptPara(0).Value Or Me.OptPara(1).Value Then
            'Clientes / proveedores
            Text1(0).Text = RecuperaValor(CadenaDevuelta, 3)
            Text1(1).Text = RecuperaValor(CadenaDevuelta, 4)
        Else
            'Trabajadores
            Text1(0).Text = RecuperaValor(CadenaDevuelta, 2)
            Text1(1).Text = RecuperaValor(CadenaDevuelta, 3)
        End If
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub Image1_Click()
    'busqueda de clientes que tiene e-mail
Dim cadSel As String
Dim cadTabla As String
Dim cadCampo As String

    'seleccionar de que tabla vamos a leer los datos
    If Me.OptPara(0).Value Then
        'leer datos de clientes
        cadTabla = "sclien"
        
        'seleccionar a que e-mail vamos a enviar
        If Me.OptMail(0).Value Then 'mail de Administracion
            'seleccionar el mail1
            cadCampo = "maiclie1"
        Else 'mail comercial
            cadCampo = "maiclie2 "
        End If
        
    ElseIf Me.OptPara(1).Value Then
        'datos de proveedores
        cadTabla = "sprove"
        
        'seleccionar a que e-mail vamos a enviar
        'seleccionar solo los proveedores que tiene valor en mail1 o mail2.
        If Me.OptMail(0).Value Then
            cadCampo = "maiprov1"
        Else
            cadCampo = "maiprov2"
        End If
    Else
        'de trabajadores
        cadTabla = "straba"
        cadCampo = "maitraba"
    End If

    cadSel = " (not isnull(" & cadCampo & ") and " & cadCampo & "<>'') "
    MandaBusquedaPrevia cadSel, cadTabla, cadCampo
    

'    Set frmC = New frmFacClientes
'    frmC.DatosADevolverBusqueda = "0|1"
''    frmC.ConfigurarBalances = 5  'NUEVO opcion
'    frmC.Show vbModal
'    Set frmC = Nothing
'    If Text1(0).Text <> "" Then PonerFoco Text1(2)
End Sub





Private Sub Option1_Click(Index As Integer)
    HabilitarText
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Function DatosOk() As Boolean
Dim i As Integer

    DatosOk = False
    'If Opcion = 0 Or Opcion = 3 Or Opcion = 4 Or Opcion = 5 Then
    If Opcion <> 1 And Opcion <> 2 Then
                'Pocas cosas a comprobar
                For i = 0 To 2
                    Text1(i).Text = Trim(Text1(i).Text)
                    If Text1(i).Text = "" Then
                        MsgBox "El campo: " & Label1(i).Caption & " no puede estar vacio.", vbExclamation
                        Exit Function
                    End If
                Next i
                
                'EL del mail tiene k tener la arroba @
                i = InStr(1, Text1(1).Text, "@")
                If i = 0 Then
                    MsgBox "Direccion e-mail erronea", vbExclamation
                    Exit Function
                End If
    Else
        Text2.Text = Trim(Text2.Text)
        'SOPORTE
        If Trim(Text3.Text) = "" Then
            MsgBox "El mensaje no puede ir en blanco", vbExclamation
            Exit Function
        End If
        If Option1(2).Value Then
            If Text2.Text = "" Then
                MsgBox "El campo 'OTRO asunto' no puede ir en blanco", vbExclamation
                Exit Function
            End If
        End If
    End If
      
    'Llegados aqui OK
    DatosOk = True
        
End Function


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then Exit Sub 'Si estamos en el de datos nos salimos
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

'El procedimiento servira para ir buscando los vbcrlf y cambiarlos por </p><p>
Private Sub FijarTextoMensaje()
Dim i As Integer
Dim J As Integer

    J = 1
    Do
        i = InStr(J, Text1(3).Text, vbCrLf)
        If i > 0 Then
              cad = cad & Mid(Text1(3).Text, J, i - J) & "</P><P>"
        Else
            cad = cad & Mid(Text1(3).Text, J)
        End If
        J = i + 2
    Loop Until i = 0
End Sub

Private Sub HabilitarText()
    If Option1(2).Value Then
        Text2.Enabled = True
        Text2.BackColor = vbWhite
    Else
        Text2.Enabled = False
        Text2.BackColor = &H80000018
    End If
End Sub



Private Function RecuperarDatosEMAILAriadna() As Boolean
Dim NF As Integer

    RecuperarDatosEMAILAriadna = False
    NF = FreeFile
    Open App.Path & "\soporte.dat" For Input As #NF
    Line Input #NF, cad
    Close #NF
    If cad <> "" Then RecuperarDatosEMAILAriadna = True
    
End Function


Private Function ObtenerValoresEnvioMail() As String
Dim miRsAux As ADODB.Recordset

    ObtenerValoresEnvioMail = ""

    Set miRsAux = New ADODB.Recordset
    
    
    '1 ver si el usuario que esta conectado tiene datos de email
    cad = "Select * from usuarios.usuarios where login = " & DBSet(vUsu.Login, "T")
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux!Dirfich, "T") <> "" Then
            cad = RecuperaValor(miRsAux!Dirfich, 2)
            cad = cad & "|" & RecuperaValor(miRsAux!Dirfich, 3)
            cad = cad & "|" & RecuperaValor(miRsAux!Dirfich, 4)
            cad = cad & "|" & RecuperaValor(miRsAux!Dirfich, 1) & "|"
            ObtenerValoresEnvioMail = cad
        Else
            cad = ""
        End If
    Else
        cad = ""
    End If
    miRsAux.Close
    
    If cad = "" Then
        cad = "Select diremail,SmtpHost, SmtpUser, SmtpPass  from spara1 where"
    '####Descomentar
'    Cad = Cad & " fechaini='" & Format(vParam.fechaini, FormatoFecha) & "';"
        cad = cad & " codigo=1;"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        If Not miRsAux.EOF Then
            cad = DBLet(miRsAux!SMTPhost)
            cad = cad & "|" & DBLet(miRsAux!SMTPuser)
            cad = cad & "|" & DBLet(miRsAux!SMTPpass)
            cad = cad & "|" & DBLet(miRsAux!DireMail) & "|"
            ObtenerValoresEnvioMail = cad
        End If
        miRsAux.Close
    End If
    Set miRsAux = Nothing
End Function

Private Sub HacerMultiEnvio()
Dim cad As String
Dim RS As ADODB.Recordset
Dim i As Integer, cont As Integer

On Error GoTo EMulti

        'Campos comunes
    'ENVIO MASIVO DE EMAILS
    Text1(2).Text = RecuperaValor(Me.DatosEnvio, 1)
    Text1(3).Text = RecuperaValor(Me.DatosEnvio, 2)
    
    Me.Refresh
    
    cad = "SELECT * from tmpMail WHERE codusu=" & vUsu.Codigo
    Set RS = New ADODB.Recordset
    RS.Open cad, conn, adOpenKeyset, adLockOptimistic, adCmdText

    cont = 0
    While Not RS.EOF
        cont = cont + 1
        RS.MoveNext
    Wend
    RS.MoveFirst
    
    
    i = 1
    Me.Refresh
    While Not RS.EOF
        Screen.MousePointer = vbHourglass
        Text1(0).Text = RS!nomprove
        Text1(1).Text = RS!email
        Caption = "Enviar E-MAIL (" & i & " de " & cont & ")"
        Me.Refresh
        
        'De momento volvemos a copiar el archivo como docum.pdf
        FileCopy App.Path & "\temp\" & RS!codProve & ".pdf", App.Path & "\docum.pdf"
        Me.Refresh
        NumRegElim = 0
        EnvioNuevo Nothing
        
        
'        If NumRegElim = 1 Then
'            'NO SE HA ENVIADO.
'            cad = "UPDATE tmp347 SET IMporte=0 WHERE codusu =" & vUsu.Codigo & " AND cliprov =0 AND cta='" & RS!cta & "'"
'            Conn.Execute cad
'        End If
        'Siguiente
        RS.MoveNext
        i = i + 1
    Wend
    RS.Close
    
EMulti:
    
End Sub








Private Sub MandaBusquedaPrevia(cadB As String, NomTabla As String, NomCampo As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Conexion As Byte

    'Llamamos a al form
    '##A mano
    cad = ""
    Select Case NomTabla
        Case "sclien"
            cad = cad & "C�digo|sclien|codclien|N|000000|9�"
            cad = cad & "Nombre|sclien|nomclien|T||29�"
            cad = cad & "Ape./Nom.Comer.|sclien|nomcomer|T||29�"
            cad = cad & "E-mail|sclien|" & NomCampo & "|T||33�"
'            Tabla = NomTabla
            Titulo = "Clientes"
        Case "sprove"
            cad = cad & "C�digo|sprove|codprove|N|000000|9�"
            cad = cad & "Nombre|sprove|nomprove|T||29�"
            cad = cad & "Nom.Comer.|sprove|nomcomer|T||29�"
            cad = cad & "E-mail|sprove|" & NomCampo & "|T||33�"
'            Tabla = NomTabla
            Titulo = "Proveedores"
        Case "straba"
            cad = cad & "C�digo|straba|codtraba|N|0000|9�"
            cad = cad & "Nombre|straba|nomtraba|T||44�"
            cad = cad & "E-mail|straba|" & NomCampo & "|T||44�"
'            Tabla = NomTabla
            Titulo = "trabajadores"
    End Select
    Tabla = NomTabla
    Conexion = conAri    'Conexi�n a BD: Aritaxi
    
'    Select Case Val(Me.imgBuscar(0).Tag)
'        Case 5  'Cuenta Contable
'            'Se llama a Busqueda desde el campo Cuenta contable
'            '#A MANO: Porque busca en la tabla cuentas
'            'de la base de datos de Contabilidad
'            cad = cad & "C�digo|cuentas|codmacta|T||30�Denominacion|cuentas|nommacta|T||70�"
'            Tabla = "cuentas"
'            Titulo = "Cuentas Contables"
'            Conexion = conConta    'Conexi�n a BD: Conta
'        Case Else   'Registro de la tabla de cabeceras: sartic
'            cad = cad & ParaGrid(Text1(0), 10, "C�digo")
'            cad = cad & ParaGrid(Text1(1), 50, "Nombre")
'            cad = cad & ParaGrid(Text1(2), 40, "Nombre Comercial")
'            Tabla = "sclien"
'            Titulo = "Clientes"
'            Conexion = conAri    'Conexi�n a BD: Aritaxi
'    End Select
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
'        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|2|3|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = Conexion
        frmB.vCargaFrame = (Conexion = 2)
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(2)
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub



'MULTIE ENVIO FACTURACION
Private Sub HacerMultiEnvioFacturacion()
Dim cad As String
Dim RS As ADODB.Recordset
Dim i As Integer, cont As Integer
Dim Lis As Collection
Dim ListaArchivos As Collection
Dim FormatoHtml As Boolean
Dim T1 As Single
On Error GoTo EMulti2

        'Campos comunes
    'ENVIO MASIVO DE EMAILS
    Text1(2).Text = RecuperaValor(Me.DatosEnvio, 1)
    
    
    Me.Refresh
    DoEvents
    cad = RecuperaValor(DatosEnvio, 4)
    'AGrupamos en el envio de facturas
    If Opcion = 4 Then cad = cad & " GROUP by codprove"
    Set RS = New ADODB.Recordset
    RS.Open cad, conn, adOpenKeyset, adLockOptimistic, adCmdText

    Set Lis = New Collection
    While Not RS.EOF
        Lis.Add CStr(RS!codProve)
        RS.MoveNext
    Wend
    RS.Close
    
    FormatoHtml = False
    If vParamAplic.ExeEnvioMail <> "" Then
        FormatoHtml = True
    Else
        If Not vParamAplic.EnvioDesdeOutlook Then FormatoHtml = True
    End If
    
    T1 = Timer
    For i = 1 To Lis.Count
        
        Caption = "Enviar E-MAIL (" & i & " de " & Lis.Count & ")"
        DoEvents
        cad = RecuperaValor(DatosEnvio, 4)
        cad = cad & " and codprove =" & Lis.item(i)
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Screen.MousePointer = vbHourglass
        Text1(0).Text = RS!nomclien
        Text1(1).Text = RS!email
        'Los meteremos en una tabla
        If FormatoHtml Then
            cad = "<BR><BR><TABLE BORDER=""1"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
            'Cuerpo del mensaje
            If Opcion = 4 Then
                cad = cad & "<TR><TD width=""274"" bgcolor=""#CCCCCC""><B>Factura</B></TD><TD width=""145"" bgcolor=""#CCCCCC""><B>Fecha</B></TD><TD width=""145"" bgcolor=""#CCCCCC""><B>Importe</B></td></TR>"
            Else
                cad = cad & "<TR><TD width=""640"" bgcolor=""#CCCCCC""><B>Documento</B></TD></TR>"
            End If
        Else
            If Opcion = 4 Then
                cad = " Factura             Fecha             Importe "
            Else
                cad = cad & "Documento "
            End If
            cad = vbCrLf & vbCrLf & vbCrLf & cad & vbCrLf & vbCrLf & String(40, "-") & vbCrLf & vbCrLf
        End If
        Text1(3).Text = RecuperaValor(Me.DatosEnvio, 2) & cad
        Set ListaArchivos = New Collection
        While Not RS.EOF
            

           
            Me.Refresh
            '
            'De momento volvemos a copiar el archivo como docum.pdf
            If Opcion = 4 Then
                'cad = App.Path & "\temp\" & RS!NumAlbar & Format(RS!codProve, "0000000") & Format(RS!codArtic, "0000000") & Format(RS!FechaAlb, "yymmdd") & ".pdf"
                cad = App.Path & "\temp\" & RS!NumAlbar & Format(RS!codArtic, "0000000") & ".pdf"
            Else
                'Opcion5: Carta renovacion
                cad = App.Path & "\temp\" & Format(RS!codProve, "0000000") & ".pdf"
            End If
            If Dir(cad, vbArchive) = "" Then
                'ERROR. El fichero ha sido eliminado
                MsgBox "No existe el fichero: " & cad & vbCrLf & "El proceso finalizara", vbExclamation
                RS.Close
                Exit Sub
            Else
                ListaArchivos.Add cad
                'En el asunto pondremos los archivos que enviamos
                cad = ""
                If Opcion = 4 Then
                    
                    If FormatoHtml Then
                        cad = "</div></TD><TD><div align=""right"">" & Format(RS!Cantidad, FormatoImporte) & "</div></TD></TR>"
                    Else
                        cad = Space(20) & Format(RS!Cantidad, FormatoImporte)
                    End If
                    
                    If FormatoHtml Then
                        cad = "</TD><TD><div align=""center"">" & Format(RS!FechaAlb, "dd/mm/yyyy") & cad
                    Else
                        cad = Space(15) & Format(RS!FechaAlb, "dd/mm/yyyy") & cad
                    End If
                    
        
                    cad = RS!NumAlbar & Format(RS!codArtic, "0000000") & cad
                                
                    If FormatoHtml Then
                        cad = "<TR><TD>" & cad
                    Else
                        cad = cad & vbCrLf
                    End If
                
                Else
                    'Opcion:5.  Carta renovacion
                    If FormatoHtml Then cad = "<TR><TD>"
                    cad = cad & "Documento" & Format(RS!codProve, "0000000")
                    If FormatoHtml Then
                        cad = cad & "</TD></TR>"
                    Else
                        cad = cad & vbCrLf
                    End If
                
                End If
                
                Text1(3).Text = Text1(3).Text & "    " & cad
            End If
            
            'Siguiente
            RS.MoveNext
            
        Wend
        RS.Close
        If FormatoHtml Then Text1(3).Text = Text1(3).Text & "</TABLE><BR><BR>"
        
        EnvioNuevo ListaArchivos
        
        Set ListaArchivos = Nothing
        
        T1 = Timer - T1
        If T1 < 3 Then
            T1 = 3 - T1
            Espera T1
        End If
        T1 = Timer
    Next i
    Set Lis = Nothing
    Exit Sub
EMulti2:
    MuestraError Err.Number
End Sub




Private Sub HacerMultiEnvioConfirPed()

Dim RS As ADODB.Recordset
Dim i As Integer, cont As Integer
Dim Lis As Collection
Dim ListaArchivos As Collection
Dim FormatoHtml As Boolean
Dim T1 As Single
On Error GoTo EMulti2

    'Campos comunes

    
    Me.Refresh
    DoEvents
    
    
    cad = "select * from tmpnlotes where codusu =" & vUsu.Codigo
    cad = cad & " GROUP by codprove order by codprove"
    Set RS = New ADODB.Recordset
    RS.Open cad, conn, adOpenKeyset, adLockOptimistic, adCmdText

    Set Lis = New Collection
    While Not RS.EOF
        Lis.Add CStr(RS!codProve)
        RS.MoveNext
    Wend
    RS.Close
    
    FormatoHtml = False
    If vParamAplic.ExeEnvioMail <> "" Then
        FormatoHtml = True
    Else
        If Not vParamAplic.EnvioDesdeOutlook Then FormatoHtml = True
    End If
    Text1(2).Text = Text1(4).Text
    T1 = Timer
    
    For i = 1 To Lis.Count
        
        Caption = "Enviar E-MAIL (" & i & " de " & Lis.Count & ")"
        DoEvents
        cad = "select * from tmpnlotes where codusu =" & vUsu.Codigo
        cad = cad & " and codprove =" & Lis.item(i) & " ORDER BY fechaalb"   'Asi nos devolvera la primera entrada para cadprovedor donde tiene el email
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Screen.MousePointer = vbHourglass
        'El email estara en numalbar+nomartic+numlote
        
        cad = RS!NumAlbar & DBLet(RS!NomArtic, "T") & DBLet(RS!numlotes, "T")
        Text1(0).Text = cad
        Text1(1).Text = cad
        Text1(3).Text = Text1(5).Text  'Body
        If FormatoHtml Then
            cad = "<P>"
            FijarTextoMensaje
            cad = cad & "</P>"
            Text1(3).Text = cad
        End If
        'Los meteremos en una tabla
        If FormatoHtml Then
            cad = "<BR><BR><TABLE BORDER=""1"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
            'Cuerpo del mensaje

                cad = cad & "<TR><TD width=""640"" bgcolor=""#CCCCCC""><B>Documento adjuntados</B></TD></TR>"
  
        Else

            cad = "Documentos adjuntados "

            cad = vbCrLf & vbCrLf & vbCrLf & cad & vbCrLf & vbCrLf & String(40, "-") & vbCrLf & vbCrLf
        End If
        Text1(3).Text = Text1(3) & cad
        Set ListaArchivos = New Collection
        While Not RS.EOF
            

           
            Me.Refresh
            '
            'De momento volvemos a copiar el archivo como docum.pdf
          
            cad = App.Path & "\temp\" & RS!codArtic
      
            If Dir(cad, vbArchive) = "" Then
                'ERROR. El fichero ha sido eliminado
                MsgBox "No existe el fichero: " & cad & vbCrLf & "El proceso finalizara", vbExclamation
                RS.Close
                Exit Sub
            Else
                ListaArchivos.Add cad
                'En el asunto pondremos los archivos que enviamos
                cad = ""

                    'Opcion:5.  Carta renovacion
                    If FormatoHtml Then cad = "<TR><TD>"
                    cad = cad & "Documento " & Format(RS!codArtic, "0000000")
                    If FormatoHtml Then
                        cad = cad & "</TD></TR>"
                    Else
                        cad = cad & vbCrLf
                    End If
                
 
                
                Text1(3).Text = Text1(3).Text & "    " & cad
            End If
            
            'Siguiente
            RS.MoveNext
            
        Wend
        RS.Close
        If FormatoHtml Then Text1(3).Text = Text1(3).Text & "</TABLE><BR><BR>"
        
        EnvioNuevo ListaArchivos
        
        Set ListaArchivos = Nothing
        
        T1 = Timer - T1
        If T1 < 3 Then
            T1 = 3 - T1
            Espera T1
        End If
        T1 = Timer
    Next i
    Set Lis = Nothing
    Exit Sub
EMulti2:
    MuestraError Err.Number
End Sub

















Private Sub EnvioDesdeExeNuestro(ListaArchivos As Collection)
Dim Lanza As String
Dim J As Integer

    If Not DatosOk Then Exit Sub
        
    'Dire email
    Lanza = Text1(1).Text & "|"
    'Asunto
    Lanza = Lanza & Text1(2).Text & "|"
    
    'Aqui pondremos lo del texto del BODY
    Lanza = Lanza & Text1(3).Text & "|"
    
    
    'Envio o mostrar
    Lanza = Lanza & "1"   '0. Display        1.send
    
    'Campos reservados para el futuro
    Lanza = Lanza & "||||"
    
    'El/los adjuntos
    If Opcion <> 1 Then   'Solo la opcion 1 NO lleva attachment
        'ADjunatmos el PDF
        If ListaArchivos Is Nothing Then
            Lanza = Lanza & App.Path & "\docum.pdf" & "|"
        Else
            
            For J = 1 To ListaArchivos.Count
                   Lanza = Lanza & ListaArchivos.item(J) & "|"
            Next J
        End If
    End If
    



    
    Lanza = App.Path & "\" & vParamAplic.ExeEnvioMail & " " & Lanza
    Shell Lanza, vbNormalFocus

End Sub





