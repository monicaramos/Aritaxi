VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLLamadasDatos2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "L L A M A D A S"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   8685
   Icon            =   "frmLLamadasDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
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
      Index           =   11
      Left            =   3600
      MaxLength       =   30
      TabIndex        =   26
      Tag             =   "Poblaci�n|T|S|||sllama|nomtraba1||N|"
      Text            =   "Text1"
      Top             =   2760
      Width           =   4845
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
      Index           =   10
      Left            =   3600
      MaxLength       =   30
      TabIndex        =   25
      Tag             =   "Poblaci�n|T|S|||sllama|nomtraba||N|"
      Text            =   "Text1"
      Top             =   2280
      Width           =   4845
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
      Height          =   2235
      Index           =   9
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Tag             =   "Poblaci�n|T|S|||sllama|observac|||"
      Text            =   "frmLLamadasDatos.frx":000C
      Top             =   3720
      Width           =   5925
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "Motivo|N|N|||sllama|codllama1||N|"
      Top             =   3240
      Width           =   5955
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
      Index           =   8
      Left            =   2520
      MaxLength       =   30
      TabIndex        =   8
      Tag             =   "Poblaci�n|T|S|||sllama|codtraba1|000|N|"
      Text            =   "Text1"
      Top             =   2760
      Width           =   1005
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
      Index           =   7
      Left            =   2520
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "Poblaci�n|T|S|||sllama|codtraba|000|N|"
      Text            =   "Text1"
      Top             =   2280
      Width           =   1005
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
      Index           =   6
      Left            =   4920
      MaxLength       =   40
      TabIndex        =   6
      Tag             =   "Poblaci�n|T|S|||sllama|mail||N|"
      Text            =   "Text1"
      Top             =   1680
      Width           =   3525
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
      Index           =   5
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Poblaci�n|T|S|||sllama|telefono||N|"
      Text            =   "Text1"
      Top             =   1680
      Width           =   1725
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
      Index           =   2
      Left            =   3480
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "Nomclien|T|N|||sllama|nomclien||N|"
      Text            =   "Text1"
      Top             =   720
      Width           =   4965
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
      Index           =   3
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   2
      Tag             =   "Cod clien|N|S|0||sllama|codclien||N|"
      Text            =   "Text1"
      Top             =   720
      Width           =   885
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
      Index           =   4
      Left            =   2520
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "Poblaci�n|T|S|||sllama|perconta||N|"
      Text            =   "Text1"
      Top             =   1200
      Width           =   3645
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
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Tag             =   "FechaHora|H|N|||sllama|feholla|dd/mm/yyyy hh:mm:ss|S|"
      Text            =   "Text"
      Top             =   150
      Width           =   2445
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
      Index           =   1
      Left            =   6270
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "usuario|T|N|||sllama|usuario||S|"
      Text            =   "Text1"
      Top             =   150
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   210
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   7440
      TabIndex        =   12
      Top             =   6120
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   6240
      TabIndex        =   11
      Top             =   6120
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   495
      Left            =   2040
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   2220
      Tag             =   "-1"
      ToolTipText     =   "Buscar poblaci�n"
      Top             =   2760
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   2220
      Tag             =   "-1"
      ToolTipText     =   "Buscar poblaci�n"
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
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
      Index           =   9
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Motivo"
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
      Index           =   8
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Atendido por"
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
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador asignado"
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
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "email"
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
      Index           =   5
      Left            =   4410
      TabIndex        =   20
      Top             =   1710
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Telefono"
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
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Persona contacto"
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
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
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
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   2220
      Tag             =   "-1"
      ToolTipText     =   "Buscar poblaci�n"
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
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
      Index           =   1
      Left            =   5340
      TabIndex        =   16
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha /  Hora"
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
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   150
      Width           =   1785
   End
End
Attribute VB_Name = "frmLLamadasDatos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmCl As frmFacClientes
Attribute frmCl.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1

Public SoloVer As Boolean

'  Variables comunes a todos los formularios
Public vModo As Byte
'       1. Buscar
'       3. Insertar
'       4. Modificar
    

Dim PrimeVez As Boolean
Dim Sql As String
Dim txtAnterior As String

Private Sub cmdAceptar_Click()
Dim b As Boolean
    b = False
    If vModo = 1 Then
        'Busqueda
        Sql = ObtenerBusqueda(Me, False)
        If Sql <> "" Then
            CadenaDesdeOtroForm = Sql
            b = True
        Else
            MsgBox "Ponga algun parametro para la busqueda", vbExclamation
            CadenaDesdeOtroForm = ""
        End If
    Else
        If DatosOk Then
            If vModo = 3 Then
                b = InsertarDesdeForm(Me)
            Else
                b = ModificaDesdeFormulario(Me, 1)
            End If
            If b Then CadenaDesdeOtroForm = "OK"
                
            End If
                
   End If
   If b Then Unload Me
End Sub

Private Sub cmdCancelar_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Form_Activate()
    If PrimeVez Then
        
        PrimeVez = False
        If vModo = 4 Then
            CadenaDesdeOtroForm = "Select * from sllama WHERE " & CadenaDesdeOtroForm
            Adodc1.ConnectionString = conn
            Adodc1.RecordSource = CadenaDesdeOtroForm
            Adodc1.Refresh
            
            'Si por algun motivo es EOF.
            If Not Me.Adodc1.Recordset.EOF Then PonerCamposForma Me, Adodc1

            CadenaDesdeOtroForm = ""
            
          
        Else
            'If vModo = 3 Then PonerFoco Text1(3)
            
        End If
        PonerModo
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Load()
Dim I As Integer

    PrimeVez = True
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I

    LimpiarCampos

    AccionesVarias
    
   Me.cmdAceptar.visible = Not SoloVer
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox del form
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


    







Private Sub frmCl_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim I As Integer
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        Set frmCl = New frmFacClientes
        frmCl.DatosADevolverBusqueda = "0"
        frmCl.Show vbModal
        Set frmCl = Nothing
        I = 3
    Else
        Set frmT = New frmAdmTrabajadores
        frmT.DatosADevolverBusqueda = "0|"
        frmT.Show vbModal
        Set frmT = Nothing
        I = 6 + Index
    End If
    
    If CadenaDesdeOtroForm <> "" Then
        Text1(I).Text = CadenaDesdeOtroForm
        PonerFoco Text1(I)
        Text1_LostFocus I
        CadenaDesdeOtroForm = ""
        If I = 3 Then
            PonerFoco Text1(7)  'traba
        ElseIf I = 7 Then
            PonerFoco Text1(8)
        Else
            PonerFocoCbo Me.Combo1
        End If
    End If
    
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), vModo
    txtAnterior = Text1(Index).Text
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 9 Then KEYpressGnral KeyAscii, vModo, False
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String


    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If vModo = 1 Then
        Text1(Index).BackColor = vbWhite
        Exit Sub
    End If
    If txtAnterior = Text1(Index).Text Then Exit Sub
    Select Case Index
        Case 0
 

        Case 3 'clien
            
            devuelve = ""
            
            If Text1(Index).Text <> "" Then
                If PonerFormatoEntero(Text1(Index)) Then
                    If PonerDatosCliente Then devuelve = Text1(2).Text
                        
                End If
            End If
            BloquearTextCliente devuelve <> ""
            Text1(2).Text = devuelve
            If devuelve = "" And Text1(Index).Text <> "" Then
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            Else
                PonerFoco Text1(5)
            End If
         Case 7, 8
            devuelve = ""
            If Text1(Index).Text <> "" Then
                If PonerFormatoEntero(Text1(Index)) Then
                    devuelve = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", Text1(Index).Text)
                    If devuelve = "" Then
                        MsgBox "No existe trabajador: " & Text1(Index).Text, vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
            End If
            Text1(Index + 3).Text = devuelve
    End Select
End Sub



Private Sub PonerCampos()
    If Adodc1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Adodc1

    '-- Esto permanece para saber donde estamos
    ' lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub









Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim cad As String

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar datos OK
    If Not b Then Exit Function
        
    DatosOk = b
End Function














Private Sub AccionesVarias()
Dim I As Integer

    On Error GoTo ER
    
    
    'Cargamos combo
    Me.Combo1.Clear
    
    Sql = "SELECT * FROM sllama1 ORDER BY nomllama1"
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!nomllama1
        Combo1.ItemData(Combo1.NewIndex) = CInt(miRsAux!codllama1)
        I = I + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close

        
    If vModo = 3 Then
        'Estamos insertando. Entonces le pongo la fecha y la hora
        Sql = "Select CURDATE() ,CURTIME() "
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO PUE SER NULL
        Sql = Format(miRsAux.Fields(0), "dd/mm/yyyy") & " " & Format(miRsAux.Fields(1), "hh:mm:ss")
        Text1(0).Text = Sql
        miRsAux.Close
        
        CadenaDesdeOtroForm = PonerTrabajadorConectado(Sql)
        If CadenaDesdeOtroForm = "" Then
            Sql = "SIN LOG."
            cmdAceptar.Enabled = False
        Else
            Sql = vUsu.Login
            Text1(8).Text = CadenaDesdeOtroForm
        End If
        Text1(1).Text = Sql
        Text1(11).Text = Sql
        
        

        
        
        CadenaDesdeOtroForm = ""
            
        
    End If
        
 
    
    
ER:
    If Err.Number <> 0 Then MuestraError Err.Number, Sql
    Set miRsAux = Nothing
End Sub


Private Sub PonerModo()
Dim b As Boolean
Dim I As Integer

    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I

    b = vModo = 1
    BloquearTxt Text1(0), Not b, True
    
    'usu
    BloquearTxt Text1(1), Not b, False
    
    BloquearTextCliente Text1(3).Text <> ""
    
End Sub


Private Sub BloquearTextCliente(NOBloquear As Boolean)
    BloquearTxt Text1(2), NOBloquear, False
    
End Sub
Private Function PonerDatosCliente() As Boolean
Dim vC As CCliente
Dim b As Boolean
    b = False
    Set vC = New CCliente
    If vC.LeerDatos(Text1(3).Text) Then
        vC.MostrarObservaciones
        Text1(2).Text = vC.Nombre
        Text1(4).Text = vC.PersonaContacto
        Text1(5).Text = vC.TfnoClien
        Text1(6).Text = vC.EMailAdm
        b = True
    Else
        MsgBox "No existe el cliente: " & Text1(3).Text, vbExclamation
    End If
    BloquearTextCliente Not b
    PonerDatosCliente = b
    Set vC = Nothing
End Function
