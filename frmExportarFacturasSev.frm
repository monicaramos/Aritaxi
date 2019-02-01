VERSION 5.00
Begin VB.Form frmExportarFacturasSev 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar Datos/Facturas"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   Icon            =   "frmExportarFacturasSev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3960
      TabIndex        =   6
      Top             =   4140
      Width           =   1695
   End
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
      Left            =   2040
      TabIndex        =   5
      Top             =   4140
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Height          =   3090
      Left            =   150
      TabIndex        =   8
      Top             =   150
      Width           =   5640
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
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
         Left            =   2790
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codsocio|||"
         Top             =   1755
         Width           =   1380
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
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
         Left            =   2790
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Código Postal|T|S|||clientes|codsocio|||"
         Top             =   2100
         Width           =   1380
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
         Index           =   1
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   270
         Width           =   3435
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
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
         Left            =   2790
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1290
         Width           =   1380
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
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
         Left            =   2790
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   915
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Socio / Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   2
         Left            =   270
         TabIndex        =   15
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1860
         TabIndex        =   14
         Top             =   2100
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1860
         TabIndex        =   13
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   270
         TabIndex        =   12
         Top             =   270
         Width           =   1605
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   1860
         TabIndex        =   11
         Top             =   930
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   1860
         TabIndex        =   10
         Top             =   1290
         Width           =   660
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2520
         ToolTipText     =   "Buscar fecha"
         Top             =   930
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   2520
         ToolTipText     =   "Buscar fecha"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   270
         TabIndex        =   9
         Top             =   720
         Width           =   1665
      End
   End
   Begin VB.Label lblInf 
      Alignment       =   2  'Center
      Caption         =   "Información del proceso"
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
      Left            =   240
      TabIndex        =   7
      Top             =   3465
      Width           =   5295
   End
End
Attribute VB_Name = "frmExportarFacturasSev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1009


Public Tipo As Byte
    'Tipo:  0 Impresion de facturas

Dim DesdeFecha As Date
Dim Hastafecha As Date
Dim frmVis As frmVisReport

Dim nompath As String
Dim ConDetalle As Byte

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

Private Sub cmdAceptar_Click()
    If Not DatosOk() Then Exit Sub
    
    
    If Combo1(1).ListIndex <= 5 Then
        CargaFacturas Combo1(1).ListIndex + 1
    ElseIf Combo1(1).ListIndex = 6 Then
        CargaSociosClientes 0
    Else
        CargaSociosClientes 1
    End If
    
    
'
'    '-- Cargar facturas  entre las fechas seleccionadas
'    Select Case Combo1(1).ListIndex
'        Case 0 ' facturas liquidacion socio
'            CargaFacturas Combo1(1).ListIndex + 1
'
'        Case 1 ' facturas de cliente
'            CargaFacturas Combo1(1).ListIndex + 1
'        Case 2 ' facturas cuotas socios
'            CargaFacturasCuotasSocio DesdeFecha, Hastafecha
'
'        Case 3 ' facturas de publicidad socios
'            CargaFacturasPublicidadSocio DesdeFecha, Hastafecha
'
'        Case 4 ' facturas de publicidad clientes
'            CargaFacturasPublicidadCliente DesdeFecha, Hastafecha
'
'        Case 5 ' facturas de venta socios
'            CargaFacturasVentaSocio DesdeFecha, Hastafecha
'
'        Case 6 ' socios
'            CargaSociosClientes 0
'
'        Case 7 ' clientes
'            CargaSociosClientes 1
'
'
'    End Select
    cmdSalir_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function DatosOk() As Boolean
Dim Codigo As String

    'Comprobacion de fechas
    DesdeFecha = CDate(txtCodigo(0).Text)
    Hastafecha = CDate(txtCodigo(1).Text)
    If DesdeFecha > Hastafecha Then
        MsgBox "La fecha desde debe ser menor que la fecha hasta", vbInformation
        Exit Function
    End If
    
    'Comprobacion de configuracion de carpeta destino
    Codigo = ""
    If vParamAplic.PathFacturaE = "" Then
        Codigo = "Falta configurar parametros"
    End If
    If Codigo <> "" Then
        MsgBox Codigo, vbExclamation
        Exit Function
    End If
    
    DatosOk = True
End Function





Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmppal.Icon

    txtCodigo(0).Text = Date
    txtCodigo(1).Text = Date
    
    CargaCombo
    
    Combo1(1).ListIndex = 0
    
    
    For I = 0 To Me.imgFec.Count - 1
        imgFec(I).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next
    
    

   '[Monica]05/09/2012: la carpeta de destino la tenemos en parámetros
    nompath = vParamAplic.PathFacturaE
        
End Sub


Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object
    
    Set frmC = New frmCal
    frmC.Fecha = Now

    esq = imgFec(Index).Left
    dalt = imgFec(Index).top

    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.top
        Set obj = obj.Container
    Wend
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.top = dalt + imgFec(Index).Parent.top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(0).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    
    PonerFormatoFecha txtCodigo(Index)
    If txtCodigo(Index).Text <> "" Then frmC.Fecha = CDate(txtCodigo(Index).Text)

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag))
    ' ***************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub


Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = teclaBuscar Then
'        Select Case Index
'            Case 0: KEYFecha KeyAscii, 0 'fecha desde
'            Case 1: KEYFecha KeyAscii, 1 'fecha hasta
'            Case 2: KEYFecha KeyAscii, 1 'fecha de factura
'        End Select
'    Else
        KEYpress KeyAscii
'    End If

End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

'    'Quitar espacios en blanco por los lados
'    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0, 1 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
    End Select
End Sub


Private Sub CargaSociosClientes(Tipo As Byte)
' tipo: 0 = socios
'       1 = clientes
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim v_cadena As String
    
Dim Tabla As String
Dim vNombre As String
    
    
On Error GoTo err_CargaDatos

    '[Monica]29/02/2012: añadida la insercion en la tabla temporal por el report
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    If Tipo = 0 Then
        Tabla = "sclien"
        vNombre = "socios"
    Else
        Tabla = "scliente"
        vNombre = "clientes"
    End If

    Sql = "select nomclien, codclien, fechaalt from " & Tabla
            
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Sql = "insert into tmpinformes (codusu, nombre1, codigo1, fecha1) "
        Sql = Sql & " select " & vUsu.Codigo & ", nomclien, codclien, fechaalt from " & Tabla
                
        conn.Execute Sql
        
        v_cadena = vParamAplic.PathFacturaE
        v_cadena = v_cadena & "/" & vNombre
        
        If Jason_GET(v_cadena) <> "" Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
        End If
    Else
        MsgBox "No hay datos entre esos límites.", vbExclamation
    End If
    
    Set Rs = Nothing
    
    Exit Sub
err_CargaDatos:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Cargar " & vNombre
    End If
End Sub


Private Sub CargaFacturas(Tipo As Byte)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim C1 As String
    Dim C2 As String
    Dim c3 As String
    Dim c4 As String
    Dim F1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim Fr As frmVisReport
    
    Dim Variedad As String
    
    Dim TipoFact1 As Byte
    Dim Gastos As Currency
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadParam As String

Dim v_cadena As String

    '[Monica]29/02/2012: añadida la insercion en la tabla temporal por el report
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql


    Select Case Tipo
        Case 1 ' liquidacion
            Sql = "select sfactusoc.*, stipom.letraser  from sfactusoc, stipom where sfactusoc.codtipom = stipom.codtipom  "
            Sql = Sql & " and sfactusoc.codtipom in ('FLI','FRL') "
            If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
            If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
            If txtCodigo(2) <> "" Then Sql = Sql & " and codsocio >= " & DBSet(txtCodigo(2), "N")
            If txtCodigo(3) <> "" Then Sql = Sql & " and codsocio <= " & DBSet(txtCodigo(3), "N")
            
            
        Case 2 ' a clientes
            Sql = "select scafaccli.*, stipom.letraser from scafaccli, stipom where scafaccli.codtipom = stipom.codtipom "
            Sql = Sql & " and scafaccli.codtipom in ('FAC','FVC','FRN') "
            If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
            If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
            If txtCodigo(2) <> "" Then Sql = Sql & " and codclien >= " & DBSet(txtCodigo(2), "N")
            If txtCodigo(3) <> "" Then Sql = Sql & " and codclien <= " & DBSet(txtCodigo(3), "N")
            
        
        Case 3 ' cuotas
            Sql = "select scafac.*, stipom.letraser from scafac, stipom where scafac.codtipom = stipom.codtipom "
            Sql = Sql & " and scafac.codtipom  in ('FCN','FCE','FRC') "
            If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
            If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
            If txtCodigo(2) <> "" Then Sql = Sql & " and codclien >= " & DBSet(txtCodigo(2), "N")
            If txtCodigo(3) <> "" Then Sql = Sql & " and codclien <= " & DBSet(txtCodigo(3), "N")
                    
        
        Case 4 ' publicidad socio
            Sql = "select sfactusoc.*, stipom.letraser from sfactusoc, stipom where sfactusoc.codtipom = stipom.codtipom "
            Sql = Sql & " and sfactusoc.codtipom = 'FPS'"
            If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
            If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
            If txtCodigo(2) <> "" Then Sql = Sql & " and codsocio >= " & DBSet(txtCodigo(2), "N")
            If txtCodigo(3) <> "" Then Sql = Sql & " and codsocio <= " & DBSet(txtCodigo(3), "N")
        
        Case 5 ' publicidad cliente
            Sql = "select scafaccli.*, stipom.letraser from scafaccli, stipom where scafaccli.codtipom = stipom.codtipom "
            Sql = Sql & " and scafaccli.codtipom in ('FPC','FRP') "
            If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
            If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
            If txtCodigo(2) <> "" Then Sql = Sql & " and codclien >= " & DBSet(txtCodigo(2), "N")
            If txtCodigo(3) <> "" Then Sql = Sql & " and codclien <= " & DBSet(txtCodigo(3), "N")
        
        Case 6 ' venta a socios
            Sql = "select scafac.*, stipom.letraser from scafac, stipom where scafac.codtipom = stipom.codtipom "
            Sql = Sql & " and scafac.codtipom  in ('FAV','FAT','FAR') "
            If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
            If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
            If txtCodigo(2) <> "" Then Sql = Sql & " and codclien >= " & DBSet(txtCodigo(2), "N")
            If txtCodigo(3) <> "" Then Sql = Sql & " and codclien <= " & DBSet(txtCodigo(3), "N")
        
        
    End Select
            
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        
        v_cadena = vParamAplic.PathFacturaE
        
        Select Case Tipo
            Case 1
                v_cadena = v_cadena & "/liquidaciones/facturas/"
            Case 2
                v_cadena = v_cadena & "/faccli/facturas/"
            Case 3
                v_cadena = v_cadena & "/cuota/facturas/"
            Case 4
                v_cadena = v_cadena & "/pubsoc/facturas/"
            Case 5
                v_cadena = v_cadena & "/pubcli/facturas/"
            Case 6
                v_cadena = v_cadena & "/ventas/facturas/"
        End Select
                        
        'desde socio/cliente
        If txtCodigo(2).Text <> "" Then
            v_cadena = v_cadena & txtCodigo(2)
        Else
            v_cadena = v_cadena & "ALL"
        End If
        v_cadena = v_cadena & "/"
        'hasta socio/cliente
        If txtCodigo(3).Text <> "" Then
            v_cadena = v_cadena & txtCodigo(3)
        Else
            v_cadena = v_cadena & "ALL"
        End If
        v_cadena = v_cadena & "/"
        'desde fecha
        If txtCodigo(0).Text <> "" Then
            v_cadena = v_cadena & Format(txtCodigo(0), "yyyy-mm-dd")
        Else
            v_cadena = v_cadena & "ALL"
        End If
        v_cadena = v_cadena & "/"
        'hasta fecha
        If txtCodigo(1).Text <> "" Then
            v_cadena = v_cadena & Format(txtCodigo(1), "yyyy-mm-dd")
        Else
            v_cadena = v_cadena & "ALL"
        End If
        
        If Jason_GET(v_cadena) <> "" Then
            ' tenemos que marcarlos todos como que han sido transpasados a la app de rafa
            Select Case Tipo
                Case 1
                    'actualizamos el pasaridoc de facturas socios
                    Sql = "update sfactusoc set exportada = 1 where codtipom in ('FLI','FRL')"
                    If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
                    If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
                    If txtCodigo(2) <> "" Then Sql = Sql & " and codsocio >= " & DBSet(txtCodigo(2), "N")
                    If txtCodigo(3) <> "" Then Sql = Sql & " and codsocio <= " & DBSet(txtCodigo(3), "N")
                    
                    conn.Execute Sql
                    
                Case 2
                    'actualizamos el pasaridoc de facturas clientes
                    Sql = "update scafaccli set exportada = 1 where codtipom in ('FAC','FVC','FRN')"
                    If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
                    If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
                    If txtCodigo(2) <> "" Then Sql = Sql & " and codclien >= " & DBSet(txtCodigo(2), "N")
                    If txtCodigo(3) <> "" Then Sql = Sql & " and codclien <= " & DBSet(txtCodigo(3), "N")
                    
                    conn.Execute Sql
                
                Case 3
                    'actualizamos el pasaridoc de facturas socios
                    Sql = "update scafac set exportada = 1 where codtipom in ('FCN','FCE','FRC')"
                    If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
                    If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
                    If txtCodigo(2) <> "" Then Sql = Sql & " and codclien >= " & DBSet(txtCodigo(2), "N")
                    If txtCodigo(3) <> "" Then Sql = Sql & " and codclien <= " & DBSet(txtCodigo(3), "N")
                    
                    conn.Execute Sql
                
                Case 4
                    'actualizamos el pasaridoc de facturas socios
                    Sql = "update sfactusoc set exportada = 1 where codtipom in ('FPS')"
                    If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
                    If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
                    If txtCodigo(2) <> "" Then Sql = Sql & " and codsocio >= " & DBSet(txtCodigo(2), "N")
                    If txtCodigo(3) <> "" Then Sql = Sql & " and codsocio <= " & DBSet(txtCodigo(3), "N")
                    
                    conn.Execute Sql
                
                Case 5
                    'actualizamos el pasaridoc de facturas clientes
                    Sql = "update scafaccli set exportada = 1 where codtipom in ('FPC','FRP')"
                    If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
                    If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
                    If txtCodigo(2) <> "" Then Sql = Sql & " and codclien >= " & DBSet(txtCodigo(2), "N")
                    If txtCodigo(3) <> "" Then Sql = Sql & " and codclien <= " & DBSet(txtCodigo(3), "N")
                    
                    conn.Execute Sql
                
                Case 6
                    'actualizamos el pasaridoc de facturas socios
                    Sql = "update scafac set exportada = 1 where codtipom in ('FAV','FAT','FAR')"
                    If txtCodigo(0) <> "" Then Sql = Sql & " and fecfactu >= " & DBSet(txtCodigo(0), "F")
                    If txtCodigo(1) <> "" Then Sql = Sql & " and fecfactu <= " & DBSet(txtCodigo(1), "F")
                    If txtCodigo(2) <> "" Then Sql = Sql & " and codclien >= " & DBSet(txtCodigo(2), "N")
                    If txtCodigo(3) <> "" Then Sql = Sql & " and codclien <= " & DBSet(txtCodigo(3), "N")
                    conn.Execute Sql
                
            End Select
                
            MsgBox "Proceso realizado correctamente.", vbExclamation
        End If
    Else
        MsgBox "No hay datos entre esos límites.", vbExclamation
    End If
    
    Set Rs = Nothing
    
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas Liquidación Socio"
    End If
End Sub




Private Sub CargaCombo()
Dim ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I
    
    Combo1(1).Clear
    
    Combo1(1).AddItem "Liquidación Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Factura Clientes"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Facturas Cuotas Socios"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    '[Monica]05/02/2015: desmarcamos los otros tipos de facturas
    Combo1(1).AddItem "Facturas Publicidad Socios"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 3
    Combo1(1).AddItem "Facturas Publicidad Clientes"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 4
    Combo1(1).AddItem "Facturas Venta Socios"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 5
    
    Combo1(1).AddItem "Socios"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 6
    
    Combo1(1).AddItem "Clientes"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 7
    
End Sub




Private Function IntentaMatar(FicheroPDF As String) As Boolean
Dim I As Integer

    On Error Resume Next
    I = 1
    IntentaMatar = False
    Do
        If Dir(FicheroPDF, vbArchive) <> "" Then
            Kill FicheroPDF
            If Err.Number <> 0 Then
                Err.Clear
                I = I + 1
            Else
                IntentaMatar = True
                I = 6
            End If
        Else
            IntentaMatar = True
            I = 6
        End If
    Loop Until I < 5 Or IntentaMatar = True
    
    
End Function






