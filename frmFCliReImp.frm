VERSION 5.00
Begin VB.Form frmFCliReImp 
   Caption         =   "Informes"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Clientes sin Facturacion Electrónica"
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
      Left            =   330
      TabIndex        =   22
      Top             =   5040
      Width           =   4635
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Rectificativas"
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
      Left            =   330
      TabIndex        =   21
      Top             =   4590
      Width           =   2025
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   1470
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "Num vehiculo|N|N|||shilla|codclien|00000|S|"
      Top             =   1800
      Width           =   885
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
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
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1800
      Width           =   3765
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   1470
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Num vehiculo|N|N|||shilla|codclien|00000|S|"
      Top             =   1455
      Width           =   885
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
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
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1455
      Width           =   3765
   End
   Begin VB.CheckBox chk_duplicado 
      Caption         =   "Duplicado"
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
      Left            =   330
      TabIndex        =   6
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtcodigo 
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
      Index           =   85
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   4
      Top             =   3510
      Width           =   1245
   End
   Begin VB.TextBox txtcodigo 
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
      Index           =   86
      Left            =   3825
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3525
      Width           =   1245
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
      Left            =   3360
      TabIndex        =   7
      Top             =   5625
      Width           =   1135
   End
   Begin VB.CommandButton cmdCancelar 
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
      Left            =   4650
      TabIndex        =   8
      Top             =   5625
      Width           =   1135
   End
   Begin VB.TextBox txtcodigo 
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
      Index           =   36
      Left            =   1470
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "codtalle|N|N|||scaord|codtalle|00|S|"
      Top             =   2490
      Width           =   1245
   End
   Begin VB.TextBox txtcodigo 
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
      Index           =   37
      Left            =   3810
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "codtalle|N|N|||scaord|codtalle|00|S|"
      Top             =   2490
      Width           =   1245
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   4
      Left            =   540
      TabIndex        =   20
      Top             =   1800
      Width           =   570
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Index           =   1
      Left            =   1170
      Top             =   1800
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   3
      Left            =   540
      TabIndex        =   19
      Top             =   1455
      Width           =   600
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   240
      Index           =   2
      Left            =   330
      TabIndex        =   18
      Top             =   1110
      Width           =   765
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Index           =   0
      Left            =   1170
      Top             =   1455
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   23
      Left            =   1170
      Top             =   3540
      Width           =   240
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Fact."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   240
      Index           =   7
      Left            =   330
      TabIndex        =   15
      Top             =   3180
      Width           =   1215
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   7
      Left            =   480
      TabIndex        =   14
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   8
      Left            =   2940
      TabIndex        =   13
      Top             =   3540
      Width           =   570
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   24
      Left            =   3540
      Top             =   3540
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   540
      TabIndex        =   12
      Top             =   2550
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   2970
      TabIndex        =   11
      Top             =   2550
      Width           =   570
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Nº Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   240
      Index           =   4
      Left            =   330
      TabIndex        =   10
      Top             =   2220
      Width           =   1140
   End
   Begin VB.Label Label10 
      Caption         =   "Reimpresión Facturas Cliente "
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
      Left            =   330
      TabIndex        =   9
      Top             =   480
      Width           =   5655
   End
End
Attribute VB_Name = "frmFCliReImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 322


Dim Tabla As String
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim codtipom As String
Dim cadSelect As String
Dim indCodigo As Long

Dim cadNombreRPT As String
Dim cadTitulo As String
Dim ConSubInforme As Boolean
Dim conSubRPT As Boolean

Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientes ' Clientes
Attribute frmCli.VB_VarHelpID = -1

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub


Private Sub cmdAceptar_Click()
Dim Codigo As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal

    InicializarVbles
    
    Tabla = "scafaccli"
    If Check1.Value = 1 Then
        codtipom = "FRN"
    Else
        codtipom = "FAC"
    End If
    
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 1
    
    'Desde/Hasta cliente
    '---------------------------------------------
    If txtCodigo(0).Text <> "" Or txtCodigo(1).Text <> "" Then
        Codigo = "{" & Tabla & ".codclien}"
        If Not PonerDesdeHasta(Codigo, "N", 0, 1, "pDHCliente=""") Then Exit Sub
    End If
    
    'Desde/Hasta numero de FACTURA
    '---------------------------------------------
    If txtCodigo(36).Text <> "" Or txtCodigo(37).Text <> "" Then
        Codigo = "{" & Tabla & ".numfactu}"
        If Not PonerDesdeHasta(Codigo, "N", 36, 37, "") Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(85).Text <> "" Or txtCodigo(86).Text <> "" Then
        Codigo = "{" & Tabla & ".fecfactu}"
        If Not PonerDesdeHasta(Codigo, "F", 85, 86, "") Then Exit Sub
    End If
    
    If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".codtipom} = """ & codtipom & """") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".codtipom} = """ & codtipom & """") Then Exit Sub
    
    If CBool(Me.chk_duplicado.Value) Then
        cadParam = cadParam & "pDuplicado=1|"
    Else
        cadParam = cadParam & "pDuplicado=0|"
    End If
    numParam = numParam + 1
    
    '[Monica]15/02/2019: solo los clientes que no tengan facturacion electronica
    If Check2.Value = 1 Then
        If Not AnyadirAFormula(cadFormula, "{scliente.tasareciclado} = 0") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "{scliente.tasareciclado} = 0") Then Exit Sub
    End If
    
    Tabla = Tabla & " inner join scliente on scafaccli.codclien = scliente.codclien"
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub

    If Check1.Value = 1 Then
        indRPT = 54 'Impresion de facturas rectificativas de clientes
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, False, "") Then Exit Sub
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
        'Nombre fichero .rpt a Imprimir
        cadTitulo = "Reimpresión Facturas Rectificativas Cliente"
        ConSubInforme = True
        
        conSubRPT = ConSubInforme
    
    Else
        '[Monica]31/03/2014: en el caso de teletaxi pedimos si imprime o no detalle
        '[Monica]19/02/2018: Entra Cordoba
            '[Monica]19/11/2018: Entra Sevilla
        If (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3) And Check1.Value = 0 Then
            If MsgBox("¿ Desea imprimir el detalle de servicios ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                cadParam = cadParam & "pDetalle=0|"
            Else
                cadParam = cadParam & "pDetalle=1|"
            End If
            numParam = numParam + 1
        End If
        'hasta aquí
    
    
        indRPT = 52 'Impresion de facturas de clientes
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, False, "") Then Exit Sub
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
        'Nombre fichero .rpt a Imprimir
        cadTitulo = "Reimpresión de Facturas a Cliente"
        ConSubInforme = True
        
        conSubRPT = ConSubInforme
    End If
    LlamarImprimir True
End Sub

Private Sub LlamarImprimir(duplicado As Boolean)
    With frmImprimir
        .Titulo = cadTitulo
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        'El nombre es el del documento
        .NombreRPT = cadNombreRPT
        .Opcion = 101
        .ConSubInforme = ConSubInforme
        .Show vbModal
    End With
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    cadFormula = ""
    numParam = 0
    cadParam = ""

    PonerFoco txtCodigo(0)

End Sub

Private Sub Form_Load()
Dim i As Integer
    'Icono del form
    Me.Icon = frmppal.Icon
    

'    CargarComboAnyo
'    Combo2.Text = Year(Date)
'    CalcularFacturas True
    For i = 0 To Me.imgBuscarOfer.Count - 1
        Me.imgBuscarOfer(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    For i = 23 To 24
        Me.imgFecha(i).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next i
    
End Sub

Private Function DatosOk() As Boolean
Dim encontrado As String
Dim Codigo As String

    DatosOk = True
    
End Function


'Private Sub CargarComboAnyo()
'Dim i As Byte
'Dim anyo As Long
'
'    anyo = Year(Date)
'    For i = 0 To 4
'        Combo2.AddItem anyo
'        Combo2.ItemData(Combo2.NewIndex) = i
'        anyo = anyo - 1
'    Next i
'
'End Sub
'
'Private Sub Option1_Click()
'    If Option1.Value = True Then CalcularFacturas True
'End Sub
'
'Private Sub Option2_Click()
'    If Option2.Value = True Then CalcularFacturas False
'End Sub

'Private Sub CalcularFacturas(cliente As Boolean)
'Dim cad As String
'
'If cliente Then
'    cad = "select min(numfactu) from scafaccli where year(fecfactu)=" & Combo2.Text & " and codtipom='FPC'"
'    cad = ejecutaselect(cad)
'    txtCodigo(36).Text = cad
'    txtCodigo(36).Text = Format(txtCodigo(36).Text, "000000")
'    cad = "select max(numfactu) from scafaccli where year(fecfactu)=" & Combo2.Text & " and codtipom='FPC'"
'    cad = ejecutaselect(cad)
'    txtCodigo(0).Text = cad
'    txtCodigo(0).Text = Format(txtCodigo(0).Text, "000000")
'Else
'    cad = "select min(numfactu) from sfactusoc where year(fecfactu)=" & Combo2.Text & " and codtipom='FPS'"
'    cad = ejecutaselect(cad)
'    txtCodigo(36).Text = cad
'    txtCodigo(36).Text = Format(txtCodigo(36).Text, "000000")
'    cad = "select max(numfactu) from sfactusoc where year(fecfactu)=" & Combo2.Text & " and codtipom='FPS'"
'    cad = ejecutaselect(cad)
'    txtCodigo(0).Text = cad
'    txtCodigo(0).Text = Format(txtCodigo(0).Text, "000000")
'End If
'
'End Sub

Private Function ejecutaselect(CADENA As String) As String
Dim Rs As Recordset
Dim C As String

ejecutaselect = ""
Set Rs = New ADODB.Recordset
Rs.Open CADENA, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not Rs.EOF Then
    If Not IsNull((Rs.Fields(0))) Then
        C = Rs.Fields(0)
    Else
        C = 0
    End If
End If
Rs.Close
Set Rs = Nothing
ejecutaselect = C


End Function

Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
'            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
'            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 0, 1 ' cliente
            indCodigo = Index
            
            Set frmCli = New frmFacClientes
            frmCli.DeConsulta = True
            frmCli.DatosADevolverBusqueda = "0|1|"
            frmCli.Show vbModal
            Set frmCli = Nothing
        
        End Select
    PonerFoco txtCodigo(indCodigo)


End Sub

Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 23, 24 'fechas de factura
            indCodigo = Index + 62
   End Select
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
End Sub

'Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
'On Error Resume Next
'    If txtCodigo(indD).Text <> "" Then
'        cad = cad & "desde " & txtCodigo(indD).Text
'        If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
'    End If
'    If txtCodigo(indH).Text <> "" Then
'        cad = cad & "  hasta " & txtCodigo(indH).Text
'        If txtNombre(indH).Text <> "" Then cad = cad & " - " & txtNombre(indH).Text
'    End If
'    AnyadirParametroDH = cad
'    If Err.Number <> 0 Then Err.Clear
'End Function

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Tabla As String
Dim codCampo As String, NomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    EsNomCod = False
    TipCampo = "N" 'Casi todos son numericos
    
    Select Case Index
        Case 0, 1 'clientes
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtnombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "scliente", "nomclien", "codclien", "N")
            End If
        
        Case 85, 86  'FECHA Desde Hasta
            If txtCodigo(Index).Text = "" Then Exit Sub
            PonerFormatoFecha txtCodigo(Index)
            
        Case 36, 37 'Nº de FACTURA
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
    End Select
    
End Sub


