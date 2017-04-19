VERSION 5.00
Begin VB.Form frmLiqReImp 
   Caption         =   "Informes"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1470
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   1
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1470
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
      Top             =   1455
      Width           =   855
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   0
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1455
      Width           =   3735
   End
   Begin VB.CheckBox chk_duplicado 
      Caption         =   "Duplicado"
      Height          =   375
      Left            =   330
      TabIndex        =   6
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   85
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   4
      Top             =   3510
      Width           =   1215
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   86
      Left            =   3705
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3525
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   4680
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   4680
      Width           =   1035
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   36
      Left            =   1470
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "codtalle|N|N|||scaord|codtalle|00|S|"
      Top             =   2490
      Width           =   1215
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   37
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "codtalle|N|N|||scaord|codtalle|00|S|"
      Top             =   2490
      Width           =   1215
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   690
      TabIndex        =   20
      Top             =   1800
      Width           =   420
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Index           =   1
      Left            =   1170
      Top             =   1830
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   690
      TabIndex        =   19
      Top             =   1455
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "V Socio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   18
      Top             =   1110
      Width           =   600
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   7
      Left            =   330
      TabIndex        =   15
      Top             =   3180
      Width           =   945
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   630
      TabIndex        =   14
      Top             =   3540
      Width           =   450
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   2940
      TabIndex        =   13
      Top             =   3540
      Width           =   420
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   24
      Left            =   3420
      Top             =   3540
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   690
      TabIndex        =   12
      Top             =   2550
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   3030
      TabIndex        =   11
      Top             =   2550
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Nº Factura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   330
      TabIndex        =   10
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label Label10 
      Caption         =   "Reimpresión Facturas Liquidación "
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
Attribute VB_Name = "frmLiqReImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private WithEvents frmMtoV As frmGesVSocio ' V socios
Attribute frmMtoV.VB_VarHelpID = -1

Dim kCampo As Integer

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
    
    Tabla = "sfactusoc"
    codtipom = "FLI"
    
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 1
    
    'Desde/Hasta numero de V
    '---------------------------------------------
    If txtcodigo(0).Text <> "" Or txtcodigo(1).Text <> "" Then
        Codigo = "{" & Tabla & ".numeruve}"
        If Not PonerDesdeHasta(Codigo, "N", 0, 1, "pDHUve=""") Then Exit Sub
    End If
    
    'Desde/Hasta numero de FACTURA
    '---------------------------------------------
    If txtcodigo(36).Text <> "" Or txtcodigo(37).Text <> "" Then
        Codigo = "{" & Tabla & ".numfactu}"
        If Not PonerDesdeHasta(Codigo, "N", 36, 37, "") Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtcodigo(85).Text <> "" Or txtcodigo(86).Text <> "" Then
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
    
    '[Monica]31/03/2014
    If vParamAplic.Cooperativa = 0 Then
        'preguntamos si quiere imprimirlo o no con los servicios
        If MsgBox("¿ Desea imprimir el detalle de servicios ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            cadParam = cadParam & "pDetalle=0|"
        Else
            cadParam = cadParam & "pDetalle=1|"
        End If
        numParam = numParam + 1
    End If
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub

    '[Monica]29/02/2012: en la impresion de factura de liquidacion de socio hemos metido el tmpinformes
    If InsertResumen(cadSelect) Then
        cadFormula = "{tmpinformes.codusu} =" & vUsu.Codigo
    End If


    indRPT = 51 'Impresion de facturas de liquidacion a socios
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, False, "") Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    cadNombreRPT = nomDocu
    'Nombre fichero .rpt a Imprimir
    cadTitulo = "Reimpresión de Facturas Liquidación a Socios"
    ConSubInforme = True
    
    conSubRPT = ConSubInforme

    LlamarImprimir True
End Sub

'Insertar Resumen
Private Function InsertResumen(cadWHERE As String) As Boolean
Dim MensError As String
Dim Sql As String
    
    On Error GoTo eInsertResumen
    
    MensError = ""
    InsertResumen = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
                                        ' codtipom, numfactu, codsocio, fecfactu
    Sql = "insert into tmpinformes (codusu, nombre1, importe1, codigo1, fecha1) select " & vUsu.Codigo & ","
    Sql = Sql & "codtipom, numfactu, codsocio, fecfactu from sfactusoc "
    If cadWHERE <> "" Then Sql = Sql & " where " & cadWHERE
    
    conn.Execute Sql
    
    InsertResumen = True
    
    Exit Function

eInsertResumen:
    MensError = "Error en la inserción de Facturas en Temporal"
    MuestraError Err.Number, MensError
End Function





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

    PonerFoco txtcodigo(0)

End Sub

Private Sub Form_Load()

    'Icono del form
    Me.Icon = frmPpal.Icon
    

'    CargarComboAnyo
'    Combo2.Text = Year(Date)
'    CalcularFacturas True

    For kCampo = 0 To 1
        Me.imgBuscarOfer(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    
    For kCampo = 23 To 24
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    
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
Dim RS As Recordset
Dim C As String

ejecutaselect = ""
Set RS = New ADODB.Recordset
RS.Open CADENA, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RS.EOF Then
    If Not IsNull((RS.Fields(0))) Then
        C = RS.Fields(0)
    Else
        C = 0
    End If
End If
RS.Close
Set RS = Nothing
ejecutaselect = C


End Function

Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, cad) Then Exit Function
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


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtcodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmMtoV_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub


Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 0, 1 ' Vsocio
            indCodigo = Index
            
            Set frmMtoV = New frmGesVSocio
            frmMtoV.DeConsulta = True
            frmMtoV.DatosADevolverBusqueda = "0|1|2|"
            frmMtoV.Show vbModal
            Set frmMtoV = Nothing
        
        End Select
    PonerFoco txtcodigo(indCodigo)


End Sub

Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 23, 24 'fechas de factura
            indCodigo = Index + 62
   End Select
   
   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtcodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)
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
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    EsNomCod = False
    TipCampo = "N" 'Casi todos son numericos
    
    Select Case Index
        Case 0, 1 'V Socio
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sclien", "nomclien", "numeruve", "N")
            End If
        
        Case 85, 86  'FECHA Desde Hasta
            If txtcodigo(Index).Text = "" Then Exit Sub
            PonerFormatoFecha txtcodigo(Index)
            
        Case 36, 37 'Nº de FACTURA
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000000")
            End If
    End Select
    
End Sub


