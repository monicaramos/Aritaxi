VERSION 5.00
Begin VB.Form frmPubliReImp 
   Caption         =   "Informes"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk_duplicado 
      Caption         =   "Duplicado"
      Height          =   375
      Left            =   330
      TabIndex        =   16
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   85
      Left            =   1710
      MaxLength       =   10
      TabIndex        =   6
      Top             =   3510
      Width           =   1215
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   86
      Left            =   3975
      MaxLength       =   10
      TabIndex        =   7
      Top             =   3525
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   4680
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   4680
      Width           =   1035
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   36
      Left            =   1710
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "codtalle|N|N|||scaord|codtalle|00|S|"
      Top             =   2490
      Width           =   1215
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   37
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "codtalle|N|N|||scaord|codtalle|00|S|"
      Top             =   2490
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   4575
      Begin VB.OptionButton Option2 
         Caption         =   "Socios"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clientes"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   23
      Left            =   1440
      Picture         =   "frmPubliReImp.frx":0000
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
      Left            =   900
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
      Left            =   3210
      TabIndex        =   13
      Top             =   3540
      Width           =   420
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   24
      Left            =   3690
      Picture         =   "frmPubliReImp.frx":008B
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
      Left            =   930
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
      Left            =   3270
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
      Caption         =   "Reimpresión Facturas Publicidad"
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
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
End
Attribute VB_Name = "frmPubliReImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tabla As String
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Integer
Dim codtipom As String
Dim cadSelect As String
Dim indCodigo As Long

Dim kCampo As Integer



Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub


Private Sub cmdAceptar_Click()
Dim Codigo As String

    InicializarVbles
    
    If Option1.Value Then
        Tabla = "scafaccli"
        codtipom = "FPC"
    End If
    If Option2.Value Then
        Tabla = "sfactusoc"
        codtipom = "FPS"
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
    
'    If CBool(Me.chk_duplicado.Value) Then
'        cadParam = "pDuplicado=1|"
'    Else
'        cadParam = "pDuplicado=0|"
'    End If
'
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub

    LlamarImprimir True
End Sub

Private Sub LlamarImprimir(duplicado As Boolean)
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        cadParam = cadParam & "pDuplicado= " & Abs(duplicado) & " |"
        numParam = 2
        
        With frmImprimir
        .Titulo = "Impresión de Facturas Publicidad"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        'El nombre es el del documento
        If Option1.Value Then
            .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "47", "N")
             '------ > Listado 47 = rFacPubli.rpt
        Else
            .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "48", "N")
            '------ > Listado 48 = rFacPubliSoc.rpt
        End If
        .Opcion = 101
        .ConSubInforme = True
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

    PonerFoco txtCodigo(36)

End Sub

Private Sub Form_Load()

    'Icono del form
    Me.Icon = frmPpal.Icon
    
    For kCampo = 23 To 24
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo

'    CargarComboAnyo
'    Combo2.Text = Year(Date)
'    CalcularFacturas True

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
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
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
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
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
        Case 85, 86  'FECHA Desde Hasta
            If txtCodigo(Index).Text = "" Then Exit Sub
            PonerFormatoFecha txtCodigo(Index)
            
        Case 36, 37 'Nº de FACTURA
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
    End Select
    
End Sub


