VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLiqDeshacerFac 
   Caption         =   "Informes"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcodigo 
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
      IMEMode         =   3  'DISABLE
      Index           =   8
      Left            =   2460
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   16
      Tag             =   "admon"
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Frame FrameProgress 
      Height          =   915
      Left            =   270
      TabIndex        =   12
      Top             =   5310
      Visible         =   0   'False
      Width           =   5835
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgess 
         Caption         =   "Deshaciendo última Liquidación:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   135
         Width           =   4215
      End
      Begin VB.Label lblProgess 
         Caption         =   "Iniciando el proceso ..."
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   350
         Width           =   4335
      End
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   1470
      TabIndex        =   2
      Top             =   4470
      Width           =   1005
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
      TabIndex        =   7
      Top             =   3045
      Width           =   3735
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
      Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
      Top             =   3045
      Width           =   855
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
      TabIndex        =   6
      Top             =   3450
      Width           =   3735
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
      Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
      Top             =   3450
      Width           =   855
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
      Left            =   3750
      TabIndex        =   3
      Top             =   4860
      Width           =   1035
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
      Left            =   4980
      TabIndex        =   4
      Top             =   4860
      Width           =   1035
   End
   Begin VB.Label Label6 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Index           =   1
      Left            =   1275
      TabIndex        =   19
      Top             =   1800
      Width           =   2235
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Actualiza contadores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   0
      Left            =   90
      TabIndex        =   18
      Top             =   1410
      Width           =   5595
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Este proceso borra la última liquidación a Socio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   195
      TabIndex        =   17
      Top             =   1080
      Width           =   5820
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   330
      TabIndex        =   11
      Top             =   4110
      Width           =   1845
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   1170
      ToolTipText     =   "Buscar fecha"
      Top             =   4470
      Width           =   240
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Index           =   0
      Left            =   1170
      Top             =   3045
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "V Socio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   330
      TabIndex        =   10
      Top             =   2700
      Width           =   750
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
      TabIndex        =   9
      Top             =   3045
      Width           =   600
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Index           =   1
      Left            =   1170
      Top             =   3450
      Width           =   240
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
      TabIndex        =   8
      Top             =   3450
      Width           =   570
   End
   Begin VB.Label Label10 
      Caption         =   "Deshacer Liquidación Socios"
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
      TabIndex        =   5
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "frmLiqDeshacerFac"
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
Private WithEvents frmFP As frmFacFormasPago ' formas de pago
Attribute frmFP.VB_VarHelpID = -1

' Importes para la Grabacion de Cabecera de Facturas de Socio
Dim TotalFac As Currency
Dim TotalLiq As Currency
Dim BaseImpo As Currency
Dim BaseReten As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim vPorcIva As String
Dim PorceIVA As Currency

Dim tipoMov As String
Dim codSocio As String

Dim kCampo As Integer


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Sub cmdAceptar_Click()
Dim Codigo As String
Dim FecFac As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    If Not DatosOk Then Exit Sub
    
    InicializarVbles
    
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 1
   
    'Desde/Hasta numero de V
    '---------------------------------------------
    If txtcodigo(0).Text <> "" Or txtcodigo(1).Text <> "" Then
        Codigo = "{" & Tabla & ".numeruve}"
        If Not PonerDesdeHasta(Codigo, "N", 0, 1, "pDHUve=""") Then Exit Sub
    End If
    
    ' las facturas correspondientes a la fecha de liquidacion que se le indique
    If Not AnyadirAFormula(cadSelect, "sfactusoc.fecfactu = " & DBSet(txtcodigo(2).Text, "F")) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "sfactusoc.codtipom = 'FLI'") Then Exit Sub
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub

    If HaySociosConLiquidacionesPosteriores(Tabla, cadSelect) Then
        ActivarCLAVE
        Exit Sub
    End If

    If HayFacturasContabilizadas(Tabla, cadSelect) Then
        ActivarCLAVE
        Exit Sub
    End If

    ' proceso de deshacer liquidacion a socios
    If ProcesoDeshacerLiquidacionSocio(cadSelect, txtcodigo(2).Text) Then
        MsgBox "Proceso realizado correctamente.", vbExclamation


        cmdCancelar_Click
    End If

End Sub

Private Function HaySociosConLiquidacionesPosteriores(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim vSocio As CSocio
Dim CadSocios As String
Dim RS As ADODB.Recordset

    On Error GoTo eHaySociosConLiquidacionesPosteriores


    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select codsocio, numfactu FROM " & QuitarCaracterACadena(cTabla, "_1")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        cWhere = Replace(cWhere, "[", "(")
        cWhere = Replace(cWhere, "]", ")")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadSocios = ""
    
    While Not RS.EOF
    
        Set vSocio = New CSocio
        If vSocio.LeerDatos(CStr(RS!codSocio)) Then
            If DBLet(RS!NumFactu, "N") <> (vSocio.ConseguirContador("FLI") - 1) Then
                CadSocios = CadSocios & Format(vSocio.Codigo, "000000") & ","
            End If
        End If
        Set vSocio = Nothing
    
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    If CadSocios <> "" Then
        ' quitamos la ultima coma
        CadSocios = Mid(CadSocios, 1, Len(CadSocios) - 1)
        MsgBox "Esta liquidación no se corresponde con la última factura de liquidación de los siguiente socios: " & vbCrLf & vbCrLf & CadSocios, vbExclamation
        HaySociosConLiquidacionesPosteriores = True
    Else
        HaySociosConLiquidacionesPosteriores = False
    End If
    
    Exit Function
    
eHaySociosConLiquidacionesPosteriores:
    MuestraError Err.Number, "Comprobando si hay socios con liquidaciones posteriores", Err.Description
    
    HaySociosConLiquidacionesPosteriores = True
End Function


Private Function HayFacturasContabilizadas(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim vSocio As CSocio
Dim CadSocios As String
Dim RS As ADODB.Recordset

    On Error GoTo eHayFacturasContabilizadas


    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        cWhere = Replace(cWhere, "[", "(")
        cWhere = Replace(cWhere, "]", ")")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Sql = Sql & " and intconta = 1    "
    
    HayFacturasContabilizadas = (DevuelveValor(Sql) > 0)
    
    If HayFacturasContabilizadas Then
        MsgBox "Hay facturas contabilizadas. No podemos deshacer proceso de liquidación", vbExclamation
    End If
    
    Exit Function
    
eHayFacturasContabilizadas:
    MuestraError Err.Number, "Comprobando si hay socios con liquidaciones posteriores", Err.Description
    
    HayFacturasContabilizadas = True
End Function





Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    cadFormula = ""
    numParam = 0
    cadParam = ""

    PonerFoco txtcodigo(8)
End Sub

Private Sub Form_Load()
    'Icono del form
    Me.Icon = frmPpal.Icon
    
    Tabla = "sfactusoc"
    
    For kCampo = 0 To Me.imgBuscarOfer.Count - 1
        Me.imgBuscarOfer(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    
    
    ActivarCLAVE
End Sub

Private Function DatosOk() As Boolean
Dim encontrado As String
Dim Codigo As String

    DatosOk = True
    If txtcodigo(2).Text = "" Then
        MsgBox "Debe introducir obligatoriamente la fecha de liquidación.", vbExclamation
        DatosOk = False
        Exit Function
    End If

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
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
            numParam = numParam + 1
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
        
        Case 1 ' forma de pago
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0|1|"
            frmFP.Show vbModal
            Set frmFP = Nothing
    
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
        Case 0
            indCodigo = 3
   End Select
   
   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtcodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)
End Sub

Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    If txtcodigo(indD).Text <> "" Then
        cad = cad & "desde " & txtcodigo(indD).Text
        If txtnombre(indD).Text <> "" Then cad = cad & " - " & txtnombre(indD).Text
    End If
    If txtcodigo(indH).Text <> "" Then
        cad = cad & "  hasta " & txtcodigo(indH).Text
        If txtnombre(indH).Text <> "" Then cad = cad & " - " & txtnombre(indH).Text
    End If
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
Private Sub KEYpress(KeyAscii As Integer)
'Dim cerrar As Boolean
'
'    KEYpressGnral KeyAscii, 2, cerrar
'    If cerrar Then Unload Me
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If

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
    
    
    Select Case Index
        Case 0, 1 'V Socio
            PonerFormatoEntero txtcodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sclien", "nomclien", "numeruve", "N")
            
        Case 2 ' fecha de liquidacion
            PonerFormatoFecha txtcodigo(Index)
            
        Case 8
            If txtcodigo(Index).Text = "" Then Exit Sub
            If Trim(txtcodigo(Index).Text) <> Trim(txtcodigo(Index).Tag) Then
                MsgBox "    ACCESO DENEGADO    ", vbExclamation
                txtcodigo(Index).Text = ""
                PonerFoco txtcodigo(Index)
            Else
                DesactivarCLAVE
                PonerFoco txtcodigo(0)
            End If
    End Select
End Sub

Private Sub ActivarCLAVE()
Dim i As Integer
    
    For i = 0 To 2
        txtcodigo(i).Enabled = False
    Next i
    txtcodigo(8).Enabled = True
    For i = 0 To 1
        imgBuscarOfer(i).Enabled = False
        imgBuscarOfer(i).visible = False
    Next i
    
    imgFecha(0).Enabled = False
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = True
    
    txtcodigo(8).Text = ""
    PonerFoco txtcodigo(8)
End Sub

Private Sub DesactivarCLAVE()
Dim i As Integer

    For i = 0 To 2
        txtcodigo(i).Enabled = True
    Next i
    txtcodigo(8).Enabled = False
    For i = 0 To 1
        imgBuscarOfer(i).Enabled = True
        imgBuscarOfer(i).visible = True
    Next i
    imgFecha(0).Enabled = True
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = False
End Sub

Private Function ProcesoDeshacerLiquidacionSocio(cadWHERE As String, FecFac As String) As Boolean
'Desde Historico de facturas deshace el proceso de liquidacion
Dim RSalb As ADODB.Recordset 'Ordenados por: codsocio
Dim b As Boolean
Dim Sql As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim antSocio As Long
Dim actSocio As Long
Dim antV As Long
Dim actV As Long

'Concatenamos todas las facturas generadas para listarlas en el informe
Dim ListFactu As String
Dim vSocio As CSocio
Dim Inc As Integer
Dim condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura

Dim nTotal As Long
Dim SQL2 As String
Dim NumFactu As Long
Dim devuelve As Long
Dim Existe As Boolean


    On Error GoTo ETraspasoAlbFac

    ProcesoDeshacerLiquidacionSocio = False

    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
    
    conn.BeginTrans

    tipoMov = "FLI"

    'comprobamos que no haya nadie desfacturando
    DesBloqueoManual ("DESLIQ") 'facturas de liquidacion
    If Not BloqueoManual("DESLIQ", "1") Then
        MsgBox "No se puede deshacer liquidación. Hay otro usuario realizando el proceso.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    
    'Bloqueamos todos los registros de facturas que vamos a eliminar (solo cabeceras)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    If Not BloqueaRegistro("sfactusoc", cadWHERE) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("DESLIQ")
        Exit Function
    End If
    
    Sql = "select numfactu, codsocio from sfactusoc where " & cadWHERE
    
    nTotal = TotalRegistrosConsulta(Sql)
    PB1.Max = nTotal
    
    FrameProgress.visible = True
    
    Set RSalb = New ADODB.Recordset
    RSalb.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    b = True
    
    While Not RSalb.EOF And b
        IncrementarProgresNew PB1, 1
        
        b = DesmarcarLLamadas(tipoMov, CStr(RSalb!NumFactu), FecFac, RSalb!codSocio)
        
        If b Then
            b = BorrarLineasFactura(tipoMov, CStr(RSalb!NumFactu), FecFac, RSalb!codSocio)
        End If
        
        If b Then
            b = EliminarFactura(tipoMov, CStr(RSalb!NumFactu), FecFac, RSalb!codSocio)
        End If
        
        
        Set vSocio = New CSocio
        
        b = vSocio.DevolverContador(RSalb!codSocio, RSalb!NumFactu, tipoMov) = 1
        
        Set vSocio = Nothing
    
        RSalb.MoveNext
    Wend
    
    Set RSalb = Nothing
    
ETraspasoAlbFac:
    If Err.Number <> 0 Or Not b Then
        If Err.Number <> 0 Then MuestraError Err.Number, "Deshacer Liquidación Socio:", Err.Description
        conn.RollbackTrans
        ProcesoDeshacerLiquidacionSocio = False
    Else
        conn.CommitTrans
        ProcesoDeshacerLiquidacionSocio = True
    End If
    DesBloqueoManual ("DESLIQ")
    TerminaBloquear
End Function


'desmarca la llamada como liquidada por socio
Private Function DesmarcarLLamadas(tipoMov As String, NumFactu As String, FecFac As String, Socio As String) As Boolean
Dim Precio As Currency
Dim Sql As String
Dim SQL2 As String
Dim SqlValues As String
Dim linea As Long
Dim MensError As String
Dim RS As ADODB.Recordset
    
    On Error GoTo eInsertLinea
    
    DesmarcarLLamadas = False
    
    MensError = ""
    
    Sql = "update shilla, sfactusoc_serv set  shilla.liquidadosocio = 0 where "
    Sql = Sql & " sfactusoc_serv.codtipom = " & DBSet(tipoMov, "T")
    Sql = Sql & " and sfactusoc_serv.numfactu = " & DBSet(NumFactu, "N")
    Sql = Sql & " and sfactusoc_serv.codsocio = " & DBSet(Socio, "N")
    Sql = Sql & " and sfactusoc_serv.fecfactu = " & DBSet(FecFac, "F")
    Sql = Sql & " and shilla.fecha = sfactusoc_serv.fecha "
    Sql = Sql & " and shilla.hora = sfactusoc_serv.hora "
    Sql = Sql & " and shilla.numeruve = sfactusoc_serv.numeruve "
    
    conn.Execute Sql
    
    DesmarcarLLamadas = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en Desmarcar LLamadas de la factura del  " & Socio
        MuestraError Err.Number, MensError, Err.descripc
    End If
End Function



'eliminar lineas de la factura
Private Function BorrarLineasFactura(tipoMov As String, NumFactu As String, FecFac As String, Socio As String) As Boolean
Dim Precio As Currency
Dim Sql As String
Dim SQL2 As String
Dim SqlValues As String
Dim linea As Long
Dim MensError As String
Dim RS As ADODB.Recordset
    
    On Error GoTo eInsertLinea
    
    BorrarLineasFactura = False
    
    MensError = ""
    
    Sql = "delete from sfactusoc_serv where "
    Sql = Sql & " sfactusoc_serv.codtipom = " & DBSet(tipoMov, "T")
    Sql = Sql & " and sfactusoc_serv.numfactu = " & DBSet(NumFactu, "N")
    Sql = Sql & " and sfactusoc_serv.codsocio = " & DBSet(Socio, "N")
    Sql = Sql & " and sfactusoc_serv.fecfactu = " & DBSet(FecFac, "F")
    
    conn.Execute Sql
    
    
    BorrarLineasFactura = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en Borrar Lineas de la factura del  " & Socio
        MuestraError Err.Number, MensError, Err.descripc
    End If
End Function


'eliminar la factura
Private Function EliminarFactura(tipoMov As String, NumFactu As String, FecFac As String, Socio As String) As Boolean
Dim Precio As Currency
Dim Sql As String
Dim SQL2 As String
Dim SqlValues As String
Dim linea As Long
Dim MensError As String
Dim RS As ADODB.Recordset
    
    On Error GoTo eInsertLinea
    
    EliminarFactura = False
    
    MensError = ""
    
    Sql = "delete from sfactusoc where "
    Sql = Sql & " sfactusoc.codtipom = " & DBSet(tipoMov, "T")
    Sql = Sql & " and sfactusoc.numfactu = " & DBSet(NumFactu, "N")
    Sql = Sql & " and sfactusoc.codsocio = " & DBSet(Socio, "N")
    Sql = Sql & " and sfactusoc.fecfactu = " & DBSet(FecFac, "F")
    
    conn.Execute Sql
    
    
    EliminarFactura = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en Eliminar Factura  "
        MuestraError Err.Number, MensError, Err.descripc
    End If
End Function


'Actualizar llamadas
Private Function ActualizarLlamadas(Uve As String, cadWHERE As String) As Boolean
Dim Precio As Currency
Dim Sql As String
Dim SQL2 As String
Dim SqlValues As String
Dim linea As Long
Dim MensError As String

    On Error GoTo eInsertLinea
    
    ActualizarLlamadas = False
    
    MensError = ""
    
    Sql = "update shilla set liquidadosocio = 1 where " & cadWHERE & " and numeruve = " & DBSet(Uve, "N")
    
    conn.Execute Sql
    
    ActualizarLlamadas = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la actualización de servicios de la factura del socio NºV " & Uve
        MuestraError Err.Number, MensError, Err.descripc
    End If
End Function


