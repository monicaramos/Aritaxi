VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFCliContaFac 
   Caption         =   "AriTaxi"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2100
      Width           =   1995
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   270
      TabIndex        =   9
      Top             =   2610
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
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
      Index           =   7
      Left            =   4680
      TabIndex        =   7
      Top             =   3660
      Width           =   1135
   End
   Begin VB.CommandButton cmdAceptarRepxDia 
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
      Left            =   3480
      TabIndex        =   6
      Top             =   3660
      Width           =   1135
   End
   Begin VB.TextBox txtCodigo 
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
      Index           =   32
      Left            =   4470
      TabIndex        =   4
      Top             =   1500
      Width           =   1345
   End
   Begin VB.TextBox txtCodigo 
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
      Index           =   31
      Left            =   1680
      TabIndex        =   2
      Top             =   1500
      Width           =   1345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
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
      Index           =   1
      Left            =   300
      TabIndex        =   11
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label lblProgess 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   1
      Left            =   270
      TabIndex        =   10
      Top             =   3270
      Width           =   5535
   End
   Begin VB.Label lblProgess 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   0
      Left            =   270
      TabIndex        =   8
      Top             =   2970
      Width           =   5535
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Contabilizar Facturas Cliente"
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
      Left            =   300
      TabIndex        =   5
      Top             =   360
      Width           =   5055
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   5
      Left            =   4140
      ToolTipText     =   "Buscar Fecha"
      Top             =   1500
      Width           =   240
   End
   Begin VB.Label Label3 
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
      Index           =   29
      Left            =   3390
      TabIndex        =   3
      Top             =   1500
      Width           =   630
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   4
      Left            =   1320
      ToolTipText     =   "Buscar Fecha"
      Top             =   1500
      Width           =   240
   End
   Begin VB.Label Label2 
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
      Left            =   570
      TabIndex        =   1
      Top             =   1500
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de factura:"
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
      Left            =   300
      TabIndex        =   0
      Top             =   1140
      Width           =   1860
   End
End
Attribute VB_Name = "frmFCliContaFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents frmCal As frmCal
Attribute frmCal.VB_VarHelpID = -1

Private Const IdPrograma = 323

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
'-----------------------------------
Private Codigo As String
Private codtipom As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Fecha As Date
Dim DtoPPago As Currency
Dim DtoGnral As Currency
Dim BaseImp As Currency
Dim TotalFac As Currency
Dim AnyoFacPr As Integer





Private Sub cmdAceptarRepxDia_Click()
'Reparaciones por Dia
Dim devuelve As String
Dim param As String
Dim TotalMante As Integer
Dim Rs As ADODB.Recordset
Dim fecha1 As String, fecha2 As String
Dim NomTabla As String
Dim bOk As Boolean


Dim ConexionContaOk As Boolean
Dim CambiaConta As Boolean
' ====

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'A�adir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    param = ""
    
    Select Case Combo1.ListIndex
        Case 0
            codtipom = "FAC"
        Case 1
            codtipom = "FRN"
        Case 2
            codtipom = "FVC"
    End Select
'    codtipom = "FAC"
'    If Check1.Value = 1 Then codtipom = "FRN"

    NomTabla = "scafaccli"
    Codigo = "{scafaccli.fecfactu}"

        
    '===================================================
    '================= FORMULA =========================
    
    '== Cadena para seleccion Desde y Hasta FECHA ==
        'comprobar que se han rellenado los dos campos de fecha
        'sino rellenar con fechaini o fechafin del ejercicio
        'que guardamos en vbles Orden1,Orden2
        
        'fechaini del ejercicio de la conta
        If txtcodigo(31).Text = "" Then txtcodigo(31).Text = Orden1
     
        'fecha fin del ejercicio de la conta
        If txtcodigo(32).Text = "" Then txtcodigo(32).Text = Orden2
     
        'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
        'contabilidad par ello mirar en la BD de la Conta los par�metros
        If Not ComprobarFechasConta(31) Then Exit Sub
        If Not ComprobarFechasConta(32) Then Exit Sub
    
    
    devuelve = CadenaDesdeHasta(txtcodigo(31).Text, txtcodigo(32).Text, Codigo, "F", "Fecha Factura")
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    'Parametro D/H Fecha
    If devuelve <> "" And param <> "" Then
        cadParam = cadParam & AnyadirParametroDH(param, 31, 32) & """|"
        numParam = numParam + 1
    End If
    
    
    '- cadena para select en BDatos
    cadSelect = CadenaDesdeHastaBD(txtcodigo(31).Text, txtcodigo(32).Text, Codigo, "F")
    
    
    '== Cadena para seleccion Desde y Hasta N�Factura ==
  
                
        '- a�adir tipo movimiento a cadena seleccion
            Codigo = "{scafaccli.codtipom}"
            devuelve = Codigo & "=" & DBSet(codtipom, "T")
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
                   

    '===================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
'    cadSelect = CadenaDesdeHastaBD(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F")
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & NomTabla & ".intconta=0 "
        
        
    
    If Not HayRegParaInforme(NomTabla, cadSelect) Then Exit Sub
    
    
    devuelve = "CLI"

        CambiaConta = False
        ConexionContaOk = True
        
           
        If ConexionContaOk Then
        ' ====
            '------------------------------------------------------------------------------
            '  LOG de acciones.                      5: Facturas compras
            Set LOG = New cLOG
            
            
            devuelve = "Contabilizar facturas " & devuelve & ":" & vbCrLf & NomTabla & vbCrLf & cadSelect
            LOG.Insertar 5, vUsu, devuelve
            Set LOG = Nothing
            '-----------------------------------------------------------------------------
            
        
            bOk = ContabilizarFacturas(NomTabla, cadSelect)
        
            TerminaBloquear
            'Eliminar la tabla TMP
            BorrarTMPFacturas
            'Desbloqueamos ya no estamos contabilizando facturas
            DesBloqueoManual ("FACCLI") 'FACturas CLIente
            If bOk Then Unload Me
        
        End If
        If CambiaConta Then AbrirConexionConta False
        ' ====
   
End Sub

Private Function ComprobarFechasConta(Ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim cad As String
Dim Rs As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtcodigo(Ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            '## LAURA 19/06/2008
'            FechaFin = DBLet(RS!FechaFin, "F") + 365
'            FechaFin = DateAdd("d", 365, DBLet(RS!FechaFin, "F"))
            FechaFin = DateAdd("yyyy", 1, DBLet(Rs!FechaFin, "F"))
            '##
            
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtcodigo(Ind).Text, FechaFin) Then
                 cad = "El per�odo de contabilizaci�n debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 cad = cad & "    Desde: " & FechaIni & vbCrLf
                 cad = cad & "    Hasta: " & FechaFin
                 MsgBox cad, vbExclamation
                 txtcodigo(Ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function

Private Function ContabilizarFacturas(cadTabla As String, cadWHERE As String) As Boolean
'Contabiliza Facturas de Clientes o de Socios
Dim Sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste2 As Byte

        '0.- Si devuelve la funcion el 0 habra CC sin confgurar en trabaja
        '1.- Todos los CC son el mismo
        '2.- Mas de un CC. Hay que agrupar

    ContabilizarFacturas = False

    Sql = "FACCLI" 'contabilizar facturas de cliente

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2
     If txtcodigo(31).Text = "" Then
        txtcodigo(31).Text = vEmpresa.FechaIni  'fechaini del ejercicio de la conta
     End If

     If txtcodigo(32).Text = "" Then
        txtcodigo(32).Text = vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
     End If


     
     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los par�metros
     If Not ComprobarFechasConta(32) Then Exit Function
     
     

    'La comprobacion solo lo hago para facturas nuestras, ya que mas adelante
    'el programa hara cdate(text1(31) cuando contabilice las facturas y dara error de tipos
    If Me.txtcodigo(31).Text = "" Then
        MsgBox "Fecha inicio incorrecta", vbExclamation
        Exit Function
    End If
    
    
    
    'comprobar si existen en Aritaxi facturas anteriores al periodo solicitado
    'sin contabilizar.
    
    If Me.txtcodigo(31).Text <> "" Then 'anteriores a fechadesde
        Sql = "SELECT COUNT(*) FROM " & cadTabla
        Sql = Sql & " WHERE codtipom=" & DBSet(codtipom, "T") & " and "
        Sql = Sql & "fecfactu <"
        Sql = Sql & DBSet(txtcodigo(31), "F") & " AND intconta=0 "
        
        
        
        If RegistrosAListar(Sql) > 0 Then
            If MsgBox("Hay Facturas anteriores sin contabilizar. " & "�Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                Exit Function
            End If
        End If
    End If
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
        
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    b = CrearTMPFacturas(cadTabla, cadWHERE)
    If Not b Then Exit Function
            
            
    Sql = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla
    Sql = Sql & ".codtipom=tmpFactu.codtipom AND "
    Sql = Sql & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    If Not BloqueaRegistro(Sql, cadWHERE) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
            
            
            
    Me.lblProgess(0).Caption = "Comprobaciones: "
    CargarProgres Me.ProgressBar1, 100
        
        
    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Aritaxi
    '--------------------------------------------------------------------------
    IncrementarProgres Me.ProgressBar1, 10
    Me.lblProgess(1).Caption = "Comprobando letras de serie ..."
    b = ComprobarLetraSerie(cadTabla)
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not b Then Exit Function
    
    
    'comprobar que no haya N� FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando N� Facturas en contabilidad ..."
    Sql = "anofaccl>=" & Year(txtcodigo(31).Text) & " AND anofaccl<= " & Year(txtcodigo(32).Text)
    b = ComprobarNumFacturas_new(cadTabla, Sql)
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Function
    
    
    'comprobar que todas las CUENTAS de los distintos clientes que vamos a
    'contabilizar existen en la Conta: sclien.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    b = ComprobarCtaContable_new(cadTabla, 1)
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Function


    'comprobar que todas las CUENTAS de venta de la familia de los articulos que vamos a
    'contabilizar existen en la Conta: sfamia.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    b = ComprobarCtaContable_new(cadTabla, 2)
    
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Function
    
    
    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: scafac.codigiv1,codigiv2,codigiv3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    b = ComprobarTiposIVA(cadTabla)
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not b Then Exit Function
    
    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de sfamia.ctaventa empiezan por el digito
    'de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    
    'No tiene analitica, NO agrupamos por codtraba
    CCoste2 = 0

    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    
    
    
    Me.lblProgess(1).Caption = "Fechas contabilizacion"
    Me.lblProgess(1).Refresh
    If cadTabla = "scafaccli" Then
        b = NuevasComprobacionesContabilizacion(False, cadWHERE)
    Else
        b = NuevasComprobacionesContabilizacion(True, cadWHERE)
    End If
    If Not b Then Exit Function
    
    
    
    
    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgess(0).Caption = "Contabilizar Facturas: "
    CargarProgres Me.ProgressBar1, 10
    Me.lblProgess(1).Caption = "Insertando Facturas en Contabilidad..."
       
    
    
    
    
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar facturas: " & vbCrLf & cadTabla & vbCrLf & cadWHERE
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    
    '---- Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTabla)
    
    '---- Pasar las Facturas a la Contabilidad
    b = PasarFacturasAContab(cadTabla, CCoste2)
    
    
    
    '---- Mostrar ListView de posibles errores (si hay)
    If Not b Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensajes.OpcionMensaje = 10
            frmMensajes.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        'Para la facturacion de TICKTS agrupada NO mostramos el mensaje de OK
            MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If
    
    '------------------------------------------------------
    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact
    ContabilizarFacturas = True
End Function

Private Function ComprobarCtaContable_local(cadTabla As String, cadWHERE As String) As Boolean
Dim cContabF As CControlFacturaContab
Dim QueCuentasSon As String
Dim CtaBloq As Collection
Dim cuenta As String
Dim Socio As String
Dim FormatSocio As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Ic As Integer
'Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim SQLcuentas As String
    
    On Error GoTo ECompCta
    ComprobarCtaContable_local = False
    cadG = ""
    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG
    Sql = "select codsocio from " & cadTabla & " where " & cadWHERE
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    b = True
    QueCuentasSon = ""
    While Not Rs.EOF And b
        Socio = Rs!codSocio
        FormatSocio = String((vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior), "0")
        cuenta = Trim(vParamAplic.Raiz_Cta_Soc_publi & Format(Socio, FormatSocio))
        
        Sql = SQLcuentas & " AND codmacta= " & DBSet(cuenta, "T")
        
        'Para comporbar si estan bloqueadas
        QueCuentasSon = QueCuentasSon & ", '" & cuenta & "'"
        
        
        If Not (RegistrosAListar(Sql, conConta) > 0) Then
        'si no lo encuentra
            b = False 'no encontrado
            Sql = cuenta & " del Cliente " & Format(Rs!codSocio, "000000")
        End If
        
        Rs.MoveNext
    Wend
        If Not b Then
            Sql = "No existe la cta contable " & Sql
            Sql = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & Sql
            
            MsgBox Sql, vbExclamation
            ComprobarCtaContable_local = False
        Else
            Sql = ""
            If QueCuentasSon <> "" Then
                QueCuentasSon = Mid(QueCuentasSon, 2)
                Set cContabF = New CControlFacturaContab
                cContabF.CuentasBloqueadas ConnConta, QueCuentasSon, Now, CtaBloq
                If CtaBloq.Count > 0 Then
                    'EXISTEN CUENTAS BLOQUEADAS
                    For Ic = 1 To CtaBloq.Count
                        QueCuentasSon = CtaBloq.Item(Ic)
                        Sql = Sql & RecuperaValor(QueCuentasSon, 1) & "   " & RecuperaValor(QueCuentasSon, 2) & vbCrLf
                    Next
                    Sql = "Cuentas bloqueadas en contabilidad: " & vbCrLf & String(30, "=") & vbCrLf & Sql
                    MsgBox Sql, vbExclamation
                Else
                    Sql = ""
                End If
                Set cContabF = Nothing
            End If
            If Sql = "" Then
                ComprobarCtaContable_local = True
            Else
                ComprobarCtaContable_local = False
            End If
        End If
        
        
        
        
    Exit Function
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function

Private Function NuevasComprobacionesContabilizacion(Proveedores As Boolean, ByVal Sql As String) As Boolean
Dim RT As ADODB.Recordset
Dim C As String
Dim F As Date
Dim Fin As Boolean
Dim ComprobacionFechaMenor As Boolean
Dim cControlFra As CControlFacturaContab
    
    On Error GoTo ENuevasComprobacionesContabilizacion
    NuevasComprobacionesContabilizacion = False
    
    
    
    Set cControlFra = New CControlFacturaContab
        'Tenemos que comprobar la fecha factura
    Set RT = New ADODB.Recordset
    ComprobacionFechaMenor = False
    If Proveedores Then
        C = "Select fecfactu from sfactusoc WHERE " & Sql
        C = C & " GROUP BY fecfactu ORDER BY fecfactu"
    Else
        C = "Select fecfactu from scafaccli WHERE " & Sql
        C = C & " GROUP BY fecfactu ORDER BY fecfactu"
    End If
    
    RT.Open C, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Fin = False
    While Not Fin
        F = RT.Fields(0)
        C = cControlFra.FechaCorrectaContabilizazion(ConnConta, F)
        If C <> "" Then
            Fin = True
        Else
            C = cControlFra.FechaCorrectaIVA(ConnConta, F)
            If C <> "" Then
                Fin = True
            End If
        End If
        RT.MoveNext
        If Not Fin Then Fin = RT.EOF
    Wend
    RT.Close
    
    If C <> "" Then
        C = C & "(" & F & ")"
        MsgBox C, vbExclamation
    Else
        NuevasComprobacionesContabilizacion = True
    End If
    
    
ENuevasComprobacionesContabilizacion:
    If Err.Number <> 0 Then MuestraError Err.Number, "Nueva Comprobacion Contabilizacion"
    Set RT = Nothing
    Set cControlFra = Nothing
End Function

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
    conSubRPT = False
    pPdfRpt = ""
End Sub

Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    
     If txtcodigo(indD).Text <> "" Then
        cad = cad & "desde " & txtcodigo(indD).Text
     End If
    If txtcodigo(indH).Text <> "" Then
        cad = cad & "  hasta " & txtcodigo(indH).Text
    End If
    
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function
'Ccoste
'   0: No tendra analitica
'   1: Solo hay un CC que tratar. NO agruparemos por trabajador
'   2: Mas de un CC. Agruparemos por trabajador
Private Function PasarFacturasAContab(cadTabla As String, miCC As Byte) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim I As Integer
Dim NumFactu As Integer
Dim Codigo1 As String
Dim cContaFra As cContabilizarFacturas


    On Error GoTo EPasarFac

    PasarFacturasAContab = False
    
    
    
    '---- Obtener el total de Facturas a Insertar en la contabilidad
    Sql = "SELECT count(*) "
    Sql = Sql & " FROM " & cadTabla & " INNER JOIN tmpFactu "
    Codigo1 = "codtipom"
    Sql = Sql & " ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1
    Sql = Sql & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        NumFactu = Rs.Fields(0)
    Else
        NumFactu = 0
    End If
    Rs.Close
    Set Rs = Nothing


    '------------------------------------------------------------
    Set cContaFra = New cContabilizarFacturas
    
    If Not cContaFra.EstablecerValoresInciales(ConnConta) Then
        'NO ha establecido los valores de la conta.  Le dejaremos seguir, avisando que
        ' obviamente, no va a contabilizar las FRAS
        Sql = "Si continua, las facturas se insertaran en el registro, pero no ser�n contabilizadas" & vbCrLf
        Sql = Sql & "en este momento. Deber�n ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
        Sql = Sql & Space(50) & "�Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    
    
    '-----------------------------------------------------------
    ' Mostraremos para cada factura de PROVEEDOR
    ' que numregis le ha asignado
    Sql = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Sql

    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If NumFactu > 0 Then
    
        Set Rs = New ADODB.Recordset
    
        CargarProgres Me.ProgressBar1, NumFactu
        
        'PreComprobacion de los asientos
        If cContaFra.RealizarContabilizacion Then
            Sql = "Select min(fecfactu) from tmpfactu"
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                If Not cContaFra.PreComprobacionNumeroAsiento(Rs.Fields(0), NumFactu) Then
                    
                    'Para que la ventana siguiente muestr bien el error
                    Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) VALUES ("
                    Sql = Sql & "'',0,'" & Format(Rs.Fields(0), FormatoFecha) & "','Error contadores')"
                    
                    conn.Execute Sql
                    Rs.Close
                    Err.Raise 6, , "Comprobacion numeros asiento"
                End If
            End If
            Rs.Close
        End If
        
        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        Sql = "SELECT * FROM tmpFactu "
            
        Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
        I = 1

        b = True
        
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not Rs.EOF
        
            'Segun sea cli o pro
                Sql = cadTabla & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " AND " & cadTabla & ".numfactu=" & Rs!NumFactu
                Sql = Sql & " and " & cadTabla & ".fecfactu=" & DBSet(Rs!FecFactu, "F")
                If cadTabla = "sfactusoc" Then
                    If PasarFacturaProv_Local(Sql, miCC, vEmpresa.FechaFin, cContaFra) = False And b Then b = False
                Else
                    If PasarFactura(Sql, miCC, False, cContaFra) = False And b Then b = False
                End If
            'Al pasar cada factura al hacer el commit desbloqueamos los registros
            'que teniamos bloqueados y los volvemos a bloquear
            Sql = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
            If Not BloqueaRegistro(Sql, cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu") Then
'                MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'                Screen.MousePointer = vbDefault
'                Exit Sub
            End If
            '----
            
            IncrementarProgres Me.ProgressBar1, 1
            Me.lblProgess(1).Caption = "Insertando Facturas en Contabilidad...   (" & I & " de " & NumFactu & ")"
            Me.Refresh
            I = I + 1
            Rs.MoveNext   'Siguiente factura
        Wend
        
        
        
        Rs.Close
        Set Rs = Nothing
    End If
    
EPasarFac:
    If Err.Number <> 0 Then b = False
    Set cContaFra = Nothing
    If b Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function

Public Function PasarFacturaProv_Local(cadWHERE As String, CodCCost As Byte, FechaFin As String, ByRef vContaFra As cContabilizarFacturas) As Boolean

Dim b As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores
Dim vLlevaRetencion As Boolean
Dim I As Integer

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    vLlevaRetencion = False 'Si llevara retencion me lo devolvera la fucion insertar
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactProv_Local(cadWHERE, cadMen, Mc, FechaFin, vLlevaRetencion, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        
        'Veremos que opcion de CC es la que hay que pasar (agrupar o no agrupar)
        
        '---- Insertar lineas de Factura en la Conta
        b = InsertarLinFact_Local("sfactusoc", cadWHERE, cadMen, vLlevaRetencion, Mc.Contador)
        cadMen = "Insertando Lin. Factura: " & cadMen

'[Monica]18/02/2011: No contabilizamos las facturas
'        If b Then
'            If vContaFra.RealizarContabilizacion Then
'                vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
'            End If
'        End If
        
        If b Then
            '---- Poner intconta=1 en aritaxi.scafac
            b = ActualizarCabFact("sfactusoc", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
        

        
    End If
    
    
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaProv_Local = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaProv_Local = False

        InsertarTMPErrFac cadMen, cadWHERE
        
        'Si es correcto entonces creo una entrada en tmp para luego listar los resultados de
        'la contabilizacion
         If Mc.Contador > 0 Then
            Sql = "DELETE from tmpinformes where codusu = " & vUsu.Codigo & " AND codigo1= " & Mc.Contador
            conn.Execute Sql
        End If
    
    End If
End Function

Private Sub InsertarTMPErrFac(MenError As String, cadWHERE As String)
Dim Sql As String

    On Error Resume Next
    Sql = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    Sql = Sql & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "scafpc", "tmpFactu")
    conn.Execute Sql
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function ActualizarCabFact(cadTabla As String, cadWHERE As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    
    Sql = "UPDATE " & cadTabla & " SET intconta=1 "
    Sql = Sql & " WHERE " & cadWHERE

    conn.Execute Sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFact = False
        cadErr = Err.Description
    Else
        ActualizarCabFact = True
    End If
End Function

Private Function InsertarLinFact_Local(cadTabla As String, cadWHERE As String, cadErr As String, vLlevaRetencion As Boolean, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim SQL2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim TotImp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim LineaCentroCoste As Boolean
Dim Socio As String
Dim FormatSocio As String
Dim cuenta As String
    'Puede ser que teniendo analitica, la cuenta no sea del grupo 6 o 7 , con lo cual nodebe poner el CC
    'Por si acaso alguna linea no es del grupo venta o grupo compras, no

    On Error GoTo EInLinea
      
        Sql = "SELECT sfactusoc.codsocio,sfactusoc.numfactu,sfactusoc.fecfactu,importel as importe  "

        Sql = Sql & " FROM sfactusoc  "
                
        Sql = Sql & " WHERE "
        
        'si tiene analitica, enlazo por con scafpa
            
        Sql = Sql & cadWHERE
        
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    TotImp = 0
    SQLaux = ""
    Aux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        ImpLinea = Rs!Importe - CCur(CalcularPorcentaje(Rs!Importe, DtoPPago, 2))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(Rs!Importe, DtoGnral, 2))
        '----
        TotImp = TotImp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        SQL2 = ""
        
            Sql = numRegis & "," & Year(Rs.Fields(2)) & "," & I & ","
            'calculo la cuenta
            Socio = Rs!codSocio
            FormatSocio = String((vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior), "0")
            cuenta = Trim(vParamAplic.Raiz_Cta_Soc_publi & Format(Socio, FormatSocio))
            Sql = Sql & DBSet(cuenta, "T")

        
        SQL2 = Sql & "," 'nos guardamos la linea sin el importe por si a la �ltima hay q descontarle para q coincida con total factura
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        
        'CENTRO DE COSTE
        LineaCentroCoste = False
            
        Sql = Sql & ValorNulo
        
        cad = cad & "(" & Sql & ")" & ","
        
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close

    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    
    
    'Facturas clientes. Ver si lleva aportacion al terminal

    Set Rs = Nothing

    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "

        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_Local = False
        cadErr = Err.Description
    Else
        InsertarLinFact_Local = True
    End If
End Function

Private Function InsertarCabFactProv_Local(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, ByRef LlevaRetencion As Boolean, ByRef vCF As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el a�o de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Socio As String
Dim FormatSocio As String
Dim cuenta As String


    On Error GoTo EInsertar
       
    

    Sql = "select * from sfactusoc"
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        Socio = Rs!codSocio
        FormatSocio = String((vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior), "0")
        cuenta = Trim(vParamAplic.Raiz_Cta_Soc_publi & Format(Socio, FormatSocio))
        
        If Mc.ConseguirContador("1", (Rs!FecFactu <= CDate(FechaFin) - 365), True) = 0 Then            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = 0
            DtoGnral = 0
            BaseImp = Rs!BaseIVA1
            TotalFac = Rs!TotalFac
            AnyoFacPr = Year(Rs!FecFactu)
            
            
            
            Nulo2 = "N"
            Nulo3 = "N"
            Sql = ""
            Sql = Mc.Contador & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!NumFactu, "T") & "," & DBSet(cuenta, "T") & ","
            
            Select Case vParamAplic.ObsFactura
            Case 0
                'Vacio
                Sql = Sql & ValorNulo
            Case 1
                'N� Factura
                Sql = Sql & "'" & DevNombreSQL("S/Fra " & Rs!NumFactu) & "'"
            Case 2
                'Fecha integracion
                Sql = Sql & "'" & Format(Now, FormatoFecha) & "'"
            End Select
            Sql = Sql & "," & DBSet(Rs!BaseIVA1, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!porciva1, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            'ANTES era dbset de Rs!totalfac, ahora lo haremos de la variabele totalfac
            Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!codiiva1, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
            
            Nulo2 = ""
            'NULOS
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!FecFactu, "F") & ",0"
            cad = cad & "(" & Sql & ")"
            
            'Insertar en la contabilidad
            Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
            Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
            Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
            
            
            
            
            'Para saber el numreo de registro que le asigna a la factrua
            Sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            Sql = Sql & ",'" & DevNombreSQL(Rs!NumFactu) & " @ " & Format(Rs!FecFactu, "dd/mm/yyyy") & "','" & DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Rs!codSocio, "T") & "'," & Rs!codSocio & ")"
            conn.Execute Sql
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactProv_Local = False
        cadErr = Err.Description
    Else
        InsertarCabFactProv_Local = True
    End If
End Function

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
Dim I As Integer
    txtcodigo(31).Text = Date
    txtcodigo(32).Text = Date
    
    'Icono del form
    Me.Icon = frmppal.Icon
    Me.ProgressBar1.visible = False
    
    For I = 4 To 5
        Me.imgFecha(I).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next I
    
    
    CargaCombo
    
    Combo1.ListIndex = 0
    
End Sub

Private Sub frmCal_Selec(vFecha As Date)
    Fecha = vFecha
End Sub

Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte

    Set frmCal = New frmCal
    frmCal.Fecha = Now


    Select Case Index
        Case 4
            indice = 31
            PonerFormatoFecha txtcodigo(indice)
            If txtcodigo(indice).Text <> "" Then frmCal.Fecha = CDate(txtcodigo(indice).Text)
        Case 5
            indice = 32
            PonerFormatoFecha txtcodigo(indice)
            If txtcodigo(indice).Text <> "" Then frmCal.Fecha = CDate(txtcodigo(indice).Text)
    End Select
    frmCal.Show vbModal
    If IsDate(Fecha) Then
        txtcodigo(indice) = Fecha
    End If
    Set frmCal = Nothing
    PonerFoco txtcodigo(indice)
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub
Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)

    Select Case Index
        Case 31
            If txtcodigo(Index).Text <> "" Then
                PonerFormatoFecha txtcodigo(Index)
            End If
        Case 32
            If txtcodigo(Index).Text <> "" Then
                PonerFormatoFecha txtcodigo(Index)
            End If
    End Select
End Sub


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim I As Byte
    
    Combo1.Clear
    
    '[Monica]11/02/2011: todo tipo de facturas excepto las de liquidacion,publicidad y cuotas de socio
    '                    y las facturas de cliente FAC y FPC
    Sql = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom in ('FAC','FRN','FVC')"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Sql = Rs!nomtipom
        Sql = Replace(Sql, "Factura", "")
        Combo1.AddItem Rs!codtipom & "-" & Sql
        Combo1.ItemData(Combo1.NewIndex) = I
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub

