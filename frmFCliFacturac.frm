VERSION 5.00
Begin VB.Form frmFCliFacturac 
   Caption         =   "Informes"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk_FactHco 
      Caption         =   "Facturación sobre Hco de llamadas"
      Height          =   375
      Left            =   150
      TabIndex        =   24
      Top             =   5070
      Width           =   3435
   End
   Begin VB.CheckBox chk_agrupados 
      Caption         =   "Clientes Agrupados"
      Height          =   375
      Left            =   150
      TabIndex        =   23
      Top             =   4620
      Width           =   2625
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   86
      Left            =   4125
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2460
      Width           =   1215
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   85
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2445
      Width           =   1065
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "Cliente|N|N|||shilla|codclien|000000|S|"
      Top             =   1755
      Width           =   855
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   1
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1755
      Width           =   3735
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Cliente|N|N|||shilla|codclien|000000|S|"
      Top             =   1410
      Width           =   855
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   0
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1410
      Width           =   3735
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2580
      TabIndex        =   15
      Top             =   4050
      Width           =   3615
   End
   Begin VB.TextBox txtcodigo 
      Height          =   285
      Index           =   5
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   6
      Tag             =   "Código de Banco Propio|N|N|0|9999|sbanpr|codbanpr|0000|S|"
      Top             =   4050
      Width           =   735
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5310
      TabIndex        =   8
      Top             =   5520
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4020
      TabIndex        =   7
      Top             =   5520
      Width           =   1035
   End
   Begin VB.TextBox txtcodigo 
      Height          =   285
      Index           =   4
      Left            =   1800
      MaxLength       =   60
      TabIndex        =   5
      Top             =   3510
      Width           =   4425
   End
   Begin VB.TextBox txtcodigo 
      Height          =   285
      Index           =   2
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "99/99/9999"
      Top             =   3030
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   5310
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label4 
      Height          =   195
      Left            =   180
      TabIndex        =   25
      Top             =   5640
      Width           =   3705
   End
   Begin VB.Image imgAyuda 
      Height          =   240
      Index           =   0
      Left            =   6210
      MousePointer    =   4  'Icon
      Tag             =   "-1"
      ToolTipText     =   "Ayuda"
      Top             =   3540
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   24
      Left            =   3840
      Picture         =   "frmFCliFacturac.frx":0000
      Top             =   2475
      Width           =   240
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
      Left            =   3360
      TabIndex        =   22
      Top             =   2505
      Width           =   420
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
      Left            =   1020
      TabIndex        =   21
      Top             =   2475
      Width           =   450
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Servicios"
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
      Left            =   180
      TabIndex        =   20
      Top             =   2250
      Width           =   1305
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   23
      Left            =   1500
      Picture         =   "frmFCliFacturac.frx":008B
      Top             =   2445
      Width           =   240
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
      Left            =   1020
      TabIndex        =   19
      Top             =   1755
      Width           =   420
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   5
      Left            =   1500
      Picture         =   "frmFCliFacturac.frx":0116
      Top             =   1755
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
      Left            =   1020
      TabIndex        =   18
      Top             =   1410
      Width           =   450
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   1500
      Picture         =   "frmFCliFacturac.frx":0218
      Top             =   1410
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   3
      Left            =   1530
      Picture         =   "frmFCliFacturac.frx":031A
      Tag             =   "-1"
      ToolTipText     =   "Buscar cuenta"
      Top             =   4080
      Width           =   240
   End
   Begin VB.Label Label7 
      Caption         =   "Cuenta Cobro"
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
      Height          =   375
      Left            =   150
      TabIndex        =   14
      Top             =   4050
      Width           =   1335
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1530
      Picture         =   "frmFCliFacturac.frx":041C
      Tag             =   "-1"
      ToolTipText     =   "Ver observaciones"
      Top             =   3540
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   2
      Left            =   1530
      Picture         =   "frmFCliFacturac.frx":051E
      ToolTipText     =   "Buscar fecha"
      Top             =   3060
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "Concepto"
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
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Factura"
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
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   3030
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
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
      Height          =   255
      Left            =   210
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Facturación Servicios de Clientes"
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
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   5325
   End
End
Attribute VB_Name = "frmFCliFacturac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public WithEvents frmCal As frmCal
Attribute frmCal.VB_VarHelpID = -1
Public WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Public WithEvents frmFP As frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Public WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios
Attribute frmMtoBancosPro.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim NumFactu As Long
Dim FecFactu As Date
Dim Modo As Byte
Dim cad As String
Dim Codigo As String
Dim CadServicios As String
Dim Salir As Boolean

Dim Tabla As String
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim codtipom As String
Dim cadSelect As String
Dim indCodigo As Long
Dim FacturasaImprimir As String






Private Sub chk_FactHco_Click()
    If chk_FactHco.Value = 1 Then
        Tabla = "shilla"
    Else
        Tabla = "sfactclitr"
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim SQL As String
Dim b As Boolean

        Set miRsAux = New ADODB.Recordset
        Screen.MousePointer = vbHourglass
        
        If Not DatosOK Then Exit Sub
         
         InicializarVbles
         
         cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
         numParam = 1
        
        
        If vParamAplic.Cooperativa = 0 Then
        
            If Not CargarTemporal2 Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            If Not AnyadirAFormula(cadSelect, Tabla & ".codclien in (select codclien from tmpcrmclien where codusu = " & vUsu.Codigo & ")") Then Exit Sub
        Else
            
            'Desde/Hasta numero de cliente
            '---------------------------------------------
            If txtcodigo(0).Text <> "" Or txtcodigo(1).Text <> "" Then
                Codigo = "{" & Tabla & ".codclien}"
                If Not PonerDesdeHasta(Codigo, "N", 0, 1, "pDHUve=""") Then Exit Sub
            End If
        
        End If
           
         
         
         '[Monica]08/09/2011: seleccionamos que servicios vamos a facturar al cliente
         If vParamAplic.Cooperativa = 1 Then
            If (txtcodigo(0).Text = txtcodigo(1).Text) And txtcodigo(0).Text <> "" Then
                Salir = False
            
                Set frmMens = New frmMensajes
                
                frmMens.OpcionMensaje = 22
                frmMens.cadWHERE = "shilla.codclien = " & DBSet(txtcodigo(0).Text, "N")
                frmMens.Show vbModal
                
                Set frmMens = Nothing
            
                If Salir Then
                    cmdCancelar_Click
                    Exit Sub
                End If
            End If
            
         End If
         
         'Cadena para seleccion Desde y Hasta FECHA
         '--------------------------------------------
         If txtcodigo(85).Text <> "" Or txtcodigo(86).Text <> "" Then
             If Tabla = "shilla" Then
                Codigo = "{" & Tabla & ".fecha}"
             Else
                Codigo = "{" & Tabla & ".fecfactu}"
             End If
             If Not PonerDesdeHasta(Codigo, "F", 85, 86, "pDHFecha=""") Then Exit Sub
         End If
         
         ' que no este facturado
         If Tabla = "shilla" Then
            If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".facturadocliente} = 0") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".facturadocliente} = 0") Then Exit Sub
         
            '[Monica]19/09/2014: añadida esta condicion por teletaxi
            ' solo servicios de credito
            If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".tipservi} = 1") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".tipservi} = 1") Then Exit Sub
            
            '[Monica]13/11/2014: solo para teletaxi, solo se facturan los servicios validados
            If vParamAplic.Cooperativa = 0 Then
                ' solo servicios validados
                If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".validado} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".validado} = 1") Then Exit Sub
            End If
         Else
            If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".facturado} = 0") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".facturado} = 0") Then Exit Sub
         End If
'         ' que no este facturado
'         If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".facturadocliente} = 0") Then Exit Sub
'         If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".facturadocliente} = 0") Then Exit Sub
'
        
        If Not HayRegParaInforme(Tabla, cadSelect) Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
       
        DesBloqueoManual ("FACCLI") 'facturas de publicidad
        If Not BloqueoManual("FACCLI", "1") Then
            MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
            
        If vParamAplic.Cooperativa = 0 Then
            ' sobre el hco de llamadas
            If Me.chk_FactHco.Value = 1 Then
                If GenerarFacturasTeletaxiNew(cadSelect, Tabla) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    HacerImpresionFacturas
                End If
            Else
                If GenerarFacturasTeletaxi(cadSelect, Tabla) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    HacerImpresionFacturas
                End If
            End If
        Else
            If GenerarFacturas(cadSelect, Tabla) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                HacerImpresionFacturas
            End If
        End If
        
        DesBloqueoManual ("FACCLI")
        TerminaBloquear
                
        Screen.MousePointer = vbDefault
End Sub


Private Function CargarTemporal2() As Boolean
Dim SQL As String
Dim SQL2 As String
    
    On Error GoTo eCargarTemporal2
    
    SQL = "delete from tmpcrmclien where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    
    SQL = "insert into tmpcrmclien (codusu,codclien) "
    
    If chk_agrupados.Value = 0 Then
        SQL2 = "select distinct " & vUsu.Codigo & ", codclien from ("
        SQL2 = SQL2 & "select distinct codclien from scliente where codclien not in (select codclienalb from scliente_albaran)  "
        If txtcodigo(0).Text <> "" Then SQL2 = SQL2 & " and codclien >= " & DBSet(txtcodigo(0).Text, "N")
        If txtcodigo(1).Text <> "" Then SQL2 = SQL2 & " and codclien <= " & DBSet(txtcodigo(1).Text, "N")
        SQL2 = SQL2 & " order by 1 ) aaaaaa "
        
    Else
        SQL2 = "select distinct " & vUsu.Codigo & ", codclien from ("
        SQL2 = SQL2 & "select distinct codclienalb codclien from scliente_albaran where (1=1) "
        If txtcodigo(0).Text <> "" Then SQL2 = SQL2 & " and codclien >= " & DBSet(txtcodigo(0).Text, "N")
        If txtcodigo(1).Text <> "" Then SQL2 = SQL2 & " and codclien <= " & DBSet(txtcodigo(1).Text, "N")
        SQL2 = SQL2 & " order by 1 ) aaaaaa "
    
    End If
    
    conn.Execute SQL & SQL2
    CargarTemporal2 = True
    Exit Function
    
    
eCargarTemporal2:
    CargarTemporal2 = False
    MuestraError Err.Number, "Cargando temporal", Err.Description
End Function






Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub



Private Function DatosOK() As Boolean
Dim SQL As String

    DatosOK = False
    '[Monica]04/10/2012: obligamos a poner la fecha hasta de servicios pq va en la fecha de albaran
    If vParamAplic.Cooperativa = 0 Then
        If txtcodigo(86).Text = "" Then
            MsgBox "Es necesario introducir fecha hasta, para fecha de albarán.", vbExclamation
            PonerFoco txtcodigo(86)
            Exit Function
        End If
    End If
    
    'fecha factu
    If txtcodigo(2).Text = "" Then
        MsgBox "Es necesario introducir fecha de factura.", vbExclamation
        PonerFoco txtcodigo(2)
        Exit Function
    End If
    
    If (Me.chk_agrupados.Value = 0 And vParamAplic.Cooperativa = 0) Or vParamAplic.Cooperativa = 1 Then
        'concepto
        If txtcodigo(4).Text = "" Then
            MsgBox "Es necesario introducir el concepto de la factura.", vbExclamation
            PonerFoco txtcodigo(4)
            Exit Function
        End If
    End If
    'banco
    If txtcodigo(5).Text = "" Then
        MsgBox "Es necesario introducir el banco de cobro.", vbExclamation
        PonerFoco txtcodigo(5)
        Exit Function
    Else
        SQL = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", txtcodigo(5).Text, "N")
        If SQL = "" Then
            MsgBox "La Cta.Contable prevista de cobro del banco debe tener valor.", vbExclamation
            Exit Function
        End If
    End If
    
    DatosOK = True

End Function

Private Sub HacerImpresionFacturas()
Dim SQL As String

    cadFormula = "({scafaccli.codtipom}= ""FAC""" & " and {scafaccli.numfactu} in [" & Mid(FacturasaImprimir, 1, Len(FacturasaImprimir) - 1) & "]"
    cadFormula = cadFormula & " and {scafaccli.fecfactu}= Date(" & Year(FecFactu) & "," & Month(FecFactu) & "," & Day(FecFactu) & "))"
    '[Monica]24/04/2014: solo las facturas que sean de socios que no tengan facturacion electronica
    If vParamAplic.Cooperativa = 0 Then
        cadFormula = cadFormula & " and {scliente.tasareciclado} = 0 "
    End If
    
    SQL = "select scafaccli.* from scafaccli, scliente   where codtipom= 'FAC' and numfactu in (" & Mid(FacturasaImprimir, 1, Len(FacturasaImprimir) - 1) & ")"
    SQL = SQL & " and fecfactu= " & DBSet(FecFactu, "F") & " and scliente.codclien = scafaccli.codclien "
    If vParamAplic.Cooperativa = 0 Then
        SQL = SQL & " and scliente.tasareciclado = 0 "
    End If
    
    If TotalRegistrosConsulta(SQL) <> 0 Then LlamarImprimir False
    
End Sub

Private Function GenerarFacturasTeletaxi(cWhere As String, cTabla As String) As Boolean
Dim vTipoMov As CTiposMov
Dim fac As CFactura
Dim TipoMovimiento As String
Dim SQL As String
Dim iva As String
Dim porIva As Currency

Dim porIvaServ As Currency
Dim porIvaGtos As Currency

Dim totfactu As Currency
Dim BaseImp As Currency
Dim ImpIVA As Currency
Dim cli As CCliente
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim o4 As String
Dim o5 As String
Dim tamanyo As Double
Dim almac As String
Dim Prove As String
Dim NomArtic As String
Dim CodTraba As String
Dim b As Boolean
Dim vDevuelve As String

Dim BaseivaGtos As Currency
Dim ImpivaGtos As Currency
Dim BaseivaServ As Currency
Dim ImpivaServ As Currency

Dim RsServ As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim devuelve As String
Dim Existe As Boolean
Dim SQL2 As String
Dim SQL3 As String

Dim linea As Long

Dim cadWHERE As String
Dim Mens As String

Dim Suplidos As Currency
Dim DtoGnral As Currency

Dim sqlLineas As String
Dim RSLineas As ADODB.Recordset
Dim NumAlbar As Long
Dim BaseivaServ2 As Currency
Dim SQLSub As String
Dim SQLSubValues As String
Dim Ampliaci As String

Dim NomLote As String

    On Error GoTo EGenFactu

    GenerarFacturasTeletaxi = False
    TipoMovimiento = "FAC"

    conn.BeginTrans
    ConnConta.BeginTrans
    
    
    
'0 VACIAMOS LA TABLA TEMPORAL
    SQL = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL
    

'1º CARGAMOS UNA TABLA TEMPORAL CON LOS CLIENTES, SUBCLIENTES DE LA QUE SACAREMOS LAS FACTURAS

    '[Monica]04/02/2015
    Label4.Caption = "Cargando tabla auxiliar... "
    DoEvents



    FecFactu = txtcodigo(2).Text
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    SQL = "Select codclien, sum(numserv) servicios, sum(if(importe is null,0,importe)) importe FROM " & QuitarCaracterACadena(cTabla, "_1")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        cWhere = Replace(cWhere, "[", "(")
        cWhere = Replace(cWhere, "]", ")")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    SQL = SQL & " group by 1 having sum(if(importe is null,0,importe)) <> 0"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Dim Cliente As String
    
    SQLSub = "Insert into tmpinformes (codusu, codigo1, importe1, importe2, importe3) values "
    SQLSubValues = ""
    While Not RS.EOF
        Cliente = ""
        Cliente = DevuelveDesdeBDNew(conAri, "scliente_albaran", "codclien", "codclienalb", RS!CodClien, "N")
        If Cliente = "" Then Cliente = RS!CodClien
        
        SQLSubValues = SQLSubValues & "(" & vUsu.Codigo & "," & DBSet(Cliente, "N") & "," & DBSet(RS!CodClien, "N") & ","
        SQLSubValues = SQLSubValues & DBSet(RS!Servicios, "N") & "," & DBSet(RS!Importe, "N")
        SQLSubValues = SQLSubValues & "),"
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    If SQLSubValues <> "" Then
        conn.Execute SQLSub & Mid(SQLSubValues, 1, Len(SQLSubValues) - 1)
    End If


'2º FACTURAMOS DE LA TABLA TEMPORAL
    FecFactu = txtcodigo(2).Text
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    SQL = "Select codigo1 codclien, sum(importe2) servicios, sum(importe3) importe FROM tmpinformes where codusu = " & vUsu.Codigo
    SQL = SQL & " group by 1"
    SQL = SQL & " order by 1"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Set miRsAux = New ADODB.Recordset
    
    'busco el minimo almacen y el minimo proveedor
    SQL = "select min(codalmac) from salmpr"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not miRsAux.EOF Then
        almac = miRsAux.Fields(0)
    End If
    
    miRsAux.Close
    
    SQL = "select min(codprove) from sprove"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not miRsAux.EOF Then
        Prove = miRsAux.Fields(0)
    End If
    
    Set miRsAux = Nothing
    
    
    b = True
    
    FacturasaImprimir = ""
    
    While Not RS.EOF And b
        Set cli = New CCliente
        Set fac = New CFactura
        
        
        If cli.LeerDatos(RS!CodClien, False) Then
            
            '[Monica]04/02/2015
            Label4.Caption = "Cliente: " & Format(RS!CodClien, "N") & " " & cli.Nombre
            DoEvents
            
            
            
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(TipoMovimiento) Then
                NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                ' si existe la factura incrementamos el contador
                Do
                    devuelve = DevuelveDesdeBDNew(conAri, "scafaccli", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", CStr(FecFactu), "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (TipoMovimiento)
                        NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
            Else
                Exit Function
            End If
        
            ' calculo de bases iva de SERVICIOS
            iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArticServ, "T")
            vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T")
            porIvaServ = 0
            If vDevuelve <> "" Then porIvaServ = CCur(vDevuelve)
            fac.TipoIVA1 = iva
            
            Suplidos = 0
            DtoGnral = 0
            BaseivaServ = Round2(RS!Importe / (1 + (porIvaServ / 100)), 2)
            ImpivaServ = Round2(RS!Importe - BaseivaServ, 2)
            
            ' calculo de base iva de GASTOS ADMON
            iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArtGastosAdmon, "T")
            vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T")
            porIvaGtos = 0
            If vDevuelve <> "" Then porIvaGtos = CCur(vDevuelve)
            
            BaseivaGtos = cli.GastosAdmon
            ImpivaGtos = Round2(BaseivaGtos * porIvaGtos / 100, 2)
            
            ' Asignamos los importes a la factura
            If BaseivaGtos <> 0 Then
                If iva = fac.TipoIVA1 Then
                    fac.TipoIVA1 = iva
                    fac.BaseIVA1 = BaseivaGtos
                    fac.PorceIVA1 = porIvaGtos
                    fac.ImpIVA1 = ImpivaGtos
                Else
                    fac.TipoIVA2 = iva
                    fac.BaseIVA2 = BaseivaGtos
                    fac.PorceIVA2 = porIvaGtos
                    fac.ImpIVA2 = ImpivaGtos
                End If
            End If
            'el tipo de iva 1 esta asignado cuando se busca en tiposiva de la conta
            fac.PorceIVA1 = porIvaServ
            fac.BaseIVA1 = fac.BaseIVA1 + BaseivaServ
            fac.ImpIVA1 = fac.ImpIVA1 + ImpivaServ
            fac.BaseImp = BaseivaServ + BaseivaGtos
            fac.ImpGnral = DtoGnral
            fac.DtoGnral = cli.DtoGnral
            fac.BrutoFac = fac.BaseImp
            fac.Suplidos = Suplidos
            fac.TotalFac = BaseivaServ + ImpivaServ + Suplidos + BaseivaGtos + ImpivaGtos
            
            fac.codtipom = TipoMovimiento
            
            fac.FecFactu = FecFactu
            fac.LetraSerie = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", TipoMovimiento, "T")
'            NumFactu = vTipoMov.Contador
            fac.NumFactu = NumFactu
            FacturasaImprimir = FacturasaImprimir & NumFactu & ","
            
            'Cuenta Prevista de Cobro de las Facturas
            fac.BancoPr = txtcodigo(5).Text
            fac.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", fac.BancoPr, "N")
        
            fac.Cliente = cli.Codigo
            fac.NombreClien = cli.Nombre
            fac.DomicilioClien = cli.Domicilio
            fac.CPostal = cli.CPostal
            fac.Poblacion = cli.Poblacion
            fac.Provincia = cli.Provincia
            fac.NIF = cli.NIF
            fac.ForPago = cli.ForPago
        
            '[Monica]22/11/2013: iban
            fac.Iban = cli.Iban
        
            fac.Banco = cli.Banco
            fac.Sucursal = cli.Sucursal
            fac.DigControl = cli.DigControl
            fac.CuentaBan = cli.CuentaBan
            
            Mens = "Insertando en cabecera factura"
            'scafaccli
            SQL = "INSERT INTO scafaccli (codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,"
            SQL = SQL & "nifclien,codagent,codforpa,dtoppago,dtognral,brutofac,impdtopp,impdtogr,baseimp1,codigiv1,porciva1,"
            SQL = SQL & "imporiv1,baseimp2,codigiv2,porciva2,imporiv2,totalfac,intconta,coddirec,codbanco,codsucur,digcontr,cuentaba, numservi, suplidos, iban) VALUES ("
            SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "'," & DBSet(fac.Cliente, "N") & ","
            SQL = SQL & DBSet(cli.Nombre, "T") & "," & DBSet(cli.Domicilio, "T") & "," & DBSet(cli.CPostal, "T") & ","
            SQL = SQL & DBSet(cli.Poblacion, "T") & "," & DBSet(cli.Provincia, "T") & "," & DBSet(cli.NIF, "T") & "," & vParamAplic.PorDefecto_Agente
            SQL = SQL & "," & cli.ForPago & ",0," & DBSet(fac.DtoGnral, "N") & "," & DBSet(fac.BrutoFac, "N") & ",0," & DBSet(fac.ImpGnral, "N") & ","
            SQL = SQL & DBSet(fac.BaseIVA1, "N") & "," & DBSet(fac.TipoIVA1, "N")
            SQL = SQL & "," & DBSet(fac.PorceIVA1, "N") & "," & DBSet(fac.ImpIVA1, "N") & ","
            SQL = SQL & DBSet(fac.BaseIVA2, "N", "S") & "," & DBSet(fac.TipoIVA2, "N", "S") & "," & DBSet(fac.PorceIVA2, "N", "S") & ","
            SQL = SQL & DBSet(fac.ImpIVA2, "N", "S") & "," & DBSet(fac.TotalFac, "N") & ",0,NULL,"
            SQL = SQL & DBSet(fac.Banco, "N") & "," & DBSet(fac.Sucursal, "N") & "," & DBSet(fac.DigControl, "T") & "," & DBSet(fac.CuentaBan, "T") & ","
            SQL = SQL & DBSet(RS!Servicios, "N") & "," & DBSet(Suplidos, "N") & "," & DBSet(fac.Iban, "T") & ")"
        
        
            conn.Execute SQL
        

            o1 = DevuelveDesdeBD(conAri, "observa1", "scliente", "codclien", cli.Codigo, "N") '("select observa1 from scliente where codclien = " & DBSet(cli.Codigo, "N"))
            
            
            CodTraba = DevuelveDesdeBD(conAri, "codtraba", "straba", "login", vUsu.Login, "T")
            If CodTraba = "" Then CodTraba = DevuelveValor("select min(codtraba) from straba")
            Mens = "Insertando Albaran"
            
            
            sqlLineas = "Select importe1 codclien, importe2 servicios, importe3 importe FROM tmpinformes where codusu = " & vUsu.Codigo & " and codigo1 = " & DBSet(RS!CodClien, "N")
    
            Set RSLineas = New ADODB.Recordset
            RSLineas.Open sqlLineas, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            NumAlbar = 0
        
            While Not RSLineas.EOF
                'NumAlbar = NumAlbar + 1
                NumAlbar = RSLineas!CodClien
            
                SQL = "INSERT INTO scafaccli1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,"
                SQL = SQL & "codenvio,codtraba,codtrab2,observa1,observa2,observa3,observa4,observa5,codtrab1) VALUES ("
                SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV'," & DBSet(NumAlbar, "N") & ",'"
                SQL = SQL & Format(txtcodigo(86).Text, FormatoFecha) & "'," & vParamAplic.PorDefecto_Envio & "," & CodTraba
                SQL = SQL & "," & CodTraba & "," & DBSet(o1, "T") & "," & DBSet(o2, "T") & "," & DBSet(o3, "T") & ","
                SQL = SQL & DBSet(o4, "T") & "," & DBSet(o5, "T") & ",NULL)"
            
                conn.Execute SQL
                'slifac
            
                Mens = "Insertando linea de articulo de Servicios"
                'busco el nombre del articulo
                NomArtic = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArticServ, "T")
            
                BaseivaServ2 = Round2(RSLineas!Importe / (1 + (porIvaServ / 100)), 2)
            
                SQL = "INSERT INTO slifacCli (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
                SQL = SQL & "numbultos,cantidad,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel) VALUES ("
                SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV'," & DBSet(NumAlbar, "N") & ",0," & almac & ","
                If RSLineas!CodClien = RS!CodClien And chk_agrupados.Value = 0 Then
                    SQL = SQL & DBSet(vParamAplic.ArticServ, "T") & "," & DBSet(NomArtic, "T") & "," & DBSet(txtcodigo(4).Text, "T") & ",1," & DBSet(RSLineas!Servicios, "N") & " ," & DBSet(BaseivaServ2, "N") & ","
                Else
                    NomLote = ""
                    NomLote = DevuelveValor("select nomenvio from senvio inner join scliente on senvio.codenvio = scliente.codenvio where scliente.codclien = " & DBSet(RS!CodClien, "N"))
                    NomArtic = Trim(NomArtic) & " " & UCase(Format(txtcodigo(86).Text, "mmmm")) & " " & Year(CDate(txtcodigo(86).Text)) & " " & NomLote
                    Ampliaci = DevuelveValor("select nomclien from scliente where codclien = " & DBSet(RSLineas!CodClien, "N"))
                    SQL = SQL & DBSet(vParamAplic.ArticServ, "T") & "," & DBSet(NomArtic, "T") & "," & DBSet(Ampliaci, "T") & ",1," & DBSet(RSLineas!Servicios, "N") & " ," & DBSet(BaseivaServ2, "N") & ","
                End If
                SQL = SQL & DBSet(BaseivaServ2, "N") & "," & DBSet(BaseivaServ2, "N") & "," & DBSet(BaseivaServ2, "N") & ","
                SQL = SQL & DBSet(BaseivaServ2, "N") & ",0,0,'M'," & Prove & "," & DBSet(BaseivaServ2, "N") & ")"
            
                conn.Execute SQL
            
                '[Monica]31/03/2014: Insertamos las lineas de servicios que han actuado
                Dim SqlLin As String
                Dim Sql4 As String
                Dim RsLin As ADODB.Recordset
                
                Sql4 = "insert into scafaccli_serv (codtipom,numfactu,fecfactu,numlinea,fecha,hora,codsocio,numeruve,"
                Sql4 = Sql4 & "dirllama,observa1,impventa,idservic,observac2,codclien) "
                
                SqlLin = "select " & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "', @rownum:=@rownum+1 AS rownum, fecha, hora, codsocio, numeruve, "
                SqlLin = SqlLin & " origen, destino, importe, nroservicio,  matricul, codclien from sfactclitr_serv,(SELECT @rownum:=0) r "
                SqlLin = SqlLin & " where (codclien, fecfactu) in (select codclien, fecfactu from sfactclitr "
                SqlLin = SqlLin & " where " & cWhere
                SqlLin = SqlLin & " and codclien = " & DBSet(RSLineas!CodClien, "N")
                SqlLin = SqlLin & " order by fecha, hora "
                SqlLin = SqlLin & ")"
                
                conn.Execute Sql4 & SqlLin
                '31/03/2014: hasta aqui
            
            
            
            
                SQL3 = "update sfactclitr set facturado = 1 where " & cWhere & " and codclien = " & DBSet(RSLineas!CodClien, "N")
                conn.Execute SQL3
                
                RSLineas.MoveNext
            Wend
            
            Set RSLineas = Nothing
            
            
             'si hay gastos de admon insertamos una linea de albaran con los gastos
             If BaseivaGtos <> 0 And NumAlbar <> 0 Then
'                 NumAlbar = NumAlbar + 1
                 
                 Mens = "Insertando linea de articulo de Gastos"
             
                 NomArtic = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtGastosAdmon, "T")
             
                 SQL = "INSERT INTO slifacCli (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
                 SQL = SQL & "numbultos,cantidad,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel) VALUES ("
                 SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV'," & DBSet(NumAlbar, "N") & ",1," & almac & ","
                 SQL = SQL & DBSet(vParamAplic.ArtGastosAdmon, "T") & "," & DBSet(NomArtic, "T") & ",1,1," & DBSet(BaseivaGtos, "N") & ","
                 SQL = SQL & DBSet(BaseivaGtos, "N") & "," & DBSet(BaseivaGtos, "N") & "," & DBSet(BaseivaGtos, "N") & ","
                 SQL = SQL & DBSet(BaseivaGtos, "N") & ",0,0,'M'," & Prove & "," & DBSet(BaseivaGtos, "N") & ")"
             
                 conn.Execute SQL
            End If
            
            'insertar en tesoreria
            fac.Agente = vParamAplic.PorDefecto_Agente
            
            b = fac.InsertarEnTesoreriaFACcli("", "Error al pasar a Tesoreria")
            'b = fac.InsertarEnTesoreriaFACcli("", "Error al pasar a tesoreria")
        
        
        
        
            If b Then vTipoMov.IncrementarContador (TipoMovimiento)
            
            Set vTipoMov = Nothing
            Set cli = Nothing
            Set fac = Nothing
        
        Else
            MsgBox "No existe el cliente " & cli.Codigo & " " & cli.Nombre
            b = False
        End If
        RS.MoveNext
    
    Wend
    
    Set RS = Nothing
    
' hasta aqui
    GenerarFacturasTeletaxi = True

EGenFactu:
    If Err.Number <> 0 Then
        Mens = "Insertando Movimiento." & vbCrLf & "----------------------------" & vbCrLf & Mens
        MuestraError Err.Number, Mens, Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        GenerarFacturasTeletaxi = True
        '[Monica]04/02/2015
        Me.Label4.Caption = ""
        DoEvents
        
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        GenerarFacturasTeletaxi = False
    End If
    
'If Err <> 0 Or Not B Then
'    conn.RollbackTrans
'    ConnConta.RollbackTrans
'    MsgBox "ERROR AL GENERAR FACTURAS:" & Err.Description
'    DesBloqueoManual ("FACCLI")
'    TerminaBloquear
'End If
End Function

Private Sub LlamarImprimir(duplicado As Boolean)
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        cadParam = cadParam & "pDuplicado= " & Abs(duplicado) & " |"
        numParam = 2
        
        '[Monica]31/03/2014: en el caso de teletaxi pedimos si imprime o no detalle
        If vParamAplic.Cooperativa = 0 Then
            '[Monica]10/09/2014: si no partimos del hco no hacemos pregunta
            If Me.chk_FactHco.Value = 1 Then
                If MsgBox("¿ Desea imprimir el detalle de servicios ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                    cadParam = cadParam & "pDetalle=0|"
                Else
                    cadParam = cadParam & "pDetalle=1|"
                End If
            Else
                cadParam = cadParam & "pDetalle=0|"
            End If
            numParam = numParam + 1
        End If
        'hasta aquí
                
        
        With frmImprimir
        .Titulo = "Impresión de Facturas de Cliente"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "52", "N")
    '------ > Listado 47 = rFacPubli.rpt
        .Opcion = 101
        .ConSubInforme = True
        .Show vbModal
    End With

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdRegresar_Click()
    Unload Me
End Sub


Private Sub Form_Load()
Dim i As Integer
    'Icono del form
    Me.Icon = frmPpal.Icon

    txtcodigo(2).Text = Date
'    Text1(4).Text = vParamAplic.ConFactuPubli
    Modo = 0
    
    Me.chk_FactHco.Value = 1
    
    chk_FactHco.visible = (vParamAplic.Cooperativa = 0)
    chk_FactHco.Enabled = (vParamAplic.Cooperativa = 0)
    
    If vParamAplic.Cooperativa = 0 Then
        If Me.chk_FactHco.Value = 1 Then
            '[Monica]10/09/2014
            Tabla = "shilla"
        Else
            Tabla = "sfactclitr"
        End If
    Else
        Tabla = "shilla"
    End If
    
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i

End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    cad = CadenaDevuelta
End Sub

Private Sub frmCal_Selec(vFecha As Date)
    txtcodigo(1).Text = vFecha
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    
    If CadenaSeleccion = "Salir" Then
        Salir = True
        Exit Sub
    End If
    
    If CadenaSeleccion <> "" Then
        CadServicios = "(fecha,hora,numeruve) in (" & CadenaSeleccion & ")"
    Else
        CadServicios = "shilla.numeruve = -1"
    End If
    
    If Not AnyadirAFormula(cadSelect, CadServicios) Then Exit Sub
    
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "El concepto únicamente aparece en la ampliacion de facturas de  " & vbCrLf & _
                      "clientes no agrupados donde indica mes año de facturacion " & vbCrLf & vbCrLf & _
                      "En el caso de los clientes agrupados no aparece en ningún lado, ni en la" & vbCrLf & _
                      "factura ni en el albarán." & vbCrLf & _
                      vbCrLf
                      
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    Select Case Index
        Case 0
            CadenaDesdeOtroForm = txtcodigo(4).Text
            frmFacClienteObser.Modificar = True
            frmFacClienteObser.Text1 = CadenaDesdeOtroForm
            frmFacClienteObser.Show vbModal
            If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then txtcodigo(4).Text = Mid(CadenaDesdeOtroForm, 3)
            CadenaDesdeOtroForm = ""
            PonerFoco txtcodigo(4)
        Case 1
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0|1|"
            frmFP.Show vbModal
            Set frmFP = Nothing
        Case 4, 5
            indCodigo = Index - 4
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0|1|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            PonerFoco txtcodigo(indCodigo)
        Case 3
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
    End Select
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    txtcodigo(5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtnombre(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
   
    Set frmCal = New frmCal
    frmCal.Fecha = Now
    PonerFormatoFecha txtcodigo(1)
    If txtcodigo(1).Text <> "" Then frmCal.Fecha = CDate(txtcodigo(1).Text)
    Screen.MousePointer = vbDefault
    frmCal.Show vbModal
    Set frmCal = Nothing
    PonerFoco txtcodigo(1)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
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
        Case 85, 86  'FECHA Desde Hasta
            PonerFormatoFecha txtcodigo(Index)
            
        Case 0, 1 'cliente
            PonerFormatoEntero txtcodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "scliente", "nomclien", "codclien", "Cliente", "N")
            
        Case 3 ' forma de pago
            PonerFormatoEntero txtcodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sforpa", "nomforpa", "codforpa", "Forma de Pago", "N")
            
        Case 2 ' fecha de factura
            PonerFormatoFecha txtcodigo(Index)
             
        Case 5 'banco propio
            If PonerFormatoEntero(txtcodigo(5)) Then
                txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sbanpr", "nombanpr", "codbanpr", , "N")
            End If
           
    End Select
End Sub

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



Private Function GenerarFacturas(cWhere As String, cTabla As String) As Boolean
Dim vTipoMov As CTiposMov
Dim fac As CFactura
Dim TipoMovimiento As String
Dim SQL As String
Dim iva As String
Dim porIva As Currency

Dim porIvaServ As Currency
Dim porIvaGtos As Currency

Dim totfactu As Currency
Dim BaseImp As Currency
Dim ImpIVA As Currency
Dim cli As CCliente
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim o4 As String
Dim o5 As String
Dim tamanyo As Double
Dim almac As String
Dim Prove As String
Dim NomArtic As String
Dim CodTraba As String
Dim b As Boolean
Dim vDevuelve As String

Dim BaseivaGtos As Currency
Dim ImpivaGtos As Currency
Dim BaseivaServ As Currency
Dim ImpivaServ As Currency

Dim RsServ As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim devuelve As String
Dim Existe As Boolean
Dim SQL2 As String
Dim SQL3 As String

Dim linea As Long

Dim cadWHERE As String
Dim Mens As String

Dim Suplidos As Currency
Dim DtoGnral As Currency


    On Error GoTo EGenFactu

    GenerarFacturas = False
    TipoMovimiento = "FAC"

    conn.BeginTrans
    ConnConta.BeginTrans

    FecFactu = txtcodigo(2).Text
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    
    If vParamAplic.Cooperativa = 0 Then
        SQL = "Select codclien, sum(numserv) servicios, sum(if(importe is null,0,importe)) importe FROM " & QuitarCaracterACadena(cTabla, "_1")
    Else
        SQL = "Select codclien, count(*) servicios, sum(if(impventa is null,0,impventa)) importe, sum(if(imppeaje is null,0,imppeaje)) suplidos FROM " & QuitarCaracterACadena(cTabla, "_1")
    End If
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        cWhere = Replace(cWhere, "[", "(")
        cWhere = Replace(cWhere, "]", ")")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    If vParamAplic.Cooperativa = 0 Then
        SQL = SQL & " group by 1 having sum(if(importe is null,0,importe)) <> 0"
    Else
        SQL = SQL & " group by 1 having sum(if(impventa is null,0,impventa)) <> 0"
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    Set miRsAux = New ADODB.Recordset
    
    'busco el minimo almacen y el minimo proveedor
    SQL = "select min(codalmac) from salmpr"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not miRsAux.EOF Then
        almac = miRsAux.Fields(0)
    End If
    
    miRsAux.Close
    
    SQL = "select min(codprove) from sprove"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not miRsAux.EOF Then
        Prove = miRsAux.Fields(0)
    End If
    
    Set miRsAux = Nothing
    
    
    b = True
    
    FacturasaImprimir = ""
    
    While Not RS.EOF And b
        Set cli = New CCliente
        Set fac = New CFactura
        
        If cli.LeerDatos(RS!CodClien, False) Then
            
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(TipoMovimiento) Then
                NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                ' si existe la factura incrementamos el contador
                Do
                    devuelve = DevuelveDesdeBDNew(conAri, "scafaccli", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", CStr(FecFactu), "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (TipoMovimiento)
                        NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
            Else
                Exit Function
            End If
        
            ' calculo de bases iva de SERVICIOS
            iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArticServ, "T")
            vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T")
            porIvaServ = 0
            If vDevuelve <> "" Then porIvaServ = CCur(vDevuelve)
            fac.TipoIVA1 = iva
            
            Suplidos = 0
            DtoGnral = 0
            If vParamAplic.Cooperativa = 0 Then
                BaseivaServ = Round2(RS!Importe / (1 + (porIvaServ / 100)), 2)
                ImpivaServ = Round2(RS!Importe - BaseivaServ, 2)
            Else
                Suplidos = RS!Suplidos
                DtoGnral = Round2((RS!Importe - Suplidos) * cli.DtoGnral / 100, 2)
                
                BaseivaServ = Round2((RS!Importe - Suplidos - DtoGnral) / (1 + (porIvaServ / 100)), 2)
'                vBruto = Round2(BaseivaServ / (1 - (cli.DtoGnral / 100)), 2)
'                DtoGnral = Round2(vBruto * cli.DtoGnral / 100, 2)
                ImpivaServ = Round2(RS!Importe - Suplidos - DtoGnral - BaseivaServ, 2)
            End If
            ' calculo de base iva de GASTOS ADMON
            iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArtGastosAdmon, "T")
            vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T")
            porIvaGtos = 0
            If vDevuelve <> "" Then porIvaGtos = CCur(vDevuelve)
            
            BaseivaGtos = cli.GastosAdmon
            ImpivaGtos = Round2(BaseivaGtos * porIvaGtos / 100, 2)
            
            ' Asignamos los importes a la factura
            If BaseivaGtos <> 0 Then
                If iva = fac.TipoIVA1 Then
                    fac.TipoIVA1 = iva
                    fac.BaseIVA1 = BaseivaGtos
                    fac.PorceIVA1 = porIvaGtos
                    fac.ImpIVA1 = ImpivaGtos
                Else
                    fac.TipoIVA2 = iva
                    fac.BaseIVA2 = BaseivaGtos
                    fac.PorceIVA2 = porIvaGtos
                    fac.ImpIVA2 = ImpivaGtos
                End If
            End If
            'el tipo de iva 1 esta asignado cuando se busca en tiposiva de la conta
            fac.PorceIVA1 = porIvaServ
            fac.BaseIVA1 = fac.BaseIVA1 + BaseivaServ
            fac.ImpIVA1 = fac.ImpIVA1 + ImpivaServ
            If vParamAplic.Cooperativa = 0 Then
                fac.BaseImp = BaseivaServ + BaseivaGtos
            Else
                fac.BaseImp = DBLet(RS!Importe, "N") - Suplidos
            End If
            fac.ImpGnral = DtoGnral
            fac.DtoGnral = cli.DtoGnral
            fac.BrutoFac = fac.BaseImp
            fac.Suplidos = Suplidos
            fac.TotalFac = BaseivaServ + ImpivaServ + Suplidos + BaseivaGtos + ImpivaGtos
            
            fac.codtipom = TipoMovimiento
            
            fac.FecFactu = FecFactu
            fac.LetraSerie = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", TipoMovimiento, "T")
'            NumFactu = vTipoMov.Contador
            fac.NumFactu = NumFactu
            FacturasaImprimir = FacturasaImprimir & NumFactu & ","
            
            'Cuenta Prevista de Cobro de las Facturas
            fac.BancoPr = txtcodigo(5).Text
            fac.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", fac.BancoPr, "N")
        
            fac.Cliente = cli.Codigo
            fac.NombreClien = cli.Nombre
            fac.DomicilioClien = cli.Domicilio
            fac.CPostal = cli.CPostal
            fac.Poblacion = cli.Poblacion
            fac.Provincia = cli.Provincia
            fac.NIF = cli.NIF
            fac.ForPago = cli.ForPago
            '[Monica]22/11/2013
            fac.Iban = cli.Iban
            fac.Banco = cli.Banco
            fac.Sucursal = cli.Sucursal
            fac.DigControl = cli.DigControl
            fac.CuentaBan = cli.CuentaBan
            
            Mens = "Insertando en cabecera factura"
            'scafaccli
            SQL = "INSERT INTO scafaccli (codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,"
            SQL = SQL & "nifclien,codagent,codforpa,dtoppago,dtognral,brutofac,impdtopp,impdtogr,baseimp1,codigiv1,porciva1,"
            SQL = SQL & "imporiv1,baseimp2,codigiv2,porciva2,imporiv2,totalfac,intconta,coddirec,codbanco,codsucur,digcontr,cuentaba, numservi, suplidos, iban) VALUES ("
            SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "'," & DBSet(fac.Cliente, "N") & ","
            SQL = SQL & DBSet(cli.Nombre, "T") & "," & DBSet(cli.Domicilio, "T") & "," & DBSet(cli.CPostal, "T") & ","
            SQL = SQL & DBSet(cli.Poblacion, "T") & "," & DBSet(cli.Provincia, "T") & "," & DBSet(cli.NIF, "T") & "," & vParamAplic.PorDefecto_Agente
            SQL = SQL & "," & cli.ForPago & ",0," & DBSet(fac.DtoGnral, "N") & "," & DBSet(fac.BrutoFac, "N") & ",0," & DBSet(fac.ImpGnral, "N") & ","
            SQL = SQL & DBSet(fac.BaseIVA1, "N") & "," & DBSet(fac.TipoIVA1, "N")
            SQL = SQL & "," & DBSet(fac.PorceIVA1, "N") & "," & DBSet(fac.ImpIVA1, "N") & ","
            SQL = SQL & DBSet(fac.BaseIVA2, "N", "S") & "," & DBSet(fac.TipoIVA2, "N", "S") & "," & DBSet(fac.PorceIVA2, "N", "S") & ","
            SQL = SQL & DBSet(fac.ImpIVA2, "N", "S") & "," & DBSet(fac.TotalFac, "N") & ",0,NULL,"
            SQL = SQL & DBSet(fac.Banco, "N") & "," & DBSet(fac.Sucursal, "N") & "," & DBSet(fac.DigControl, "T") & "," & DBSet(fac.CuentaBan, "T") & ","
            SQL = SQL & DBSet(RS!Servicios, "N") & "," & DBSet(Suplidos, "N") & "," & DBSet(fac.Iban, "T") & ")"
        
        
            conn.Execute SQL
        
'            'scafac1
'            'acoplamos el concepto a las observaciones de la scafac1
'            Tamanyo = Len(txtcodigo(4).Text)
'            Tamanyo = Tamanyo / 80
'            Select Case Tamanyo
'                Case Is <= 1
'                    o1 = txtcodigo(4).Text
'                Case Is <= 2
'                    o1 = Mid(txtcodigo(4).Text, 1, 80)
'                    o2 = Mid(txtcodigo(4).Text, 81)
'                Case Is <= 3
'                    o1 = Mid(txtcodigo(4).Text, 1, 80)
'                    o2 = Mid(txtcodigo(4).Text, 81, 160)
'                    o3 = Mid(txtcodigo(4).Text, 161)
'                Case Is <= 4
'                    o1 = Mid(txtcodigo(4).Text, 1, 80)
'                    o2 = Mid(txtcodigo(4).Text, 81, 160)
'                    o3 = Mid(txtcodigo(4).Text, 161, 240)
'                    o4 = Mid(txtcodigo(4).Text, 241)
'                Case Else
'                    o1 = Mid(txtcodigo(4).Text, 1, 80)
'                    o2 = Mid(txtcodigo(4).Text, 81, 160)
'                    o3 = Mid(txtcodigo(4).Text, 161, 240)
'                    o4 = Mid(txtcodigo(4).Text, 241, 320)
'                    o5 = Mid(txtcodigo(4).Text, 321, 400)
'            End Select

            o1 = DevuelveDesdeBD(conAri, "observa1", "scliente", "codclien", cli.Codigo, "N") '("select observa1 from scliente where codclien = " & DBSet(cli.Codigo, "N"))
            
            
            CodTraba = DevuelveDesdeBD(conAri, "codtraba", "straba", "login", vUsu.Login, "T")
            If CodTraba = "" Then CodTraba = DevuelveValor("select min(codtraba) from straba")
            Mens = "Insertando Albaran"
            
            SQL = "INSERT INTO scafaccli1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,"
            SQL = SQL & "codenvio,codtraba,codtrab2,observa1,observa2,observa3,observa4,observa5,codtrab1) VALUES ("
            SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0,'"
            SQL = SQL & Format(FecFactu, FormatoFecha) & "'," & vParamAplic.PorDefecto_Envio & "," & CodTraba
            SQL = SQL & "," & CodTraba & "," & DBSet(o1, "T") & "," & DBSet(o2, "T") & "," & DBSet(o3, "T") & ","
            SQL = SQL & DBSet(o4, "T") & "," & DBSet(o5, "T") & ",NULL)"
            
            conn.Execute SQL
            'slifac
            
            
            Mens = "Insertando linea de articulo de Servicios"
            'busco el nombre del articulo
            NomArtic = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArticServ, "T")
            
            SQL = "INSERT INTO slifacCli (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
            SQL = SQL & "numbultos,cantidad,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel) VALUES ("
            SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0,0," & almac & ","
            SQL = SQL & DBSet(vParamAplic.ArticServ, "T") & "," & DBSet(NomArtic, "T") & "," & DBSet(txtcodigo(4).Text, "T") & ",1,1," & DBSet(BaseivaServ, "N") & ","
            SQL = SQL & DBSet(BaseivaServ, "N") & "," & DBSet(BaseivaServ, "N") & "," & DBSet(BaseivaServ, "N") & ","
            SQL = SQL & DBSet(BaseivaServ, "N") & ",0,0,'M'," & Prove & "," & DBSet(BaseivaServ, "N") & ")"
            
            conn.Execute SQL
            
            'si hay gastos de admon insertamos una linea de albaran con los gastos
            If BaseivaGtos <> 0 Then
                Mens = "Insertando linea de articulo de Gastos"
                
                NomArtic = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtGastosAdmon, "T")
                
                SQL = "INSERT INTO slifacCli (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
                SQL = SQL & "numbultos,cantidad,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel) VALUES ("
                SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0,1," & almac & ","
                SQL = SQL & DBSet(vParamAplic.ArtGastosAdmon, "T") & "," & DBSet(NomArtic, "T") & ",1,1," & DBSet(BaseivaGtos, "N") & ","
                SQL = SQL & DBSet(BaseivaGtos, "N") & "," & DBSet(BaseivaGtos, "N") & "," & DBSet(BaseivaGtos, "N") & ","
                SQL = SQL & DBSet(BaseivaGtos, "N") & ",0,0,'M'," & Prove & "," & DBSet(BaseivaGtos, "N") & ")"
                
                conn.Execute SQL
           End If
            
            If vParamAplic.Cooperativa = 1 Then
            
                ' insertamos los servicios
                Mens = "Insertando Servicios"
    
                SQL = "INSERT INTO scafaccli_serv (codtipom,numfactu,fecfactu,numlinea,fecha,hora,codsocio,numeruve,dirllama,numllama,puerllama,"
                SQL = SQL & "ciudadre , Telefono, impventa, idservic, observac2, observa1) values "
    
                SQL2 = "select * from shilla where codclien = " & fac.Cliente & " and " & cWhere
                SQL2 = SQL2 & " order by fecha, hora "
    
                Set RsServ = New ADODB.Recordset
                RsServ.Open SQL2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                linea = 0
                cadWHERE = ""
                While Not RsServ.EOF
                    linea = linea + 1
    
                    cadWHERE = cadWHERE & "(" & DBSet(TipoMovimiento, "T") & "," & NumFactu & "," & DBSet(FecFactu, "F") & "," & DBSet(linea, "N") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!Fecha, "F") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!hora, "H") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!codSocio, "N") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!NumerUve, "N") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!dirllama, "T") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!numllama, "T") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!puerllama, "T") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!ciudadre, "T") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!Telefono, "T") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!impventa, "N") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!idservic, "T") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!observac2, "T") & ","
                    cadWHERE = cadWHERE & DBSet(RsServ!observa1, "T") & "),"
    
                    ' a la vez que guardamos los servicios los marcamos como que ya han sido cobrados
                    SQL3 = "update shilla set facturadocliente = 1 where fecha = " & DBSet(RsServ!Fecha, "F")
                    SQL3 = SQL3 & " and hora = " & DBSet(RsServ!hora, "H")
                    SQL3 = SQL3 & " and numeruve = " & DBSet(RsServ!NumerUve, "N")
    
                    conn.Execute SQL3
    
                    RsServ.MoveNext
                Wend
                Set RsServ = Nothing
    
                If linea <> 0 Then
                    SQL = SQL & Mid(cadWHERE, 1, Len(cadWHERE) - 1)
    
                    conn.Execute SQL
                End If
                Set RsServ = Nothing
            End If
            
            'insertar en tesoreria
            fac.Agente = vParamAplic.PorDefecto_Agente
            
            b = fac.InsertarEnTesoreriaFACcli("", "Error al pasar a Tesoreria")
            'b = fac.InsertarEnTesoreriaFACcli("", "Error al pasar a tesoreria")
        
            If b Then vTipoMov.IncrementarContador (TipoMovimiento)
            
            ' marcamos los registros de sfactclitr
            If b And vParamAplic.Cooperativa = 0 Then
                SQL3 = "update sfactclitr set facturado = 1 where " & cWhere & " and codclien = " & DBSet(RS!CodClien, "N")
                conn.Execute SQL3
            End If
            
            
            Set vTipoMov = Nothing
            Set cli = Nothing
            Set fac = Nothing
        
        Else
            MsgBox "No existe el cliente " & cli.Codigo & " " & cli.Nombre
            b = False
        End If
        RS.MoveNext
    
    Wend
    
    Set RS = Nothing
    
    
    GenerarFacturas = True

EGenFactu:
    If Err.Number <> 0 Then
        Mens = "Insertando Movimiento." & vbCrLf & "----------------------------" & vbCrLf & Mens
        MuestraError Err.Number, Mens, Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        GenerarFacturas = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        GenerarFacturas = False
    End If
    
'If Err <> 0 Or Not B Then
'    conn.RollbackTrans
'    ConnConta.RollbackTrans
'    MsgBox "ERROR AL GENERAR FACTURAS:" & Err.Description
'    DesBloqueoManual ("FACCLI")
'    TerminaBloquear
'End If
End Function

'###################################################################################################################
'#############
'############# NUEVA FACTURACION DE TELETAXI PARTIENDO DEL HISTORICO DE LLAMADAS
'#############
'###################################################################################################################

Private Function GenerarFacturasTeletaxiNew(cWhere As String, cTabla As String) As Boolean
Dim vTipoMov As CTiposMov
Dim fac As CFactura
Dim TipoMovimiento As String
Dim SQL As String
Dim iva As String
Dim porIva As Currency

Dim porIvaServ As Currency
Dim porIvaGtos As Currency

Dim totfactu As Currency
Dim BaseImp As Currency
Dim ImpIVA As Currency
Dim cli As CCliente
Dim o1 As String
Dim o2 As String
Dim o3 As String
Dim o4 As String
Dim o5 As String
Dim tamanyo As Double
Dim almac As String
Dim Prove As String
Dim NomArtic As String
Dim CodTraba As String
Dim b As Boolean
Dim vDevuelve As String

Dim BaseivaGtos As Currency
Dim ImpivaGtos As Currency
Dim BaseivaServ As Currency
Dim ImpivaServ As Currency

Dim RsServ As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim devuelve As String
Dim Existe As Boolean
Dim SQL2 As String
Dim SQL3 As String

Dim linea As Long

Dim cadWHERE As String
Dim Mens As String

Dim Suplidos As Currency
Dim DtoGnral As Currency

Dim sqlLineas As String
Dim RSLineas As ADODB.Recordset
Dim NumAlbar As Long
Dim BaseivaServ2 As Currency
Dim SQLSub As String
Dim SQLSubValues As String
Dim Ampliaci As String

Dim NomLote As String

    On Error GoTo EGenFactu

    GenerarFacturasTeletaxiNew = False
    TipoMovimiento = "FAC"

    conn.BeginTrans
    ConnConta.BeginTrans
    
    
    
'0 VACIAMOS LA TABLA TEMPORAL
    SQL = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL
    

'1º CARGAMOS UNA TABLA TEMPORAL CON LOS CLIENTES, SUBCLIENTES DE LA QUE SACAREMOS LAS FACTURAS
    
    '[Monica]04/02/2015
    Label4.Caption = "Cargando tabla auxiliar... "
    DoEvents
    
    
    FecFactu = txtcodigo(2).Text
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    SQL = "Select codclien, count(*) servicios, sum(if(impventa is null,0,impventa)) importe FROM " & QuitarCaracterACadena(cTabla, "_1")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        cWhere = Replace(cWhere, "[", "(")
        cWhere = Replace(cWhere, "]", ")")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    SQL = SQL & " group by 1 having sum(if(impventa is null,0,impventa)) <> 0"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Dim Cliente As String
    
    SQLSub = "Insert into tmpinformes (codusu, codigo1, importe1, importe2, importe3) values "
    SQLSubValues = ""
    While Not RS.EOF
        Cliente = ""
        Cliente = DevuelveDesdeBDNew(conAri, "scliente_albaran", "codclien", "codclienalb", RS!CodClien, "N")
        If Cliente = "" Then Cliente = RS!CodClien
        
        SQLSubValues = SQLSubValues & "(" & vUsu.Codigo & "," & DBSet(Cliente, "N") & "," & DBSet(RS!CodClien, "N") & ","
        SQLSubValues = SQLSubValues & DBSet(RS!Servicios, "N") & "," & DBSet(RS!Importe, "N")
        SQLSubValues = SQLSubValues & "),"
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    If SQLSubValues <> "" Then
        conn.Execute SQLSub & Mid(SQLSubValues, 1, Len(SQLSubValues) - 1)
    End If


'2º FACTURAMOS DE LA TABLA TEMPORAL
    FecFactu = txtcodigo(2).Text
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    SQL = "Select codigo1 codclien, sum(importe2) servicios, sum(importe3) importe FROM tmpinformes where codusu = " & vUsu.Codigo
    SQL = SQL & " group by 1"
    SQL = SQL & " order by 1"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Set miRsAux = New ADODB.Recordset
    
    'busco el minimo almacen y el minimo proveedor
    SQL = "select min(codalmac) from salmpr"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not miRsAux.EOF Then
        almac = miRsAux.Fields(0)
    End If
    
    miRsAux.Close
    
    SQL = "select min(codprove) from sprove"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not miRsAux.EOF Then
        Prove = miRsAux.Fields(0)
    End If
    
    Set miRsAux = Nothing
    
    
    b = True
    
    FacturasaImprimir = ""
    
    While Not RS.EOF And b
        Set cli = New CCliente
        Set fac = New CFactura
        
        
        If cli.LeerDatos(RS!CodClien, False) Then
            
            '[Monica]04/02/2015
            Label4.Caption = "Cliente: " & Format(RS!CodClien, "N") & " " & cli.Nombre
            DoEvents
            
            
            
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(TipoMovimiento) Then
                NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                ' si existe la factura incrementamos el contador
                Do
                    devuelve = DevuelveDesdeBDNew(conAri, "scafaccli", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", CStr(FecFactu), "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (TipoMovimiento)
                        NumFactu = vTipoMov.ConseguirContador(TipoMovimiento)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
            Else
                Exit Function
            End If
        
            ' calculo de bases iva de SERVICIOS
            iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArticServ, "T")
            vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T")
            porIvaServ = 0
            If vDevuelve <> "" Then porIvaServ = CCur(vDevuelve)
            fac.TipoIVA1 = iva
            
            Suplidos = 0
            DtoGnral = 0
            BaseivaServ = Round2(RS!Importe / (1 + (porIvaServ / 100)), 2)
            ImpivaServ = Round2(RS!Importe - BaseivaServ, 2)
            
            ' calculo de base iva de GASTOS ADMON
            iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArtGastosAdmon, "T")
            vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T")
            porIvaGtos = 0
            If vDevuelve <> "" Then porIvaGtos = CCur(vDevuelve)
            
            BaseivaGtos = cli.GastosAdmon
            ImpivaGtos = Round2(BaseivaGtos * porIvaGtos / 100, 2)
            
            ' Asignamos los importes a la factura
            If BaseivaGtos <> 0 Then
                If iva = fac.TipoIVA1 Then
                    fac.TipoIVA1 = iva
                    fac.BaseIVA1 = BaseivaGtos
                    fac.PorceIVA1 = porIvaGtos
                    fac.ImpIVA1 = ImpivaGtos
                Else
                    fac.TipoIVA2 = iva
                    fac.BaseIVA2 = BaseivaGtos
                    fac.PorceIVA2 = porIvaGtos
                    fac.ImpIVA2 = ImpivaGtos
                End If
            End If
            'el tipo de iva 1 esta asignado cuando se busca en tiposiva de la conta
            fac.PorceIVA1 = porIvaServ
            fac.BaseIVA1 = fac.BaseIVA1 + BaseivaServ
            fac.ImpIVA1 = fac.ImpIVA1 + ImpivaServ
            fac.BaseImp = BaseivaServ + BaseivaGtos
            fac.ImpGnral = DtoGnral
            fac.DtoGnral = cli.DtoGnral
            fac.BrutoFac = fac.BaseImp
            fac.Suplidos = Suplidos
            fac.TotalFac = BaseivaServ + ImpivaServ + Suplidos + BaseivaGtos + ImpivaGtos
            
            fac.codtipom = TipoMovimiento
            
            fac.FecFactu = FecFactu
            fac.LetraSerie = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", TipoMovimiento, "T")
'            NumFactu = vTipoMov.Contador
            fac.NumFactu = NumFactu
            FacturasaImprimir = FacturasaImprimir & NumFactu & ","
            
            'Cuenta Prevista de Cobro de las Facturas
            fac.BancoPr = txtcodigo(5).Text
            fac.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", fac.BancoPr, "N")
        
            fac.Cliente = cli.Codigo
            fac.NombreClien = cli.Nombre
            fac.DomicilioClien = cli.Domicilio
            fac.CPostal = cli.CPostal
            fac.Poblacion = cli.Poblacion
            fac.Provincia = cli.Provincia
            fac.NIF = cli.NIF
            fac.ForPago = cli.ForPago
        
            '[Monica]22/11/2013: iban
            fac.Iban = cli.Iban
        
            fac.Banco = cli.Banco
            fac.Sucursal = cli.Sucursal
            fac.DigControl = cli.DigControl
            fac.CuentaBan = cli.CuentaBan
            
            Mens = "Insertando en cabecera factura"
            'scafaccli
            SQL = "INSERT INTO scafaccli (codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,"
            SQL = SQL & "nifclien,codagent,codforpa,dtoppago,dtognral,brutofac,impdtopp,impdtogr,baseimp1,codigiv1,porciva1,"
            SQL = SQL & "imporiv1,baseimp2,codigiv2,porciva2,imporiv2,totalfac,intconta,coddirec,codbanco,codsucur,digcontr,cuentaba, numservi, suplidos, iban) VALUES ("
            SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "'," & DBSet(fac.Cliente, "N") & ","
            SQL = SQL & DBSet(cli.Nombre, "T") & "," & DBSet(cli.Domicilio, "T") & "," & DBSet(cli.CPostal, "T") & ","
            SQL = SQL & DBSet(cli.Poblacion, "T") & "," & DBSet(cli.Provincia, "T") & "," & DBSet(cli.NIF, "T") & "," & vParamAplic.PorDefecto_Agente
            SQL = SQL & "," & cli.ForPago & ",0," & DBSet(fac.DtoGnral, "N") & "," & DBSet(fac.BrutoFac, "N") & ",0," & DBSet(fac.ImpGnral, "N") & ","
            SQL = SQL & DBSet(fac.BaseIVA1, "N") & "," & DBSet(fac.TipoIVA1, "N")
            SQL = SQL & "," & DBSet(fac.PorceIVA1, "N") & "," & DBSet(fac.ImpIVA1, "N") & ","
            SQL = SQL & DBSet(fac.BaseIVA2, "N", "S") & "," & DBSet(fac.TipoIVA2, "N", "S") & "," & DBSet(fac.PorceIVA2, "N", "S") & ","
            SQL = SQL & DBSet(fac.ImpIVA2, "N", "S") & "," & DBSet(fac.TotalFac, "N") & ",0,NULL,"
            SQL = SQL & DBSet(fac.Banco, "N") & "," & DBSet(fac.Sucursal, "N") & "," & DBSet(fac.DigControl, "T") & "," & DBSet(fac.CuentaBan, "T") & ","
            SQL = SQL & DBSet(RS!Servicios, "N") & "," & DBSet(Suplidos, "N") & "," & DBSet(fac.Iban, "T") & ")"
        
        
            conn.Execute SQL
        

            o1 = DevuelveDesdeBD(conAri, "observa1", "scliente", "codclien", cli.Codigo, "N") '("select observa1 from scliente where codclien = " & DBSet(cli.Codigo, "N"))
            
            
            CodTraba = DevuelveDesdeBD(conAri, "codtraba", "straba", "login", vUsu.Login, "T")
            If CodTraba = "" Then CodTraba = DevuelveValor("select min(codtraba) from straba")
            Mens = "Insertando Albaran"
            
            
            sqlLineas = "Select importe1 codclien, importe2 servicios, importe3 importe FROM tmpinformes where codusu = " & vUsu.Codigo & " and codigo1 = " & DBSet(RS!CodClien, "N")
    
            Set RSLineas = New ADODB.Recordset
            RSLineas.Open sqlLineas, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            NumAlbar = 0
        
            While Not RSLineas.EOF
                'NumAlbar = NumAlbar + 1
                NumAlbar = RSLineas!CodClien
            
                SQL = "INSERT INTO scafaccli1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,"
                SQL = SQL & "codenvio,codtraba,codtrab2,observa1,observa2,observa3,observa4,observa5,codtrab1) VALUES ("
                SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV'," & DBSet(NumAlbar, "N") & ",'"
                SQL = SQL & Format(txtcodigo(86).Text, FormatoFecha) & "'," & vParamAplic.PorDefecto_Envio & "," & CodTraba
                SQL = SQL & "," & CodTraba & "," & DBSet(o1, "T") & "," & DBSet(o2, "T") & "," & DBSet(o3, "T") & ","
                SQL = SQL & DBSet(o4, "T") & "," & DBSet(o5, "T") & ",NULL)"
            
                conn.Execute SQL
                'slifac
            
                Mens = "Insertando linea de articulo de Servicios"
                'busco el nombre del articulo
                NomArtic = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArticServ, "T")
            
                BaseivaServ2 = Round2(RSLineas!Importe / (1 + (porIvaServ / 100)), 2)
            
                SQL = "INSERT INTO slifacCli (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
                SQL = SQL & "numbultos,cantidad,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel) VALUES ("
                SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV'," & DBSet(NumAlbar, "N") & ",0," & almac & ","
                If RSLineas!CodClien = RS!CodClien And chk_agrupados.Value = 0 Then
                    SQL = SQL & DBSet(vParamAplic.ArticServ, "T") & "," & DBSet(NomArtic, "T") & "," & DBSet(txtcodigo(4).Text, "T") & ",1," & DBSet(RSLineas!Servicios, "N") & " ," & DBSet(BaseivaServ2, "N") & ","
                Else
                    NomLote = ""
                    NomLote = DevuelveValor("select nomenvio from senvio inner join scliente on senvio.codenvio = scliente.codenvio where scliente.codclien = " & DBSet(RS!CodClien, "N"))
                    NomArtic = Trim(NomArtic) & " " & UCase(Format(txtcodigo(86).Text, "mmmm")) & " " & Year(CDate(txtcodigo(86).Text)) & " " & NomLote
                    Ampliaci = DevuelveValor("select nomclien from scliente where codclien = " & DBSet(RSLineas!CodClien, "N"))
                    SQL = SQL & DBSet(vParamAplic.ArticServ, "T") & "," & DBSet(NomArtic, "T") & "," & DBSet(Ampliaci, "T") & ",1," & DBSet(RSLineas!Servicios, "N") & " ," & DBSet(BaseivaServ2, "N") & ","
                End If
                SQL = SQL & DBSet(BaseivaServ2, "N") & "," & DBSet(BaseivaServ2, "N") & "," & DBSet(BaseivaServ2, "N") & ","
                SQL = SQL & DBSet(BaseivaServ2, "N") & ",0,0,'M'," & Prove & "," & DBSet(BaseivaServ2, "N") & ")"
            
                conn.Execute SQL
            
'[Monica]10/09/2014: partimos de la shilla y grabamos el destino pq lo hemos insertado
                '[Monica]31/03/2014: Insertamos las lineas de servicios que han actuado
                Dim SqlLin As String
                Dim Sql4 As String
                Dim RsLin As ADODB.Recordset
                
                Sql4 = "insert into scafaccli_serv (codtipom,numfactu,fecfactu,numlinea,fecha,hora,codsocio,numeruve,"
                Sql4 = Sql4 & "dirllama,observa1,impventa,idservic,observac2,codclien, destino, codautor, licencia, fecfinal, horfinal, codusuar) " '[Monica]03/10/2014: insertamos el destino
                                                                                                                                                    '[Monica]12/12/2014: faltaba insertar el codusuar
                SqlLin = "select " & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "', @rownum:=@rownum+1 AS rownum, fecha, hora, codsocio, numeruve, "
                SqlLin = SqlLin & " concat(dirllama,' ',numllama), observa2, impventa, idservic,  matricul, codclien, destino, codautor, licencia, fecfinal, horfinal, codusuar from shilla,(SELECT @rownum:=0) r "  '[Monica]03/10/2014: insertamos el destino
                SqlLin = SqlLin & " where " & cWhere
                SqlLin = SqlLin & " and codclien = " & DBSet(RSLineas!CodClien, "N") & " and facturadocliente = 0 and validado = 1 "
                SqlLin = SqlLin & " order by fecha, hora "
                
                
                conn.Execute Sql4 & SqlLin
                '31/03/2014: hasta aqui
            
            
            
'[Monica]10/09/2014: actualizamos la shilla
                SQL3 = "update shilla set facturadocliente = 1, facturad = 1 where " & cWhere & " and codclien = " & DBSet(RSLineas!CodClien, "N")
                conn.Execute SQL3
                
                RSLineas.MoveNext
            Wend
            
            Set RSLineas = Nothing
            
            
             'si hay gastos de admon insertamos una linea de albaran con los gastos
             If BaseivaGtos <> 0 And NumAlbar <> 0 Then
'                 NumAlbar = NumAlbar + 1
                 
                 Mens = "Insertando linea de articulo de Gastos"
             
                 NomArtic = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtGastosAdmon, "T")
             
                 SQL = "INSERT INTO slifacCli (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
                 SQL = SQL & "numbultos,cantidad,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel) VALUES ("
                 SQL = SQL & DBSet(TipoMovimiento, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV'," & DBSet(NumAlbar, "N") & ",1," & almac & ","
                 SQL = SQL & DBSet(vParamAplic.ArtGastosAdmon, "T") & "," & DBSet(NomArtic, "T") & ",1,1," & DBSet(BaseivaGtos, "N") & ","
                 SQL = SQL & DBSet(BaseivaGtos, "N") & "," & DBSet(BaseivaGtos, "N") & "," & DBSet(BaseivaGtos, "N") & ","
                 SQL = SQL & DBSet(BaseivaGtos, "N") & ",0,0,'M'," & Prove & "," & DBSet(BaseivaGtos, "N") & ")"
             
                 conn.Execute SQL
            End If
            
            'insertar en tesoreria
            fac.Agente = vParamAplic.PorDefecto_Agente
            
            b = fac.InsertarEnTesoreriaFACcli("", "Error al pasar a Tesoreria")
            'b = fac.InsertarEnTesoreriaFACcli("", "Error al pasar a tesoreria")
        
        
        
        
            If b Then vTipoMov.IncrementarContador (TipoMovimiento)
            
            Set vTipoMov = Nothing
            Set cli = Nothing
            Set fac = Nothing
        
        Else
            MsgBox "No existe el cliente " & cli.Codigo & " " & cli.Nombre
            b = False
        End If
        RS.MoveNext
    
    Wend
    
    Set RS = Nothing
    
' hasta aqui
    GenerarFacturasTeletaxiNew = True

EGenFactu:
    If Err.Number <> 0 Then
        Mens = "Insertando Movimiento." & vbCrLf & "----------------------------" & vbCrLf & Mens
        MuestraError Err.Number, Mens, Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        GenerarFacturasTeletaxiNew = True
        '[Monica]04/02/2015
        Me.Label4.Caption = ""
        DoEvents
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        GenerarFacturasTeletaxiNew = False
    End If
    
'If Err <> 0 Or Not B Then
'    conn.RollbackTrans
'    ConnConta.RollbackTrans
'    MsgBox "ERROR AL GENERAR FACTURAS:" & Err.Description
'    DesBloqueoManual ("FACCLI")
'    TerminaBloquear
'End If
End Function

