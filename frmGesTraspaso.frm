VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGesTraspaso 
   Caption         =   "Traspaso TaxiTronic"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Desde Fichero Excel"
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
      Index           =   0
      Left            =   270
      TabIndex        =   17
      Tag             =   "Facturado|N|N|0|1|sclien|essocio|||"
      Top             =   1590
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   210
      TabIndex        =   13
      Top             =   750
      Width           =   7485
      Begin VB.OptionButton Option1 
         Caption         =   "Servicios Socios"
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
         Left            =   4620
         TabIndex        =   16
         Top             =   300
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Llamadas"
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
         Left            =   630
         TabIndex        =   15
         Top             =   300
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Servicios Clientes"
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
         Left            =   2280
         TabIndex        =   14
         Top             =   300
         Width           =   2115
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3960
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
   Begin VB.CommandButton cmdCancel 
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
      Left            =   6510
      TabIndex        =   4
      Top             =   4920
      Width           =   1135
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
      Left            =   5220
      TabIndex        =   3
      Top             =   4920
      Width           =   1135
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   3870
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   195
      TabIndex        =   6
      Top             =   2040
      Width           =   7455
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Top             =   780
         Width           =   1215
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
         Index           =   4
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1230
         Width           =   4605
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   600
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   930
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Concepto"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1230
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   930
         ToolTipText     =   "Buscar Fichero"
         Top             =   390
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6840
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   2
      Left            =   210
      TabIndex        =   10
      Top             =   4590
      Width           =   7425
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   9
      Top             =   4260
      Width           =   7425
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Traspaso TaxiTronic"
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
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmGesTraspaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 316

Private menErrProceso As String 'mensaje final del proceso actualizacion de precios
Dim Vehiculo As String
Dim Fecha As String
Dim hora As String
Dim pulsadoCancelar As Boolean
Dim procesoFinalizado As Boolean
Dim procesoCancelado As Boolean
Dim Contador As Currency

Dim indCodigo As Long


Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim CargarServicios As Boolean

Dim kCampo As Integer


Private Function rsContador(CADENA As String) As Currency
    
    rsContador = 0
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open CADENA, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        rsContador = miRsAux.Fields(0)
    End If
    miRsAux.Close
    
End Function



Private Sub cmdAceptar_Click()
Dim cadSel As String
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim encontrado As String
Dim total As Currency
Dim Contador As Currency
Dim Sql As String
Dim cadTabla As String


    ' solo puede haber una persona ejecutando el proceso de traspaso

    If Not DatosOk Then Exit Sub
    
    If Text1.Text = "" Then
        cadSel = "Seleccione un fichero de importacion"
    Else
        If Dir(Text1.Text, vbArchive) = "" Then cadSel = "No existe el archivo: " & Text1.Text
    End If
    If cadSel <> "" Then
        MsgBox cadSel, vbExclamation
        Exit Sub
    End If
        
        
        
'[Monica]10/09/2014: de momento comentado pq partimos de la shilla
'    '[Monica]31/03/2014: miramos si existe el fichero de traspaso de contabilidad
'    CargarServicios = False
'    If Dir(CurDir(Text1.Text) & "\Traspaso contabilidad.txt", vbArchive) = "" Then
'        If MsgBox("No existe el fichero de Contabilidad. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbYesNo) = vbNo Then
'            Exit Sub
'        End If
'    Else
'        CargarServicios = True
'
'        ' borramos la tabla auxiliar
'        SQL = "delete from tmpservicios"
'        conn.Execute SQL
'        ' la cargamos con los servicios que nos llegan
'        SQL = "load data local infile '" & Replace(CurDir(Text1.Text) & "\Traspaso contabilidad.txt", "\", "\\") & "' into table `tmpservicios` fields escaped by '\\' terminated by ';' enclosed by '""'  lines terminated by '\r\n' ( `nroservicio`, `numeruve`, `campo3`, `matricul`, `codclien`, `nomclien`, `campo7`, `fecha`, `hora`, `origen`, `destino`, `campo12`, `importe`, `tipo` )"
'        conn.Execute SQL
'    End If
        
        
    If Me.Option1(1).Value Or Option1(2).Value Then   ' traspaso de facturacion de clientes
        If ComprobarFichero(Option1(1).Value) Then
            cadTabla = "tmpinformes"
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            'Añadir el parametro de Empresa
            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
            
            Sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
            
            If TotalRegistros(Sql) <> 0 Then
'                If HayRegParaInforme(cadTABLA, cadSelect) Then
                MsgBox "Hay errores en el Traspaso de Datos. Debe corregirlos previamente.", vbExclamation
                cadTitulo = "Errores de Traspaso de Datos"
                cadNombreRPT = "rErroresTraspaso.rpt"
                
                
                LlamarImprimir
                
                '[Monica]31/03/2014: dejamos continuar
                If MsgBox("¿ Desea continuar con el proceso ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
                
                If TraspasoFichero(Option1(1).Value) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click
                End If
                'hasta aqui
            Else
                If TraspasoFichero(Option1(1).Value) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click
                End If
            End If
        End If
    
    
    ElseIf Option1(0).Value Then ' traspaso de llamadas
        
            '[Monica]26/12/2017: solo si viene de excel, añado esta condicion
            If Check1(0).visible And Check1(0).Value = 1 Then
                
                '[Monica]13/12/2017: para el caso d que la tabla intermedia esté cargada
                
                Dim NRegs As Long
                Dim Empezar As Boolean
                
                NRegs = DevuelveValor("select count(*) from tmptaxi where error1 = 1")
                Empezar = False
                If NRegs <> 0 Then
                    Label1(0).Caption = ""
                    Label1(2).Caption = ""
                    If MsgBox("Hay registros en la tabla intermedia." & vbCrLf & vbCrLf & "¿ Desea empezar el proceso ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then Empezar = True
                    If Not Empezar Then
        '                MostrarTablas
                        VerErrores
                        If MsgBox("¿Desea continuar la actualización de las tablas?.", vbQuestion + vbYesNo) = vbYes Then
                            ActualizarTabla
                            BorrarTablas
                        End If
                        Exit Sub
                    End If
                End If
            End If
            
            BorrarTablas
            'hasta aqui


            'Llegados aqui, procesamos el fichero
            Screen.MousePointer = vbHourglass
        '    b = ProcesarFichero

            '[Monica] Para el caso de radiotaxi se trabaja con un fichero excel
            If Check1(0).visible And Check1(0).Value = 1 Then
                If Dir(App.Path & "\trasaritaxi.z") <> "" Then Kill App.Path & "\trasaritaxi.z"

                Shell App.Path & "\trasaritaxi.exe /I|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|" & Text1.Text & "|", vbNormalFocus

                While Dir(App.Path & "\trasaritaxi.z") = ""
                    Me.Label1(0).Caption = "Procesando Insercion "
                    DoEvents

                    Espera 1
                Wend

                b = True
            Else
                b = ProcesarFichero_new
            End If
        

            If b Then
                'verificamos que los numeruve esten asociados a algun socio
                ProgressBar1.Value = 0
                Contador = 0
                Label1(0).Caption = ""
                Set Rs = New ADODB.Recordset
                Sql = "select * from tmptaxi where error1 = 0 group by numeruve"
                Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
                total = rsContador("select count(distinct(numeruve)) from tmptaxi where error1=0")
                Label1(2).Caption = "Verificando códigos de socios."
                Label1(2).Refresh
        
        
                '[Monica]26/12/2017: solo si viene de excel, añado esta condicion
                If Check1(0).visible And Check1(0).Value = 1 Then
        
                    While Not Rs.EOF
                        Contador = Contador + 1
                        ProgressBar1.Value = (Contador * 100) / total
                        DoEvents
                        'Label1(0).Caption = Round(ProgressBar1.Value, 2) & "%"
                        Label1(0).Caption = Round2(ProgressBar1.Value, 0) & " %"
                        Label1(0).Refresh
                        
'[Monica]09/02/2018: quito la condicion Trim(vParam.CifEmpresa) = "B98877806"
'                    pq ahora tele es igual que radio en el numeruve lleva la licencia
                        
'                        ' me viene la licencia (caso de Radio Taxi en la V llevo la licencia)
'                        If Trim(vParam.CifEmpresa) = "B98877806" Then
                        
'[Monica]18/06/2018: obtenemos el codigo de socio pero comparando la fecha de servicio con la de alta y baja de socio
'                            encontrado = DevuelveDesdeBD(conAri, "codclien", "sclien", "numeruve", Rs!NumerUve, "T")
                            encontrado = EncontradoSocio(DBLet(Rs!NumerUve, "N"), Rs!Fecha)
                            
                            '[Monica]28/02/2018: para el caso de cordoba si no lo encuentra lo marco como error
                            'B = Updatear(Rs!NumerUve, encontrado, False)
                            '[Monica]19/11/2018: Entra Sevilla
                            b = Updatear(Rs!NumerUve, encontrado, vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3)
                        
'                        Else
'
'                            ' Para el caso de TELETAXI, busco los codigos de socio que llevan o llevaron esa licencia
'                            '
'                            '   si los encuentro:
'                            '                   busco el codigo de socio que tenga la v activa
'                            '                       si lo encuentro: es ese codigo
'                            '                       si no          : lo marco como erroneo, pq es de tele pero no está activo
'                            '   si no los encuentro:
'                            '                   el socio es de RADIO, y lo marco como tal
'
'                            encontrado = DevuelveDesdeBD(conAri, "codclien", "sclien", "licencia", Rs!NumerUve, "T")
'
'                            If encontrado <> "" Then
'                                ' pq me viene la licencia
'                                Dim rs4 As ADODB.Recordset
'                                Dim Sql4 As String
'                                Set rs4 = New ADODB.Recordset
'                                Sql4 = "select codclien from sclien where licencia = " & DBSet(Rs!NumerUve, "N") & " and not numeruve is null and numeruve <> 0"
'                                rs4.Open Sql4, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'                                encontrado = ""
'
'                                If Not rs4.EOF Then
'                                    encontrado = DBLet(rs4.Fields(0))
'                                End If
'                                Set rs4 = Nothing
'
'                                b = Updatear(Rs!NumerUve, encontrado, True)
'                            Else
'                                b = Updatear(Rs!NumerUve, encontrado, False)
'                            End If
'
'                        End If
                                                
                        Rs.MoveNext
                    Wend
                    
                Else
                
                    While Not Rs.EOF
                        Contador = Contador + 1
                        ProgressBar1.Value = (Contador * 100) / total
                        DoEvents
                        'Label1(0).Caption = Round(ProgressBar1.Value, 2) & "%"
                        Label1(0).Caption = Round2(ProgressBar1.Value, 0) & " %"
                        Label1(0).Refresh
                        
'[Monica]18/06/2018: obtenemos el codigo de socio pero comparando la fecha de servicio con la de alta y baja de socio
'                       encontrado = DevuelveDesdeBD(conAri, "codclien", "sclien", "numeruve", Rs!NumerUve, "T")
                        encontrado = EncontradoSocio(DBLet(Rs!NumerUve, "N"), Rs!Fecha)
                        
                        '[Monica]28/02/2018
                        'B = Updatear(Rs!NumerUve, encontrado, False)
                        '[Monicax]19/11/2018: Entra Sevilla
                        b = Updatear(Rs!NumerUve, encontrado, vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3)
                        Rs.MoveNext
                    Wend
                
                End If
                
                Rs.Close
                Label1(0).Caption = ""
                Label1(0).Refresh
                
                '[Monica]12/12/2017: por el tema de fusion de empresas, SOLO SI VIENE DE EXCEL
                '                    si el fichero es de la otra empresa ponemos que el cliente es el gros
                If Check1(0).Value = 1 Then
                    If b Then
                        If ComprobarCero(vParamAplic.EmpresaTaxitronic) <> 0 Then
                            Label1(2).Caption = "Modificando códigos de cliente de otra empresa"
                            Label1(2).Refresh
                            
                            Sql = "update tmptaxi set codclien = " & DBSet(vParamAplic.ClienteCooperativa, "N")
                            Sql = Sql & " where error1 = 0 and empresa <> " & vParamAplic.EmpresaTaxitronic
                            Sql = Sql & " and not codclien is null "
                            b = EjecutarSQL(Sql)
                        End If
                    End If
                    '[Monica]12/12/2017: eliminamos todos aquellas llamadas que no son de nuestros clientes ni lo ha hecho un asociado nuestro
                    If b Then
                        Label1(2).Caption = "Eliminando registros que no se tienen que procesar"
                        Label1(2).Refresh
                        
                        Sql = "delete from tmptaxi where codclien = " & DBSet(vParamAplic.ClienteCooperativa, "N")
                        Sql = Sql & " and codsocio = " & DBSet(vParamAplic.SocioCooperativa, "N")
                        Sql = Sql & " and empresa <> " & vParamAplic.EmpresaTaxitronic
                    
                        b = EjecutarSQL(Sql)
                    End If
                End If
                
                'buscamos en la misma tabla que los registros no esten duplicados
                If b Then
                    ProgressBar1.Value = 0
                    Contador = 0
        
                    Set Rs = New ADODB.Recordset
                    Sql = "select numeruve,fecha,hora, count(*) from tmptaxi where error1 = 0 group by 1,2,3 having count(*) > 1"
                    Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
                    total = rsContador("select count(*) from (" & Sql & ") aaalias ")  'tmptaxi where error1 <> 1")
                    Label1(2).Caption = "eliminando(II) duplicidad de registros en el fichero."
                    Label1(2).Refresh
                    While Not Rs.EOF
                        Contador = Contador + 1
                        ProgressBar1.Value = Round2((Contador * 100) / total, 0)
                        DoEvents
                        Label1(0).Caption = Round(ProgressBar1.Value, 0) & " %"
                        Label1(0).Refresh
                        '[Monica]23/03/2018: ya no funcionamos con segundos en la hora
                        Sql = "numeruve=" & Rs!NumerUve & " and fecha=" & DBSet(Rs!Fecha, "F") & " and mid(hora,1,5)='" & Format(Rs!hora, "hh:mm") & "' "
                        Sql = Sql & " and impventa = 0 and codclien =0 "
                        
                        Dim Ident As Long
                        
                        Ident = DevuelveValor("select id from tmptaxi where " & Sql)
                        
                        If Ident <> 0 Then
                            Sql = "delete from tmptaxi where id = " & DBSet(Ident, "N")
                            conn.Execute Sql
                        Else
                            'Stop
                        End If
       
        '                End If
                        Rs.MoveNext
                    Wend
                    Rs.Close
        
                    
                    '
                    Sql = "select numeruve,fecha,hora, count(*) from tmptaxi where error1 = 0 group by 1,2,3 having count(*) > 1"
                    Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
                    total = rsContador("select count(*) from (" & Sql & ") aaalias ")  'tmptaxi where error1 <> 1")
                    Label1(2).Caption = "Verificando duplicidad de registros en el fichero."
                    Label1(2).Refresh
                    While Not Rs.EOF
                        Contador = Contador + 1
                       ' ProgressBar1.Value = Round2((Contador * 100) / total, 0)
                        DoEvents
                        Label1(0).Caption = Contador
                        Label1(0).Refresh
                        '[Monica]23/03/2018: no funcionamos con segundos
                        Sql = "numeruve=" & Rs!NumerUve & " and fecha=" & DBSet(Rs!Fecha, "F") & " and mid(hora,1,5)='" & Format(Rs!hora, "hh:mm") & "' "
                        
                        
                        
        '                If SituarDataMULTI(Adodc1, SQL, encontrado) Then
        
                            'esta, entonces es repetido
                            Sql = "UPDATE tmptaxi set error1=1,error='Registro duplicado' where " & Sql
                            conn.Execute Sql
        '                End If
                        Rs.MoveNext
                    Wend
                    Rs.Close
        
        
'[Monica]18/06/2018: quito esto ya van todos por licencia
'                    '[Monica]28/12/2017: para el caso de Tele y Alfa 6 pongo el numero de V correcto y que no sea cordoba
'                    If Trim(vParam.CifEmpresa) <> "B98877806" And Check1(0).Value = 1 And vParamAplic.Cooperativa <> 2 Then
'                        Dim nUve As Long
'
'                        Sql = "select codsocio from tmptaxi where error1 = 0 and codsocio <> " & vParamAplic.SocioCooperativa & " group by 1"
'
'                        Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
'
'                        total = rsContador("select count(*) from (" & Sql & ") aaalias ")  'tmptaxi where error1 <> 1")
'                        Label1(2).Caption = "Modificando Vehículo en registros del fichero."
'                        Label1(2).Refresh
'
'                        While Not Rs.EOF
'                            Contador = Contador + 1
'                           ' ProgressBar1.Value = Round2((Contador * 100) / total, 0)
'                            DoEvents
'                            Label1(0).Caption = Contador
'                            Label1(0).Refresh
'
'                            Sql = "select numeruve from sclien where codclien = " & DBSet(Rs!codSocio, "N")
'                            nUve = DevuelveValor(Sql)
'
'                            Sql = "UPDATE tmptaxi set numeruve = " & DBSet(nUve, "N") & " where codsocio = " & DBSet(Rs!codSocio, "N") & " and error1 = 0 "
'                            conn.Execute Sql
'
'                            Rs.MoveNext
'                        Wend
'                        Rs.Close
'
'                    End If
                    
        
                    'ahora vamos a buscar en la tabla shilla
                    Sql = "select * from tmptaxi where error1 = 0"
                    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    ProgressBar1.Value = 0
                    Contador = 0
                    total = rsContador("select count(*) from tmptaxi where error1 = 0")
                    Label1(2).Caption = "Verificando duplicidad de registros en la tabla."
                    Label1(2).Refresh

                    While Not Rs.EOF
                        Contador = Contador + 1
                        ProgressBar1.Value = Round2((Contador * 100) / total, 0)

                        'Label1(0).Caption = Round2(ProgressBar1.Value, 2) & "%"
                        Label1(0).Caption = Round2(ProgressBar1.Value, 0) & "%"
                        Label1(0).Refresh
                        DoEvents
                        
                        '[Monica]11/11/2014: si aparece en la shilla no damos error, updateamos (antes no la introducíamos la marcabamos como errónea)
                        '                    SOLO EN EL CASO DE QUE NO ESTE LIQUIDADA NI FACTURADA
                        '                    En el caso de que esté liquidada o facturada la marcamos como erronea
'                        Sql = "fecha='" & Format(RS!Fecha, FormatoFecha) & "' and hora='" & Format(RS!hora, FormatoHora) & "' and numeruve"
'                        encontrado = DevuelveDesdeBD(conAri, "codsocio", "shilla", Sql, RS!NumerUve, "N")
'                        If encontrado <> "" Then

                        '[Monica]23/03/2018: sin segundos
                        Sql = "select count(*) from shilla where numeruve = " & DBSet(Rs!NumerUve, "N") & " and fecha = " & DBSet(Rs!Fecha, "F") & " and mid(hora,1,5) = " & DBSet(Rs!hora, "HH") & " and (facturad=1 and abonados=1 and validado=1)"
                        If TotalRegistros(Sql) <> 0 Then
                            '[Monica]31/10/2017: los marco como 2 para no mostrarlos
                            'esta entonces es repetido
                            Sql = "UPDATE tmptaxi set error1=2,error='Registro duplicado' where id=" & Rs!Id
                            conn.Execute Sql
                        End If
                        Rs.MoveNext
                    Wend
                    Rs.Close
                End If
            End If
            If procesoCancelado Then
                MsgBox "Traspaso cancelado", vbInformation
            ElseIf procesoFinalizado Then
                MsgBox "Traspaso finalizado.", vbInformation
            End If
            cmdAceptar.Enabled = False
            cmdCancel.Caption = "Salir"
            Label1(0).Caption = ""
            Label1(2).Caption = ""
        
            menErrProceso = "" 'Reestablezco esta variable reutilizada
            Screen.MousePointer = vbDefault
            If Not b Then
                EnTomaDeDatos
                Exit Sub
            End If
            MostrarTablas
'[Monica]13/12/2017: el borrar tablas lo dejo dentro de mostrar tablas si actualizan
'            BorrarTablas
            cmdCancel.Caption = "&Cancelar"
     End If
End Sub

Private Function EncontradoSocio(nUve As Long, FechaServ As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eEncontradoSocio
    
    EncontradoSocio = ""

    Sql = "select codclien from sclien where numeruve = " & DBSet(nUve, "N") & " and fechaalt <= " & DBSet(FechaServ, "F")
    Sql = Sql & " and (fechabaj is null or fechabaj >=" & DBSet(FechaServ, "F") & ")"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        EncontradoSocio = DBLet(Rs!CodClien, "N")
    End If
    Set Rs = Nothing
    Exit Function
    
eEncontradoSocio:
    MuestraError Err.Number, "Encontrado Socio", Err.Description
End Function


Private Function Updatear(Vehiculo, encontrado As String, LicenciaSinV As Boolean) As Boolean
Dim Sql As String

On Error GoTo EUp

Updatear = False

If encontrado = "" Then
'[Monica]12/12/2017: ahora si no encuentro el socio que lleva ese numero de vehiculo es que es de la otra empresa
'                    si viene de fichero plano lo marco como error
    If Check1(0).Value = 0 Or LicenciaSinV Then
        Sql = "UPDATE tmptaxi set error1=1,error='Ningun socio tiene asociado este codigo de vehiculo' where numeruve=" & Vehiculo
    Else
        Sql = "UPDATE tmptaxi set codsocio=" & vParamAplic.SocioCooperativa & " where numeruve=" & Vehiculo
    End If
Else
    Sql = "UPDATE tmptaxi set codsocio=" & CInt(encontrado) & " where numeruve=" & Vehiculo
End If

conn.Execute Sql

Updatear = True

EUp:
If Err.Number <> 0 Then
    Updatear = False
End If

End Function

Private Sub MostrarTablas()
Dim Rs As ADODB.Recordset
Dim Sql As String

    Set Rs = New ADODB.Recordset
    Sql = "select * from tmptaxi where error1=1"
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        If MsgBox("Se ha procesado el fichero correctamente. ¿Desea continuar?.", vbQuestion + vbYesNo) = vbYes Then
            ActualizarTabla
            BorrarTablas
        End If
    Else
        If MsgBox("Ha habido errores en el proceso del fichero. ¿Desea ver los errores?.", vbQuestion + vbYesNo) = vbYes Then
            VerErrores
        End If
        If MsgBox("¿Desea continuar la actualización de las tablas?.", vbQuestion + vbYesNo) = vbYes Then
            ActualizarTabla
            BorrarTablas
        End If
    End If
    Rs.Close
    Set Rs = Nothing
End Sub

Private Sub ActualizarTabla()
Dim Sql As String
Dim SQL1 As String
Dim Rs As ADODB.Recordset
Dim linea As String
Dim values As String
Dim Contador As Currency
Dim total As Currency
Dim SqlUpdate As String
Dim cWhere As String

Dim ImporteCompra As Currency
Dim Suplemento As Currency
Dim DtoCli As String

    On Error GoTo EActua
    
    Screen.MousePointer = vbHourglass
    
    
    
    Set Rs = New ADODB.Recordset
    Sql = "select fecha,hora,codsocio,numeruve,codclien,codusuar,nomclien,dirllama,"
    Sql = Sql & "numllama,puerllama,ciudadre,tipservi,telefono,observac2,codautor,observa1,licencia,"
    Sql = Sql & "matricul,idservic,opereser,opedespa,estado,observa2,fecreser,horreser,fecaviso,"           '[Monica]08/02/2019: añadimos el codigo de usuario para Sevilla
    Sql = Sql & "horaviso,fecllega,horllega,fecocupa,horocupa,fecfinal,horfinal,importtx,impcompr,"         '[Monica]03/10/2014: añadimos el destino
    '[Monica]28/12/2017: al añadir la situacion 2, ésta tambien es erronea, luego no debe entrar, solo entran situacion = 0
    Sql = Sql & "extcompr,impventa,extventa,distanci,suplemen,imppeaje,imppropi,facturad,abonados,validado, destino, empresa, codigousu from tmpTaxi where error1=0"
    
    Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
    total = rsContador("select count(*) from tmpTaxi where error1=0")
    If total = 0 Then
        MsgBox "No hay datos para actualizar.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    conn.BeginTrans
    
    Contador = 0
    linea = ""
    values = ""
    ProgressBar1.Value = 0
    Label1(2).Caption = "Actualizando Bases de datos"
    Label1(2).Refresh
    
    Sql = "INSERT INTO shilla (fecha,hora,codsocio,numeruve,codclien,codusuar,nomclien,dirllama,"
    Sql = Sql & "numllama,puerllama,ciudadre,tipservi,telefono,observac2,codautor,observa1,licencia,"
    Sql = Sql & "matricul,idservic,opereser,opedespa,estado,observa2,fecreser,horreser,fecaviso,"
    Sql = Sql & "horaviso,fecllega,horllega,fecocupa,horocupa,fecfinal,horfinal,importtx,impcompr,"
    Sql = Sql & "extcompr,impventa,extventa,distanci,suplemen,imppeaje,imppropi,facturad,abonados,validado, destino, empresa, codigousu) values "
    
    '[Monica]11/11/2014: dejamos actualizar si no esta liquidada ni facturada
    SqlUpdate = "update shilla set "
    
    
    While Not Rs.EOF
        Contador = Contador + 1
        ProgressBar1.Value = Round2((Contador * 100) / total, 0)
        
        cWhere = "numeruve = " & DBSet(Rs!NumerUve, "N") & " and fecha = " & DBSet(Rs!Fecha, "F") & " and mid(hora,1,5) = " & DBSet(Rs!hora, "HH")
        
        If ExisteEnShilla(cWhere) Then
        
'[Monica]02/02/2018: levanto la condicion de que solo se actualicen las llamadas que sean de credito (ahora se actualizan todas)
            
            '[Monica]13/11/2014: sólo en el caso de que sea de credito actualizamos
'            If EsdeCredito(cWhere) Then
'                linea = " fecha = " & DBSet(Rs!Fecha, "F")
'                linea = linea & ",hora = " & DBSet(Rs!hora, "H")

'[Monica]18/06/2018: no modificamos el codigo de socio si es una actualizacion (update pq el registro ya estaba en la shilla)
'                linea = "codsocio = " & DBSet(Rs!codSocio, "N") & ","

            '[Monica]18/02/2019: en el caso de que esté facturado y sea teletaxi then
            If EsTeletaxi And EstaFacturadoCliente(cWhere) Then
                '[Monica]28/02/2018: si esta marcado que se marcan como validados, independientemente de lo que venga
                linea = ""
                If vParamAplic.MarcarValidados Then
                    linea = linea & " validado = 1 "
                Else
                    If Not IsNull(Rs!CodClien) Then
                        If MarcarValidadosEnCliente(Rs!CodClien) Then
                            linea = linea & " validado = 1 "
                        Else
                            linea = linea & " validado = " & DBSet(Rs!validado, "N")
                        End If
                    Else
                        linea = linea & " validado = " & DBSet(Rs!validado, "N")
                    End If
                End If
                linea = linea & " where " & cWhere
                
                conn.Execute SqlUpdate & linea
                
            Else
                linea = "numeruve = " & DBSet(Rs!NumerUve, "N")
                linea = linea & ",codclien = " & DBSet(Rs!CodClien, "N")
                linea = linea & ",codusuar = " & DBSet(Rs!codusuar, "T")
                linea = linea & ",nomclien = " & DBSet(Rs!nomclien, "T")
                linea = linea & ",dirllama = " & DBSet(Rs!dirllama, "T")
                linea = linea & ",numllama = " & DBSet(Rs!numllama, "T")
                linea = linea & ",puerllama = " & DBSet(Rs!puerllama, "T")
                linea = linea & ",ciudadre = " & DBSet(Rs!ciudadre, "T")
                linea = linea & ",tipservi = " & DBSet(Rs!tipservi, "N")
                linea = linea & ",telefono = " & DBSet(Rs!Telefono, "T")
                linea = linea & ",observac2 = " & DBSet(Rs!observac2, "T")
                linea = linea & ",codautor = " & DBSet(Rs!codautor, "T")
                
                '[Monica]15/01/2018: cambio el orden de las observaciones si estamos en traspaso del alfa 6
                If EsTeletaxi And Check1(0).Value = 1 Then
                    linea = linea & ",observa1 = " & DBSet(Rs!observa2, "T")
                Else
                    linea = linea & ",observa1 = " & DBSet(Rs!observa1, "T")
                End If
                
                linea = linea & ",licencia = " & DBSet(Rs!Licencia, "T")
                linea = linea & ",matricul = " & DBSet(Rs!matricul, "T")
                linea = linea & ",idservic = " & DBSet(Rs!idservic, "T")
                linea = linea & ",opereser = " & DBSet(Rs!opereser, "T")
                linea = linea & ",opedespa = " & DBSet(Rs!opedespa, "T")
                linea = linea & ",estado = " & DBSet(Rs!Estado, "T")
                
                '[Monica]15/01/2018: cambio el orden de las observaciones si estamos en traspaso del alfa 6
                If EsTeletaxi And Check1(0).Value = 1 Then
                    linea = linea & ",observa2 = " & DBSet(Rs!observa1, "T")
                Else
                    linea = linea & ",observa2 = " & DBSet(Rs!observa2, "T")
                End If
                linea = linea & ",fecreser = " & DBSet(Rs!fecreser, "F")
                linea = linea & ",horreser = " & DBSet(Rs!horreser, "H")
                linea = linea & ",fecaviso = " & DBSet(Rs!fecaviso, "F")
                linea = linea & ",horaviso = " & DBSet(Rs!horaviso, "H")
                linea = linea & ",fecllega = " & DBSet(Rs!fecllega, "F")
                linea = linea & ",horllega = " & DBSet(Rs!horllega, "H")
                linea = linea & ",fecocupa = " & DBSet(Rs!fecocupa, "F")
                linea = linea & ",horocupa = " & DBSet(Rs!horocupa, "H")
                linea = linea & ",fecfinal = " & DBSet(Rs!fecfinal, "F")
                linea = linea & ",horfinal = " & DBSet(Rs!horfinal, "H")
                linea = linea & ",importtx = " & DBSet(Rs!importtx, "N")
                
                '[Monica]28/05/2018: para el caso de que el cliente tenga descuento se le descuenta al socio
'                linea = linea & ",impcompr = " & DBSet(Rs!impcompr, "N")
                '[Monica]04/02/2019: todos menos Sevilla
                If vParamAplic.Cooperativa <> 3 Then
                    DtoCli = DevuelveDesdeBDNew(conAri, "scliente", "dtognral", "codclien", DBLet(Rs!CodClien, "N"), "N")
                    ImporteCompra = Round2(DBLet(Rs!impcompr, "N") * (100 - ImporteSinFormato(ComprobarCero(DtoCli))) / 100, 2)
                    linea = linea & ",impcompr = " & DBSet(ImporteCompra, "N")
                Else
                    linea = linea & ",impcompr = " & DBSet(Rs!impcompr, "N")
                End If
                
                linea = linea & ",extcompr = " & DBSet(Rs!extcompr, "N")
                linea = linea & ",impventa = " & DBSet(Rs!impventa, "N")
                linea = linea & ",extventa = " & DBSet(Rs!extventa, "N")
                linea = linea & ",distanci = " & DBSet(Rs!distanci, "N")
                
                '[Monica]06/07/2018: el descuento en cordoba tambien se aplica en el suplemento
                Suplemento = Round2(DBLet(Rs!suplemen, "N") * (100 - ImporteSinFormato(ComprobarCero(DtoCli))) / 100, 2)
                linea = linea & ",suplemen = " & DBSet(Suplemento, "N")
                
                linea = linea & ",imppeaje = " & DBSet(Rs!imppeaje, "N")
                linea = linea & ",imppropi = " & DBSet(Rs!imppropi, "N")
                linea = linea & ",facturad = " & DBSet(Rs!facturad, "N")
                linea = linea & ",abonados = " & DBSet(Rs!abonados, "N")
                
                '[Monica]28/02/2018: si esta marcado que se marcan como validados, independientemente de lo que venga
                If vParamAplic.MarcarValidados Then
                    linea = linea & ",validado = 1 "
                Else
                    If Not IsNull(Rs!CodClien) Then
                        If MarcarValidadosEnCliente(Rs!CodClien) Then
                            linea = linea & ",validado = 1 "
                        Else
                            linea = linea & ",validado = " & DBSet(Rs!validado, "N")
                        End If
                    Else
                        linea = linea & ",validado = " & DBSet(Rs!validado, "N")
                    End If
                End If
                
                linea = linea & ",destino = " & DBSet(Rs!Destino, "T")
                linea = linea & ",empresa = " & DBSet(Rs!Empresa, "N")
                '[Monica]08/02/2019:
                linea = linea & ",codigousu = " & DBSet(Rs!codigousu, "N")
                
                
                linea = linea & " where " & cWhere
                
                conn.Execute SqlUpdate & linea
            End If
        Else
            
            If IsNull(Rs!Fecha) Then
                linea = "(NULL,"
            Else
                linea = "(" & DBSet(Rs!Fecha, "F") & ","
            End If
            If IsNull(Rs!hora) Then
                linea = linea & "NULL,"
            Else
                linea = linea & "'" & Format(Rs!hora, FormatoHora2) & "',"
            End If
            If IsNull(Rs!codSocio) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!codSocio, "N") & ","
            End If
            If IsNull(Rs!NumerUve) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!NumerUve, "N") & ","
            End If
            If IsNull(Rs!CodClien) Or Rs!CodClien = 0 Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!CodClien, "N") & ","
            End If
            If IsNull(Rs!codusuar) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!codusuar, "T") & ","
            End If
            If IsNull(Rs!nomclien) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!nomclien, "T") & ","
            End If
            If IsNull(Rs!dirllama) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!dirllama, "T") & ","
            End If
            If IsNull(Rs!numllama) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!numllama, "T") & ","
            End If
            If IsNull(Rs!puerllama) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!puerllama, "T") & ","
            End If
            If IsNull(Rs!ciudadre) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!ciudadre, "T") & ","
            End If
            If IsNull(Rs!tipservi) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!tipservi, "N") & ","
            End If
            If IsNull(Rs!Telefono) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!Telefono, "T") & ","
            End If
            If IsNull(Rs!observac2) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!observac2, "T") & ","
            End If
            If IsNull(Rs!codautor) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!codautor, "T") & ","
            End If
            
            '[Monica]15/01/2018: cambio el orden de las observaciones si estamos en traspaso del alfa 6, solo para tele
            '                    pq tenia entradas anteriores que tenia q facturar
            If EsTeletaxi And Check1(0).Value = 1 Then
                If IsNull(Rs!observa2) Then
                    linea = linea & "NULL,"
                Else
                    linea = linea & DBSet(Rs!observa2, "T") & ","
                End If
            Else
            ' se queda como estaba
                If IsNull(Rs!observa1) Then
                    linea = linea & "NULL,"
                Else
                    linea = linea & DBSet(Rs!observa1, "T") & ","
                End If
            End If
            If IsNull(Rs!Licencia) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!Licencia, "T") & ","
            End If
            If IsNull(Rs!matricul) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!matricul, "T") & ","
            End If
            If IsNull(Rs!idservic) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!idservic, "T") & ","
            End If
            If IsNull(Rs!opereser) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!opereser, "T") & ","
            End If
            If IsNull(Rs!opedespa) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!opedespa, "T") & ","
            End If
            If IsNull(Rs!Estado) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!Estado, "T") & ","
            End If
            
            '[Monica]15/01/2018: cambio el orden de las observaciones si estamos en traspaso del alfa 6, solo para tele
            '                    pq tenia entradas anteriores que tenia q facturar
            If EsTeletaxi And Check1(0).Value = 1 Then
                If IsNull(Rs!observa1) Then
                    linea = linea & "NULL,"
                Else
                    linea = linea & DBSet(Rs!observa1, "T") & ","
                End If
            Else
            ' se queda como estaba
                If IsNull(Rs!observa2) Then
                    linea = linea & "NULL,"
                Else
                    linea = linea & DBSet(Rs!observa2, "T") & ","
                End If
            End If
            
            If IsNull(Rs!fecreser) Then
                linea = linea & "NULL,"
            Else
                linea = linea & "'" & Format(Rs!fecreser, FormatoFecha) & "',"
            End If
            If IsNull(Rs!horreser) Then
                linea = linea & "NULL,"
            Else
                linea = linea & "'" & Format(Rs!horreser, FormatoHora) & "',"
            End If
            If IsNull(Rs!fecaviso) Then
                linea = linea & "NULL,"
            Else
                linea = linea & "'" & Format(Rs!fecaviso, FormatoFecha) & "',"
            End If
            If IsNull(Rs!horaviso) Then
                linea = linea & "NULL,"
            Else
                linea = linea & "'" & Format(Rs!horaviso, FormatoHora) & "',"
            End If
            If IsNull(Rs!fecllega) Then
                linea = linea & "NULL,"
            Else
                linea = linea & "'" & Format(Rs!fecllega, FormatoFecha) & "',"
            End If
            If IsNull(Rs!horllega) Then
                linea = linea & "NULL,"
            Else
                linea = linea & "'" & Format(Rs!horllega, FormatoHora) & "',"
            End If
            If IsNull(Rs!fecocupa) Then
                linea = linea & "NULL,"
            Else
                linea = linea & "'" & Format(Rs!fecocupa, FormatoFecha) & "',"
            End If
            If IsNull(Rs!horocupa) Then
                linea = linea & "NULL,"
            Else
                linea = linea & "'" & Format(Rs!horocupa, FormatoHora) & "',"
            End If
            If IsNull(Rs!fecfinal) Then
                linea = linea & "NULL,"
            Else
                linea = linea & "'" & Format(Rs!fecfinal, FormatoFecha) & "',"
            End If
            If IsNull(Rs!horfinal) Then
                linea = linea & "NULL,"
            Else
                linea = linea & "'" & Format(Rs!horfinal, FormatoHora) & "',"
            End If
            If IsNull(Rs!importtx) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!importtx, "N") & ","
            End If
            
            
            
            
            If IsNull(Rs!impcompr) Then
                linea = linea & "NULL,"
            Else
                '[Monica]28/05/2018: para el caso de que el cliente tenga descuento se le descuenta al socio
'                linea = linea & DBSet(Rs!impcompr, "N") & ","
                '[Monica]04/02/2019: todos menos Sevilla
                If vParamAplic.Cooperativa <> 3 Then
                    DtoCli = DevuelveDesdeBDNew(conAri, "scliente", "dtognral", "codclien", DBLet(Rs!CodClien, "N"), "N")
                    ImporteCompra = Round2(DBLet(Rs!impcompr, "N") * (100 - ImporteSinFormato(ComprobarCero(DtoCli))) / 100, 2)
                    linea = linea & DBSet(ImporteCompra, "N") & ","
                Else
                    linea = linea & DBSet(Rs!impcompr, "N") & ","
                End If
                
            End If
            
            
            
            If IsNull(Rs!extcompr) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!extcompr, "N") & ","
            End If
            If IsNull(Rs!impventa) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!impventa, "N") & ","
            End If
            If IsNull(Rs!extventa) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!extventa, "N") & ","
            End If
            If IsNull(Rs!distanci) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!distanci, "N") & ","
            End If
            If IsNull(Rs!suplemen) Then
                linea = linea & "NULL,"
            Else
                '[Monica]06/07/2018: para el caso de que el cliente tenga descuento se le descuenta al socio
'                linea = linea & DBSet(Rs!suplemen, "N") & ","
                DtoCli = DevuelveDesdeBDNew(conAri, "scliente", "dtognral", "codclien", DBLet(Rs!CodClien, "N"), "N")
                Suplemento = Round2(DBLet(Rs!suplemen, "N") * (100 - ImporteSinFormato(ComprobarCero(DtoCli))) / 100, 2)
                linea = linea & DBSet(Suplemento, "N") & ","

            End If
            If IsNull(Rs!imppeaje) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!imppeaje, "N") & ","
            End If
            If IsNull(Rs!imppropi) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!imppropi, "N") & ","
            End If
            If IsNull(Rs!facturad) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!facturad, "N") & ","
            End If
            If IsNull(Rs!abonados) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!abonados, "N") & ","
            End If
            
            '[Monica]28/02/2018: marcar como validados
'            If vParamAplic.MarcarValidados Or MarcarValidadosEnCliente(Rs!CodClien) Then
'                linea = linea & "1,"
'            Else
'                If IsNull(Rs!validado) Then
'                    linea = linea & "NULL,"
'                Else
'                    linea = linea & DBSet(Rs!validado, "N") & ","
'                End If
'            End If
            
            If vParamAplic.MarcarValidados Then
                linea = linea & "validado = 1, "
            Else
                If Not IsNull(Rs!CodClien) Then
                    If MarcarValidadosEnCliente(Rs!CodClien) Then
                        linea = linea & "1, "
                    Else
                        linea = linea & DBSet(Rs!validado, "N") & ","
                    End If
                Else
                    linea = linea & DBSet(Rs!validado, "N") & ","
                End If
            End If
            
            If IsNull(Rs!Destino) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!Destino, "T") & ","
            End If
            If IsNull(Rs!Empresa) Then
                linea = linea & "NULL,"
            Else
                linea = linea & DBSet(Rs!Empresa, "N") & ","
            End If
            
            If IsNull(Rs!codigousu) Then
                linea = linea & "NULL)"
            Else
                linea = linea & DBSet(Rs!codigousu, "N") & ")"
            End If
            
            
            values = values & linea & ","
            'If Len(values) > 100000 Then
                'quitamos la ultima coma
                values = Mid(values, 1, Len(values) - 1)
                SQL1 = Sql & values
                conn.Execute SQL1
                values = ""
            'End If
        End If
            
        DoEvents
        
        Label1(0).Caption = Round2(ProgressBar1.Value, 0) & " %"
        Label1(0).Refresh
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    'quitamos la ultima coma
'    values = Mid(values, 1, Len(values) - 1)
'    SQL = SQL & values
'    conn.Execute SQL
    Screen.MousePointer = vbDefault
EActua:
If Err.Number <> 0 Then
    conn.RollbackTrans
    MsgBox "Error al actualizar BD: " & Err.Description
Else
    conn.CommitTrans
    MsgBox "El proceso ha finalizado con éxito.", vbExclamation
End If
End Sub

Private Function MarcarValidadosEnCliente(vCodClien As String)
Dim Sql As String

    Sql = "select marcarvalidados from scliente where codclien = " & DBSet(vCodClien, "N")
    
    MarcarValidadosEnCliente = (DevuelveValor(Sql) = 1)

End Function



Private Function ExisteEnShilla(vWhere As String) As Boolean
Dim Sql As String

    On Error Resume Next
    
    
    Sql = "select count(*) from shilla where " & vWhere
    ExisteEnShilla = (TotalRegistros(Sql) <> 0)


End Function

Private Function EstaFacturadoCliente(vWhere As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error Resume Next
    
    EstaFacturadoCliente = False
    
    
    Sql = "select facturadocliente from shilla where " & vWhere
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        EstaFacturadoCliente = (DBLet(Rs!facturadocliente, "N") = 1)
    End If
    Set Rs = Nothing
    

End Function




Private Function EsdeCredito(vWhere As String) As Boolean
Dim Sql As String

    On Error Resume Next
    
    Sql = "select count(*) from shilla where " & vWhere & " and tipservi = 1"
    EsdeCredito = (TotalRegistros(Sql) <> 0)

End Function

Private Sub VerErrores()
    frmFacturas.Socio = False
    frmFacturas.Sql = "select numeruve,fecha,hora,error, id from tmptaxi where error1=1"
    frmFacturas.Caption = "Errores en el fichero de traspaso"
    frmFacturas.deExcel = (Check1(0).Value = 1)
    frmFacturas.Show vbModal

End Sub

Private Sub cmdCancel_Click()
    If cmdCancel.Caption <> "&Cancelar" Then
        pulsadoCancelar = True
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    procesoCancelado = False
    procesoFinalizado = False
    pulsadoCancelar = False
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon


    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 23 To 23
        Me.imgFecha(kCampo).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next kCampo


    txtcodigo(85).Text = Format(Now, "dd/mm/yyyy")

    EnTomaDeDatos
'    BorrarTablas
    procesoCancelado = False
    procesoFinalizado = False
End Sub

Private Sub EnTomaDeDatos()
    Me.ProgressBar1.visible = False
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    cd1.InitDir = Mid(App.Path, 1, 3)
    cd1.ShowOpen
    If cd1.FileName <> "" Then Text1.Text = cd1.FileName
End Sub

Private Function ProcesarFichero_new() As Boolean
Dim NF As Integer
Dim LlevoFichero As Currency
Dim values As String
Dim linea As String
Dim Sql As String


    On Error GoTo EProcesarFichero
    ProcesarFichero_new = False
'-- iniciar la barra de progreso
    Label1(2).Caption = "Preparando tablas."
    Label1(2).Refresh
    CargarProgresNew Me.ProgressBar1, 100
    LlevoFichero = 0
    Me.ProgressBar1.visible = True
    DoEvents
    linea = "(id,telefono,codclien,codautor,codusuar,nomclien,tipservi,observa1,numeruve,licencia,matricul,"
    linea = linea & "dirllama,ciudadre,numllama,puerllama,fecha,hora,idservic,opereser,opedespa,estado,"
    linea = linea & "observa2,fecreser,horreser,fecaviso,horaviso,fecllega,horllega,fecocupa,horocupa,fecfinal,"
    linea = linea & "horfinal,importtx,impcompr,extcompr,impventa,extventa,distanci,suplemen,imppeaje,imppropi,facturad,"
    '[Monica]03/10/2014: añadimos el taxi del destino
    linea = linea & "abonados,validado,destino,error1,error)"
    values = ""
    Contador = 0
    'Empezamos
    Label1(2).Caption = "Procesando fichero."
    Label1(2).Refresh
    NF = FreeFile
    NumRegElim = FileLen(Text1.Text)
    Open Text1.Text For Input As #NF
    
    
    Line Input #NF, menErrProceso
    
    
    Dim CadenaAux As String
    Dim s As String
    
    
    
    
    While Not EOF(NF)
    
        Line Input #NF, CadenaAux
        
        
        s = Mid(CadenaAux, 1, 1)
        If IsNumeric(s) Then
            ArmarCadena menErrProceso, values
            menErrProceso = CadenaAux
        Else
            menErrProceso = menErrProceso & " " & CadenaAux
        End If
        
        
        LlevoFichero = LlevoFichero + Len(menErrProceso)
        Me.ProgressBar1.Value = Round2((LlevoFichero * 100) / NumRegElim, 2)
        Me.Label1(0).Caption = Round2(Me.ProgressBar1.Value, 0) & " %"
        Me.Label1(0).Refresh
        'aux = Trim(Mid(menErrProceso, 1, 11))
        'If aux = "963713908" Then Stop
        
        'If LlevoFichero = 5080986 Then Stop
'        ArmarCadena menErrProceso, values
        DoEvents
        If Len(values) > 100000 Then
            'quitamos la ultima coma
            values = Mid(values, 1, Len(values) - 1)
            Sql = "INSERT INTO tmptaxi " & linea & " VALUES " & values
            conn.Execute Sql
            values = ""
        End If
        If pulsadoCancelar Then
            procesoCancelado = True
            Close #NF
            Exit Function
        End If
        Contador = Contador + 1
    Wend
    Close #NF
    
    If menErrProceso <> "" Then ArmarCadena menErrProceso, values
    
    
    If values <> "" Then
        'quitamos la ultima coma
        values = Mid(values, 1, Len(values) - 1)
    
        Sql = "INSERT INTO tmptaxi " & linea & " VALUES " & values
        conn.Execute Sql
    End If
    
    ProcesarFichero_new = True
    If Not procesoCancelado Then
        procesoFinalizado = True
    End If
    Exit Function
EProcesarFichero:
If Err.Number <> 0 Then
    ProcesarFichero_new = False
    MsgBox "Error al procesar fichero: " & Err.Description
    procesoFinalizado = False
End If
If NF > 0 Then Close #NF
End Function

Private Sub ArmarCadena(CADENA As String, ByRef Values1 As String)
Dim Telefono As String
Dim values As String
Dim Error As String
Dim Error1 As Byte

Dim Valor As Double

    Fecha = Mid(menErrProceso, 397, 10)
    hora = Mid(menErrProceso, 407, 8)
    Vehiculo = Mid(menErrProceso, 293, 4)


Error1 = 0
Error = ""
'armamos los registros segun la cadena
Telefono = Trim(Mid(CADENA, 1, 11))
'telefono

values = Contador

If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 11, 4))
'codclien
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
        values = values & ",NULL"
        Error1 = 1
        Error = "codclien con formato incorrecto"
Else
        values = values & "," & CInt(Telefono)
End If

Telefono = Trim(Mid(CADENA, 15, 14))
'codautor"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 29, 30))
'codusuar"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 59, 30))
'nomclien"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 89, 1))
'tipservi"
If Telefono = "" Then
    values = values & ",NULL"
Else
    If Telefono = "N" Then
        values = values & ",0"
    ElseIf Telefono = "S" Then
            values = values & ",1"
    Else
        values = values & ",NULL"
        Error1 = "1"
        Error = "tipservi con formato incorrecto"
    End If
End If

Telefono = Trim(Mid(CADENA, 93, 200))
'observa1"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

'numeruve"
If Not IsNumeric(Vehiculo) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "Vehiculo con formato incorrecto"
Else
    values = values & "," & CInt(Vehiculo)
End If

Telefono = Trim(Mid(CADENA, 297, 10))
'licencia"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 307, 10))
'matricul"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 317, 30))
'dirllama"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 347, 30))
'ciudadre"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 377, 10))
'numllama"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 387, 10))
'puerllama"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

'fecha"
If Fecha = "" Then
    values = values & ",NULL"
    Error1 = 1
    Error = "Falta fecha"
ElseIf Not IsDate(Fecha) Then
        values = values & ",NULL"
        Error1 = 1
        Error = "Fecha con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Fecha), FormatoFecha), "T")
End If

'hora"
If hora = "" Then
    values = values & ",NULL"
    Error1 = 1
    Error = "Falta hora"
ElseIf Not IsDate(hora) Then
        values = values & ",NULL"
        Error1 = 1
        Error = "Hora con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(hora), FormatoHora), "T")
End If

Telefono = Trim(Mid(CADENA, 415, 6))
'idservic"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If


Telefono = Trim(Mid(CADENA, 421, 30))
'opereser"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 451, 30))
'opedespa"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 481, 4))
'estado"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 485, 200))
'observa2"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 685, 10))
'fecreser"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "fecha reserva con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoFecha), "T")
End If

Telefono = Trim(Mid(CADENA, 695, 8))
'horreser"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "hora reserva con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoHora), "T")
End If

Telefono = Trim(Trim(Mid(CADENA, 721, 10)))
'fecaviso"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "fecha aviso con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoFecha), "T")
End If

Telefono = Trim(Mid(CADENA, 731, 8))
'horaviso"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "hora aviso con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoHora), "T")
End If

Telefono = Trim(Mid(CADENA, 739, 10))
'fecllega"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "fecha llegada con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoFecha), "T")
End If

Telefono = Trim(Mid(CADENA, 749, 8))
'horllega"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "hora llegada con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoHora), "T")
End If

Telefono = Trim(Mid(CADENA, 757, 10))
'fecocupa"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "fecha ocupa con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoFecha), "T")
End If

Telefono = Trim(Mid(CADENA, 767, 8))
'horocupa"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "hora ocupa con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoHora), "T")
End If

Telefono = Trim(Mid(CADENA, 775, 10))
'fecfinal"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "fecha final con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoFecha), "T")
End If

Telefono = Trim(Mid(CADENA, 785, 8))
'horfinal"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "hora final con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoHora), "T")
End If

Telefono = Trim(Mid(CADENA, 793, 15))
'importtx"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "importe tx con formato incorrecto"
Else
    If InStr(1, Telefono, ",") > 0 Then
        Valor = ImporteFormateado(Telefono)
    Else
        Valor = CDbl(TransformaPuntosComas(Telefono))
    End If

    values = values & "," & DBSet(Valor, "N")
End If

Telefono = Trim(Mid(CADENA, 808, 15))
'impcompr"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "importe compra con formato incorrecto"
Else
    If InStr(1, Telefono, ",") > 0 Then
        Valor = ImporteFormateado(Telefono)
    Else
        Valor = CDbl(TransformaPuntosComas(Telefono))
    End If

    values = values & "," & DBSet(Valor, "N")
End If

Telefono = Trim(Mid(CADENA, 823, 15))
'extcompr"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "extcompr con formato incorrecto"
Else
    If InStr(1, Telefono, ",") > 0 Then
        Valor = ImporteFormateado(Telefono)
    Else
        Valor = CDbl(TransformaPuntosComas(Telefono))
    End If
    
    values = values & "," & DBSet(Valor, "N")
End If

Telefono = Trim(Mid(CADENA, 838, 15))
'impventa"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "importe venta con formato incorrecto"
Else
    If InStr(1, Telefono, ",") > 0 Then
        Valor = ImporteFormateado(Telefono)
    Else
        Valor = CDbl(TransformaPuntosComas(Telefono))
    End If

    values = values & "," & DBSet(Valor, "N")
End If

Telefono = Trim(Mid(CADENA, 853, 15))
'extventa"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "extventa con formato incorrecto"
Else
    If InStr(1, Telefono, ",") > 0 Then
        Valor = ImporteFormateado(Telefono)
    Else
        Valor = CDbl(TransformaPuntosComas(Telefono))
    End If
    
    values = values & "," & DBSet(Valor, "N")
End If

Telefono = Trim(Mid(CADENA, 868, 15))
'distanci"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "distancia con formato incorrecto"
Else
    If InStr(1, Telefono, ",") > 0 Then
        Valor = ImporteFormateado(Telefono)
    Else
        Valor = CDbl(TransformaPuntosComas(Telefono))
    End If
    values = values & "," & DBSet(Valor, "N")
End If

Telefono = Trim(Mid(CADENA, 883, 15))
'suplemen"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "suplemento con formato incorrecto"
Else
    If InStr(1, Telefono, ",") > 0 Then
        Valor = ImporteFormateado(Telefono)
    Else
        Valor = CDbl(TransformaPuntosComas(Telefono))
    End If
    
    values = values & "," & DBSet(Valor, "N")
End If

Telefono = Trim(Mid(CADENA, 898, 15))
'imppeaje"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "importe peaje con formato incorrecto"
Else
    If InStr(1, Telefono, ",") > 0 Then
        Valor = ImporteFormateado(Telefono)
    Else
        Valor = CDbl(TransformaPuntosComas(Telefono))
    End If

    values = values & "," & DBSet(Valor, "N")
End If

Telefono = Trim(Mid(CADENA, 913, 15))
'imppropi"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "importe propina con formato incorrecto"
Else
    If InStr(1, Telefono, ",") > 0 Then
        Valor = ImporteFormateado(Telefono)
    Else
        Valor = CDbl(TransformaPuntosComas(Telefono))
    End If

    values = values & "," & DBSet(Valor, "N")
End If

Telefono = Trim(Mid(CADENA, 931, 1))
'facturad"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "facturado con formato incorrecto"
Else
    values = values & "," & CInt(Telefono)
End If

Telefono = Trim(Mid(CADENA, 935, 1))
'abonados"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "abonado con formato incorrecto"
Else
    values = values & "," & CInt(Telefono)
End If

Telefono = Trim(Mid(CADENA, 939, 1))
'validado"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "validado con formato incorrecto"
Else
    values = values & "," & CInt(Telefono)
End If

'[Monica]03/10/2014: añadimos el destino del servicio
Telefono = Trim(Trim(Mid(CADENA, 940, 30)) & " " & Trim(Mid(CADENA, 970, 30)) & " " & Trim(Mid(CADENA, 1000, 10)) & " " & Trim(Mid(CADENA, 1010, 10)))
'destino"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If
'++


'error1,error
values = "(" & values & "," & Error1 & "," & DBSet(Error, "T") & ")"
Values1 = Values1 & values & ","

EInsert:
If Err.Number <> 0 Then
    MsgBox "Error al insertar en error1. " & Err.Description
End If

End Sub
Private Function ProcesarFichero() As Boolean
Dim NF As Integer
Dim LlevoFichero As Currency


    On Error GoTo EProcesarFichero
    ProcesarFichero = False
'-- iniciar la barra de progreso
    Label1(0).Caption = "Preparando tablas."
    CargarProgresNew Me.ProgressBar1, 100
    LlevoFichero = 0
    Me.ProgressBar1.visible = True
    DoEvents
    
    'Empezamos
    
    
    
    Label1(0).Caption = "Procesando fichero."
    Label1(0).Refresh
    NF = FreeFile
    NumRegElim = FileLen(Text1.Text)
    Open Text1.Text For Input As #NF
    
    While Not EOF(NF)
        Line Input #NF, menErrProceso
        
        LlevoFichero = LlevoFichero + Len(menErrProceso)
        'If LlevoFichero > NumRegElim Then LlevoFichero = NumRegElim
        Me.ProgressBar1.Value = Round((LlevoFichero * 100) / NumRegElim, 2)
        Me.Label1(0).Caption = Me.ProgressBar1.Value & " %"
        
        'busco si ya existe en nuestra tabla tmp el registro
        Fecha = Mid(menErrProceso, 397, 10)
        hora = Mid(menErrProceso, 407, 8)
        Vehiculo = Mid(menErrProceso, 293, 4)
        If Fecha = "" Then
            Insertar menErrProceso, "tmperr", "falta fecha"
        ElseIf Not IsDate(Fecha) Then
                Insertar menErrProceso, "tmperr", "fecha formato incorrecto"
            ElseIf hora = "" Then
                    Insertar menErrProceso, "tmperr", "falta hora"
                ElseIf Not IsDate(hora) Then
                    Insertar menErrProceso, "tmperr", "hora formato incorrecto"
                ElseIf Vehiculo = "" Then
                    Insertar menErrProceso, "tmperr", "falta vehiculo"
                ElseIf Not IsNumeric(Vehiculo) Then
                    Insertar menErrProceso, "tmperr", "vehiculo formato incorrecto"
            Else
                DoEvents
                If BuscarRegistro(Vehiculo, Fecha, hora) Then
                    Insertar menErrProceso, "tmpErr", "Registro duplicado"
                Else
                    Insertar menErrProceso, "tmpTaxi"
                End If

        End If
        DoEvents
        If pulsadoCancelar Then
            procesoCancelado = True
            Close #NF
            Exit Function
        End If
    Wend
    Close #NF
    
    ProcesarFichero = True
    If Not procesoCancelado Then
        procesoFinalizado = True
    End If
    Exit Function
EProcesarFichero:
If Err.Number <> 0 Then
    ProcesarFichero = False
    MsgBox "Error al procesar fichero: " & Err.Description
    procesoFinalizado = False
End If
If NF > 0 Then Close #NF
End Function

Public Sub BorrarTablas()
On Error Resume Next

    conn.Execute "DELETE from tmpTaxi"
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Function BuscarRegistro(ByRef V As String, ByRef F As String, ByRef H As String) As Boolean
Dim fec As Date
Dim hor As Date
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo EBusqueda
    
    Set Rs = New ADODB.Recordset
    fec = CDate(F)
    hora = CDate(H)

    Sql = "select * from tmptaxi where numeruve=" & CInt(V) & " and fecha='" & Format(fec, FormatoFecha)
    Sql = Sql & "' and hora='" & Format(hora, FormatoHora) & "'"

    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        Rs.Close
        Sql = DevuelveDesdeBD(conAri, "codclien", "sclien", "numeruve", V, "T")
        If Sql <> "" Then
            Sql = "select * from shilla where codsocio=" & CInt(Sql) & " and numeruve=" & CInt(V) & " and fecha='" & Format(fec, FormatoFecha)
            Sql = Sql & "' and hora='" & Format(hora, FormatoHora) & "'"
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Rs.EOF Then
                BuscarRegistro = False
            Else
                BuscarRegistro = True
            End If
            Rs.Close
        Else
            BuscarRegistro = False
        End If
    Else
        BuscarRegistro = True
        Rs.Close
    End If
    
EBusqueda:
If Err.Number <> 0 Then
    MsgBox "Error al buscar registro:" & Err.Description
    BuscarRegistro = False
End If
    
End Function
Private Sub Insertar(CADENA As String, Tabla As String, Optional Error As String)
Dim Telefono As String
Dim linea As String, values As String
Dim Sql As String
Dim Socio As String

On Error GoTo EInsert

'armamos los registros segun la cadena
Telefono = Trim(Mid(CADENA, 1, 11))
linea = "telefono"

If Telefono = "" Then
    values = "NULL"
Else
    values = DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 11, 4))
linea = linea & "," & "codclien"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
        values = values & ",NULL"
        Tabla = "tmpErr"
        Error = "codclien con formato incorrecto"
Else
        values = values & "," & CInt(Telefono)
End If

Telefono = Trim(Mid(CADENA, 15, 14))
linea = linea & "," & "codautor"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 29, 30))
linea = linea & "," & "codusuar"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 59, 30))
linea = linea & "," & "nomclien"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 89, 1))
linea = linea & "," & "tipservi"
If Telefono = "" Then
    values = values & ",NULL"
Else
    If Telefono = "N" Then
        values = values & ",0"
    ElseIf Telefono = "S" Then
            values = values & ",1"
    Else
        values = values & ",NULL"
        Tabla = "tmpErr"
        Error = "tipservi con formato incorrecto"
    End If
End If

Telefono = Trim(Mid(CADENA, 93, 200))
linea = linea & "," & "observa1"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

linea = linea & "," & "numeruve"
If Not IsNumeric(Vehiculo) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "Vehiculo con formato incorrecto"
    values = values & ",NULL"
Else
    values = values & "," & CInt(Vehiculo)
    'con el número de vehiculo buscamos el socio,si no lo encontramos
    'preparamos para agregar en la tabla de errores
    Socio = DevuelveDesdeBD(conAri, "codclien", "sclien", "numeruve", Vehiculo, "T")
    If Socio = "" Then
        Tabla = "tmpErr"
        Error = "Ningun socio tiene asociado este codigo de vehiculo"
        values = values & ",NULL"
    Else
        values = values & "," & CInt(Socio)
    End If
End If
linea = linea & "," & "codsocio"

Telefono = Trim(Mid(CADENA, 297, 10))
linea = linea & "," & "licencia"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 307, 10))
linea = linea & "," & "matricul"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 317, 30))
linea = linea & "," & "dirllama"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 347, 30))
linea = linea & "," & "ciudadre"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 377, 10))
linea = linea & "," & "numllama"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 387, 10))
linea = linea & "," & "puerllama"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

linea = linea & "," & "fecha"
If Fecha = "" Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "Falta fecha"
ElseIf Not IsDate(Fecha) Then
        values = values & ",NULL"
        Tabla = "tmpErr"
        Error = "Fecha con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Fecha), FormatoFecha), "T")
End If

linea = linea & "," & "hora"
If hora = "" Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "Falta hora"
ElseIf Not IsDate(hora) Then
        values = values & ",NULL"
        Tabla = "tmpErr"
        Error = "Hora con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(hora), FormatoHora), "T")
End If

Telefono = Trim(Mid(CADENA, 415, 6))
linea = linea & "," & "idservic"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If


Telefono = Trim(Mid(CADENA, 421, 30))
linea = linea & "," & "opereser"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 451, 30))
linea = linea & "," & "opedespa"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 481, 4))
linea = linea & "," & "estado"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 485, 200))
linea = linea & "," & "observa2"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 685, 10))
linea = linea & "," & "fecreser"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "fecha reserva con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoFecha), "T")
End If

Telefono = Trim(Mid(CADENA, 695, 8))
linea = linea & "," & "horreser"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "hora reserva con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoHora), "T")
End If

Telefono = Trim(Trim(Mid(CADENA, 721, 10)))
linea = linea & "," & "fecaviso"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "fecha aviso con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoFecha), "T")
End If

Telefono = Trim(Mid(CADENA, 731, 8))
linea = linea & "," & "horaviso"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "hora aviso con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoHora), "T")
End If

Telefono = Trim(Mid(CADENA, 739, 10))
linea = linea & "," & "fecllega"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "fecha llegada con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoFecha), "T")
End If

Telefono = Trim(Mid(CADENA, 749, 8))
linea = linea & "," & "horllega"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "hora llegada con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoHora), "T")
End If

Telefono = Trim(Mid(CADENA, 757, 10))
linea = linea & "," & "fecocupa"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "fecha ocupa con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoFecha), "T")
End If

Telefono = Trim(Mid(CADENA, 767, 8))
linea = linea & "," & "horocupa"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "hora ocupa con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoHora), "T")
End If

Telefono = Trim(Mid(CADENA, 775, 10))
linea = linea & "," & "fecfinal"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "fecha final con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoFecha), "T")
End If

Telefono = Trim(Mid(CADENA, 785, 8))
linea = linea & "," & "horfinal"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsDate(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "hora final con formato incorrecto"
Else
    values = values & "," & DBSet(Format(CDate(Telefono), FormatoHora), "T")
End If

Telefono = Trim(Mid(CADENA, 793, 15))
linea = linea & "," & "importtx"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "importe tx con formato incorrecto"
Else
    values = values & "," & DBSet(Telefono, "N")
End If

Telefono = Trim(Mid(CADENA, 808, 15))
linea = linea & "," & "impcompr"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "importe compra con formato incorrecto"
Else
    values = values & "," & DBSet(Telefono, "N")
End If

Telefono = Trim(Mid(CADENA, 823, 15))
linea = linea & "," & "extcompr"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "extcompr con formato incorrecto"
Else
    values = values & "," & DBSet(Telefono, "N")
End If

Telefono = Trim(Mid(CADENA, 838, 15))
linea = linea & "," & "impventa"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "importe venta con formato incorrecto"
Else
    values = values & "," & DBSet(Telefono, "N")
End If

Telefono = Trim(Mid(CADENA, 853, 15))
linea = linea & "," & "extventa"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "extventa con formato incorrecto"
Else
    values = values & "," & DBSet(Telefono, "N")
End If

Telefono = Trim(Mid(CADENA, 868, 15))
linea = linea & "," & "distanci"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "distancia con formato incorrecto"
Else
    values = values & "," & DBSet(Telefono, "N")
End If

Telefono = Trim(Mid(CADENA, 883, 15))
linea = linea & "," & "suplemen"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "suplemento con formato incorrecto"
Else
    values = values & "," & DBSet(Telefono, "N")
End If

Telefono = Trim(Mid(CADENA, 898, 15))
linea = linea & "," & "imppeaje"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "importe peaje con formato incorrecto"
Else
    values = values & "," & DBSet(Telefono, "N")
End If

Telefono = Trim(Mid(CADENA, 913, 15))
linea = linea & "," & "imppropi"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "importe propina con formato incorrecto"
Else
    values = values & "," & DBSet(Telefono, "N")
End If

Telefono = Trim(Mid(CADENA, 931, 1))
linea = linea & "," & "facturad"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "facturado con formato incorrecto"
Else
    values = values & "," & CInt(Telefono)
End If

Telefono = Trim(Mid(CADENA, 935, 1))
linea = linea & "," & "abonados"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "abonado con formato incorrecto"
Else
    values = values & "," & CInt(Telefono)
End If

Telefono = Trim(Mid(CADENA, 939, 1))
linea = linea & "," & "validado"
If Telefono = "" Then
    values = values & ",NULL"
ElseIf Not IsNumeric(Telefono) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "validado con formato incorrecto"
Else
    values = values & "," & CInt(Telefono)
End If

If Tabla = "tmpTaxi" Then
    linea = linea & ",error1"
    values = values & ",0"
Else
   linea = linea & ",error1,error"
   values = values & ",1," & DBSet(Error, "T")
End If
Sql = "INSERT INTO tmptaxi (" & linea & ") VALUES ("
Sql = Sql & values & ")"
conn.Execute Sql

EInsert:
If Err.Number <> 0 Then
    MsgBox "Error al insertar en tabla. " & Err.Description
End If

End Sub

Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 23 'fechas de factura
            indCodigo = Index + 62
   End Select
   
   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtcodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)

End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Check1(0).Enabled = True
        Check1(0).visible = True
    Else
        Check1(0).Enabled = False
        Check1(0).visible = False
    End If
End Sub

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
        Case 85  'FECHA Desde Hasta
            If txtcodigo(Index).Text = "" Then Exit Sub
            PonerFormatoFecha txtcodigo(Index)
            
    End Select
    
End Sub


Private Function DatosOk() As Boolean
Dim encontrado As String
Dim Codigo As String

    DatosOk = True
    If Me.Option1(1).Value Or Option1(2).Value Then
        If txtcodigo(85).Text = "" Then
            MsgBox "Debe introducir obligatoriamente la fecha de  traspaso.", vbExclamation
            DatosOk = False
            Exit Function
        End If
        If txtcodigo(4).Text = "" Then
            MsgBox "Debe introducir obligatoriamente un concepto.", vbExclamation
            DatosOk = False
            Exit Function
        End If
    End If

End Function



Private Function ComprobarFichero(Escliente As Boolean) As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim SQL1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    On Error GoTo eComprobarFichero
    
    ComprobarFichero = False
    
    NF = FreeFile
    Open Text1.Text For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    
    Label1(0).Caption = "Insertando en Tabla temporal: " & Text1.Text
    longitud = FileLen(Text1.Text)
    
    ProgressBar1.visible = True
    Me.ProgressBar1.Max = longitud
    Me.Refresh
    Me.ProgressBar1.Value = 0
    ' PROCESO DEL FICHERO VENTAS.TXT

    b = True

    While Not EOF(NF) And b
        i = i + 1
        
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Len(Cad)
        Label1(2).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarRegistro(Cad, Escliente)
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" Then
        i = i + 1
        
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Len(Cad)
        Label1(2).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarRegistro(Cad, Escliente)
    End If
    
    ProgressBar1.visible = False
    Label1(2).Caption = ""
    Label1(0).Caption = ""

    ComprobarFichero = b
    Exit Function

eComprobarFichero:
    ComprobarFichero = False
End Function


Private Function ComprobarRegistro(Cad As String, EsClien As Boolean) As Boolean
Dim Sql As String
Dim c_Importe As Currency
Dim Mens As String
Dim CodSoc As String
Dim Id As String
Dim Importe As String
Dim NServicios As String

Dim vServicios As Long
Dim vImporte As Currency

Dim Rs As ADODB.Recordset

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True

    If EsClien Then ' facturacion a clientes
        Id = Mid(Cad, 1, 6)
        Importe = Mid(Cad, 352, 10)
        NServicios = Mid(Cad, 362, 5)
    Else ' liquidacion a socios
        Id = Mid(Cad, 1, 6)
        Importe = Mid(Cad, 375, 10)
        NServicios = Mid(Cad, 385, 5)
    End If
    
    c_Importe = Replace(ComprobarCero(Importe), ".", ",")
    
    If EsClien Then
        'Comprobamos que el cliente existe
        If Id <> "" Then
            Sql = ""
            Sql = DevuelveDesdeBDNew(conAri, "scliente", "codclien", "codclien", Id, "N")
            If Sql = "" Then
                Mens = "No existe el cliente"
                Sql = "insert into tmpinformes (codusu, importe1, fecha1, importe2, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & "," & Importe & "," & _
                      DBSet(Mens, "T") & ")"
                conn.Execute Sql
            End If
            ' comprobamos que no exista el registro en la tabla
            Sql = ""
            Sql = DevuelveDesdeBDNew(conAri, "sfactclitr", "codclien", "codclien", Id, "N", , "fecfactu", txtcodigo(85).Text, "F")
            If Sql <> "" Then
                Mens = "Existe el registro en las facturas"
                Sql = "insert into tmpinformes (codusu, importe1, fecha1, importe2, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & "," & Importe & "," & _
                      DBSet(Mens, "T") & ")"
                conn.Execute Sql
            End If
            
'[Monica]10/09/2014: de momento comentado pq partimos de la shilla
'            '[Monica]31/03/2014: comprobamos que el nro de servicios se corresponde con el fichero de contabilidad
'            If CargarServicios Then
'                SQL = "select sum(importe) importe, count(*) servicios from tmpservicios where codclien = " & DBSet(Id, "N")
'
'                Set RS = New ADODB.Recordset
'                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                vServicios = 0
'                vImporte = 0
'                If Not RS.EOF Then
'                    vServicios = DBLet(RS!Servicios, "N")
'                    vImporte = DBLet(RS!Importe, "N")
'                End If
'
'                If vServicios <> CLng(NServicios) Then
'                    Mens = "Nro de servicios diferente."
'                    SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, nombre1) values (" & _
'                          vUsu.Codigo & "," & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & "," & NServicios & "," & _
'                          DBSet(Mens, "T") & ")"
'                    conn.Execute SQL
'                End If
'                If vImporte <> c_Importe Then
'                    Mens = "Suma importes diferente."
'                    SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, nombre1) values (" & _
'                          vUsu.Codigo & "," & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & "," & Importe & "," & _
'                          DBSet(Mens, "T") & ")"
'                    conn.Execute SQL
'                End If
'            End If
        End If
    Else
        If c_Importe <> 0 Then
            If Id <> "" Then
                'Comprobamos que la v del socio existe
                Sql = ""
                Sql = DevuelveDesdeBDNew(conAri, "sclien", "codclien", "numeruve", Id, "N")
                CodSoc = Sql
                If Sql = "" Then
                    Mens = "No existe VSocio"
                    Sql = "insert into tmpinformes (codusu, importe1, fecha1, importe2, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & "," & Importe & "," & _
                          DBSet(Mens, "T") & ")"
                    conn.Execute Sql
                End If
                
                'Comprobamos que el socio existe
                Sql = ""
                Sql = DevuelveDesdeBDNew(conAri, "sclien", "codclien", "codclien", CodSoc, "N")
                If Sql = "" Then
                    Mens = "No existe el Socio"
                    Sql = "insert into tmpinformes (codusu, importe1, fecha1, importe2, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(CodSoc, "N") & "," & DBSet(txtcodigo(85).Text, "F") & "," & Importe & "," & _
                          DBSet(Mens, "T") & ")"
                    conn.Execute Sql
                End If
            
                ' comprobamos que no exista el registro en la tabla
                Sql = ""
                Sql = DevuelveDesdeBDNew(conAri, "sfactsoctr", "numeruve", "numeruve", Id, "N", , "fecfactu", txtcodigo(85).Text, "F")
                If Sql <> "" Then
                    Mens = "Existe el registro en las facturas"
                    Sql = "insert into tmpinformes (codusu, importe1, fecha1, importe2, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & "," & Importe & "," & _
                          DBSet(Mens, "T") & ")"
                    conn.Execute Sql
                End If
                
                
                
'[Monica]10/09/2014: de momento comentado pq partimos de la shilla
'                '[Monica]31/03/2014: comprobamos que el nro de servicios se corresponde con el fichero de contabilidad
'                '                    tendiramos que hacer la misma comprobacion que para los clientes
'                If CargarServicios Then
'                    SQL = "select sum(importe) importe, count(*) servicios from tmpservicios where codclien = " & DBSet(Id, "N")
'
'                    Set RS = New ADODB.Recordset
'                    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                    vServicios = 0
'                    vImporte = 0
'                    If Not RS.EOF Then
'                        vServicios = DBLet(RS!Servicios, "N")
'                        vImporte = DBLet(RS!Importe, "N")
'                    End If
'
'                    If vServicios <> CLng(NServicios) Then
'                        Mens = "Nro de servicios diferente."
'                        SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, nombre1) values (" & _
'                              vUsu.Codigo & "," & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & "," & NServicios & "," & _
'                              DBSet(Mens, "T") & ")"
'                        conn.Execute SQL
'                    End If
'                    If vImporte <> c_Importe Then
'                        Mens = "Suma importes diferente."
'                        SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, nombre1) values (" & _
'                              vUsu.Codigo & "," & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & "," & Importe & "," & _
'                              DBSet(Mens, "T") & ")"
'                        conn.Execute SQL
'                    End If
'                End If
            End If
        End If
    End If
    
eComprobarRegistro:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Registro", Err.Descripcion
        ComprobarRegistro = False
    End If
End Function


Private Function TraspasoFichero(EsClien As Boolean) As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim SQL1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String
Dim CadValues As String
Dim Id As String
Dim Importe As String
Dim NServicios As String
Dim Socio As Long
Dim c_Importe As Currency
Dim SqlServ As String


    On Error GoTo eTraspasoFichero
    
    conn.BeginTrans

    TraspasoFichero = False
    
    NF = FreeFile
    
    Open Text1.Text For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    Label1(0).Caption = "Procesando Fichero: " & Text1.Text
    
    longitud = FileLen(Text1.Text)
    
    Me.ProgressBar1.visible = True
    Me.ProgressBar1.Max = longitud
    Me.Refresh
    Me.ProgressBar1.Value = 0
        
    If EsClien Then
        Sql = "insert into sfactclitr (codclien,fecfactu,importe,numserv,concepto,facturado) values "
    Else
        Sql = "insert into sfactsoctr (numeruve,codsocio,fecfactu,importe,numserv,concepto,facturado) values "
    End If
        
    CadValues = ""
        
    b = True
    While Not EOF(NF)
        i = i + 1
        
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Len(Cad)
        Label1(2).Caption = "Linea " & i
        Me.Refresh
        
        If EsClien Then ' facturacion a clientes
            Id = Mid(Cad, 1, 6)
            Importe = Mid(Cad, 352, 10)
            NServicios = Mid(Cad, 362, 5)
        Else ' liquidacion a socios
            Id = Mid(Cad, 1, 6)
            Importe = Mid(Cad, 375, 10)
            NServicios = Mid(Cad, 385, 5)
        End If
    
        If EsClien Then
            CadValues = CadValues & "(" & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & ","
            CadValues = CadValues & Importe & "," & DBSet(NServicios, "N") & ","
            CadValues = CadValues & DBSet(txtcodigo(4).Text, "T") & ",0),"

'[Monica]10/09/2014: de momento comentado pq partimos de la shilla
'            '[Monica]31/03/2014: en el caso de que se carguen los servicios los metemos en la tabla auxiliar
'            SqlServ = "insert into sfactclitr_serv (codclien, fecfactu, fecha, hora, origen, destino, importe, nroservicio, numeruve, matricul) "
'            SqlServ = SqlServ & "select " & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & ",concat('20', mid(fecha,7,2),'-',mid(fecha,4,2),'-',mid(fecha,1,2)), hora, origen, destino, importe, nroservicio, numeruve, matricul from tmpservicios where codclien = " & DBSet(Id, "N")
'
'            conn.Execute SqlServ
'
'            SqlServ = "update sfactclitr_serv, sclien set sfactclitr_serv.codsocio = sclien.codclien where sfactclitr_serv.codclien = " & DBSet(Id, "N") & " and fecfactu = " & DBSet(txtcodigo(85).Text, "F")
'            SqlServ = SqlServ & " and sclien.numeruve = sfactclitr_serv.numeruve "
'
'            conn.Execute SqlServ
        Else
            
            c_Importe = ComprobarCero(Importe)
            If c_Importe <> 0 Then
                Socio = DevuelveValor("select codclien from sclien where numeruve = " & DBSet(Id, "N"))
                
                CadValues = CadValues & "(" & DBSet(Id, "N") & "," & DBSet(Socio, "N") & "," & DBSet(txtcodigo(85).Text, "F") & ","
                CadValues = CadValues & Importe & "," & DBSet(NServicios, "N") & ","
                CadValues = CadValues & DBSet(txtcodigo(4).Text, "T") & ",0),"
            
'[Monica]10/09/2014: de momento comentado pq partimos de la shilla
'                '[Monica]31/03/2014: en el caso de que se carguen los servicios los metemos en la tabla auxiliar
'                SqlServ = "insert into sfactsoctr_serv (numeruve,fecfactu, fecha, hora, origen, destino, importe, nroservicio, codclien, matricul) "
'                SqlServ = SqlServ & "select " & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & ",concat('20', mid(fecha,7,2),'-',mid(fecha,4,2),'-',mid(fecha,1,2)), hora, origen, destino, importe, nroservicio, codclien, matricul from tmpservicios where numeruve = " & DBSet(Id, "N")
'
'                conn.Execute SqlServ
'
'                SqlServ = "update sfactsoctr_serv, sclien set sfactsoctr_serv.codsocio = sclien.codclien where sfactsoctr_serv.numeruve = " & DBSet(Id, "N") & " and fecfactu = " & DBSet(txtcodigo(85).Text, "F")
'                SqlServ = SqlServ & " and sclien.numeruve = sfactsoctr_serv.numeruve "
'
'                conn.Execute SqlServ
            End If
        End If
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Len(Cad)
        Label1(2).Caption = "Linea " & i
        Me.Refresh
        
        If EsClien Then ' facturacion a clientes
            Id = Mid(Cad, 1, 6)
            Importe = Mid(Cad, 352, 10)
            NServicios = Mid(Cad, 362, 5)
        Else ' liquidacion a socios
            Id = Mid(Cad, 1, 6)
            Importe = Mid(Cad, 375, 10)
            NServicios = Mid(Cad, 385, 5)
        End If
        
        
        If EsClien Then
            CadValues = CadValues & "(" & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & ","
            CadValues = CadValues & Importe & "," & DBSet(NServicios, "N") & ","
            CadValues = CadValues & DBSet(txtcodigo(4).Text, "T") & ",0),"
            
'[Monica]10/09/2014: de momento comentado pq partimos de la shilla
'            '[Monica]31/03/2014: en el caso de que se carguen los servicios los metemos en la tabla auxiliar
'            SqlServ = "insert into sfactclitr_serv (codclien, fecfactu, fecha, hora, origen, destino, importe, nroservicio, numeruve, matricul) "
'            SqlServ = SqlServ & "select " & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & ",concat('20', mid(fecha,7,2),'-',mid(fecha,4,2),'-',mid(fecha,1,2)), hora, origen, destino, importe, nroservicio, numeruve, matricul from tmpservicios where codclien = " & DBSet(Id, "N")
'
'            conn.Execute SqlServ
'
'            SqlServ = "update sfactclitr_serv, sclien set sfactclitr_serv.codsocio = sclien.codclien where sfactclitr_serv.codclien = " & DBSet(Id, "N") & " and fecfactu = " & DBSet(txtcodigo(85).Text, "F")
'            SqlServ = SqlServ & " and sclien.numeruve = sfactclitr_serv.numeruve "
'
'            conn.Execute SqlServ
            
            
            
        Else
            c_Importe = ComprobarCero(Importe)
            If c_Importe <> 0 Then
                Socio = DevuelveValor("select codclien from sclien where numeruve = " & DBSet(Id, "N"))
                
                CadValues = CadValues & "(" & DBSet(Id, "N") & "," & DBSet(Socio, "N") & "," & DBSet(txtcodigo(85).Text, "F") & ","
                CadValues = CadValues & Importe & "," & DBSet(NServicios, "N") & ","
                CadValues = CadValues & DBSet(txtcodigo(4).Text, "T") & ",0),"
            
'[Monica]10/09/2014: de momento comentado pq partimos de la shilla
'                '[Monica]31/03/2014: en el caso de que se carguen los servicios los metemos en la tabla auxiliar
'                SqlServ = "insert into sfactsoctr_serv (numeruve,fecfactu, fecha, hora, origen, destino, importe, nroservicio, codclien, matricul) "
'                SqlServ = SqlServ & "select " & DBSet(Id, "N") & "," & DBSet(txtcodigo(85).Text, "F") & ",concat('20', mid(fecha,7,2),'-',mid(fecha,4,2),'-',mid(fecha,1,2)), hora, origen, destino, importe, nroservicio, codclien, matricul from tmpservicios where numeruve = " & DBSet(Id, "N")
'
'                conn.Execute SqlServ
'
'                SqlServ = "update sfactsoctr_serv, sclien set sfactsoctr_serv.codsocio = sclien.codclien where sfactsoctr_serv.numeruve = " & DBSet(Id, "N") & " and fecfactu = " & DBSet(txtcodigo(85).Text, "F")
'                SqlServ = SqlServ & " and sclien.numeruve = sfactsoctr_serv.numeruve "
'
'                conn.Execute SqlServ
            
            
            End If
        End If
    
    End If
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute Sql & CadValues
    End If
    
    TraspasoFichero = b
    
    Me.ProgressBar1.visible = False
    Label1(0).Caption = ""
    Label1(2).Caption = ""

    conn.CommitTrans
    Exit Function

eTraspasoFichero:
    TraspasoFichero = False
    MuestraError Err.Number, "Traspaso Fichero", Err.Description
    conn.RollbackTrans
End Function


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .ConSubInforme = True
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .Show vbModal
    End With
End Sub

