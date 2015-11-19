VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGesTraspaso 
   Caption         =   "Traspaso TaxiTronic"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   615
      Left            =   210
      TabIndex        =   13
      Top             =   750
      Width           =   6975
      Begin VB.OptionButton Option1 
         Caption         =   "Servicios Socios"
         Height          =   255
         Index           =   2
         Left            =   4620
         TabIndex        =   16
         Top             =   210
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Llamadas"
         Height          =   255
         Index           =   0
         Left            =   780
         TabIndex        =   15
         Top             =   210
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Servicios Clientes"
         Height          =   255
         Index           =   1
         Left            =   2430
         TabIndex        =   14
         Top             =   210
         Width           =   1755
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
      Height          =   375
      Left            =   6210
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5130
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   3450
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   210
      TabIndex        =   6
      Top             =   1590
      Width           =   6975
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   85
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   1
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   4
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1230
         Width           =   4605
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   810
         Picture         =   "frmGesTraspaso.frx":0000
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Concepto"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1230
         Width           =   885
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   810
         Picture         =   "frmGesTraspaso.frx":008B
         ToolTipText     =   "Buscar Fichero"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
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
      Top             =   4170
      Width           =   6945
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
      Top             =   3840
      Width           =   6945
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


Private menErrProceso As String 'mensaje final del proceso actualizacion de precios
Dim vehiculo As String
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

Private Function RScontador(CADENA As String) As Currency
    
    RScontador = 0
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open CADENA, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        RScontador = miRsAux.Fields(0)
    End If
    miRsAux.Close
    
End Function

Private Sub cmdAceptar_Click()
Dim cadSel As String
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim encontrado As String
Dim total As Currency
Dim Contador As Currency
Dim Sql As String
Dim cadTabla As String


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
        
            'Llegados aqui, procesamos el fichero
            Screen.MousePointer = vbHourglass
        '    b = ProcesarFichero
            b = ProcesarFichero_new
            If b Then
                'verificamos que los numeruve esten asociados a algun socio
                ProgressBar1.Value = 0
                Contador = 0
                Label1(0).Caption = ""
                Set RS = New ADODB.Recordset
                Sql = "select * from tmptaxi where error1 <> 1 group by numeruve"
                RS.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
                total = RScontador("select count(distinct(numeruve)) from tmptaxi where error1<>1")
                Label1(2).Caption = "Verificando códigos de socios."
                Label1(2).Refresh
        
                While Not RS.EOF
                    Contador = Contador + 1
                    ProgressBar1.Value = (Contador * 100) / total
                    DoEvents
                    'Label1(0).Caption = Round(ProgressBar1.Value, 2) & "%"
                    Label1(0).Caption = Round2(ProgressBar1.Value, 0) & " %"
                    Label1(0).Refresh
                    encontrado = DevuelveDesdeBD(conAri, "codclien", "sclien", "numeruve", RS!NumerUve, "T")
                    b = Updatear(RS!NumerUve, encontrado)
                    RS.MoveNext
                Wend
                RS.Close
                Label1(0).Caption = ""
                Label1(0).Refresh
                'buscamos en la misma tabla que los registros no esten duplicados
                If b Then
                    ProgressBar1.Value = 0
                    Contador = 0
        
                    Set RS = New ADODB.Recordset
                    Sql = "select fecha,hora,numeruve, count(*) from tmptaxi where error1 <> 1 group by 1,2,3 having count(*) > 1"
                    RS.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
                    total = RScontador("select count(*) from (" & Sql & ") aaalias ")  'tmptaxi where error1 <> 1")
                    Label1(2).Caption = "eliminando(II) duplicidad de registros en el fichero."
                    Label1(2).Refresh
                    While Not RS.EOF
                        Contador = Contador + 1
                        ProgressBar1.Value = Round2((Contador * 100) / total, 0)
                        DoEvents
                        Label1(0).Caption = Round(ProgressBar1.Value, 0) & " %"
                        Label1(0).Refresh
        
                        Sql = "fecha=" & DBSet(RS!Fecha, "F") & " and hora='" & Format(RS!hora, "hh:mm:ss") & "' and numeruve=" & RS!NumerUve
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
                        RS.MoveNext
                    Wend
                    RS.Close
        
                    
                    '
                    Sql = "select fecha,hora,numeruve, count(*) from tmptaxi where error1 <> 1 group by 1,2,3 having count(*) > 1"
                    RS.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
                    total = RScontador("select count(*) from (" & Sql & ") aaalias ")  'tmptaxi where error1 <> 1")
                    Label1(2).Caption = "Verificando duplicidad de registros en el fichero."
                    Label1(2).Refresh
                    While Not RS.EOF
                        Contador = Contador + 1
                       ' ProgressBar1.Value = Round2((Contador * 100) / total, 0)
                        DoEvents
                        Label1(0).Caption = Contador
                        Label1(0).Refresh
        
                        Sql = "fecha=" & DBSet(RS!Fecha, "F") & " and hora='" & Format(RS!hora, "hh:mm:ss") & "' and numeruve=" & RS!NumerUve
                        
                        
                        
        '                If SituarDataMULTI(Adodc1, SQL, encontrado) Then
        
                            'esta, entonces es repetido
                            Sql = "UPDATE tmptaxi set error1=1,error='Registro duplicado' where " & Sql
                            conn.Execute Sql
        '                End If
                        RS.MoveNext
                    Wend
                    RS.Close
        
        
        
                    'ahora vamos a buscar en la tabla shilla
                    Sql = "select * from tmptaxi where error1 <> 1"
                    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    ProgressBar1.Value = 0
                    Contador = 0
                    total = RScontador("select count(*) from tmptaxi where error1 <> 1")
                    Label1(2).Caption = "Verificando duplicidad de registros en la tabla."
                    Label1(2).Refresh

                    While Not RS.EOF
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
                        
                        Sql = "select count(*) from shilla where fecha = " & DBSet(RS!Fecha, "F") & " and hora = " & DBSet(RS!hora, "H") & " and numeruve = " & DBSet(RS!NumerUve, "N") & " and (facturad=1 and abonados=1 and validado=1)"
                        If TotalRegistros(Sql) <> 0 Then
                            'esta entonces es repetido
                            Sql = "UPDATE tmptaxi set error1=1,error='Registro duplicado' where id=" & RS!Id
                            conn.Execute Sql
                        End If
                        RS.MoveNext
                    Wend
                    RS.Close
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
            BorrarTablas
            cmdCancel.Caption = "&Cancelar"
     End If
End Sub
Private Function Updatear(vehiculo, encontrado As String) As Boolean
Dim Sql As String

On Error GoTo EUp

Updatear = False

If encontrado = "" Then
    Sql = "UPDATE tmptaxi set error1=1,error='Ningun socio tiene asociado este codigo de vehiculo' where numeruve=" & vehiculo
Else
    Sql = "UPDATE tmptaxi set codsocio=" & CInt(encontrado) & " where numeruve=" & vehiculo
End If

conn.Execute Sql

Updatear = True

EUp:
If Err.Number <> 0 Then
    Updatear = False
End If

End Function

Private Sub MostrarTablas()
Dim RS As ADODB.Recordset
Dim Sql As String

    Set RS = New ADODB.Recordset
    Sql = "select * from tmptaxi where error1=1"
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        If MsgBox("Se ha procesado el fichero correctamente. ¿Desea continuar?.", vbQuestion + vbYesNo) = vbYes Then
            ActualizarTabla
        End If
    Else
        If MsgBox("Ha habido errores en el proceso del fichero. ¿Desea ver los errores?.", vbQuestion + vbYesNo) = vbYes Then
            VerErrores
        End If
        If MsgBox("¿Desea continuar la actualización de las tablas?.", vbQuestion + vbYesNo) = vbYes Then
            ActualizarTabla
        End If
    End If
    RS.Close
    Set RS = Nothing
End Sub

Private Sub ActualizarTabla()
Dim Sql As String
Dim SQL1 As String
Dim RS As ADODB.Recordset
Dim Linea As String
Dim values As String
Dim Contador As Currency
Dim total As Currency
Dim SqlUpdate As String
Dim cWhere As String

    On Error GoTo EActua
    
    Screen.MousePointer = vbHourglass
    
    conn.BeginTrans
    
    Set RS = New ADODB.Recordset
    Sql = "select fecha,hora,codsocio,numeruve,codclien,codusuar,nomclien,dirllama,"
    Sql = Sql & "numllama,puerllama,ciudadre,tipservi,telefono,observac2,codautor,observa1,licencia,"
    Sql = Sql & "matricul,idservic,opereser,opedespa,estado,observa2,fecreser,horreser,fecaviso,"
    Sql = Sql & "horaviso,fecllega,horllega,fecocupa,horocupa,fecfinal,horfinal,importtx,impcompr,"         '[Monica]03/10/2014: añadimos el destino
    Sql = Sql & "extcompr,impventa,extventa,distanci,suplemen,imppeaje,imppropi,facturad,abonados,validado, destino from tmpTaxi where error1<>1"
    RS.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
    total = RScontador("select count(*) from tmpTaxi where error1<>1")
    If total = 0 Then
        MsgBox "No hay datos para actualizar.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Contador = 0
    Linea = ""
    values = ""
    ProgressBar1.Value = 0
    Label1(2).Caption = "Actualizando Bases de datos"
    Label1(2).Refresh
    
    Sql = "INSERT INTO shilla (fecha,hora,codsocio,numeruve,codclien,codusuar,nomclien,dirllama,"
    Sql = Sql & "numllama,puerllama,ciudadre,tipservi,telefono,observac2,codautor,observa1,licencia,"
    Sql = Sql & "matricul,idservic,opereser,opedespa,estado,observa2,fecreser,horreser,fecaviso,"
    Sql = Sql & "horaviso,fecllega,horllega,fecocupa,horocupa,fecfinal,horfinal,importtx,impcompr,"
    Sql = Sql & "extcompr,impventa,extventa,distanci,suplemen,imppeaje,imppropi,facturad,abonados,validado, destino) values "
    
    '[Monica]11/11/2014: dejamos actualizar si no esta liquidada ni facturada
    SqlUpdate = "update shilla set "
    
    
    While Not RS.EOF
        Contador = Contador + 1
        ProgressBar1.Value = Round2((Contador * 100) / total, 0)
        
        cWhere = " fecha = " & DBSet(RS!Fecha, "F") & " and hora = " & DBSet(RS!hora, "H") & " and numeruve = " & DBSet(RS!NumerUve, "N")
        
        If ExisteEnShilla(cWhere) Then
            '[Monica]13/11/2014: sólo en el caso de que sea de credito actualizamos
            If EsdeCredito(cWhere) Then
                Linea = " fecha = " & DBSet(RS!Fecha, "F")
                Linea = Linea & ",hora = " & DBSet(RS!hora, "H")
                Linea = Linea & ",codsocio = " & DBSet(RS!codSocio, "N")
                Linea = Linea & ",numeruve = " & DBSet(RS!NumerUve, "N")
                Linea = Linea & ",codclien = " & DBSet(RS!CodClien, "N")
                Linea = Linea & ",codusuar = " & DBSet(RS!codusuar, "T")
                Linea = Linea & ",nomclien = " & DBSet(RS!nomclien, "T")
                Linea = Linea & ",dirllama = " & DBSet(RS!dirllama, "T")
                Linea = Linea & ",numllama = " & DBSet(RS!numllama, "T")
                Linea = Linea & ",puerllama = " & DBSet(RS!puerllama, "T")
                Linea = Linea & ",ciudadre = " & DBSet(RS!ciudadre, "T")
                Linea = Linea & ",tipservi = " & DBSet(RS!tipservi, "N")
                Linea = Linea & ",telefono = " & DBSet(RS!Telefono, "T")
                Linea = Linea & ",observac2 = " & DBSet(RS!observac2, "T")
                Linea = Linea & ",codautor = " & DBSet(RS!codautor, "T")
                Linea = Linea & ",observa1 = " & DBSet(RS!observa1, "T")
                Linea = Linea & ",licencia = " & DBSet(RS!Licencia, "T")
                Linea = Linea & ",matricul = " & DBSet(RS!matricul, "T")
                Linea = Linea & ",idservic = " & DBSet(RS!idservic, "T")
                Linea = Linea & ",opereser = " & DBSet(RS!opereser, "T")
                Linea = Linea & ",opedespa = " & DBSet(RS!opedespa, "T")
                Linea = Linea & ",estado = " & DBSet(RS!Estado, "T")
                Linea = Linea & ",observa2 = " & DBSet(RS!observa2, "T")
                Linea = Linea & ",fecreser = " & DBSet(RS!fecreser, "F")
                Linea = Linea & ",horreser = " & DBSet(RS!horreser, "H")
                Linea = Linea & ",fecaviso = " & DBSet(RS!fecaviso, "F")
                Linea = Linea & ",horaviso = " & DBSet(RS!horaviso, "H")
                Linea = Linea & ",fecllega = " & DBSet(RS!fecllega, "F")
                Linea = Linea & ",horllega = " & DBSet(RS!horllega, "H")
                Linea = Linea & ",fecocupa = " & DBSet(RS!fecocupa, "F")
                Linea = Linea & ",horocupa = " & DBSet(RS!horocupa, "H")
                Linea = Linea & ",fecfinal = " & DBSet(RS!fecfinal, "F")
                Linea = Linea & ",horfinal = " & DBSet(RS!horfinal, "H")
                Linea = Linea & ",importtx = " & DBSet(RS!importtx, "N")
                Linea = Linea & ",impcompr = " & DBSet(RS!impcompr, "N")
                Linea = Linea & ",extcompr = " & DBSet(RS!extcompr, "N")
                Linea = Linea & ",impventa = " & DBSet(RS!impventa, "N")
                Linea = Linea & ",extventa = " & DBSet(RS!extventa, "N")
                Linea = Linea & ",distanci = " & DBSet(RS!distanci, "N")
                Linea = Linea & ",suplemen = " & DBSet(RS!suplemen, "N")
                Linea = Linea & ",imppeaje = " & DBSet(RS!imppeaje, "N")
                Linea = Linea & ",imppropi = " & DBSet(RS!imppropi, "N")
                Linea = Linea & ",facturad = " & DBSet(RS!facturad, "N")
                Linea = Linea & ",abonados = " & DBSet(RS!abonados, "N")
                Linea = Linea & ",validado = " & DBSet(RS!validado, "N")
                Linea = Linea & ",destino = " & DBSet(RS!Destino, "T")
                Linea = Linea & " where " & cWhere
                
                conn.Execute SqlUpdate & Linea
            End If
        Else
            
            If IsNull(RS!Fecha) Then
                Linea = "(NULL,"
            Else
                Linea = "(" & DBSet(RS!Fecha, "F") & ","
            End If
            If IsNull(RS!hora) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & "'" & Format(RS!hora, FormatoHora) & "',"
            End If
            If IsNull(RS!codSocio) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!codSocio, "N") & ","
            End If
            If IsNull(RS!NumerUve) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!NumerUve, "N") & ","
            End If
            If IsNull(RS!CodClien) Or RS!CodClien = 0 Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!CodClien, "N") & ","
            End If
            If IsNull(RS!codusuar) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!codusuar, "T") & ","
            End If
            If IsNull(RS!nomclien) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!nomclien, "T") & ","
            End If
            If IsNull(RS!dirllama) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!dirllama, "T") & ","
            End If
            If IsNull(RS!numllama) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!numllama, "T") & ","
            End If
            If IsNull(RS!puerllama) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!puerllama, "T") & ","
            End If
            If IsNull(RS!ciudadre) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!ciudadre, "T") & ","
            End If
            If IsNull(RS!tipservi) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!tipservi, "N") & ","
            End If
            If IsNull(RS!Telefono) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!Telefono, "T") & ","
            End If
            If IsNull(RS!observac2) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!observac2, "T") & ","
            End If
            If IsNull(RS!codautor) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!codautor, "T") & ","
            End If
            If IsNull(RS!observa1) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!observa1, "T") & ","
            End If
            If IsNull(RS!Licencia) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!Licencia, "T") & ","
            End If
            If IsNull(RS!matricul) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!matricul, "T") & ","
            End If
            If IsNull(RS!idservic) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!idservic, "T") & ","
            End If
            If IsNull(RS!opereser) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!opereser, "T") & ","
            End If
            If IsNull(RS!opedespa) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!opedespa, "T") & ","
            End If
            If IsNull(RS!Estado) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!Estado, "T") & ","
            End If
            If IsNull(RS!observa2) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!observa2, "T") & ","
            End If
            If IsNull(RS!fecreser) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & "'" & Format(RS!fecreser, FormatoFecha) & "',"
            End If
            If IsNull(RS!horreser) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & "'" & Format(RS!horreser, FormatoHora) & "',"
            End If
            If IsNull(RS!fecaviso) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & "'" & Format(RS!fecaviso, FormatoFecha) & "',"
            End If
            If IsNull(RS!horaviso) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & "'" & Format(RS!horaviso, FormatoHora) & "',"
            End If
            If IsNull(RS!fecllega) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & "'" & Format(RS!fecllega, FormatoFecha) & "',"
            End If
            If IsNull(RS!horllega) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & "'" & Format(RS!horllega, FormatoHora) & "',"
            End If
            If IsNull(RS!fecocupa) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & "'" & Format(RS!fecocupa, FormatoFecha) & "',"
            End If
            If IsNull(RS!horocupa) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & "'" & Format(RS!horocupa, FormatoHora) & "',"
            End If
            If IsNull(RS!fecfinal) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & "'" & Format(RS!fecfinal, FormatoFecha) & "',"
            End If
            If IsNull(RS!horfinal) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & "'" & Format(RS!horfinal, FormatoHora) & "',"
            End If
            If IsNull(RS!importtx) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!importtx, "N") & ","
            End If
            If IsNull(RS!impcompr) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!impcompr, "N") & ","
            End If
            If IsNull(RS!extcompr) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!extcompr, "N") & ","
            End If
            If IsNull(RS!impventa) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!impventa, "N") & ","
            End If
            If IsNull(RS!extventa) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!extventa, "N") & ","
            End If
            If IsNull(RS!distanci) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!distanci, "N") & ","
            End If
            If IsNull(RS!suplemen) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!suplemen, "N") & ","
            End If
            If IsNull(RS!imppeaje) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!imppeaje, "N") & ","
            End If
            If IsNull(RS!imppropi) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!imppropi, "N") & ","
            End If
            If IsNull(RS!facturad) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!facturad, "N") & ","
            End If
            If IsNull(RS!abonados) Then
                Linea = Linea & "NULL,"
            Else
                Linea = Linea & DBSet(RS!abonados, "N") & ","
            End If
            If IsNull(RS!validado) Then
                Linea = Linea & "NULL)"
            Else
                Linea = Linea & DBSet(RS!validado, "N") & ","
            End If
            If IsNull(RS!Destino) Then
                Linea = Linea & "NULL)"
            Else
                Linea = Linea & DBSet(RS!Destino, "T") & ")"
            End If
            values = values & Linea & ","
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
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
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

Private Function ExisteEnShilla(vWhere As String) As Boolean
Dim Sql As String

    On Error Resume Next
    
    
    Sql = "select count(*) from shilla where " & vWhere
    ExisteEnShilla = (TotalRegistros(Sql) <> 0)


End Function

Private Function EsdeCredito(vWhere As String) As Boolean
Dim Sql As String

    On Error Resume Next
    
    Sql = "select count(*) from shilla where " & vWhere & " and tipservi = 1"
    EsdeCredito = (TotalRegistros(Sql) <> 0)

End Function

Private Sub VerErrores()
    frmFacturas.Socio = False
    frmFacturas.Sql = "select numeruve,fecha,hora,error from tmptaxi where error1=1"
    frmFacturas.Caption = "Errores en el fichero de traspaso"
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
    Me.Icon = frmPpal.Icon

    txtcodigo(85).Text = Format(Now, "dd/mm/yyyy")

    EnTomaDeDatos
    BorrarTablas
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
Dim Linea As String
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
    Linea = "(id,telefono,codclien,codautor,codusuar,nomclien,tipservi,observa1,numeruve,licencia,matricul,"
    Linea = Linea & "dirllama,ciudadre,numllama,puerllama,fecha,hora,idservic,opereser,opedespa,estado,"
    Linea = Linea & "observa2,fecreser,horreser,fecaviso,horaviso,fecllega,horllega,fecocupa,horocupa,fecfinal,"
    Linea = Linea & "horfinal,importtx,impcompr,extcompr,impventa,extventa,distanci,suplemen,imppeaje,imppropi,facturad,"
    '[Monica]03/10/2014: añadimos el taxi del destino
    Linea = Linea & "abonados,validado,destino,error1,error)"
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
            Sql = "INSERT INTO tmptaxi " & Linea & " VALUES " & values
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
    
        Sql = "INSERT INTO tmptaxi " & Linea & " VALUES " & values
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
Private Sub ArmarCadena(CADENA As String, ByRef values1 As String)
Dim Telefono As String
Dim values As String
Dim Error As String
Dim Error1 As Byte

Dim Valor As Double

    Fecha = Mid(menErrProceso, 397, 10)
    hora = Mid(menErrProceso, 407, 8)
    vehiculo = Mid(menErrProceso, 293, 4)


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
If Not IsNumeric(vehiculo) Then
    values = values & ",NULL"
    Error1 = 1
    Error = "Vehiculo con formato incorrecto"
Else
    values = values & "," & CInt(vehiculo)
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
values1 = values1 & values & ","

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
        vehiculo = Mid(menErrProceso, 293, 4)
        If Fecha = "" Then
            Insertar menErrProceso, "tmperr", "falta fecha"
        ElseIf Not IsDate(Fecha) Then
                Insertar menErrProceso, "tmperr", "fecha formato incorrecto"
            ElseIf hora = "" Then
                    Insertar menErrProceso, "tmperr", "falta hora"
                ElseIf Not IsDate(hora) Then
                    Insertar menErrProceso, "tmperr", "hora formato incorrecto"
                ElseIf vehiculo = "" Then
                    Insertar menErrProceso, "tmperr", "falta vehiculo"
                ElseIf Not IsNumeric(vehiculo) Then
                    Insertar menErrProceso, "tmperr", "vehiculo formato incorrecto"
            Else
                DoEvents
                If BuscarRegistro(vehiculo, Fecha, hora) Then
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
Dim RS As ADODB.Recordset

    On Error GoTo EBusqueda
    
    Set RS = New ADODB.Recordset
    fec = CDate(F)
    hora = CDate(H)

    Sql = "select * from tmptaxi where numeruve=" & CInt(V) & " and fecha='" & Format(fec, FormatoFecha)
    Sql = Sql & "' and hora='" & Format(hora, FormatoHora) & "'"

    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RS.EOF Then
        RS.Close
        Sql = DevuelveDesdeBD(conAri, "codclien", "sclien", "numeruve", V, "T")
        If Sql <> "" Then
            Sql = "select * from shilla where codsocio=" & CInt(Sql) & " and numeruve=" & CInt(V) & " and fecha='" & Format(fec, FormatoFecha)
            Sql = Sql & "' and hora='" & Format(hora, FormatoHora) & "'"
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RS.EOF Then
                BuscarRegistro = False
            Else
                BuscarRegistro = True
            End If
            RS.Close
        Else
            BuscarRegistro = False
        End If
    Else
        BuscarRegistro = True
        RS.Close
    End If
    
EBusqueda:
If Err.Number <> 0 Then
    MsgBox "Error al buscar registro:" & Err.Description
    BuscarRegistro = False
End If
    
End Function
Private Sub Insertar(CADENA As String, Tabla As String, Optional Error As String)
Dim Telefono As String
Dim Linea As String, values As String
Dim Sql As String
Dim Socio As String

On Error GoTo EInsert

'armamos los registros segun la cadena
Telefono = Trim(Mid(CADENA, 1, 11))
Linea = "telefono"

If Telefono = "" Then
    values = "NULL"
Else
    values = DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 11, 4))
Linea = Linea & "," & "codclien"
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
Linea = Linea & "," & "codautor"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 29, 30))
Linea = Linea & "," & "codusuar"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 59, 30))
Linea = Linea & "," & "nomclien"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 89, 1))
Linea = Linea & "," & "tipservi"
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
Linea = Linea & "," & "observa1"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Linea = Linea & "," & "numeruve"
If Not IsNumeric(vehiculo) Then
    values = values & ",NULL"
    Tabla = "tmpErr"
    Error = "Vehiculo con formato incorrecto"
    values = values & ",NULL"
Else
    values = values & "," & CInt(vehiculo)
    'con el número de vehiculo buscamos el socio,si no lo encontramos
    'preparamos para agregar en la tabla de errores
    Socio = DevuelveDesdeBD(conAri, "codclien", "sclien", "numeruve", vehiculo, "T")
    If Socio = "" Then
        Tabla = "tmpErr"
        Error = "Ningun socio tiene asociado este codigo de vehiculo"
        values = values & ",NULL"
    Else
        values = values & "," & CInt(Socio)
    End If
End If
Linea = Linea & "," & "codsocio"

Telefono = Trim(Mid(CADENA, 297, 10))
Linea = Linea & "," & "licencia"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 307, 10))
Linea = Linea & "," & "matricul"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 317, 30))
Linea = Linea & "," & "dirllama"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 347, 30))
Linea = Linea & "," & "ciudadre"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 377, 10))
Linea = Linea & "," & "numllama"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 387, 10))
Linea = Linea & "," & "puerllama"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Linea = Linea & "," & "fecha"
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

Linea = Linea & "," & "hora"
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
Linea = Linea & "," & "idservic"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If


Telefono = Trim(Mid(CADENA, 421, 30))
Linea = Linea & "," & "opereser"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 451, 30))
Linea = Linea & "," & "opedespa"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 481, 4))
Linea = Linea & "," & "estado"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 485, 200))
Linea = Linea & "," & "observa2"
If Telefono = "" Then
    values = values & ",NULL"
Else
    values = values & "," & DBSet(Telefono, "T")
End If

Telefono = Trim(Mid(CADENA, 685, 10))
Linea = Linea & "," & "fecreser"
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
Linea = Linea & "," & "horreser"
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
Linea = Linea & "," & "fecaviso"
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
Linea = Linea & "," & "horaviso"
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
Linea = Linea & "," & "fecllega"
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
Linea = Linea & "," & "horllega"
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
Linea = Linea & "," & "fecocupa"
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
Linea = Linea & "," & "horocupa"
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
Linea = Linea & "," & "fecfinal"
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
Linea = Linea & "," & "horfinal"
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
Linea = Linea & "," & "importtx"
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
Linea = Linea & "," & "impcompr"
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
Linea = Linea & "," & "extcompr"
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
Linea = Linea & "," & "impventa"
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
Linea = Linea & "," & "extventa"
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
Linea = Linea & "," & "distanci"
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
Linea = Linea & "," & "suplemen"
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
Linea = Linea & "," & "imppeaje"
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
Linea = Linea & "," & "imppropi"
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
Linea = Linea & "," & "facturad"
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
Linea = Linea & "," & "abonados"
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
Linea = Linea & "," & "validado"
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
    Linea = Linea & ",error1"
    values = values & ",0"
Else
   Linea = Linea & ",error1,error"
   values = values & ",1," & DBSet(Error, "T")
End If
Sql = "INSERT INTO tmptaxi (" & Linea & ") VALUES ("
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
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
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
    
    Line Input #NF, cad
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
        
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Len(cad)
        Label1(2).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarRegistro(cad, Escliente)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        i = i + 1
        
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Len(cad)
        Label1(2).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarRegistro(cad, Escliente)
    End If
    
    ProgressBar1.visible = False
    Label1(2).Caption = ""
    Label1(0).Caption = ""

    ComprobarFichero = b
    Exit Function

eComprobarFichero:
    ComprobarFichero = False
End Function


Private Function ComprobarRegistro(cad As String, EsClien As Boolean) As Boolean
Dim Sql As String
Dim c_Importe As Currency
Dim Mens As String
Dim CodSoc As String
Dim Id As String
Dim Importe As String
Dim NServicios As String

Dim vServicios As Long
Dim vImporte As Currency

Dim RS As ADODB.Recordset

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True

    If EsClien Then ' facturacion a clientes
        Id = Mid(cad, 1, 6)
        Importe = Mid(cad, 352, 10)
        NServicios = Mid(cad, 362, 5)
    Else ' liquidacion a socios
        Id = Mid(cad, 1, 6)
        Importe = Mid(cad, 375, 10)
        NServicios = Mid(cad, 385, 5)
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
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
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
    
    Line Input #NF, cad
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
        
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Len(cad)
        Label1(2).Caption = "Linea " & i
        Me.Refresh
        
        If EsClien Then ' facturacion a clientes
            Id = Mid(cad, 1, 6)
            Importe = Mid(cad, 352, 10)
            NServicios = Mid(cad, 362, 5)
        Else ' liquidacion a socios
            Id = Mid(cad, 1, 6)
            Importe = Mid(cad, 375, 10)
            NServicios = Mid(cad, 385, 5)
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
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Len(cad)
        Label1(2).Caption = "Linea " & i
        Me.Refresh
        
        If EsClien Then ' facturacion a clientes
            Id = Mid(cad, 1, 6)
            Importe = Mid(cad, 352, 10)
            NServicios = Mid(cad, 362, 5)
        Else ' liquidacion a socios
            Id = Mid(cad, 1, 6)
            Importe = Mid(cad, 375, 10)
            NServicios = Mid(cad, 385, 5)
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

