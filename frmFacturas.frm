VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacturas 
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameBotonGnral 
      Height          =   600
      Left            =   270
      TabIndex        =   6
      Top             =   135
      Width           =   1230
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Comprobar"
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Left            =   5295
      TabIndex        =   1
      Top             =   5760
      Width           =   1135
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
      Left            =   6555
      TabIndex        =   2
      Top             =   5760
      Width           =   1135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3795
      Left            =   270
      TabIndex        =   0
      Top             =   900
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6694
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   225
      TabIndex        =   3
      Top             =   4770
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
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
      Height          =   210
      Index           =   0
      Left            =   255
      TabIndex        =   5
      Top             =   5130
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
      Height          =   210
      Index           =   2
      Left            =   255
      TabIndex        =   4
      Top             =   5400
      Width           =   7425
   End
End
Attribute VB_Name = "frmFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sql As String
Public Socio As Boolean
Public deExcel As Boolean

Private Const IdPrograma = 206


Private Sub cmdAceptar_Click()
    CadenaDesdeOtroForm = "DATOS"
    Unload Me
End Sub


Private Sub cmdCancelar_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub


Private Sub Comprobaciones()
Dim b As Boolean
Dim Contador As Long
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim total As Long
Dim encontrado As String

     b = True
     If b Then
     
         Sql = "update tmptaxi set error1 = 3 where error1 = 1"
         conn.Execute Sql
     
         ComprobacionDatos
     
         'verificamos que los numeruve esten asociados a algun socio
         ProgressBar1.Value = 0
         Contador = 0
         Label1(0).Caption = ""
         Set Rs = New ADODB.Recordset
         Sql = "select * from tmptaxi where error1 = 3 group by numeruve"
         Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
         total = rsContador("select count(distinct(numeruve)) from tmptaxi where error1=3")
         Label1(2).Caption = "Verificando códigos de socios."
         Label1(2).Refresh
 
'         While Not RS.EOF
'             Contador = Contador + 1
'             ProgressBar1.Value = (Contador * 100) / total
'             DoEvents
'             'Label1(0).Caption = Round(ProgressBar1.Value, 2) & "%"
'             Label1(0).Caption = Round2(ProgressBar1.Value, 0) & " %"
'             Label1(0).Refresh
'             encontrado = DevuelveDesdeBD(conAri, "codclien", "sclien", "numeruve", RS!NumerUve, "T")
'             b = Updatear(RS!NumerUve, encontrado)
'             RS.MoveNext
'         Wend
'[Monica]28/12/2017: cambiado por
        If deExcel Then

            While Not Rs.EOF
                Contador = Contador + 1
                ProgressBar1.Value = (Contador * 100) / total
                DoEvents
                'Label1(0).Caption = Round(ProgressBar1.Value, 2) & "%"
                Label1(0).Caption = Round2(ProgressBar1.Value, 0) & " %"
                Label1(0).Refresh
                
'[Monica]09/02/2018: quito la concicion Trim(vParam.CifEmpresa) = "B98877806"
'                    pq ahora tele es igual que radio en el numeruve lleva la licencia
'                '??????????
'                ' me viene la licencia (caso de Radio Taxi en la V llevo la licencia)
'                If Trim(vParam.CifEmpresa) = "B98877806" Then
                
                    encontrado = DevuelveDesdeBD(conAri, "codclien", "sclien", "numeruve", Rs!NumerUve, "T")
                    
                    b = Updatear(Rs!NumerUve, encontrado, False)
                
'                Else
'                    encontrado = DevuelveDesdeBD(conAri, "codclien", "sclien", "licencia", Rs!NumerUve, "T")
'
'                    If encontrado <> "" Then
'                        ' pq me viene la licencia
'                        Dim rs4 As ADODB.Recordset
'                        Dim Sql4 As String
'                        Set rs4 = New ADODB.Recordset
'                        Sql4 = "select codclien from sclien where licencia = " & DBSet(Rs!NumerUve, "N") & " and not numeruve is null and numeruve <> 0"
'                        rs4.Open Sql4, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'                        encontrado = ""
'
'                        If Not rs4.EOF Then
'                            encontrado = DBLet(rs4.Fields(0))
'                        End If
'                        Set rs4 = Nothing
'
'                        b = Updatear(Rs!NumerUve, encontrado, True)
'                    Else
'                        b = Updatear(Rs!NumerUve, encontrado, False)
'                    End If
'
'                End If
                                        
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
                
                encontrado = DevuelveDesdeBD(conAri, "codclien", "sclien", "numeruve", Rs!NumerUve, "T")
                b = Updatear(Rs!NumerUve, encontrado, False)
                Rs.MoveNext
            Wend
        
        End If
        

         Rs.Close
         Label1(0).Caption = ""
         Label1(0).Refresh
         
         '[Monica]12/12/2017: por el tema de fusion de empresas, SOLO SI VIENE DE EXCEL
         '                    si el fichero es de la otra empresa ponemos que el cliente es el gros
         If deExcel Then
             If b Then
                 If ComprobarCero(vParamAplic.EmpresaTaxitronic) <> 0 Then
                     Label1(2).Caption = "Modificando códigos de cliente de otra empresa"
                     Label1(2).Refresh
                     
                     Sql = "update tmptaxi set codclien = " & DBSet(vParamAplic.ClienteCooperativa, "N")
                     Sql = Sql & " where error1 = 3 and empresa <> " & vParamAplic.EmpresaTaxitronic
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
             Sql = "select numeruve,fecha,hora, count(*) from tmptaxi where error1 = 3 group by 1,2,3 having count(*) > 1"
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
 
                 Sql = "numeruve=" & Rs!NumerUve & " and fecha=" & DBSet(Rs!Fecha, "F") & " and hora='" & Format(Rs!hora, "hh:mm:ss") & "' "
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
             Sql = "select numeruve,fecha,hora, count(*) from tmptaxi where error1 = 3 group by 1,2,3 having count(*) > 1"
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
 
                 Sql = "numeruve=" & Rs!NumerUve & " and fecha=" & DBSet(Rs!Fecha, "F") & " and hora='" & Format(Rs!hora, "hh:mm:ss") & "' "
                 
                 
                 
 '                If SituarDataMULTI(Adodc1, SQL, encontrado) Then
 
                     'esta, entonces es repetido
                     Sql = "UPDATE tmptaxi set error1=1,error='Registro duplicado' where " & Sql
                     conn.Execute Sql
 '                End If
                 Rs.MoveNext
             Wend
             Rs.Close
 
             '[Monica]28/12/2017: para el caso de Tele y Alfa 6 pongo el numero de V correcto
            If Trim(vParam.CifEmpresa) <> "B98877806" And deExcel Then
                Dim NUve As Long
            
                Sql = "select codsocio from tmptaxi where error1 = 3 and codsocio <> " & vParamAplic.SocioCooperativa & " group by 1"
                
                Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
                
                total = rsContador("select count(*) from (" & Sql & ") aaalias ")  'tmptaxi where error1 <> 1")
                Label1(2).Caption = "Modificando Vehículo en registros del fichero."
                Label1(2).Refresh
                
                While Not Rs.EOF
                    Contador = Contador + 1
                   ' ProgressBar1.Value = Round2((Contador * 100) / total, 0)
                    DoEvents
                    Label1(0).Caption = Contador
                    Label1(0).Refresh
                
                    Sql = "select numeruve from sclien where codclien = " & DBSet(Rs!codSocio, "N")
                    NUve = DevuelveValor(Sql)
                
                    Sql = "UPDATE tmptaxi set numeruve = " & DBSet(NUve, "N") & " where codsocio = " & DBSet(Rs!codSocio, "N") & " and error1 = 3 "
                    conn.Execute Sql
                    
                    Rs.MoveNext
                Wend
                
                Rs.Close
            End If


 
 
             'ahora vamos a buscar en la tabla shilla
             Sql = "select * from tmptaxi where error1 = 3"
             Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
             ProgressBar1.Value = 0
             Contador = 0
             total = rsContador("select count(*) from tmptaxi where error1 = 3")
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
                 
                 Sql = "select count(*) from shilla where numeruve = " & DBSet(Rs!NumerUve, "N") & " and fecha = " & DBSet(Rs!Fecha, "F") & " and hora = " & DBSet(Rs!hora, "H") & " and  (facturad=1 and abonados=1 and validado=1)"
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

     'los que continuan con 3 es que ya no tienen error
     Sql = "update tmptaxi set error1 = 0 where error1 = 3"
     conn.Execute Sql




     Label1(0).Caption = ""
     Label1(2).Caption = ""
     Me.ProgressBar1.visible = False
     
     CargaGrid DataGrid1, Adodc1

End Sub
Private Function rsContador(CADENA As String) As Currency
    
    rsContador = 0
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open CADENA, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        rsContador = miRsAux.Fields(0)
    End If
    miRsAux.Close
    
End Function

Private Sub Imprimir()

Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim ImprimeDirecto As Boolean


    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el parametro de Empresa
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 1
    
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = "rErroresTraspaso2.rpt"
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Factura
    '---------------------------------------------------
    devuelve = "{tmptaxi.error1} = 1"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
    cadSelect = cadFormula
   
    If Not HayRegParaInforme("tmptaxi", cadSelect) Then Exit Sub
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .NombrePDF = pPdfRpt
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 4
        .Titulo = ""
        .Show vbModal
    End With

End Sub

Private Sub DataGrid1_Click()

    frmGesHisLlamTMP.DatosADevolverBusqueda = "id = " & Adodc1.Recordset!Id
    frmGesHisLlamTMP.Show vbModal
    
    CargaGrid Me.DataGrid1, Me.Adodc1
    
End Sub





Private Sub Form_Load()

    'Icono del formulario
    Me.Icon = frmppal.Icon

    Screen.MousePointer = vbDefault
    Adodc1.ConnectionString = conn
    Adodc1.RecordSource = Sql
    Adodc1.Refresh
    
    If Not Adodc1.Recordset.EOF Then
        CargaGrid DataGrid1, Adodc1
    End If

    With Me.Toolbar3(0)
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.imgListComun1
        .Buttons(1).Image = 16  'impresora
        .Buttons(2).Image = 18  'Comprobar
    End With

    '[Monica]13/12/2017: para comprobar sin salir
    FrameBotonGnral.visible = Not Socio
    FrameBotonGnral.Enabled = Not Socio
    
    Label1(0).visible = Not Socio
    Label1(2).visible = Not Socio
    Me.ProgressBar1.visible = False
    
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Integer

On Error GoTo ECargaGrid

    vData.Refresh
    Set vDataGrid.DataSource = vData
    If Socio Then
        vDataGrid.Columns(0).Caption = "Socio"
        vDataGrid.Columns(0).Width = 3100
        vDataGrid.Columns(1).Caption = "Importes"
        vDataGrid.Columns(1).Width = 1200
        vDataGrid.Columns(1).NumberFormat = "#,###,###,##0.00"
        vDataGrid.Columns(2).Caption = "Desde"
        vDataGrid.Columns(2).Width = 1000
        vDataGrid.Columns(3).Caption = "Hasta"
        vDataGrid.Columns(3).Width = 1000
    Else
        vDataGrid.Columns(0).Caption = "Vehiculo"
        vDataGrid.Columns(0).Width = 1000
        vDataGrid.Columns(1).Caption = "Fecha"
        vDataGrid.Columns(1).Width = 1200
        vDataGrid.Columns(2).Caption = "Hora"
        vDataGrid.Columns(2).Width = 1000
        vDataGrid.Columns(2).NumberFormat = "hh:mm:ss"
        vDataGrid.Columns(3).Caption = "Error"
        vDataGrid.Columns(3).Width = 3600
        vDataGrid.Columns(4).Caption = "Id"
        vDataGrid.Columns(4).Width = 0
    End If
    
    vDataGrid.RowHeight = 350
    
    
    vDataGrid.Enabled = True
    For I = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(I).Locked = True
        vDataGrid.Columns(I).AllowSizing = False
    Next I

    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Function Updatear(Vehiculo, encontrado As String, LicenciaSinV As Boolean) As Boolean
Dim Sql As String

On Error GoTo EUp

Updatear = False

If encontrado = "" Then
'[Monica]12/12/2017: ahora si no encuentro el socio que lleva ese numero de vehiculo es que es de la otra empresa
'                    si viene de fichero plano lo marco como error
    If Not deExcel Or LicenciaSinV Then
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


Private Sub ComprobacionDatos()
Dim Rs As ADODB.Recordset
Dim Telefono As String
Dim Values1 As String
Dim Error As String
Dim Error1 As Byte
Dim FechaHora As String

Dim Valor As Double
Dim Fecha As String
Dim hora As String
Dim Vehiculo As String

Dim Contador As Long
Dim total As Long


    Contador = 0
    Label1(0).Caption = ""
    total = rsContador("select count(*) from tmptaxi where error1=3")
    Label1(2).Caption = "Comprobación de datos."
    Label1(2).Refresh

    Sql = "select * from tmptaxi where error1 = 3"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Contador = Contador + 1
        ProgressBar1.Value = (Contador * 100) / total
        DoEvents
    
        Error1 = 0
        Error = ""
    
        Fecha = DBLet(Rs!Fecha, "F")
        hora = DBLet(Rs!hora, "F")
        
        If hora = "" Then hora = "00:00:00"
        
        Vehiculo = DBLet(Rs!NumerUve, "N")
    
        Error1 = 0
        Error = ""
        'armamos los registros segun la cadena
        Telefono = DBLet(Rs!Telefono, "T")
        'telefono
    
        Telefono = DBLet(Rs!CodClien, "N")
        'codclien
        
        Telefono = DBLet(Rs!codautor, "T")
        'codautor"
    
        Telefono = DBLet(Rs!codusuar, "T")
        'codusuar"
        
        Telefono = DBLet(Rs!nomclien, "T")
        'nomclien"
    
        Telefono = DBLet(Rs!tipservi, "N")
        'tipservi"
        If Telefono <> "0" And Telefono <> "1" Then
            Error1 = "1"
            Error = "tipservi con formato incorrecto"
        End If
        
        Telefono = DBLet(Rs!observa1, "T")
        'observa1"
    
        'numeruve"
        If ComprobarCero(Vehiculo) = 0 Then
            Error1 = 1
            Error = "Vehiculo con formato incorrecto"
        End If
    
        Telefono = DBLet(Rs!Licencia, "T")
        'licencia"
    
        Telefono = DBLet(Rs!matricul, "T")
        'matricul"
    
        Telefono = DBLet(Rs!dirllama, "T")
        'dirllama"
        
        Telefono = DBLet(Rs!ciudadre, "T")
        'ciudadre"
        
    
        'fecha"
        If Fecha = "" Then
            Error1 = 1
            Error = "Falta fecha"
        End If
        
        'hora"
        If hora = "" Then
            Error1 = 1
            Error = "Falta hora"
        End If
    
        Telefono = DBLet(Rs!idservic, "T")
        'idservic"
    
        Telefono = DBLet(Rs!opereser, "T")
        'opereser"
        
        Telefono = DBLet(Rs!opedespa, "T")
        'opedespa"
    
    
        '****** NO HE ENCONTRADO EL ESTADO
        '******
        Telefono = "" 'Trim(Mid(Cadena, 481, 4))
        'estado"
    
        Telefono = DBLet(Rs!observa2, "T")
        'observa2"
    
        
        '[Monica]02/08/2017: añadido
        Fecha = DBLet(Rs!fecreser, "F")
        hora = DBLet(Rs!horreser, "F")
        
        'fecreser"
        If Fecha = "" Then
        ElseIf Not IsDate(Fecha) Then
            Error1 = 1
            Error = "fecha reserva con formato incorrecto"
        End If
    
        'horreser"
        If hora = "" Then
        ElseIf Not IsDate(hora) Then
            Error1 = 1
            Error = "hora reserva con formato incorrecto"
        End If
        
        
        Fecha = DBLet(Rs!fecaviso, "F")
        hora = DBLet(Rs!horaviso, "F")
        'fecaviso"
        If Fecha = "" Then
        ElseIf Not IsDate(Fecha) Then
            Error1 = 1
            Error = "fecha aviso con formato incorrecto"
        End If
    
        'horaviso"
        If hora = "" Then
        ElseIf Not IsDate(hora) Then
            Error1 = 1
            Error = "hora aviso con formato incorrecto"
        End If
    
    
        Fecha = DBLet(Rs!fecllega, "F")
        hora = DBLet(Rs!horllega, "F")
        'fecllega"
        If Fecha = "" Then
        ElseIf Not IsDate(Fecha) Then
            Error1 = 1
            Error = "fecha llegada con formato incorrecto"
        End If
        
        'horllega"
        If hora = "" Then
        ElseIf Not IsDate(hora) Then
            Error1 = 1
            Error = "hora llegada con formato incorrecto"
        End If
    
        Fecha = DBLet(Rs!fecocupa, "F")
        hora = DBLet(Rs!horocupa, "F")
        'fecocupa"
        If Fecha = "" Then
        ElseIf Not IsDate(Fecha) Then
            Error1 = 1
            Error = "fecha ocupa con formato incorrecto"
        End If
        
        'horocupa"
        If hora = "" Then
        ElseIf Not IsDate(hora) Then
            Error1 = 1
            Error = "hora ocupa con formato incorrecto"
        End If
    
    
        Fecha = DBLet(Rs!fecfinal, "F")
        hora = DBLet(Rs!horfinal, "F")
        'fecfinal"
        If Fecha = "" Then
        ElseIf Not IsDate(Fecha) Then
            Error1 = 1
            Error = "fecha final con formato incorrecto"
        End If
        
        'horfinal"
        If hora = "" Then
        ElseIf Not IsDate(hora) Then
            Error1 = 1
            Error = "hora final con formato incorrecto"
        End If
    
    
        Telefono = DBLet(Rs!importtx, "N")
        'importtx"
    
    
        Telefono = DBLet(Rs!impcompr, "N")
        'impcompr"
    
    
        Telefono = DBLet(Rs!extcompr, "N")
        'extcompr"
    
        Telefono = DBLet(Rs!impventa, "N")
        'impventa"
    
        Telefono = DBLet(Rs!extventa, "N")
        'extventa"
    
        Telefono = DBLet(Rs!distanci, "T")
        'distanci"
    
        '[Monica]30/11/2017: ya tenemos el suplemento y peaje
        Telefono = DBLet(Rs!suplemen, "N")
        'suplemen"
    
        Telefono = DBLet(Rs!imppeaje, "N")
        'imppeaje"
        
        Telefono = DBLet(Rs!imppropi, "N")
        'imppropi"
    
        Telefono = DBLet(Rs!facturad, "N")
        'facturad"
        If Telefono <> "0" And Telefono <> "1" Then
            Error1 = 1
            Error = "facturado con formato incorrecto"
        End If
    
        Telefono = DBLet(Rs!abonados, "N")
        'abonados"
        If Telefono <> "0" And Telefono <> "1" Then
            Error1 = 1
            Error = "abonado con formato incorrecto"
        End If
    
        Telefono = DBLet(Rs!validado, "N")
        'validado"
        If Telefono <> "0" And Telefono <> "1" Then
            Error1 = 1
            Error = "validado con formato incorrecto"
        End If
    
        '[Monica]03/10/2014: añadimos el destino del servicio
        Telefono = DBLet(Rs!Destino, "T")
        'destino"
        
        '[Monica]12/12/2017: quien coge la llamada (radio o tele)
        Telefono = DBLet(Rs!Empresa, "N")
        'empresa"
        
        'error1,error
        If Error1 = 1 Then
            Sql = "update tmptaxi set error1 = 1, error = " & DBSet(Error, "T") & " where id = " & DBSet(Rs!Id, "N")
            conn.Execute Sql
        End If
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
EInsert:
    If Err.Number <> 0 Then
        MsgBox "Error en comprobación de datos. " & Err.Description
    End If
End Sub



Private Sub Toolbar3_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' Imprimir
            Imprimir
        Case 2 ' Comprobar
            Comprobaciones
    End Select
End Sub
