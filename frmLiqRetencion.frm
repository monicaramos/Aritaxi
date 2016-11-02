VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLiqRetencion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retenciones de Socio"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11085
   ClipControls    =   0   'False
   Icon            =   "frmLiqRetencion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   5010
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmLiqRetencion.frx":000C
      Left            =   7890
      List            =   "frmLiqRetencion.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Tipo Retencion|N|N|0|4|sreten|tiporeten|||"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   5400
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   16
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   14
      ToolTipText     =   "Buscar socio"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   6990
      TabIndex        =   13
      ToolTipText     =   "Buscar fecha"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   360
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Cod. Socio|N|N|0|999999|sreten|codsocio|000000|S|"
      Text            =   "socio"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   5310
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "Uve|N|N|0|9999|sreten|numeruve|0000|N|"
      Text            =   "Uve"
      Top             =   3630
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   7230
      MaxLength       =   7
      TabIndex        =   3
      Tag             =   "Factura|N|N|||sreten|numfactu|0000000|S|"
      Text            =   "Factu"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   5970
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||sreten|fecfactu|dd/mm/yyyy|S|"
      Text            =   "fecha"
      Top             =   3630
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   5400
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9675
      TabIndex        =   7
      Top             =   5400
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9675
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   11
      Text            =   "Text2 Text2 Text2 Text2 Text2 Text2 Text"
      Top             =   3600
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   9690
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Importe|N|N|||sreten|impreten|###,##0.00|N|"
      Text            =   "Impo"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recibos de Retenciones"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Reimprimir Recibo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5040
      Top             =   5400
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLiqRetencion.frx":0010
      Height          =   4125
      Left            =   240
      TabIndex        =   17
      Top             =   720
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   7276
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin VB.Label Label2 
      Caption         =   "IMPORTE TOTAL: "
      Height          =   195
      Left            =   7560
      TabIndex        =   19
      Top             =   5040
      Width           =   1395
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnRecibos 
         Caption         =   "&Recibos de Retenciones"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnReimprimir 
         Caption         =   "Reimprimir Recibo"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmLiqRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmSoc As frmGesSocios  'Form Mantenimiento socios
Attribute frmSoc.VB_VarHelpID = -1

Dim PrimeraVez As Boolean


Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Dim kCampo As Integer

Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Dim CadenaConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Private HaDevueltoDatos As Boolean

' Reimpresion de recibos
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


Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim NumReg As Long
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 1 'BUSQUEDA
        HacerBusqueda
    Case 3 'INSERTAR
        If DatosOk Then
            If InsertarDesdeForm(Me) Then
                CargaGrid True
                BotonAnyadir
            End If
        End If
    Case 4 'MODIFICAR
        If DatosOk And BLOQUEADesdeFormulario(Me) Then
             If ModificaDesdeFormulario(Me, 3) Then
                 TerminaBloquear
                 NumReg = Data1.Recordset.AbsolutePosition
                 PonerModo 2
                 CancelaADODC Me.Data1
                 CargaGrid True
                 LLamaLineas 10
                 SituarDataPosicion Data1, NumReg, Indicador
             End If
             lblIndicador.Caption = Indicador
             DataGrid1.SetFocus
         End If
    End Select
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Cod socio
            Set frmSoc = New frmGesSocios
            frmSoc.DatosADevolverBusqueda = "0"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
        
        Case 1 ' fecha de factura
            Screen.MousePointer = vbHourglass
            Set frmF = New frmCal
            frmF.Fecha = Now
            If txtAux(Index).Text <> "" Then frmF.Fecha = CDate(txtAux(Index).Text)
            Screen.MousePointer = vbDefault
            frmF.Show vbModal
            Set frmF = Nothing
    End Select
    PonerFoco txtAux(Index)
End Sub


Private Sub cmdCancelar_Click()
    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
            LLamaLineas 10
            
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            
            If Not Data1.Recordset.EOF Then
                PonerModo 2
                Data1.Recordset.MoveFirst
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Else 'No hay Registros en la Tabla
                PonerModo 0
            End If
            
            
            LLamaLineas 10
            
        Case 4  'MODIFICAR
            TerminaBloquear
            DeseleccionaGrid Me.DataGrid1
'            CargaTxtAux False, False
            PonerModo 2
            LLamaLineas 10
            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
    End Select
    
    CalcularTotales Data1.RecordSource


ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Not Data1.Recordset.EOF And (Modo <> 3 And Modo <> 4) Then
        
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        
        PonerModoOpcionesMenu
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    PrimeraVez = True
    
    'ICONOS de laLa toolbar
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(8).Image = 31 'Recibos de Retenciones
        
        .Buttons(10).Image = 40 'Reimresion de Recibos
        .Buttons(11).Image = 16 'Imprimir
        
        .Buttons(13).Image = 15 'Salir
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CargaCombo
    
    PonerModo 0
    
    NombreTabla = "sreten" 'Tabla de retenciones
    Ordenacion = " ORDER BY fecfactu desc, codsocio, numfactu "
    
    CadenaConsulta = "Select sreten.codsocio, sclien.nomclien, sreten.numeruve, sreten.fecfactu, "
    CadenaConsulta = CadenaConsulta & " sreten.numfactu, sreten.tiporeten,"
    CadenaConsulta = CadenaConsulta & " CASE sreten.tiporeten WHEN 0 THEN ""Liquidación"" WHEN 1 THEN ""Traspaso""  WHEN 2 THEN ""Rectificativa""  END, "
    CadenaConsulta = CadenaConsulta & " impreten "
    CadenaConsulta = CadenaConsulta & " from " & NombreTabla & " INNER JOIN sclien ON sreten.codsocio = sclien.codclien "
    CadenaConsulta = CadenaConsulta & " WHERE sreten.codsocio is null " 'No recupera datos
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    CargaGrid (Modo = 2)
    Screen.MousePointer = vbDefault
    PrimeraVez = False
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim SQL As String
On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, False

    CargaGrid2

    CalcularTotales SQL
    
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


Private Sub CargaGrid2()
Dim tots As String
On Error GoTo ECarga2

    '"SELECT codprove, " & NombreTabla & ".codfamia, nomfamia, " & NombreTabla & ".codmarca, nommarca, fechadto, dtoline1, dtoline2 "
    tots = "S|txtAux(0)|T|Socio|850|;S|cmdAux(0)|B||0|;S|Text2(0)|T|Nombre|3500|;S|txtAux(4)|T|V-Socio|800|;"
    tots = tots & "S|txtAux(3)|T|Fecha|1150|;S|cmdAux(1)|B||0|;S|txtAux(2)|T|Factura|1200|;"
    tots = tots & "N||||0|;S|Combo1(0)|C|Tipo|1220|;S|txtAux(1)|T|Importe|1350|;N||||0|;"
    
    arregla tots, DataGrid1, Me

    'dtos alineados a la dcha
    DataGrid1.Columns(6).Alignment = dbgRight
    DataGrid1.Columns(7).Alignment = dbgRight
    
ECarga2:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Estamos en Cabecera
            'Recupera todo el registro de Tarifas de Precios
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(txtAux(0), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(txtAux(1), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
Dim indice As Byte
    indice = 3
    txtAux(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento socios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    BotonImprimir
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnRecibos_Click()
    BotonRecibos
End Sub

Private Sub mnReimprimir_Click()
    BotonReimprimir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            mnBuscar_Click
        Case 2 'Ver Todos
            mnVerTodos_Click
        Case 5 'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7 'Eliminar
            mnEliminar_Click
        Case 8 'Recibos de retenciones
            mnRecibos_Click
        Case 10 ' Reimprimir recibo de retenciones
            mnReimprimir_Click
        Case 11 'Imprimir
            BotonImprimir
        Case 13  'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
      
    Select Case Kmodo
        Case 1 'Modo Buscar
            PonerFoco txtAux(0)
            txtAux(0).BackColor = vbYellow
        Case 2    'Preparamos para que pueda Modificar
            Me.cmdRegresar.visible = False
    End Select
           
     'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.visible = b
'        If Modo = 1 Then Me.lblIndicador.Caption = "BUSQUEDA"
'    Else
'        cmdRegresar.visible = False
'    End If
                 
    
    '-----------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    Me.DataGrid1.Enabled = (Modo <> 3 And Modo <> 4)
    BloquearClavesP (Modo = 4) ' si modificar


    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    
    b = (Modo = 2)
    'Insertar
    Toolbar1.Buttons(5).Enabled = (b Or (Modo = 0))
    Me.mnNuevo.Enabled = (b Or (Modo = 0))
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    Toolbar1.Buttons(10).Enabled = b 'And (Data1.Recordset!tiporeten = 1)
    Me.mnReimprimir.Enabled = b 'And (Data1.Recordset!tiporeten = 1)
    If b Then
        Toolbar1.Buttons(10).Enabled = Data1.Recordset!tiporeten
        Me.mnReimprimir.Enabled = Data1.Recordset!tiporeten
    End If


    b = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1(0).ListIndex = -1
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "Select sreten.codsocio, sclien.nomclien, sreten.numeruve, sreten.fecfactu, "
    SQL = SQL & " sreten.numfactu, sreten.tiporeten,"
    SQL = SQL & " CASE sreten.tiporeten WHEN 0 THEN ""Liquidación"" WHEN 1 THEN ""Traspaso"" WHEN 2 THEN ""Rectificativa""   END, "
    SQL = SQL & " impreten, hastafec "
    SQL = SQL & " from " & NombreTabla & " INNER JOIN sclien ON sreten.codsocio = sclien.codclien "

    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then SQL = SQL & CadenaBusqueda
    Else
        SQL = SQL & " WHERE codsocio = -1"
    End If
    SQL = SQL & Ordenacion
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    EsBusqueda = True
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
'        CargaTxtAux True, True
        anc = ObtenerAlto(Me.DataGrid1)
        LLamaLineas anc
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco txtAux(0)
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            txtAux(kCampo).Text = ""
            txtAux(kCampo).BackColor = vbYellow
            PonerFoco txtAux(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    EsBusqueda = False
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
'    CargaGrid False
    
'    If chkVistaPrevia.Value = 1 Then
'        MandaBusquedaPrevia ""
'    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
'    End If
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
       
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
    
'    CargaTxtAux True, True
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc
    
    Combo1(0).ListIndex = 1
    txtAux(2).Text = 0
    txtAux(3).Text = Format(Now, "dd/mm/yyyy")
    
    
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
Dim i As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
'    CargaTxtAux True, False
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc
    
    
    '---- poner valores grabados
    'codsocio
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "N")
    FormateaCampo txtAux(i)
    
    'nomsocio
    Text2(0).Text = DBLet(DataGrid1.Columns(1).Value, "T")
    
    'uve
    txtAux(4).Text = DBLet(DataGrid1.Columns(2).Value, "N")
    FormateaCampo txtAux(4)
    
    'fecha
    txtAux(3).Text = DBLet(DataGrid1.Columns(3).Value, "F")
    FormateaCampo txtAux(3)
    
    'factura
    txtAux(2).Text = DBLet(DataGrid1.Columns(4).Value, "N")
    FormateaCampo txtAux(2)
    
    ' ***** canviar-ho pel nom del camp del combo *********
    i = Me.Data1.Recordset!tiporeten
    ' *****************************************************
    PosicionarCombo Me.Combo1(0), i
    
    'Importe
    txtAux(1).Text = DBLet(DataGrid1.Columns(7).Value, "N")
    FormateaCampo txtAux(1)
    
    
    
    '-----
    If BLOQUEADesdeFormulario(Me) Then
        PonerFoco txtAux(2)
    Else
        cmdCancelar_Click
    End If
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        'Ciertas comprobaciones
        If Data1.Recordset.EOF Then Exit Function
        
        SQL = "¿Seguro que desea eliminar la Retención para?" & vbCrLf
        SQL = SQL & vbCrLf & "Socio: " & Format(Data1.Recordset.Fields(0).Value, "000000") & " - " & Data1.Recordset!nomclien
        SQL = SQL & vbCrLf & "Fecha: " & Data1.Recordset.Fields(3).Value
        SQL = SQL & vbCrLf & "Factura : " & Format(Data1.Recordset.Fields(4).Value, "0000000")
        
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            'Hay que eliminar
            NumRegElim = Me.Data1.Recordset.AbsolutePosition
            SQL = "Delete from " & NombreTabla & " where codsocio=" & Val(Data1.Recordset!codSocio)
            SQL = SQL & " and fecfactu=" & DBSet(Data1.Recordset!FecFactu, "F") & " and numfactu=" & Val(Data1.Recordset!NumFactu)
            conn.Execute SQL
            CancelaADODC Me.Data1
            CargaGrid True
            CancelaADODC Me.Data1
            SituarDataPosicion Me.Data1, NumRegElim, SQL
        End If
        
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Descuento", Err.Description
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String

    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        SQL = "select count(*) from sreten where codsocio = " & txtAux(0).Text & " and fecfactu = "
        SQL = SQL & DBSet(txtAux(3).Text, "F") & " and numfactu = " & DBSet(txtAux(2).Text, "N")
        If TotalRegistros(SQL) > 0 Then
            MsgBox "Ya existe la factura para este socio en esta fecha.", vbExclamation
            PonerFoco txtAux(0)
            b = False
            Exit Function
        End If
    End If
    DatosOk = True
    
End Function



Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & Ordenacion
        CadenaBusqueda = " WHERE " & cadB
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        PonerModo Modo
        CargaGrid False
        MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        PonerCampos
    End If
    LLamaLineas 10
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub



Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub

Private Sub BotonImprimir()
    frmLiqListReten.OpcionListado = 0
    frmLiqListReten.Show vbModal
End Sub

Private Sub BotonRecibos()
    frmLiqListReten.OpcionListado = 1
    frmLiqListReten.Show vbModal
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Sub BotonReimprimir()
Dim SQL As String
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    InicializarVbles
    
    If vParamAplic.Cooperativa = 0 Then
    
        cadNombreRPT = "rRecRetenciones.rpt"
        cadTitulo = "Reimpresion Recibos Retenciones"
        
        ConSubInforme = False
        
        SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
        conn.Execute SQL
        
        SQL = "insert into tmpinformes (codusu, codigo1, importe1, importe2, fecha1) values ("
        SQL = SQL & vUsu.Codigo & "," & DBSet(Me.Data1.Recordset!codSocio, "N") & "," & DBSet(Data1.Recordset!NumerUve, "N") & ","
        SQL = SQL & DBSet(Data1.Recordset!impreten * (-1), "N") & "," & DBSet(Data1.Recordset!hastafec, "F") & ")"
        
        conn.Execute SQL
        
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        
        LlamarImprimir False
    Else
    
        
        indRPT = 12 'Facturas Clientes
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, False, pPdfRpt) Then Exit Sub
    
    
        cadFormula = "{scafac.codtipom} = 'FAV' and {scafac.numfactu}= " & Me.Data1.Recordset!NumFactu & " and "
        cadFormula = cadFormula & "{scafac.fecfactu}= Date(" & Year(DBLet(Data1.Recordset!FecFactu, "F")) & "," & Month(DBLet(Data1.Recordset!FecFactu, "F")) & "," & Day(DBLet(Data1.Recordset!FecFactu, "F")) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
    
        devuelve = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", "FAV", "T")
     
     
         With frmImprimir
                'Nuevo. Febrero 2010
                .outClaveNombreArchiv = devuelve & Format(Me.Data1.Recordset!NumFactu, "000")
                .outCodigoCliProv = Me.Data1.Recordset!codSocio
                .outTipoDocumento = 100
                
                .FormulaSeleccion = cadFormula
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                .NombrePDF = pPdfRpt
                .NombreRPT = pPdfRpt
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = 53 'OpcionListado
                .Titulo = ""
                .Show vbModal
        End With
    
    
    
    End If

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


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub

Private Sub BloquearClavesP(bol As Boolean)
'Si BloquearClavesPrimarias=true deshablilita los textbox de codigos y lo pone amarillo
'y habilita el resto de campos para introducir nuevos valores
'Si BloquearClavesPrimarias=false habilita los textbox de codigos para introducir
Dim i As Byte

    For i = 0 To 0 'Codigo socio
        BloquearTxt txtAux(i), bol
        Me.cmdAux(i).Enabled = Not bol
    Next i
    BloquearTxt txtAux(2), bol
    BloquearTxt txtAux(3), bol
    
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim campo As String, Tabla As String
Dim campo2 As String

  
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'cod. socio
            If PonerFormatoEntero(txtAux(Index)) Then
                campo2 = "numeruve"
                Text2(Index).Text = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", txtAux(Index).Text, "N", campo2)
                If Modo = 3 Then
                    txtAux(4).Text = campo2
                    If campo2 = "" Then txtAux(4).Text = "0"
                    PonerFoco txtAux(3)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 3 'Fecha
            PonerFormatoFecha txtAux(Index)
            
        Case 4, 2 ' uve
            PonerFormatoEntero txtAux(Index)
            
        Case 1 'Importe
            PonerFormatoDecimal txtAux(Index), 6
    
    End Select
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Integer
Dim b As Boolean

        DeseleccionaGrid Me.DataGrid1
        'PonerModo xModo + 1

        b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

        For jj = 0 To txtAux.Count - 1
            txtAux(jj).Height = DataGrid1.RowHeight
            txtAux(jj).Top = alto
            txtAux(jj).visible = b
        Next jj
        
        For jj = 0 To Text2.Count - 1
            Text2(jj).Height = Me.DataGrid1.RowHeight
            Text2(jj).Top = alto
            Text2(jj).visible = b
        Next jj
        
        For jj = 0 To Me.cmdAux.Count - 1
            Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
            Me.cmdAux(jj).Top = alto
            Me.cmdAux(jj).visible = b
        Next jj
        
        For jj = 0 To Combo1.Count - 1
            Combo1(jj).Top = alto
            Combo1(jj).visible = b
        Next jj
        
        
End Sub


Private Sub CargaCombo()

    On Error GoTo ErrCarga
    
    'Tipo de Calidad
    Combo1(0).Clear
    
    Combo1(0).AddItem "Liquidación"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Traspaso"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Rectificativa"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub


Private Sub CalcularTotales(CADENA As String)
Dim Importe  As Currency
Dim Compleme As Currency
Dim Penaliza As Currency

Dim RS As ADODB.Recordset
Dim SQL As String

    On Error Resume Next
    
    SQL = "select sum(impreten) importe  from (" & CADENA & ") aaaaa"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Importe = 0
    Text1.Text = ""
    
    If TotalRegistrosConsulta(CADENA) = 0 Then Exit Sub
    
    If Not RS.EOF Then
        If RS.Fields(0).Value <> 0 Then Importe = DBLet(RS.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
    
        Text1.Text = Format(Importe, "###,###,##0.00")
    End If
    RS.Close
    Set RS = Nothing

    
    DoEvents
    
End Sub


