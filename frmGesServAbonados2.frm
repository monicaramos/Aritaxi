VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGesServAbonados2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicios Abonados"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   13185
   ClipControls    =   0   'False
   Icon            =   "frmGesServAbonados2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   10350
      MaxLength       =   6
      TabIndex        =   7
      Tag             =   "Numero V|N|N|||sfactclitr_serv|numeruve|000000|N|"
      Text            =   "Nro V"
      Top             =   3570
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   9240
      MaxLength       =   10
      TabIndex        =   6
      Tag             =   "NroServicios|N|N|||sfactclitr_serv|nroservicio|0000000|N|"
      Text            =   "NroServici"
      Top             =   3570
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   7020
      MaxLength       =   100
      TabIndex        =   5
      Tag             =   "Destino|T|N|||sfactclitr_serv|destino|||"
      Text            =   "Destino"
      Top             =   3570
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   4350
      MaxLength       =   100
      TabIndex        =   4
      Tag             =   "Origen|T|N|||sfactclitr_serv|origen|||"
      Text            =   "Origen"
      Top             =   3570
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   2430
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||sfactclitr_serv|fecha|dd/mm/yyyy||"
      Text            =   "fecha"
      Top             =   3570
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4980
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   5400
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   3450
      TabIndex        =   15
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
      Tag             =   "Cod.Cliente|N|N|0|999999|sfactclitr_serv|codclien|000000|S|"
      Text            =   "client"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   3660
      MaxLength       =   8
      TabIndex        =   3
      Tag             =   "Hora|H|N|||sfactclitr_serv|hora|hh:mm:ss||"
      Text            =   "hora"
      Top             =   3570
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||sfactclitr_serv|fecfactu|dd/mm/yyyy|S|"
      Text            =   "fecfactu"
      Top             =   3570
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10860
      TabIndex        =   9
      Top             =   5400
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12045
      TabIndex        =   10
      Top             =   5400
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12030
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   11460
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "Importe|N|N|||sfactclitr|importe|###,##0.00|N|"
      Text            =   "Importe"
      Top             =   3570
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   14
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
      Bindings        =   "frmGesServAbonados2.frx":000C
      Height          =   4110
      Left            =   240
      TabIndex        =   18
      Top             =   720
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   7250
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
      Left            =   9870
      TabIndex        =   20
      Top             =   5010
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
      TabIndex        =   13
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
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
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
Attribute VB_Name = "frmGesServAbonados2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Public Cliente As String
Public FecFactura As String

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmCli As frmFacClientes  'Form Mantenimiento clientes
Attribute frmCli.VB_VarHelpID = -1

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

' utilizado para buscar por checks
Private BuscaChekc As String


Private HaDevueltoDatos As Boolean
Dim cadB1 As String


Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub

Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



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


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Not Data1.Recordset.EOF And (Modo <> 3 And Modo <> 4) Then
        
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmppal.Icon
    
    PrimeraVez = True
    
    'ICONOS de laLa toolbar
    With Toolbar1
        .ImageList = frmppal.ImgListComun1
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(11).Image = 15 'Salir
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
 

    PonerModo 0
    
    NombreTabla = "sfactclitr_serv" 'Tabla de facturas asociados
    Ordenacion = " ORDER BY fecha, hora "
    
    CadenaConsulta = "Select sfactclitr_serv.codclien, sfactclitr_serv.fecfactu, sfactclitr_serv.fecha, sfactclitr_serv.hora, "
    CadenaConsulta = CadenaConsulta & " sfactclitr_serv.origen, sfactclitr_serv.destino, sfactclitr_serv.importe"
    CadenaConsulta = CadenaConsulta & " from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " WHERE sfactclitr_serv.codclien = " & DBSet(Cliente, "N")
    CadenaConsulta = CadenaConsulta & " and sfactclitr_serv.fecfactu = " & DBSet(FecFactura, "F")
    
    

    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    CargaGrid True ' (Modo = 2)
    Screen.MousePointer = vbDefault
    PrimeraVez = False

End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim Sql As String
On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, Sql, False

    CargaGrid2

    CalcularTotales Sql
    
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


Private Sub CargaGrid2()
Dim tots As String
On Error GoTo ECarga2

    '"SELECT codprove, " & NombreTabla & ".codfamia, nomfamia, " & NombreTabla & ".codmarca, nommarca, fechadto, dtoline1, dtoline2 "
    tots = "N|txtAux(0)|T|Cliente|850|;"
    tots = tots & "N|txtAux(1)|T|FecFac|1150|;S|txtAux(2)|T|Fecha|1150|;S|cmdAux(1)|B||0|;S|txtAux(3)|T|Hora|850|;S|txtAux(4)|T|Origen|3450|;"
    tots = tots & "S|txtAux(5)|T|Destino|3450|;S|txtAux(7)|T|NroServicio|1200|;S|txtAux(8)|T|Nro V|900|;S|txtAux(6)|T|Importe|1250|;"
    
    arregla tots, DataGrid1, Me

    'dtos alineados a la dcha
'    DataGrid1.Columns(6).Alignment = dbgRight
'    DataGrid1.Columns(7).Alignment = dbgRight
    
ECarga2:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Estamos en Cabecera
            'Recupera todo el registro de Tarifas de Precios
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            CadB = ""
            Aux = ValorDevueltoFormGrid(txtAux(0), CadenaDevuelta, 1)
            CadB = Aux
            Aux = ValorDevueltoFormGrid(txtAux(1), CadenaDevuelta, 2)
            CadB = CadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
Dim indice As Byte
    indice = 2
    txtAux(indice).Text = Format(vFecha, "dd/mm/yyyy")
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
        Case 10 'Imprimir
            BotonImprimir
        Case 11  'Salir
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
Dim i As Integer

    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    For i = 0 To txtAux.Count - 1
        txtAux(i).BackColor = vbWhite
    Next i
      
    Select Case Kmodo
        Case 1 'Modo Buscar
            PonerFoco txtAux(2)
            txtAux(2).BackColor = vbLightBlue 'vbYellow
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
    
    b = False '(Modo = 2)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b '(b Or (Modo = 0))
    Me.mnNuevo.Enabled = b '(b Or (Modo = 0))
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'imprimir
    Toolbar1.Buttons(10).Enabled = b
    Me.mnImprimir.Enabled = b
    
    
    
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
Dim Sql As String
    
    Sql = "Select sfactclitr_serv.codclien, sfactclitr_serv.fecfactu, sfactclitr_serv.fecha, sfactclitr_serv.hora, "
    Sql = Sql & " sfactclitr_serv.origen, sfactclitr_serv.destino,"
    Sql = Sql & " sfactclitr_serv.nroservicio, sfactclitr_serv.numeruve, sfactclitr_serv.importe "
    Sql = Sql & " from " & NombreTabla
    Sql = Sql & " where  sfactclitr_serv.codclien = " & DBSet(Cliente, "N")
    Sql = Sql & " and  sfactclitr_serv.fecfactu = " & DBSet(FecFactura, "F")
    

    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then Sql = Sql & CadenaBusqueda
    Else
        Sql = Sql & " and sfactclitr_serv.codclien = -1"
    End If
    Sql = Sql & Ordenacion
    MontaSQLCarga = Sql
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
        PonerFoco txtAux(2)
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            txtAux(kCampo).Text = ""
            txtAux(kCampo).BackColor = vbLightBlue 'vbYellow
            PonerFoco txtAux(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    EsBusqueda = False
    LimpiarCampos
    
    CadenaConsulta = "Select * from " & NombreTabla & " where codclien = " & DBSet(Cliente, "N") & " and fecfactu = " & DBSet(FecFactura, "F") & Ordenacion
    PonerCadenaBusqueda
    
    cadB1 = ""

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
    
    
    txtAux(0).Text = Cliente
    txtAux(1).Text = FecFactura
    txtAux(2).Text = 0
    txtAux(3).Text = Format(Now, "dd/mm/yyyy")

    
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
Dim i As Integer
Dim anc As Single

    If CInt(Data1.Recordset!facturado) = 1 Then
        MsgBox "La línea está facturada, no se permite ni modificar ni eliminar.", vbExclamation
        Exit Sub
    End If

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
    'codclien
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "N")
    FormateaCampo txtAux(i)
    
    'nomclien
'    Text2(0).Text = DBLet(DataGrid1.Columns(1).Value, "T")
    
    'fecha
    txtAux(3).Text = DBLet(DataGrid1.Columns(2).Value, "F")
    FormateaCampo txtAux(3)
    
    'concepto
    txtAux(4).Text = DBLet(DataGrid1.Columns(3).Value, "N")
    
    'servicios
    txtAux(2).Text = DBLet(DataGrid1.Columns(4).Value, "N")
    FormateaCampo txtAux(2)
    
    
    'Importe
    txtAux(1).Text = DBLet(DataGrid1.Columns(5).Value, "N")
    FormateaCampo txtAux(1)
    
    
    
    '-----
    If BLOQUEADesdeFormulario(Me) Then
        PonerFoco txtAux(4)
    Else
        cmdCancelar_Click
    End If
End Sub


Private Function BotonEliminar() As Boolean
Dim Sql As String
On Error GoTo FinEliminar
        
        'Ciertas comprobaciones
        If Data1.Recordset.EOF Then Exit Function
        
        If CInt(Data1.Recordset.Fields(6).Value) = 1 Then
            MsgBox "La línea está facturada, no se permite ni modificar ni eliminar.", vbExclamation
            Exit Function
        End If
        
        Sql = "¿Seguro que desea eliminar el registro?" & vbCrLf
        Sql = Sql & vbCrLf & "Cliente: " & Format(Data1.Recordset.Fields(0).Value, "000000") & " - " '& Text2(0).Text
        Sql = Sql & vbCrLf & "Fecha: " & Data1.Recordset.Fields(2).Value
        
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
            'Hay que eliminar
            NumRegElim = Me.Data1.Recordset.AbsolutePosition
            Sql = "Delete from " & NombreTabla & " where codclien=" & Val(Data1.Recordset!CodClien)
            Sql = Sql & " and fecfactu=" & DBSet(Data1.Recordset!FecFactu, "F")
            conn.Execute Sql
            CancelaADODC Me.Data1
            CargaGrid True
            CancelaADODC Me.Data1
            SituarDataPosicion Me.Data1, NumRegElim, Sql
        End If
        
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Descuento", Err.Description
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String

    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        Sql = "select count(*) from sfactclitr where codsocio = " & txtAux(0).Text & " and fecfactu = "
        Sql = Sql & DBSet(txtAux(3).Text, "F") & " and numfactu = " & DBSet(txtAux(2).Text, "N")
        If TotalRegistros(Sql) > 0 Then
            MsgBox "Ya existe la factura para este socio en esta fecha.", vbExclamation
            PonerFoco txtAux(0)
            b = False
            Exit Function
        End If
    End If
    DatosOk = True
    
End Function



Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False, BuscaChekc)
    cadB1 = ObtenerBusqueda(Me, True, BuscaChekc)
    
    If CadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " where codclien = " & DBSet(Cliente, "N") & " and fecfactu = " & DBSet(FecFactura, "F") & " and  " & CadB & Ordenacion
        CadenaBusqueda = " and " & CadB
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
    printNou
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

    For i = 1 To 1 'Codigo socio
        BloquearTxt txtAux(i), bol
        Me.cmdAux(i).Enabled = Not bol
    Next i
    ' fecha bloqueada
    Me.cmdAux(1).Enabled = Not bol
    BloquearTxt txtAux(3), bol
    
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim campo As String, Tabla As String
Dim campo2 As String

  
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 2 ' fecha
            PonerFormatoFecha txtAux(Index)
        
        Case 3 'hora
            PonerFormatoHora txtAux(Index)
            
        Case 6 'Importe
            PonerFormatoDecimal txtAux(Index), 6
    
    End Select
End Sub


Private Sub LLamaLineas(alto As Single)
Dim jj As Integer
Dim b As Boolean

        DeseleccionaGrid Me.DataGrid1
        'PonerModo xModo + 1

        b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar
        
        txtAux(0).Text = Cliente
        txtAux(1).Text = FecFactura
        

        For jj = 2 To txtAux.Count - 1
            txtAux(jj).Height = DataGrid1.RowHeight
            txtAux(jj).top = alto
            txtAux(jj).visible = b
        Next jj
        
        cmdAux(1).Height = DataGrid1.RowHeight
        cmdAux(1).top = alto
        cmdAux(1).visible = b
        
End Sub



Private Sub CalcularTotales(CADENA As String)
Dim Importe  As Currency
Dim Compleme As Currency
Dim Penaliza As Currency

Dim RS As ADODB.Recordset
Dim Sql As String

    On Error Resume Next
    
    Sql = "select sum(importe) importe  from (" & CADENA & ") aaaaa"
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "sfactclitr"
        .Informe2 = "rGesServAbonados.rpt"
        If cadB1 <> "" Then
            .cadRegSelec = cadB1 'SQL2SF(cadB1)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(Data1, Me)
        .cadTodosReg = ""
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|pOrden={tarjbanc.nomtarje}|"
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|"
        .NumeroParametros2 = 1
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub


