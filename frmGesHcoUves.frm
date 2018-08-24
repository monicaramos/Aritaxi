VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGesHcoUves 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Uves"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ClipControls    =   0   'False
   Icon            =   "frmGesHcoUves.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5100
      TabIndex        =   17
      Top             =   300
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   210
      TabIndex        =   15
      Top             =   120
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   16
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   360
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Cod. Socio|N|S|0|999999|shiuve|codsocio|000000|S|"
      Text            =   "cod.so"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   6990
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Fecha Baja|F|S|||shiuve|fechabaja|dd/mm/yyyy|N|"
      Text            =   "fecha baja"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   210
      TabIndex        =   13
      Top             =   5580
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Left            =   240
         TabIndex        =   14
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
      TabIndex        =   12
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
      Left            =   6750
      TabIndex        =   11
      ToolTipText     =   "Buscar fecha"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   2
      Left            =   7950
      TabIndex        =   10
      ToolTipText     =   "Buscar fecha"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   6150
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha Alta|F|N|||shiuve|fechaalta|dd/mm/yyyy|S|"
      Text            =   "Fecha Alta"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
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
      Left            =   6960
      TabIndex        =   4
      Top             =   5640
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
      Left            =   8235
      TabIndex        =   5
      Top             =   5640
      Width           =   1135
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   8220
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   5310
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "NumUve|N|S|0|999999|shiuve|numeruve|000000|S|"
      Text            =   "uve"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1230
      Top             =   4920
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
      Bindings        =   "frmGesHcoUves.frx":10CA
      Height          =   4545
      Left            =   210
      TabIndex        =   7
      Top             =   975
      Width           =   9155
      _ExtentX        =   16140
      _ExtentY        =   8017
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
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
      TabIndex        =   8
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
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
         Caption         =   "Baja de &V"
         HelpContextID   =   2
         Shortcut        =   ^V
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
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
Attribute VB_Name = "frmGesHcoUves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Private Const IdPrograma = 405


Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmC As frmGesSocios 'Form Mantenimiento socios
Attribute frmC.VB_VarHelpID = -1

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Dim kCampo As Integer


Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Dim CadenaConsulta As String
Dim WhereConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Dim PrimeraVez As Boolean
Dim indCodigo As Integer

Dim cadB1 As String

Private HaDevueltoDatos As Boolean

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
            If InsertaRegistro() Then ' antes InsertarDesdeForm(Me)
                CargaGrid True
                BotonAnyadir
            End If
        End If
    Case 4 'MODIFICAR
           If DatosOk And BLOQUEADesdeFormulario(Me) Then
                'Marzo 2010
                'antes cuando habia clvae ppal sin valores nul
                'If ModificaDesdeFormulario(Me, 3) Then
                If ModificaRegistro() Then
                    TerminaBloquear
                    NumReg = Data1.Recordset.AbsolutePosition
                    PonerModo 2
                    CancelaADODC Me.Data1
                    CargaGrid True
'                    CargaTxtAux False, False
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
        Case 0 'codigo de socio
            Set frmC = New frmGesSocios
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
        Case 1, 2 'Fecha de alta y fecha de baja
            indCodigo = Index + 1
            Screen.MousePointer = vbHourglass
            Set frmF = New frmCal
            frmF.Fecha = Now
            If txtAux(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtAux(indCodigo).Text)
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
            
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                PonerModo 2
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Else 'No hay Registros en la Tabla
                PonerModo 0
            End If
            LLamaLineas 10
            
        Case 4  'Modificar
            TerminaBloquear
            DeseleccionaGrid Me.DataGrid1
            PonerModo 2
            LLamaLineas 10
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.AbsolutePosition > 0 Then
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        End If
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon

    'ICONOS de La toolbar
'    With Toolbar1
'        .ImageList = frmPpal.imgListComun
'        'ASignamos botones
'        .Buttons(1).Image = 1   'Buscar
'        .Buttons(2).Image = 2 'Ver Todos
'        .Buttons(5).Image = 3 'Añadir
'        .Buttons(6).Image = 4 'Modificar
'        .Buttons(7).Image = 5 'Eliminar
'        .Buttons(10).Image = 16 'Imprimir
'        .Buttons(11).Image = 15 'Salir
'    End With
    
    
    
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        'el 1 es separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(8).Image = 16  'imprimir
    End With
    
'    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    PrimeraVez = True
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    PonerModo 0
    NombreTabla = "shiuve" 'Historico de Uves
    Ordenacion = " ORDER BY shiuve.codsocio, fechaalta"
    WhereConsulta = " WHERE shiuve.codsocio = -1"
    
    CargaGrid (Modo = 2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim Sql As String
On Error GoTo ECarga

    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, Sql, PrimeraVez
    
    CargaGrid2
    DataGrid1.Enabled = (Modo = 2)
    Me.DataGrid1.ScrollBars = dbgAutomatic
    PrimeraVez = False
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub CargaGrid2()
Dim tots As String
On Error GoTo ECarga2

    'SQL = "SELECT codsocio, nomsocio, numeruve, fechaalta, fechabaja
    tots = "S|txtAux(0)|T|Socio|1050|;S|cmdAux(0)|B||0|;S|txtAux2|T|Nombre|3800|;S|txtAux(1)|T|Vehiculo|1000|;S|txtAux(2)|T|Fecha Alta|1350|;S|cmdAux(1)|B||0|;"
    tots = tots & "S|txtAux(3)|T|Fecha Baja|1350|;S|cmdAux(2)|B||0|;"
    
    arregla tots, DataGrid1, Me, 350

ECarga2:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


Private Sub LLamaLineas(alto As Single)
Dim jj As Integer
Dim b As Boolean

        DeseleccionaGrid Me.DataGrid1
        b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

        For jj = 0 To txtAux.Count - 1
            txtAux(jj).Height = DataGrid1.RowHeight
            txtAux(jj).top = alto
            txtAux(jj).visible = b
        Next jj
        txtAux2.Height = Me.DataGrid1.RowHeight
        txtAux2.top = alto
        txtAux2.visible = b
        
        For jj = 0 To Me.cmdAux.Count - 1
            Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
            Me.cmdAux(jj).top = alto
            Me.cmdAux(jj).visible = b
        Next jj
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
    txtAux(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Socios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtAux2.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    printNou
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
        Case 1 'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3 'Eliminar
            mnEliminar_Click
        Case 5 'Busqueda
            mnBuscar_Click
        Case 6 'Ver Todos
            mnVerTodos_Click
        Case 8 'Imprimir
            mnImprimir_Click
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim I As Integer

    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).BackColor = vbWhite
    Next I
    
    
    Select Case Kmodo
        Case 1 'Modo Buscar
            PonerFoco txtAux(0)
        Case 2    'Preparamos para que pueda Modificar
            Me.cmdRegresar.visible = False
    End Select
                            
    
    BloquearClavesP (Modo = 4) ' si modificar
           
    '-----------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    Me.DataGrid1.Enabled = (Modo = 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
    PonerModoUsuarioGnral Modo, "aritaxi"
    
End Sub

Private Sub PonerModoUsuarioGnral(Modo As Byte, Aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(Aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = Toolbar1.Buttons(1).Enabled And DBLet(Rs!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = Toolbar1.Buttons(2).Enabled And DBLet(Rs!Modificar, "N")
        Toolbar1.Buttons(3).Enabled = Toolbar1.Buttons(3).Enabled And DBLet(Rs!creareliminar, "N")
        
        Toolbar1.Buttons(5).Enabled = Toolbar1.Buttons(5).Enabled And DBLet(Rs!ver, "N")
        Toolbar1.Buttons(6).Enabled = Toolbar1.Buttons(6).Enabled And DBLet(Rs!ver, "N")
        
        Toolbar1.Buttons(8).Enabled = Toolbar1.Buttons(8).Enabled And DBLet(Rs!Imprimir, "N")
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub
Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    'Modo 2. Hay datos y estamos visualizandolos
    b = (Modo = 2)
    'Insertar
    Toolbar1.Buttons(1).Enabled = (b Or (Modo = 0))
    Me.mnNuevo.Enabled = (b Or (Modo = 0))
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'imprimir
    Toolbar1.Buttons(8).Enabled = b
    Me.mnImprimir.Enabled = b
    
    
    b = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b
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
Dim Tabla As String
    
    Tabla = "shiuve"
    Sql = "SELECT shiuve.codsocio, sclien.nomclien, shiuve.numeruve, shiuve.fechaalta, shiuve.fechabaja "
    Sql = Sql & " FROM " & Tabla & " INNER JOIN sclien ON " & Tabla & ".codsocio ="
    Sql = Sql & " sclien.codclien"

    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then
            Sql = Sql & CadenaBusqueda
        ElseIf CadenaConsulta = "" Then
            If CadenaBusqueda <> "" Then
                CadenaBusqueda = CadenaBusqueda & " OR (" & MontaWHERE(True, True) & ")"
            Else
                'CadenaBusqueda = " WHERE (codclien=" & txtAux(0).Text & " and codfamia=" & txtAux(1).Text & " and codmarca=" & txtAux(2).Text & ")"
                CadenaBusqueda = " WHERE (" & MontaWHERE(True, True) & ")"
            End If
            Sql = Sql & CadenaBusqueda
        End If
    Else
        Sql = Sql & " WHERE shiuve.codsocio = -1"
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
        anc = ObtenerAlto(Me.DataGrid1)
        LLamaLineas anc
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco txtAux(0)
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            txtAux(kCampo).Text = ""
            PonerFoco txtAux(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    EsBusqueda = False
    LimpiarCampos
    
'    If chkVistaPrevia.Value = 1 Then
'        MandaBusquedaPrevia ""
'    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
'        CadenaConsulta = ""
'    End If
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    LimpiarCampos 'Vacía los TextBox
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
Dim I As Integer
Dim anc As Single
Dim Sql As String
    
    
    '[Monica]01/04/2014: comprobamos si tiene choferes dados de alta
    Sql = "select count(*) from sclien_chofer where codsocio = " & DBSet(Me.Data1.Recordset!codSocio, "N") & " and fechabaj is null"
    If TotalRegistros(Sql) <> 0 Then
        Sql = "El socio tiene choferes dados de alta. ¿ Desea continuar ?."
        If MsgBox(Sql, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
    End If
        
    
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc
    
    txtAux(0).Text = DBLet(Data1.Recordset.Fields(0).Value, "N")
    FormateaCampo txtAux(0)
    txtAux2.Text = DBLet(Data1.Recordset.Fields(1), "T")
    
    txtAux(1).Text = DBLet(Data1.Recordset.Fields(2).Value, "N")
    FormateaCampo txtAux(1)
    
    txtAux(2).Text = DBLet(Data1.Recordset.Fields(3).Value, "F")
    txtAux(3).Text = DBLet(Data1.Recordset.Fields(4).Value, "F")
    
    PonerFoco txtAux(3)
End Sub


Private Function BotonEliminar() As Boolean
Dim Sql As String
On Error GoTo FinEliminar
        
        'Ciertas comprobaciones
        If Data1.Recordset.EOF Then Exit Function
        
        Sql = "¿Seguro que desea eliminar el Registro para?" & vbCrLf
        Sql = Sql & vbCrLf & "Socio: " & Format(Data1.Recordset.Fields(0).Value, "000000") & " - " & Data1.Recordset.Fields(1).Value
        Sql = Sql & vbCrLf & "Uve: " & Format(Data1.Recordset.Fields(2).Value, "000000")
        Sql = Sql & vbCrLf & "Fecha Alta : " & Format(Data1.Recordset.Fields(3).Value, "dd/mm/yyyy")
        
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
            'Hay que eliminar
            NumRegElim = Me.Data1.Recordset.AbsolutePosition
            Sql = "Delete from shiuve where codsocio " & vDBSET(Data1.Recordset!codSocio, True, True, False)
            Sql = Sql & " and numeruve " & vDBSET(Data1.Recordset!NumerUve, True, True, False)
            Sql = Sql & " and fechaalta = " & DBSet(Data1.Recordset!fechaalta, "F")
            conn.Execute Sql
            CancelaADODC Me.Data1
            CargaGrid True
            CancelaADODC Me.Data1
            SituarDataTrasEliminar Data1, NumRegElim
            CargaGrid2
        End If
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Descuento", Err.Description
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim C As String
Dim C2 As String
Dim Sql As String
Dim NueDesFec As Date

    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    'Como NO hay clave primaria tengo que comprobar que NO exista un valor
    Set Rs = New ADODB.Recordset
    
    ' libera la V y activa la fecha de baja
    If b And Modo = 4 Then
        ' comprobamos que la fecha de alta es inferior a la fecha de baja
        If txtAux(3).Text = "" Then
            MsgBox "Debe introducir una fecha de baja. Revise.", vbExclamation
            PonerFoco txtAux(3)
            b = False
        Else
            If CDate(txtAux(2).Text) > CDate(txtAux(3).Text) Then
                MsgBox "La fecha de baja debe ser superior a la fecha de alta. Revise.", vbExclamation
                PonerFoco txtAux(2)
                b = False
            End If
        End If
    End If
    
    
    
    If b And Modo = 3 Then
        ' comprobamos que el socio no este dado de alta con otra uve (sclien)
        Sql = "select count(*) from sclien where codclien = " & txtAux(0).Text & " and fechabaj is null"
        If TotalRegistros(Sql) <> 0 Then
            MsgBox "Socio dado de alta con otra uve. Revise.", vbExclamation
            PonerFoco txtAux(0)
            b = False
        End If
        
        ' comprobamos que el socio no esta ya introducido en el hco con fecha de baja null (shiuve)
        If b Then
            Sql = "select count(*) from shiuve where codsocio = " & txtAux(0).Text & " and fechabaj is null"
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Socio ya introducido en el histórico. Revise.", vbExclamation
                PonerFoco txtAux(0)
                b = False
            End If
        End If
        
        ' comprobamos que la uve no este asignada a otro socio activo (ssocio)
        If b Then
            Sql = "select count(*) from sclien where numeruve = " & txtAux(1).Text & " and fechabaj is null"
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "V asignada a otro Socio activo. Revise.", vbExclamation
                PonerFoco txtAux(0)
                b = False
            End If
        End If
    End If
    
    '[Monica]11/06/2015: si vamos a darlo de baja comprobamos que no tenga servicios por facturar
    If b And Modo = 4 Then
        NueDesFec = DateAdd("yyyy", -1, CDate(txtAux(3).Text))
    
        Sql = "select count(*) from shilla where numeruve = " & DBSet(txtAux(1).Text, "N") & " and codsocio = " & DBSet(txtAux(0).Text, "N") & " and liquidadosocio = 0 and impcompr <> 0 and tipservi = 1 "
        
        '[Monica]28/10/2016: que no tengan servicios pendientes en el ultimo año. Miro la fecha de baja
        'SQL = SQL & " and fecha >= '2015-01-01'"
        Sql = Sql & " and fecha >= " & DBSet(NueDesFec, "F")
        
        '[Monica]18/06/2018: sean ademas servicios de fecha inferior a la fecha de baja
        Sql = Sql & " and fecha <= " & DBSet(txtAux(3).Text, "F")
        
        If TotalRegistros(Sql) <> 0 Then
            MsgBox "Este Socio tiene servicios por liquidar. Revise.", vbExclamation
            PonerFoco txtAux(0)
            b = False
        End If
    End If
    
    DatosOk = b

End Function




Private Function MontaWHERE(ConLosTxt As Boolean, ComprobarConFecha As Boolean) As String
Dim s As String
    
    If ConLosTxt Then
        s = " codsocio " & vDBSET(txtAux(0).Text, True, True, ConLosTxt)
        s = s & " and numeruve " & vDBSET(txtAux(1).Text, True, True, ConLosTxt)
        s = s & " and fechaalta = " & DBSet(txtAux(2).Text, "F")
    Else
        
        'Contra el DATA1
        s = " codsocio " & vDBSET(Data1.Recordset!codSocio, True, True, ConLosTxt)
        s = s & " and numeruve " & vDBSET(Data1.Recordset!NumerUve, True, True, ConLosTxt)
        s = s & " and fechaalta = " & DBSet(Data1.Recordset!fechaalta, "F")
    End If
    MontaWHERE = s
End Function


Private Function vDBSET(Valor As Variant, EsNumerico As Boolean, EsNulo As Boolean, DesdeTextos As Boolean) As Variant
Dim eNulo As Boolean
    If DesdeTextos Then
        eNulo = Valor = ""
    Else
        eNulo = IsNull(Valor)
    End If
    
    If eNulo Then
        vDBSET = " is null"
    Else
        If EsNumerico Then
            vDBSET = " = " & Val(Valor)
        Else
            vDBSET = " = '" & Format(Valor, FormatoFecha) & "'"
        End If
    End If
End Function



Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    Cad = Cad & ParaGrid(txtAux(0), 40, "Cod. Clien.")
    Cad = Cad & ParaGrid(txtAux(1), 20, "Cod. Artic")
    Tabla = NombreTabla
    Titulo = "Precios Especiales"

    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexión a BD: Aritaxi
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            If Modo = 5 Then
'                PonerFoco txtAux(0)
'            Else
                PonerFoco txtAux(kCampo)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    cadB1 = ObtenerBusqueda(Me, True)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & Ordenacion
        CadenaBusqueda = " WHERE " & CadB
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
        If EsBusqueda Then CadenaConsulta = ""
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


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
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
Dim I As Byte

    For I = 0 To 2 'Codigos
        BloquearTxt txtAux(I), bol
        Me.cmdAux(I).Enabled = Not bol
    Next I
'    Me.cmdAux(4).Enabled = Not bol
'    BloquearTxt txtAux(8), bol
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
'    If txtAux(Index).Text = "" Then Exit Sub
    If Modo = 1 Then Exit Sub
    Select Case Index
        Case 0 'cod. Cliente
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2.Text = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", txtAux(Index).Text, "N")
            Else
                txtAux2.Text = ""
            End If
        Case 1 'Nro de Uve
            PonerFormatoEntero txtAux(Index)
        Case 2 'Fecha de Alta
            PonerFormatoFecha txtAux(Index)
        Case 3 'Fecha de Baja
            PonerFormatoFecha txtAux(Index)
    End Select
End Sub

Private Function InsertaRegistro() As Boolean
Dim C2 As String
    On Error GoTo EInsertaRegistro
    InsertaRegistro = False
    
    conn.BeginTrans
    
    ' actualizamos el socio
    C2 = "UPDATE sclien SET "
    C2 = C2 & " numeruve = " & DBSet(txtAux(1).Text, "N")
    C2 = C2 & ", fechabaj = " & DBSet(txtAux(3), "F", "S")
    C2 = C2 & " WHERE codclien = " & DBSet(txtAux(0).Text, "N")
    conn.Execute C2
    
    ' Insertamos en el hco
    C2 = "insert into shiuve (codsocio,numeruve,fechaalta,fechabaja) values ("
    C2 = C2 & DBSet(txtAux(0).Text, "N") & "," & DBSet(txtAux(1).Text, "N") & ","
    C2 = C2 & DBSet(txtAux(2).Text, "F") & "," & DBSet(txtAux(3).Text, "F", "S") & ")"
    conn.Execute C2
    
    conn.CommitTrans
    InsertaRegistro = True
    Exit Function
    
EInsertaRegistro:
    conn.RollbackTrans
    MuestraError Err.Number, "Inserta Registro"
End Function




Private Function ModificaRegistro() As Boolean
Dim C2 As String
    On Error GoTo EModificaRegistro
    ModificaRegistro = False
    
    conn.BeginTrans
    
    ' actualizamos el hco
    C2 = "UPDATE shiuve SET "
    C2 = C2 & " fechabaja = " & DBSet(txtAux(3), "F", "N")
    C2 = C2 & " WHERE " & MontaWHERE(True, False)
    conn.Execute C2
    
    ' liberamos la v
    C2 = "UPDATE sclien SET "
    C2 = C2 & "numeruve = null "
    C2 = C2 & ", fechabaj = " & DBSet(txtAux(3).Text, "F")
    C2 = C2 & " where codclien = " & DBSet(txtAux(0).Text, "N")
    conn.Execute C2
    
    conn.CommitTrans
    ModificaRegistro = True
    Exit Function
EModificaRegistro:
    conn.RollbackTrans
    MuestraError Err.Number, "Modifica Registro"
End Function



Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "shiuve"
        .Informe2 = "rGesHcoUves.rpt"
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


