VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmManTraReali 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trabajos realizados"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   5475
   Icon            =   "frmManTraReali.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   9
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
      Left            =   120
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "C�digo|N|N|0|9999|smantr|codtrabajo|0000|S|"
      Text            =   "Dat"
      Top             =   5040
      Width           =   800
   End
   Begin VB.TextBox txtAux 
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
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "Descr|T|N|||smantr|nomtrabajo||N|"
      Text            =   "Dato2"
      Top             =   5040
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManTraReali.frx":000C
      Height          =   4545
      Left            =   90
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   870
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   8017
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
      Left            =   4290
      TabIndex        =   3
      Top             =   5730
      Width           =   1035
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
      Left            =   3090
      TabIndex        =   2
      Top             =   5730
      Width           =   1035
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
      Left            =   4290
      TabIndex        =   6
      Top             =   5730
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   90
      TabIndex        =   4
      Top             =   5550
      Width           =   1845
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1200
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   495
      Left            =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManTraReali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

Public Event DatoSeleccionado(CadenaSeleccion As String)

Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos

Dim FormatoCod As String 'formato del campo de codigo
Dim Modo As Byte
'-------------------------------------------------------
'Se distinguen varios MODOS
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas


Private Sub PonerModo(vModo As Byte)
Dim b As Boolean
Dim i As Integer

    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    For i = 0 To txtAux.Count - 1
        txtAux(i).BackColor = vbWhite
    Next i
    
    txtAux(0).visible = Not b
    txtAux(1).visible = Not b
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b

    'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    End If

    'Si estamos insertando o busqueda
    BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2)
    'Buscar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
   
    b = b And Not DeConsulta
    'Insertar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    Toolbar1.Buttons(8).Enabled = False
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub BotonAnyadir()
Dim anc As Single
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Adodc1
      
    anc = ObtenerAlto(DataGrid1, 10)
    
    'Obtenemos la siguiente numero de c�digo de Marca
    txtAux(0).Text = SugerirCodigoSiguienteStr("smantr", "codtrabajo")
    FormateaCampo txtAux(0)
    txtAux(1).Text = ""
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub


Private Sub BotonBuscar()
    CargaGrid "codtrabajo= -1"
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""

    LLamaLineas DataGrid1.Top + 250, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
On Error Resume Next
    CargaGrid ""
    If Adodc1.Recordset.RecordCount <= 0 Then
         'MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
         MsgBox "No hay ning�n registro en la tabla smantr", vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
        DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificar()
Dim Cad As String
Dim anc As Single
Dim i As Integer
On Error GoTo EModificar

    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    'El registro de codigo 0 no se puede Modificar ni Eliminar
    If EsCodigoCero(CStr(Adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    anc = ObtenerAlto(DataGrid1, 10)
    
    Cad = ""
    For i = 0 To 1
        Cad = Cad & DataGrid1.Columns(i).Text & "|"
    Next i
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    LLamaLineas anc, 4
   
    Screen.MousePointer = vbDefault
EModificar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar", Err.Description
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(0).Left = DataGrid1.Left + 340
    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
End Sub


Private Sub BotonEliminar()
Dim Sql As String
On Error GoTo Error2

    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    If EsCodigoCero(CStr(Adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    
    '### a mano
    Sql = "�Seguro que desea eliminar la Marca?" & vbCrLf
    Sql = Sql & vbCrLf & "C�digo: " & Format(Adodc1.Recordset.Fields(0), FormatoCod)
    Sql = Sql & vbCrLf & "Denominaci�n: " & Adodc1.Recordset.Fields(1)
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Adodc1.Recordset.AbsolutePosition
        Sql = "Delete from smantr where codtrabajo=" & Adodc1.Recordset!codtrabajo
        conn.Execute Sql
        CancelaADODC Adodc1
        CargaGrid ""
        CancelaADODC Me.Adodc1
        SituarDataPosicion Me.Adodc1, NumRegElim, Sql
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Marca", Err.Description
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim CadB As String
On Error Resume Next

    Select Case Modo
        Case 3   'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    BotonAnyadir
                End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk And BLOQUEADesdeFormulario(Me) Then
                If ModificaDesdeFormulario(Me, 3) Then
                    TerminaBloquear
                    i = Adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CancelaADODC Me.Adodc1
                    CargaGrid
                    Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & i)
                End If
                DataGrid1.SetFocus
            End If
            
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me, False)
            If CadB <> "" Then
                'Encuentra registros
                PonerModo 2
                CargaGrid CadB
                DataGrid1.SetFocus
            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            Me.lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
        Case 1 'Buscar
            CargaGrid
    End Select
    PonerModo 2
    DataGrid1.SetFocus
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Adodc1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If

    Cad = Adodc1.Recordset.Fields(0) & "|"
    Cad = Cad & Adodc1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Adodc1.Recordset.EOF Then 'And Modo = 0 Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
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
    
    FormatoCod = FormatoCampo(txtAux(0))
    
    '## A mano
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    DataGrid1.ClearFields

    CadAncho = False
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    PonerModo 2
      
    'Cadena consulta
    CadenaConsulta = "Select * from smantr"
    CargaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnNuevo_Click
        Case 2: mnModificar_Click
        Case 3: mnEliminar_Click
        Case 5: mnBuscar_Click
        Case 6: mnVerTodos_Click
        Case 8:  'Imprimir Listado de Marcas
           ' Me.Hide
           ' AbrirListado (1) 'OpcionListado=1
           ' Me.Show vbModal
    End Select
End Sub



Private Sub CargaGrid(Optional Sql As String)
Dim i As Byte
Dim b As Boolean
    
    b = DataGrid1.Enabled

    If Sql <> "" Then
        Sql = CadenaConsulta & " WHERE " & Sql
    Else
        Sql = CadenaConsulta
    End If
    Sql = Sql & " ORDER BY codtrabajo"

    CargaGridGnral DataGrid1, Me.Adodc1, Sql, False

    
    'Nombre producto
    i = 0
        DataGrid1.Columns(i).Caption = "C�digo"
        DataGrid1.Columns(i).Width = 900
        DataGrid1.Columns(i).NumberFormat = FormatoCod
    
    'Leemos del vector en 2
    i = 1
        DataGrid1.Columns(i).Caption = "Denominaci�n"
        DataGrid1.Columns(i).Width = 3730
            
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        CadAncho = True
    End If
   
   'No permitir cambiar tama�o de columnas
   For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
   Next i
   
   DataGrid1.RowHeight = 350
   
   
    'Habilitamos botones Modificar y Eliminar
   If Toolbar1.Buttons(2).Enabled = True Then
        Toolbar1.Buttons(2).Enabled = Not Adodc1.Recordset.EOF
        Toolbar1.Buttons(3).Enabled = Not Adodc1.Recordset.EOF
        mnModificar.Enabled = Not Adodc1.Recordset.EOF
        mnEliminar.Enabled = Not Adodc1.Recordset.EOF
    End If
   DataGrid1.Enabled = b
   DataGrid1.ScrollBars = dbgAutomatic
   
   'Actualizar indicador
   If Not Adodc1.Recordset.EOF And (Modo = 0) Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
   PonerOpcionesMenu
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


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    If Index = 0 Then PonerFormatoEntero txtAux(Index) 'codtrabajo
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean

    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    'Comprobar si ya existe el cod de marca en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(txtAux(0)) Then b = False
    End If
    
    DatosOk = b
End Function


Private Function SepuedeBorrar() As Boolean
    If SePuedeEliminarRelReparacione(3, CStr(Adodc1.Recordset.Fields(0))) Then
        SepuedeBorrar = True
    Else
        MsgBox "Existen registros relacionados", vbExclamation
        SepuedeBorrar = False
    End If
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub
