VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCPostal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C�digos Postales"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmCPostal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4140
      TabIndex        =   3
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3015
      TabIndex        =   2
      Top             =   5520
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "C�digo Postal|T|N|||scpostal|cpostal||S|"
      Text            =   "Da"
      Top             =   4800
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   1260
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Provincia|T|N|||scpostal|provincia||S|"
      Text            =   "Dato2"
      Top             =   4800
      Width           =   1395
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   4140
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   5
      Top             =   5400
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2280
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3240
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCPostal.frx":000C
      Height          =   4725
      Left            =   150
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   540
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8334
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
         HelpContextID   =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   1
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
Attribute VB_Name = "frmCPostal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private NombreTabla As String
Private CadenaConsulta As String
Private cadB As String 'Cadena de Busqueda

Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos
Dim Modo As Byte

'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas  INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo As Byte)
Dim b As Boolean

    Modo = vModo
    b = (Modo = 0)
    If b Then Me.lblIndicador.Caption = ""
    
    txtAux(0).visible = Not b
    txtAux(1).visible = Not b
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    End If

    'Si estamos mod or insert
    If Modo = 2 Then
       txtAux(0).BackColor = &H80000018
    Else
        txtAux(0).BackColor = &H80000005
    End If
    txtAux(0).Enabled = (Modo <> 2)
    
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                            'de permisos del usuario
End Sub


Private Sub BotonAnyadir()
Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
    lblIndicador.Caption = "INSERTAR"
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
    
    anc = ObtenerAlto(DataGrid1, 20)
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    LLamaLineas anc, 0
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub


Private Sub BotonBuscar()
    CargaGrid "cpostal='-1'"  'para vaciar los datos del Grid
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    lblIndicador.Caption = "BUSQUEDA"
    LLamaLineas 770, 2
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
    cadB = ""
    CargaGrid "length(cpostal)>2"
    If adodc1.Recordset.RecordCount <= 0 Then
         'MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
         MsgBox "No hay ning�n registro en la tabla CPostal", vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
'        PonerModo 2
        adodc1.Recordset.MoveFirst
'        PonerCampos
    End If
End Sub


Private Sub BotonModificar()
'Dim cad As String
Dim anc As Single
Dim i As Byte

    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
'    cad = ""
'    For i = 0 To 1
'        cad = cad & DataGrid1.Columns(i).Text & "|"
'    Next i

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    anc = ObtenerAlto(DataGrid1, 20)
    LLamaLineas anc, 1
   
    'Como es modificar
    PonerFoco txtAux(1)
   
    Screen.MousePointer = vbDefault
End Sub



Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo + 1
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(0).Left = DataGrid1.Left + 340
    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
End Sub


Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub

    Screen.MousePointer = vbHourglass
    '### a mano
    SQL = "Seguro que desea eliminar el Cod. Postal:"
    SQL = SQL & vbCrLf & "C�digo: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Provincia: " & adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from cpostal where cpostal='" & adodc1.Recordset!CPostal & "'"
        SQL = SQL & " AND provincia= '" & adodc1.Recordset!Provincia & "'"
        Conn.Execute SQL
        CancelaADODC Me.adodc1
        Espera 0.5
        If cadB <> "" Then
            CargaGrid cadB & "and length(cpostal)>2"
        Else
            CargaGrid "length(cpostal)>2"
        End If
        CancelaADODC Me.adodc1
    End If
    Screen.MousePointer = vbDefault
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar C.Postal", Err.Description
End Sub



Private Sub cmdAceptar_Click()
Dim i As String

Screen.MousePointer = vbHourglass
Select Case Modo
    Case 1 'INSERTAR
        If DatosOk Then
            If InsertarDesdeForm(Me) Then
                Espera 0.5
                If cadB <> "" Then
                    CargaGrid cadB & "and length(cpostal)>2"
                Else
                    CargaGrid "length(cpostal)>2"
                End If
                BotonAnyadir
            End If
        End If
    Case 2  'MODIFICAR
         If DatosOk Then
             If ModificarCPostal(Me) Then
'             If ModificaDesdeFormulario(Me) Then
                  Espera 0.5
                  i = adodc1.Recordset.Fields(0)
                  PonerModo 0
                  CancelaADODC Me.adodc1
                  If cadB <> "" Then
                      CargaGrid cadB & "and length(cpostal)>2"
                  Else
                      CargaGrid "length(cpostal)>2"
                  End If
                  adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
              End If
'              adodc1.Recordset.MoveFirst
              lblIndicador.Caption = ""
        End If
    Case 3
        'HacerBusqueda
        cadB = ObtenerBusqueda(Me, False)
        If cadB <> "" Then
            PonerModo 0
            CargaGrid cadB & " and length(cpostal)>2"
        End If
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1 'Insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        
        Case 2 'Modificar
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 3 'Buscar
            CargaGrid "length(cpostal)>2"
    End Select
    PonerModo 0
    PonerOpcionesMenu
    DataGrid1.SetFocus
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = adodc1.Recordset.Fields(0) & "|"
    cad = cad & adodc1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(2).Image = 1   'Bot�n Busqueda
        .Buttons(3).Image = 2   'Bot�n Recuperar Todos
        .Buttons(7).Image = 3   'Bot�n A�adir Nuevo Registro
        .Buttons(8).Image = 4   'Bot�n Modificar Registro
        .Buttons(9).Image = 5   'Bot�n Borrar Registro
        .Buttons(11).Image = 15  'Bot�n Salir
    End With

    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)

    CadAncho = False
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    PonerModo 0
'    PonerOpcionesMenu
    
    'Cadena consulta
    NombreTabla = "scpostal"
    CadenaConsulta = "Select * from " & NombreTabla
    CargaGrid "length(cpostal)>2"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
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
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 2
            mnBuscar_Click
    Case 3
           mnVerTodos_Click
    Case 7
            mnNuevo_Click
    Case 8
            mnModificar_Click
    Case 9
            mnEliminar_Click
    Case 11 'Salir
        mnSalir_Click
    End Select
End Sub


Private Sub CargaGrid(Optional SQL As String)
Dim i As Byte
Dim b As Boolean

    On Error GoTo ErrGrid
    
    b = DataGrid1.Enabled

    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY cpostal"
    

    CargaGridGnral DataGrid1, Me.adodc1, SQL, False

    
    'Nombre producto
    i = 0
        DataGrid1.Columns(i).Caption = "C.Postal"
        DataGrid1.Columns(i).Width = 1200
        DataGrid1.Columns(i).Alignment = dbgCenter
    
    'Leemos del vector en 2
    i = 1
        DataGrid1.Columns(i).Caption = "Provincia"
        DataGrid1.Columns(i).Width = 3250
            
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
   
    'Habilitamos botones Modificar y Eliminar
   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
   Toolbar1.Buttons(9).Enabled = Not adodc1.Recordset.EOF
   mnModificar.Enabled = Not adodc1.Recordset.EOF
   mnEliminar.Enabled = Not adodc1.Recordset.EOF
   DataGrid1.Enabled = b
   DataGrid1.ScrollBars = dbgAutomatic
   PonerOpcionesMenu
   Exit Sub
   
ErrGrid:
    MuestraError Err.Number, "Cargagrid", Err.Description
   
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), 3
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then Exit Sub
    'If Modo = 3 Then Exit Sub 'Busquedas
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    b = CompForm(Me, 3)
    If Not b Then Exit Function

'If Modo = 1 Then
'    'Estamos insertando
'     'BD 1: conexion a BD
'     Datos = DevuelveDesdeBD(1, NombreTabla, "cpostal", "cpostal", txtAux(0).Text, "T")
'     If Datos <> "" Then
'        MsgBox "Ya existe el C.Postal : " & txtAux(0).Text, vbInformation, "Comprobador de Campos"
'        b = False
'    End If
'End If
    DatosOk = b
End Function


Private Function SepuedeBorrar() As Boolean
    SepuedeBorrar = True
End Function


Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Sub PonerModoOpcionesMenu()
'segun Modo en el que estemos
Dim b As Boolean

    b = (Modo = 0 Or Modo = 2)
    'Buscar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ber Todos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
'     b = b And Not DeConsulta
    'A�adir
    Toolbar1.Buttons(7).Enabled = b
    Me.mnNuevo.Enabled = b
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
End Sub



Public Function ModificarCPostal(ByRef formulario As Form) As Boolean
'Funcion creada a partir De: ** ModificaDesdeFormulario **
'Pero en este caso no nos sirve la funci�n anterior ya que los dos campos: cpostal y provincia
'son clave primaria y vamos a modificar el nombre de provincia que es clave primaria
Dim Control As Object
Dim mTag As CTag
Dim Aux As String, aux2 As String
Dim cadWHERE As String
Dim cadUPDATE As String
Dim b As Boolean

On Error GoTo EModificaDesdeFormulario
    ModificarCPostal = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             If mTag.columna = "provincia" Then
                                aux2 = "'" & adodc1.Recordset!Provincia & "'"
                                cadWHERE = cadWHERE & "(" & mTag.columna & " = " & aux2 & ")"
                             Else
                                cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                             End If
'                        Else
'                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            If mTag.columna = "provincia" Then
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
       End If
    Next Control
    'Construimos el SQL
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    Conn.Execute Aux, , adCmdText

    
    ModificarCPostal = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

