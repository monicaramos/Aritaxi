VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacFormasEnvio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Textos Clientes Agrupados"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   6075
   Icon            =   "frmFacFormasEnvio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   10
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
   Begin VB.ComboBox CboEtiqueta 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmFacFormasEnvio.frx":000C
      Left            =   2400
      List            =   "frmFacFormasEnvio.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Opci�n Imprimir etiqueta|N|N|||senvio|impetiqu||N|"
      Top             =   4920
      Width           =   975
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
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "C�digo|N|N|0|999|senvio|codenvio|000|S|"
      Text            =   "Codigo"
      Top             =   4920
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
      Left            =   960
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "Descripci�n|T|N|||senvio|nomenvio||N|"
      Text            =   "Descripcion"
      Top             =   4920
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacFormasEnvio.frx":0010
      Height          =   4545
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   825
      Width           =   5815
      _ExtentX        =   10266
      _ExtentY        =   8017
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
      Left            =   4800
      TabIndex        =   4
      Top             =   5550
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
      Left            =   3600
      TabIndex        =   3
      Top             =   5550
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
      Left            =   4800
      TabIndex        =   7
      Top             =   5550
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   5
      Top             =   5400
      Width           =   1995
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
         Left            =   420
         TabIndex        =   6
         Top             =   210
         Width           =   1200
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   1080
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
Attribute VB_Name = "frmFacFormasEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 302

Private CadenaConsulta As String

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
'-------------------------------------------------------


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(vModo As Byte)
Dim b As Boolean
Dim I As Integer

    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).BackColor = vbWhite
    Next I
    txtAux(0).visible = Not b
    txtAux(1).visible = Not b
    CboEtiqueta.visible = Not b
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
    PonerModoUsuarioGnral Modo, "aritaxi"
End Sub

 
Private Sub PonerModoUsuarioGnral(Modo As Byte, Aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(Aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = Toolbar1.Buttons(1).Enabled And DBLet(Rs!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = Toolbar1.Buttons(2).Enabled And DBLet(Rs!Modificar, "N")
        Toolbar1.Buttons(3).Enabled = Toolbar1.Buttons(3).Enabled And DBLet(Rs!creareliminar, "N")
        
        Toolbar1.Buttons(5).Enabled = Toolbar1.Buttons(5).Enabled And DBLet(Rs!Ver, "N")
        Toolbar1.Buttons(6).Enabled = Toolbar1.Buttons(6).Enabled And DBLet(Rs!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = Toolbar1.Buttons(8).Enabled And DBLet(Rs!Imprimir, "N")
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub







Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2)
    'Buscar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ber Todos
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    b = b And Not DeConsulta
    'A�adir
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(8).Enabled = b
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub BotonAnyadir()
Dim anc As Single

    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
    
    'Obtenemos la siguiente numero de factura
    txtAux(0).Text = SugerirCodigoSiguienteStr("senvio", "codenvio")
    txtAux(0).Text = Format(txtAux(0).Text, FormatoCod)
    txtAux(1).Text = ""
    CboEtiqueta.ListIndex = 1
    
    anc = ObtenerAlto(DataGrid1, 10)
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub


Private Sub BotonBuscar()
    CargaGrid "codenvio= -1"
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    Me.CboEtiqueta.ListIndex = -1
    LLamaLineas DataGrid1.top + 240, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
On Error Resume Next

    CargaGrid ""
    If adodc1.Recordset.RecordCount <= 0 Then
         MsgBox "No hay ning�n registro en la tabla senvio", vbInformation
         Screen.MousePointer = vbDefault
          Exit Sub
    Else
        PonerModo 2
         DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificar()
Dim anc As Single
Dim I As Integer

    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    Select Case DataGrid1.Columns(2).Value
        Case "Si"
            CboEtiqueta.ListIndex = 0
        Case "No"
            CboEtiqueta.ListIndex = 1
    End Select
    
    anc = ObtenerAlto(DataGrid1, 10)
    LLamaLineas anc, 4
   
    'Como es modificar
'    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).top = alto
    txtAux(1).top = alto
    CboEtiqueta.top = alto - 15
    txtAux(0).Left = DataGrid1.Left + 340
    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
    CboEtiqueta.Left = txtAux(1).Left + txtAux(1).Width + 55
End Sub

Private Sub BotonEliminar()
Dim Sql As String
On Error GoTo Error2
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub

    '### a mano
    Sql = "�Seguro que desea eliminar el C�digo?" & vbCrLf
    Sql = Sql & vbCrLf & "C�digo: " & Format(adodc1.Recordset.Fields(0), FormatoCod)
    Sql = Sql & vbCrLf & "Denominaci�n: " & adodc1.Recordset.Fields(1)
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.adodc1.Recordset.AbsolutePosition
        Sql = "Delete from senvio where codenvio=" & adodc1.Recordset!CodEnvio
        conn.Execute Sql
        CancelaADODC Me.adodc1
        CargaGrid ""
        CancelaADODC Me.adodc1
        SituarDataPosicion Me.adodc1, NumRegElim, Sql
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Textos Clientes Agrupados", Err.Description
End Sub

Private Sub CboEtiqueta_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
On Error GoTo EAceptar

    Select Case Modo
        Case 3 'Insertar
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    BotonAnyadir
                End If
            End If
        Case 4  'Modificar
            If DatosOk And BLOQUEADesdeFormulario(Me) Then
                If ModificaDesdeFormulario(Me, 3) Then
                    TerminaBloquear
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CancelaADODC Me.adodc1
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                End If
                DataGrid1.SetFocus
            End If
        Case 1 'HacerBusqueda
            CadB = ObtenerBusqueda(Me, False)
            If CadB <> "" Then
                PonerModo 2
                CargaGrid CadB
                DataGrid1.SetFocus
            End If
    End Select
    
EAceptar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
On Error Resume Next

    Select Case Modo
        Case 1 'Busqueda
            CargaGrid
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            TerminaBloquear
            Me.lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    End Select
    PonerModo 2
    DataGrid1.SetFocus
    
    If Err.Number <> 0 Then Err.Clear
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

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     If Not adodc1.Recordset.EOF Then 'And Modo = 0 Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .Buttons(1).Image = 3   'Bot�n A�adir Nuevo Registro
        .Buttons(2).Image = 4   'Bot�n Modificar Registro
        .Buttons(3).Image = 5   'Bot�n Borrar Registro
        .Buttons(5).Image = 1   'Bot�n Buscar
        .Buttons(6).Image = 2   'Bot�n Recuperar Todos
        .Buttons(8).Image = 16  'Bot�n Imprimir
    End With
    
'    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With
    
    
    FormatoCod = FormatoCampo(txtAux(0))
    
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    PonerModo 2
    
    'Cadena consulta
    CadenaConsulta = "Select codenvio,nomenvio,If(impetiqu=1,""Si"",""No"") AS etiq from senvio"
    CargaGrid
    CargaCombo
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
        Case 5: BotonBuscar
        Case 6: BotonVerTodos
        Case 1: BotonAnyadir
        Case 2: BotonModificar
        Case 3: BotonEliminar
        Case 8 'Imprimir listado de Formas de Env�o
                Me.Hide
                AbrirListado (23) 'OpcionListado=23
                Me.Show vbModal
    End Select
End Sub


Private Sub CargaGrid(Optional Sql As String)
Dim b As Boolean
Dim tots As String

    b = DataGrid1.Enabled

    If Sql <> "" Then
        Sql = CadenaConsulta & " WHERE " & Sql
    Else
        Sql = CadenaConsulta
    End If
    Sql = Sql & " ORDER BY codenvio"
    
    CargaGridGnral DataGrid1, Me.adodc1, Sql, False

    '### a mano
    tots = "S|txtAux(0)|T|C�digo|900|;S|txtAux(1)|T|Denominaci�n|3300|;S|CboEtiqueta|C|Etiqueta|1020|;"
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
   
   'Actualizar indicador
   If Not adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
End Sub

Private Sub CargaCombo()
    'Carga la lista de impresi�n de etiquetas
    CboEtiqueta.Clear
    CboEtiqueta.AddItem "Si"
    CboEtiqueta.ItemData(CboEtiqueta.NewIndex) = 1
    
    CboEtiqueta.AddItem "No"
    CboEtiqueta.ItemData(CboEtiqueta.NewIndex) = 0
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
    If Index = 0 Then PonerFormatoEntero txtAux(0)
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    
    b = CompForm(Me, 3)
    If Not b Then Exit Function

    'comprobamos si ya existe la forma de envio
    If Modo = 3 Then 'Insertar
        If ExisteCP(txtAux(0)) Then b = False
    End If

    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub
