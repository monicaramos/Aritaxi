VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LOG de acciones"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   11115
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   11
      Top             =   30
      Width           =   1545
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
               Object.Tag             =   "0"
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
      Index           =   3
      Left            =   7200
      TabIndex        =   4
      Tag             =   "Desc|T|N|||slog|descripcion||N|"
      Text            =   "Descripcion"
      Top             =   4920
      Width           =   2355
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
      Index           =   2
      Left            =   5640
      TabIndex        =   3
      Tag             =   "PC|T|N|||slog|pc||N|"
      Text            =   "Descripcion"
      Top             =   4920
      Width           =   1395
   End
   Begin VB.ComboBox CboTipoSitu 
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
      ItemData        =   "frmLog.frx":000C
      Left            =   2880
      List            =   "frmLog.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Accion|N|N|||slog|accion||N|"
      Top             =   4920
      Width           =   1215
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
      TabIndex        =   0
      Tag             =   "Fecha |F|N|||slog|fecha|dd/mm/yyyy hh:mm|S|"
      Text            =   "Codigo"
      Top             =   4920
      Width           =   2235
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
      Left            =   4200
      TabIndex        =   2
      Tag             =   "usuario|T|N|||slog|usuario||N|"
      Text            =   "Descripcion"
      Top             =   4920
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLog.frx":0010
      Height          =   4545
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   780
      Width           =   10735
      _ExtentX        =   18944
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
      Left            =   9840
      TabIndex        =   6
      Top             =   5490
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
      Left            =   8640
      TabIndex        =   5
      Top             =   5490
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
      Left            =   9840
      TabIndex        =   9
      Top             =   5490
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   7
      Top             =   5355
      Width           =   2115
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
         Left            =   180
         TabIndex        =   8
         Top             =   180
         Width           =   1800
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
         Visible         =   0   'False
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
         Visible         =   0   'False
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
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

Public Event DatoSeleccionado(CadenaSeleccion As String)

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
    txtAux(2).visible = Not b
    txtAux(3).visible = Not b
    
    CboTipoSitu.visible = Not b
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    End If
    
    'Si estamos insertando o busqueda
    
    BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    BloquearTxt txtAux(1), (Modo <> 3 And Modo <> 1)
    BloquearTxt txtAux(2), (Modo <> 3 And Modo <> 1)
    BloquearTxt txtAux(3), (Modo <> 3 And Modo <> 1)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    'PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2)
    'Buscar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ber Todos
    Toolbar1.Buttons(2).Enabled = b
    Me.mnVerTodos.Enabled = b
    
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


'Private Sub BotonAnyadir()
'Dim anc As Single
'
'    'Situamos el grid al final
'    AnyadirLinea DataGrid1, adodc1
'
'    'Obtenemos la siguiente numero de codigo de Situaciones
'    txtAux(0).Text = SugerirCodigoSiguienteStr("ssitua", "codsitua")
'    FormateaCampo txtAux(0)
'    txtAux(1).Text = ""
'    CboTipoSitu.ListIndex = 1
'
'    anc = ObtenerAlto(DataGrid1)
'    LLamaLineas anc, 3
'
'    'Ponemos el foco
'    PonerFoco txtAux(0)
'End Sub


Private Sub BotonBuscar()
    CargaGrid "accion= -1"
    limpiar Me
    Me.CboTipoSitu.ListIndex = -1
    LLamaLineas DataGrid1.Top + 250, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
On Error Resume Next

    CargaGrid ""
    If Adodc1.Recordset.RecordCount <= 0 Then
         MsgBox "No hay ning�n registro en la tabla ssitua", vbInformation
         Screen.MousePointer = vbDefault
          Exit Sub
    Else
        PonerModo 2
         DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


'Private Sub BotonModificar()
'Dim anc As Single
'Dim I As Integer
'
'    If adodc1.Recordset.EOF Then Exit Sub
'    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
'
'    'El registro de codigo 0 no se puede Modificar ni Eliminar
'    If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub
'
'    Screen.MousePointer = vbHourglass
'
'    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
'        I = DataGrid1.Bookmark - DataGrid1.FirstRow
'        DataGrid1.Scroll 0, I
'        DataGrid1.Refresh
'    End If
'
'    'Llamamos al form
'    txtAux(0).Text = DataGrid1.Columns(0).Text
'    txtAux(1).Text = DataGrid1.Columns(1).Text
'    Select Case DataGrid1.Columns(2).Value
'        Case "Si"
'            CboTipoSitu.ListIndex = 1
'        Case "No"
'            CboTipoSitu.ListIndex = 0
'    End Select
'
'    anc = ObtenerAlto(DataGrid1)
'    LLamaLineas anc, 4
'
'    'Como es modificar
''    PonerFoco txtAux(1)
'   Screen.MousePointer = vbDefault
'End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
'Pone posicion TOP y LEFT de los controles en el form
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    txtAux(3).Top = alto
    CboTipoSitu.Top = alto - 15
    txtAux(0).Left = DataGrid1.Left + 320
    CboTipoSitu.Left = txtAux(0).Left + txtAux(0).Width + 45
    txtAux(1).Left = CboTipoSitu.Left + CboTipoSitu.Width + 45
    txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 45
    txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 45
    
End Sub


'Private Sub BotonEliminar()
'Dim SQL As String
'On Error GoTo Error2
'    'Ciertas comprobaciones
'    If adodc1.Recordset.EOF Then Exit Sub
'
'    'El registro de codigo 0 no se puede Modificar ni Eliminar
'    If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub
'
'    '### a mano
'    SQL = "�Seguro que desea eliminar la Situaci�n?" & vbCrLf
'    SQL = SQL & vbCrLf & "C�digo: " & Format(adodc1.Recordset.Fields(0), FormatoCod)
'    SQL = SQL & vbCrLf & "Denominaci�n: " & adodc1.Recordset.Fields(1)
'    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
'        'Hay que eliminar
'        NumRegElim = Me.adodc1.Recordset.AbsolutePosition
'        SQL = "Delete from ssitua where codsitua=" & adodc1.Recordset!codsitua
'        Conn.Execute SQL
'        CancelaADODC Me.adodc1
'        CargaGrid ""
'        CancelaADODC Me.adodc1
'        SituarDataPosicion Me.adodc1, NumRegElim, SQL
'    End If
'
'Error2:
'    Screen.MousePointer = vbDefault
'    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Situaciones Especiales", Err.Description
'End Sub


Private Sub CboTipoSitu_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
'Dim I As Integer
Dim CadB As String
Dim Aux As String
On Error Resume Next

    Select Case Modo
        Case 1 'HacerBusqueda
            'COMO ES UN CAMPO FECHA HORA LO TRATARE DE FORMA ESPECIAL
            Aux = ""
            If txtAux(0).Text <> "" Then
                'SI lo que han puesto es una fecha
                Aux = txtAux(0).Text
                If EsFechaOK(Aux) Then
                   Aux = Format(Aux, FormatoFecha)
                   Aux = "slog.fecha  >=  '" & Aux & "' AND slog.fecha <= '" & Aux & " 23:59:59'"
                   'Pongo txtaux(0)=""
                   txtAux(0).Text = ""
                Else
                    Aux = ""
                End If
            End If
        
        
            CadB = ObtenerBusqueda(Me, False)
            If CadB <> "" And Aux <> "" Then Aux = " AND " & Aux
            CadB = CadB & Aux
            If CadB <> "" Then
                PonerModo 2
                CargaGrid CadB
                DataGrid1.SetFocus
            End If
'        Case 3 'Insertar
'            If DatosOk Then
'                If InsertarDesdeForm(Me) Then
'                    CargaGrid
'                    BotonAnyadir
'                End If
'            End If
'        Case 4 'Modificar
'            If DatosOk And BLOQUEADesdeFormulario(Me) Then
'                If ModificaDesdeFormulario(Me, 3) Then
'                   TerminaBloquear
'                   I = adodc1.Recordset.Fields(0)
'                   PonerModo 2
'                   CancelaADODC Me.adodc1
'                   CargaGrid
'                   adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
'                End If
'                DataGrid1.SetFocus
'            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    Select Case Modo
        Case 1 'Busqueda
            CargaGrid
        Case 3
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            TerminaBloquear
            Me.lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    End Select
    PonerModo 2
    DataGrid1.SetFocus
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
    If cmdRegresar.visible = True Then
        cmdRegresar_Click
    Else
        If Not (Adodc1.Recordset Is Nothing) Then
            If Not Adodc1.Recordset.EOF Then
                CadenaDesdeOtroForm = "Fecha: " & Adodc1.Recordset!Fecha & vbCrLf
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Usuario / PC : " & Adodc1.Recordset!Usuario & " - " & Adodc1.Recordset!PC & vbCrLf
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Accion: " & DBLet(Adodc1.Recordset!NomArtic, "T") & vbCrLf & vbCrLf
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Replace(Space(80), " ", "-") & vbCrLf
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Descripci�n:" & vbCrLf & Adodc1.Recordset!Descripcion
                MsgBox CadenaDesdeOtroForm, vbInformation
                CadenaDesdeOtroForm = ""
            End If
        End If
    End If
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
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun1
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2   'Todos
    End With

    
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    PonerModo 2
    
    
    CargaCombo
    
    'Cadena consulta
'    CadenaConsulta = "select slog.fecha,nomartic,usuario,pc,descripcion from slog,tmpnlotes "
'    CadenaConsulta = CadenaConsulta & " where tmpnlotes.codusu=" & vUsu.Codigo & " and slog.accion=tmpnlotes.codprove"
    CadenaConsulta = "select slog.fecha,nombre2 nomartic,usuario,pc,descripcion from slog,tmpinformes "
    CadenaConsulta = CadenaConsulta & " where tmpinformes.codusu=" & vUsu.Codigo & " and slog.accion=tmpinformes.codigo1"
    
    CargaGrid
    FormatoCod = FormatoCampo(txtAux(0))
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
   ' BotonEliminar
End Sub

Private Sub mnModificar_Click()
  '  BotonModificar
End Sub

Private Sub mnNuevo_Click()
  '  BotonAnyadir
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
        Case 1: BotonBuscar
        Case 2: BotonVerTodos
    End Select
End Sub


Private Sub CargaGrid(Optional Sql As String)
Dim b As Boolean
Dim tots As String
    
    b = DataGrid1.Enabled
    
    If Sql <> "" Then
        Sql = CadenaConsulta & " AND " & Sql
        Else
        Sql = CadenaConsulta
    End If
    Sql = Sql & " ORDER BY fecha desc"
    
    CargaGridGnral DataGrid1, Me.Adodc1, Sql, False
    
    '### a mano
    tots = "S|txtAux(0)|T|Fecha|2000|;S|CboTipoSitu|C|Accion|2200|;S|txtAux(1)|T|Usuario|1300|;"
    tots = tots & "S|txtAux(2)|T|PC|1400|;S|txtAux(3)|T|Descripcion|3250|;"
    
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
   
   'Actualizar indicador
   If Not Adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If

'    'Habilitamos botones Modificar y Eliminar
'   If Toolbar1.Buttons(6).Enabled Then
'        Toolbar1.Buttons(6).Enabled = Not adodc1.Recordset.EOF
'        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
'        mnModificar.Enabled = Not adodc1.Recordset.EOF
'        mnEliminar.Enabled = Not adodc1.Recordset.EOF
'   End If
'   DataGrid1.Enabled = b
'   DataGrid1.ScrollBars = dbgAutomatic
'
'   PonerOpcionesMenu
End Sub

Private Sub CargaCombo()
Dim L As Collection
    
    conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo
    CboTipoSitu.Clear
    FormatoCod = ""
    Set L = New Collection
    Set LOG = New cLOG
    If LOG.DevuelveAcciones(L) Then
        'Carga la lista de impresi�n de etiquetas
        If L.Count > 0 Then
            For NumRegElim = 1 To L.Count
                
                CboTipoSitu.AddItem RecuperaValor(L.item(NumRegElim), 2)
                CboTipoSitu.ItemData(CboTipoSitu.NewIndex) = Val(RecuperaValor(L.item(NumRegElim), 1))
                FormatoCod = FormatoCod & ",(" & vUsu.Codigo & ",0,'2007-07-04'," & CboTipoSitu.ItemData(CboTipoSitu.NewIndex) & ",0,'" & DevNombreSQL(CboTipoSitu.List(CboTipoSitu.NewIndex)) & "')"
            Next NumRegElim
        End If
    End If
    Set LOG = Nothing
    If FormatoCod <> "" Then
        FormatoCod = Mid(FormatoCod, 2) 'quito la coma
        FormatoCod = "Insert into tmpinformes (codusu,nombre1,fecha1,codigo1,campo1,nombre2) VALUES " & FormatoCod & ";"
        conn.Execute FormatoCod
    End If
    
'    conn.Execute "DELETE FROM tmpnlotes where codusu =" & vUsu.Codigo
'    CboTipoSitu.Clear
'    FormatoCod = ""
'    Set L = New Collection
'    Set LOG = New cLOG
'    If LOG.DevuelveAcciones(L) Then
'        'Carga la lista de impresi�n de etiquetas
'        If L.Count > 0 Then
'            For NumRegElim = 1 To L.Count
'
'                CboTipoSitu.AddItem RecuperaValor(L.item(NumRegElim), 2)
'                CboTipoSitu.ItemData(CboTipoSitu.NewIndex) = Val(RecuperaValor(L.item(NumRegElim), 1))
'                FormatoCod = FormatoCod & ",(" & vUsu.Codigo & ",0,'2007-07-04'," & CboTipoSitu.ItemData(CboTipoSitu.NewIndex) & ",0,'" & DevNombreSQL(CboTipoSitu.List(CboTipoSitu.NewIndex)) & "')"
'            Next NumRegElim
'        End If
'    End If
'    Set LOG = Nothing
'    If FormatoCod <> "" Then
'        FormatoCod = Mid(FormatoCod, 2) 'quito la coma
'        FormatoCod = "Insert into tmpnlotes (codusu,numalbar,fechaalb,codprove,numlinea,nomartic) VALUES " & FormatoCod & ";"
'        conn.Execute FormatoCod
'    End If



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
    'If Index = 0 Then PonerFormatoEntero txtAux(Index)
End Sub


'Private Function DatosOk() As Boolean
'Dim b As Boolean
'
'    b = CompForm(Me, 3)
'    If Not b Then Exit Function
'
'    'comprobar si ya existe el codigo de situacion en la tabla
'    If Modo = 3 Then 'Insertar
'        If ExisteCP(txtAux(0)) Then b = False
'    End If
'
'    DatosOk = b
'End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

