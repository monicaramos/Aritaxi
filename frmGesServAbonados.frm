VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGesServAbonados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicios Abonados"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
   ClipControls    =   0   'False
   Icon            =   "frmGesServAbonados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   210
      TabIndex        =   18
      Top             =   60
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   19
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
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   10530
      TabIndex        =   5
      Tag             =   "Int.Contable|N|N|||sfactclitr|facturado|||"
      Top             =   3630
      Visible         =   0   'False
      Width           =   255
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
      Height          =   315
      Index           =   4
      Left            =   6240
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "Concepto|T|S|||sfactclitr|concepto|||"
      Text            =   "concept"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   10770
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5460
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   5760
      Width           =   2535
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
      Left            =   1080
      TabIndex        =   12
      ToolTipText     =   "Buscar cliente"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
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
      Left            =   6000
      TabIndex        =   11
      ToolTipText     =   "Buscar fecha"
      Top             =   3630
      Visible         =   0   'False
      Width           =   195
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
      Left            =   360
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Cod.Cliente|N|N|0|999999|sfactclitr|codclien|000000|S|"
      Text            =   "client"
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
      Height          =   315
      Index           =   2
      Left            =   8640
      MaxLength       =   7
      TabIndex        =   3
      Tag             =   "Servicios|N|N|||sfactclitr|numserv|###,##0||"
      Text            =   "servici"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
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
      Height          =   315
      Index           =   3
      Left            =   5010
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||sfactclitr|fecfactu|dd/mm/yyyy|S|"
      Text            =   "fecha"
      Top             =   3630
      Visible         =   0   'False
      Width           =   975
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
      Left            =   10530
      TabIndex        =   6
      Top             =   5880
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
      Left            =   11775
      TabIndex        =   7
      Top             =   5880
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
      Left            =   11790
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
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
      Index           =   0
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   10
      Text            =   "Text2 Text2 Text2 Text2 Text2 Text2 Text"
      Top             =   3600
      Visible         =   0   'False
      Width           =   3615
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
      Height          =   315
      Index           =   1
      Left            =   9300
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Importe|N|N|||sfactclitr|importe|###,##0.00|N|"
      Text            =   "Impo"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1065
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
      Bindings        =   "frmGesServAbonados.frx":000C
      Height          =   4545
      Left            =   240
      TabIndex        =   15
      Top             =   810
      Width           =   12665
      _ExtentX        =   22331
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
   Begin VB.Image imgDoc 
      Height          =   330
      Index           =   0
      Left            =   3420
      MousePointer    =   4  'Icon
      Tag             =   "-1"
      ToolTipText     =   "Detalle de Servicios"
      Top             =   5850
      Width           =   360
   End
   Begin VB.Label Label2 
      Caption         =   "IMPORTE TOTAL: "
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
      Left            =   8760
      TabIndex        =   17
      Top             =   5490
      Width           =   1815
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
      TabIndex        =   9
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
Attribute VB_Name = "frmGesServAbonados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Private Const IdPrograma = 318

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
        Case 0 'Cod socio
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0"
            frmCli.Show vbModal
            Set frmCli = Nothing
        
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


Private Sub DataGrid1_DblClick()
'[Monica]10/09/2014: de momento comentado pq partimos de la shilla
'    imgDoc_Click (0)
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
Dim I As Integer
    'Icono del formulario
    Me.Icon = frmppal.Icon
    
    PrimeraVez = True
    
    'ICONOS de laLa toolbar
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .Buttons(1).Image = 3   'Botón Añadir Nuevo Registro
        .Buttons(2).Image = 4   'Botón Modificar Registro
        .Buttons(3).Image = 5   'Botón Borrar Registro
        .Buttons(5).Image = 1   'Botón Buscar
        .Buttons(6).Image = 2   'Botón Recuperar Todos
        .Buttons(8).Image = 16  'Botón Imprimir
    End With
    
'    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With
    
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
 
    
'[Monica]10/09/2014: de momento comentado pq partimos de la shilla
'    For i = 0 To imgDoc.Count - 1
'        imgDoc(i).Picture = frmPpal.ImageListTPV.ListImages(8).Picture
'    Next i

    PonerModo 0
    
    NombreTabla = "sfactclitr" 'Tabla de facturas asociados
    Ordenacion = " ORDER BY fecfactu desc, codclien "
    
    CadenaConsulta = "Select sfactclitr.codclien, scliente.nomclien, sfactclitr.fecfactu, "
    CadenaConsulta = CadenaConsulta & " sfactclitr.concepto, sfactclitr.numserv, sfactclitr.importe, sfactclitr.facturado,"
    CadenaConsulta = CadenaConsulta & "  IF(facturado=1,'*','') as factur "
    CadenaConsulta = CadenaConsulta & " from " & NombreTabla & " INNER JOIN scliente ON sfactclitr.codclien = scliente.codclien "
    CadenaConsulta = CadenaConsulta & " WHERE sfactclitr.codclien is null " 'No recupera datos
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    CargaGrid (Modo = 2)
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
    tots = "S|txtAux(0)|T|Cliente|1050|;S|cmdAux(0)|B||0|;S|Text2(0)|T|Nombre|3400|;"
    tots = tots & "S|txtAux(3)|T|Fecha|1450|;S|cmdAux(1)|B||0|;S|txtAux(4)|T|Concepto|3450|;S|txtAux(2)|T|Serv.|900|;"
    tots = tots & "S|txtAux(1)|T|Importe|1450|;N||||0|;S|chkAux(0)|CB|Fa|360|;"
    
    arregla tots, DataGrid1, Me, 350
    

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

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento clientes
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
Dim indice As Byte
    indice = 3
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
        Case 5 'Busqueda
            mnBuscar_Click
        Case 6 'Ver Todos
            mnVerTodos_Click
        Case 1 'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3 'Eliminar
            mnEliminar_Click
        Case 8 'Imprimir
            BotonImprimir
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
            txtAux(0).BackColor = vbLightBlue 'vbYellow
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
                 
    '[Monica]01/04/2014: para ver los servicios del cliente
    imgDoc(0).visible = (Modo = 2)
    imgDoc(0).Enabled = (Modo = 2)
    
    '-----------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    Me.DataGrid1.Enabled = (Modo <> 3 And Modo <> 4)
    BloquearClavesP (Modo = 4) ' si modificar
    
    BloquearChk Me.chkAux(0), (Modo = 4) Or (Modo = 3)
    Me.chkAux(0).visible = (Modo = 1)


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
    Toolbar1.Buttons(1).Enabled = (b Or (Modo = 0))
    Me.mnnuevo.Enabled = (b Or (Modo = 0))
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
    Me.mnvertodos.Enabled = Not b
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
    chkAux(0).Value = 0
    
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
    
    Sql = "Select sfactclitr.codclien, scliente.nomclien, sfactclitr.fecfactu, sfactclitr.concepto, "
    Sql = Sql & " sfactclitr.numserv, sfactclitr.importe, sfactclitr.facturado,"
    Sql = Sql & "  IF(facturado=1,'*','') as factur "
    Sql = Sql & " from " & NombreTabla & " INNER JOIN scliente ON sfactclitr.codclien = scliente.codclien "

    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then Sql = Sql & CadenaBusqueda
    Else
        Sql = Sql & " WHERE sfactclitr.codclien = -1"
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
        anc = ObtenerAlto(Me.DataGrid1, 10)
        LLamaLineas anc
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco txtAux(0)
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
    
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
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
    anc = ObtenerAlto(Me.DataGrid1, 5)
    LLamaLineas anc
    
    txtAux(2).Text = 0
    txtAux(3).Text = Format(Now, "dd/mm/yyyy")
    chkAux(0).Value = 0

    
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
Dim I As Integer
Dim anc As Single

    If CInt(Data1.Recordset!facturado) = 1 Then
        MsgBox "La línea está facturada, no se permite ni modificar ni eliminar.", vbExclamation
        Exit Sub
    End If

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
'    CargaTxtAux True, False
    anc = ObtenerAlto(Me.DataGrid1, 5)
    LLamaLineas anc
    
    
    '---- poner valores grabados
    'codclien
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "N")
    FormateaCampo txtAux(I)
    
    'nomclien
    Text2(0).Text = DBLet(DataGrid1.Columns(1).Value, "T")
    
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
    
    
    Me.chkAux(0).Value = Me.Data1.Recordset!facturado

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
        Sql = Sql & vbCrLf & "Cliente: " & Format(Data1.Recordset.Fields(0).Value, "000000") & " - " & Data1.Recordset.Fields(1).Value
        Sql = Sql & vbCrLf & "Fecha: " & Data1.Recordset.Fields(2).Value
        
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
            'Hay que eliminar
            NumRegElim = Me.Data1.Recordset.AbsolutePosition
            
            '[Monica]02/04/2014: borramos los servicios asociados si existen
            Sql = "Delete from sfactclitr_serv where codclien= " & DBSet(Data1.Recordset!CodClien, "N")
            Sql = Sql & " and fecfactu = " & DBSet(Data1.Recordset!FecFactu, "F")
            conn.Execute Sql
            
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
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cliente
            Case 3: KEYBusqueda KeyAscii, 1 'fecha
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    cmdAux_Click (indice)
End Sub
    



Private Sub BloquearClavesP(bol As Boolean)
'Si BloquearClavesPrimarias=true deshablilita los textbox de codigos y lo pone amarillo
'y habilita el resto de campos para introducir nuevos valores
'Si BloquearClavesPrimarias=false habilita los textbox de codigos para introducir
Dim I As Byte

    For I = 0 To 0 'Codigo socio
        BloquearTxt txtAux(I), bol
        Me.cmdAux(I).Enabled = Not bol
    Next I
    ' fecha bloqueada
    Me.cmdAux(1).Enabled = Not bol
    BloquearTxt txtAux(3), bol
    
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim campo As String, Tabla As String
Dim campo2 As String

  
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'cod. cliente
            If PonerFormatoEntero(txtAux(Index)) Then
'                campo2 = "numeruve"
                Text2(Index).Text = DevuelveDesdeBDNew(conAri, "scliente", "nomclien", "codclien", txtAux(Index).Text, "N") ' , campo2)
            Else
                Text2(Index).Text = ""
            End If
            
        Case 3 'Fecha
            PonerFormatoFecha txtAux(Index)
            
        Case 2 ' numero de servicios
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
            txtAux(jj).top = alto
            txtAux(jj).visible = b
        Next jj
        
        For jj = 0 To Text2.Count - 1
            Text2(jj).Height = Me.DataGrid1.RowHeight
            Text2(jj).top = alto
            Text2(jj).visible = b
        Next jj
        
        For jj = 0 To Me.cmdAux.Count - 1
            Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
            Me.cmdAux(jj).top = alto
            Me.cmdAux(jj).visible = b
        Next jj
        
        Me.chkAux(0).top = alto
        Me.chkAux(0).visible = b
        

        
End Sub



Private Sub CalcularTotales(CADENA As String)
Dim Importe  As Currency
Dim Compleme As Currency
Dim Penaliza As Currency

Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error Resume Next
    
    Sql = "select sum(importe) importe  from (" & CADENA & ") aaaaa"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Importe = 0
    Text1.Text = ""
    
    If TotalRegistrosConsulta(CADENA) = 0 Then Exit Sub
    
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then Importe = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
    
        Text1.Text = Format(Importe, "###,###,##0.00")
    End If
    Rs.Close
    Set Rs = Nothing

    
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

Private Sub imgDoc_Click(Index As Integer)
Dim vCadena As String
'[Monica]10/09/2014: de momento comentado pq partimos de la shilla
'    Select Case Index
'        Case 0
'            frmGesServAbonados2.Cliente = Me.Data1.Recordset!CodClien
'            frmGesServAbonados2.FecFactura = Me.Data1.Recordset!FecFactu
'            frmGesServAbonados2.Caption = "Servicios Cliente " & Format(Me.Data1.Recordset!CodClien, "000000") & " " & Me.Data1.Recordset!nomclien
'            frmGesServAbonados2.Show vbModal
'
'    End Select
    
End Sub

