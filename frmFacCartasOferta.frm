VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacCartasOferta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartas"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   ClipControls    =   0   'False
   Icon            =   "frmFacCartasOferta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3390
      TabIndex        =   25
      Top             =   60
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   210
         TabIndex        =   26
         Top             =   210
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "�ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   180
      TabIndex        =   23
      Top             =   60
      Width           =   3075
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   24
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
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
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
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
      Left            =   7950
      TabIndex        =   22
      Top             =   270
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Index           =   3
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "P�rrafo 1|T|S|||scartas|parrafo1||N|"
      Text            =   "frmFacCartasOferta.frx":000C
      Top             =   2220
      Width           =   9615
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   700
      Index           =   4
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Tag             =   "P�rrafo 2|T|S|||scartas|parrafo2||N|"
      Text            =   "frmFacCartasOferta.frx":010E
      Top             =   2925
      Visible         =   0   'False
      Width           =   8835
   End
   Begin VB.TextBox Text1 
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
      Index           =   2
      Left            =   150
      MaxLength       =   80
      TabIndex        =   2
      Tag             =   "Saludos|T|S|||scartas|saludos||N|"
      Text            =   "Text1"
      Top             =   1575
      Width           =   9615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   6
      Left            =   150
      MaxLength       =   110
      MultiLine       =   -1  'True
      TabIndex        =   6
      Tag             =   "Despedida|T|S|||scartas|desped||N|"
      Text            =   "frmFacCartasOferta.frx":0114
      Top             =   5250
      Width           =   9615
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   3300
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "Descripci�n|T|S|||scartas|descarta||N|"
      Text            =   "Text1"
      Top             =   930
      Width           =   4275
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
      Left            =   7410
      TabIndex        =   7
      Top             =   5850
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
      Left            =   8655
      TabIndex        =   8
      Top             =   5850
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
      Left            =   8655
      TabIndex        =   9
      Top             =   5850
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   150
      TabIndex        =   12
      Top             =   5730
      Width           =   2655
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
         TabIndex        =   13
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   1050
      MaxLength       =   5
      TabIndex        =   0
      Tag             =   "Cod. Carta|N|N|0|999|scartas|codcarta|000|S|"
      Text            =   "Text1"
      Top             =   930
      Width           =   780
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3270
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   700
      Index           =   5
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "P�rrafo 3|T|S|||scartas|parrafo3||N|"
      Text            =   "frmFacCartasOferta.frx":011A
      Top             =   3930
      Visible         =   0   'False
      Width           =   8835
   End
   Begin VB.Label Label1 
      Caption         =   "Vista Preliminar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7680
      TabIndex        =   21
      Top             =   1020
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Documento Alta"
      Height          =   255
      Left            =   210
      TabIndex        =   20
      Top             =   90
      Width           =   1230
   End
   Begin VB.Image imgDoc 
      Height          =   375
      Index           =   1
      Left            =   9360
      ToolTipText     =   "Vista Preliminar"
      Top             =   960
      Width           =   390
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1290
      Picture         =   "frmFacCartasOferta.frx":0120
      Tag             =   "-1"
      ToolTipText     =   "Ver observaciones"
      Top             =   2700
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1290
      Tag             =   "-1"
      ToolTipText     =   "Ver observaciones"
      Top             =   1980
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   1290
      Picture         =   "frmFacCartasOferta.frx":0222
      Tag             =   "-1"
      ToolTipText     =   "Ver observaciones"
      Top             =   3690
      Width           =   240
   End
   Begin VB.Label Label22 
      Caption         =   "Parrafo 3"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   19
      Top             =   3720
      Width           =   795
   End
   Begin VB.Label Label22 
      Caption         =   "Saludos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   18
      Top             =   1305
      Width           =   840
   End
   Begin VB.Label Label22 
      Caption         =   "Despedida"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   150
      TabIndex        =   17
      Top             =   5010
      Width           =   1110
   End
   Begin VB.Label Label22 
      Caption         =   "Descripci�n"
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
      Index           =   6
      Left            =   2160
      TabIndex        =   16
      Top             =   960
      Width           =   1110
   End
   Begin VB.Label Label22 
      Caption         =   "Parrafo 2"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   15
      Top             =   2715
      Width           =   795
   End
   Begin VB.Label Label22 
      Caption         =   "Cuerpo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   14
      Top             =   1965
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "C�digo"
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
      Left            =   180
      TabIndex        =   11
      Top             =   960
      Width           =   1065
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacCartasOferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String

Private HaDevueltoDatos As Boolean

Private frmObs As frmFacClienteObser


Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'cadena con los parametros q se pasan a Crystal R.
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas




Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 1 'BUSQUEDA
        HacerBusqueda
    Case 3 'INSERTAR
        If DatosOk Then
            If InsertarDesdeForm(Me) Then PosicionarData
        End If
    Case 4 'MODIFICAR
        If DatosOk Then
             If ModificaDesdeFormulario(Me, 1) Then
                 TerminaBloquear
                 PosicionarData
             End If
         End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
        Case 3 'Insertar
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If

    Cad = Data1.Recordset.Fields(0) & "|"
    Cad = Cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

'    'ICONOS de La toolbar
'    btnPrimero = 13 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
'    With Toolbar1
'        .ImageList = frmPpal.imgListComun
'        'ASignamos botones
'        .Buttons(1).Image = 1   'Buscar
'        .Buttons(2).Image = 2 'Ver Todos
'        .Buttons(5).Image = 3 'A�adir
'        .Buttons(6).Image = 4 'Modificar
'        .Buttons(7).Image = 5 'Eliminar
'        .Buttons(10).Image = 15 'Salir
'        .Buttons(btnPrimero).Image = 6 'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
'    End With
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Eliminar
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun1
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
'    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With
    
    
    
    
    LimpiarCampos   'Limpia los campos TextBox
    
    Me.imgDoc(1).Picture = frmPpal.ImageListTPV.ListImages(8).Picture

    For I = 0 To Me.imgBuscar.Count - 1
        imgBuscar(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next



    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "scartas" 'Tabla Cartas Oferta
    Ordenacion = " ORDER BY codcarta"
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE codcarta = -1" 'No recupera datos
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    PonerModo 0
    Screen.MousePointer = vbDefault
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
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    Select Case Index
        Case 0
            If Modo = 3 Or Modo = 4 Then
                CadenaDesdeOtroForm = Text1(3).Text
            Else
                CadenaDesdeOtroForm = ""
                If Not Data1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Data1.Recordset!parrafo1, "T")
            End If
            
            Set frmObs = New frmFacClienteObser
            
            frmObs.Modificar = (Modo >= 3)
            frmObs.Text1 = CadenaDesdeOtroForm
            frmObs.Caption = "Cuerpo"
            frmObs.Show vbModal
            'Llevara DOS VALORES.
            'Si modifica y el texto
            If Modo = 3 Or Modo = 4 Then
                If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(3).Text = Mid(CadenaDesdeOtroForm, 3)
            End If
            CadenaDesdeOtroForm = ""
        
            Set frmObs = Nothing
        
        Case 1
            If Modo = 3 Or Modo = 4 Then
                CadenaDesdeOtroForm = Text1(4).Text
            Else
                CadenaDesdeOtroForm = ""
                If Not Data1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Data1.Recordset!parrafo2, "T")
            End If
            
            Set frmObs = New frmFacClienteObser
            
            frmObs.Modificar = (Modo >= 3)
            frmObs.Text1 = CadenaDesdeOtroForm
            frmObs.Caption = "P�rrafo 2"
            frmObs.Show vbModal
            'Llevara DOS VALORES.
            'Si modifica y el texto
            If Modo = 3 Or Modo = 4 Then
                If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(4).Text = Mid(CadenaDesdeOtroForm, 3)
            End If
            CadenaDesdeOtroForm = ""
        
            Set frmObs = Nothing
        
        Case 2
            If Modo = 3 Or Modo = 4 Then
                CadenaDesdeOtroForm = Text1(5).Text
            Else
                CadenaDesdeOtroForm = ""
                If Not Data1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Data1.Recordset!parrafo3, "T")
            End If
            
            Set frmObs = New frmFacClienteObser
            
            frmObs.Modificar = (Modo >= 3)
            frmObs.Text1 = CadenaDesdeOtroForm
            frmObs.Caption = "P�rrafo 3"
            frmObs.Show vbModal
            'Llevara DOS VALORES.
            'Si modifica y el texto
            If Modo = 3 Or Modo = 4 Then
                If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(5).Text = Mid(CadenaDesdeOtroForm, 3)
            End If
            CadenaDesdeOtroForm = ""
        
            Set frmObs = Nothing
        
    End Select

End Sub

Private Sub imgDoc_Click(Index As Integer)
    
    TerminaBloquear
    
    If Text1(1).Text = "" Then Exit Sub
    
    Select Case Index
        Case 1 'vista preliminar de la impresion del documento
            InicializarVbles

        
            'Parametro cod. carta
            cadParam = "|pCodCarta= " & Text1(0).Text & "|"
            numParam = numParam + 1
            
            'Nombre fichero .rpt a Imprimir
            nomRPT = "rFacSocioCarta.rpt" '"rComProveCarta.rpt"
            Titulo = "Cartas a Socios"
            conSubRPT = True
            
        LlamarImprimir False, False
            
    End Select
End Sub

Private Sub LlamarImprimir(PonerNombrePDF As Boolean, EnviaPorEmail As Boolean, Optional DatosEnvioMail As String)
     With frmImprimir
        .FormulaSeleccion = "{sclien.codclien} = 1"
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = EnviaPorEmail
        .Opcion = 190
        .Titulo = Titulo
        .NombreRPT = nomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With

End Sub



Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub


Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
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

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
  
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
      
    With Text1(Index)
        'C�digo de Carta
        If Index = 0 Then
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
        End If
    End With
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
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim NumReg As Byte
Dim I As Integer

    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I
    
    Modo = Kmodo
    PonerIndicador Me.lblIndicador, Modo
    
    '===========================================
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
        If Modo = 1 Then Me.lblIndicador.Caption = "BUSQUEDA"
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Me.Data1.Recordset.RecordCount > 1 ' Me.Toolbar1, btnPrimero, b, NumReg
       
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
     ' solo si tenemos registro cargado podemos imprimir documentos
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    Me.imgDoc(1).visible = b
    Me.imgDoc(1).Enabled = b
    Me.Label1.visible = b
    
    
         
    '==============================
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    chkVistaPrevia.Enabled = (Modo <= 2)

    '===============================
    PonerModoOpcionesMenu 'Activa las Opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b

    b = (Modo >= 3)
    'Insertar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnNuevo.Enabled = Not b
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue 'vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue 'vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    LimpiarCampos
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos 'Vac�a los TextBox
    ModoAnterior = Modo 'Para el bot�n Cancelar en Modo Insertar
    PonerModo 3
              
    Text1(0).Text = SugerirCodigoSiguienteStr(NombreTabla, "codcarta")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim Sql As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
       
    Sql = Sql & "�Desea Eliminar la Carta para Ofertas? " & vbCrLf
    Sql = Sql & vbCrLf & "C�digo : " & Format(Text1(0).Text, "000")
    Sql = Sql & vbCrLf & "Descripci�n : " & Text1(1).Text
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Carta de Oferta.", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
On Error GoTo FinEliminar
    
    If Data1.Recordset.EOF Then
        Eliminar = False
        Exit Function
    End If
    
    conn.Execute "Delete  from " & NombreTabla & ObtenerWhereCP
                      
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Eliminar = False
    Else
        Eliminar = True
    End If
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
On Error Resume Next

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    DatosOk = True
End Function


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: scapla
    Cad = Cad & ParaGrid(Text1(0), 20, "Cod. Carta")
    Cad = Cad & ParaGrid(Text1(1), 80, "Descripci�n")
    
    Tabla = NombreTabla
    Titulo = "Cartas de Oferta"
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri 'Conexion a BD Aritaxi
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
                PonerFoco Text1(kCampo)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Dim cadMen As String
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        cadMen = "No hay ning�n registro en la tabla " & NombreTabla
        If Modo = 1 Then
            MsgBox cadMen & " para ese criterio de B�squeda.", vbInformation
        Else
            MsgBox cadMen, vbInformation
        End If
        Screen.MousePointer = vbDefault
        PonerModo Modo
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
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
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = Mid(ObtenerWhereCP, 7)
    If SituarData(Data1, vWhere, Indicador) Then
        PonerModo 2
        Indicador = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP() As String
On Error Resume Next
    ObtenerWhereCP = " WHERE codcarta= " & Text1(0).Text
End Function

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub
