VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacTarifasPrecios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tarifas de Art�culos"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   8460
   ClipControls    =   0   'False
   Icon            =   "frmFacTarifasPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameNuevo 
      Caption         =   "Valores Nuevos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1095
      Left            =   3120
      TabIndex        =   28
      Top             =   1440
      Width           =   4695
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Fecha Cambio|F|S|||slista|fechanue|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   680
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   6
         Tag             =   "Precio Caja Nuevo|N|S|0|999999.0000|slista|precion1|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   680
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   5
         Tag             =   "Precio Nuevo|N|S|0|999999.0000|slista|precionu|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   290
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Cambio"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   450
         Width           =   1080
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   4125
         ToolTipText     =   "Buscar fecha"
         Top             =   405
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Precio Caja"
         Height          =   255
         Left            =   300
         TabIndex        =   30
         Top             =   680
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Precio"
         Height          =   255
         Left            =   300
         TabIndex        =   29
         Top             =   290
         Width           =   615
      End
   End
   Begin VB.Frame FrameActuales 
      Caption         =   "Valores Actuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1095
      Left            =   480
      TabIndex        =   25
      Top             =   1440
      Width           =   2535
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1020
         MaxLength       =   12
         TabIndex        =   4
         Tag             =   "Precio Caja Actual|N|S|0|999999.0000|slista|precioa1|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   680
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1020
         MaxLength       =   12
         TabIndex        =   3
         Tag             =   "Precio Actual|N|N|0|999999.0000|slista|precioac|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   290
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Precio Caja"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   680
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Precio"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   290
         Width           =   735
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   1
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   22
      Text            =   "Fecha Cambio"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   21
      Text            =   "NumLinea"
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkPermiteDto 
      Caption         =   "Permite Descuento"
      Height          =   195
      Left            =   5880
      TabIndex        =   2
      Tag             =   "Permite Descuento|N|N|||slista|dtopermi||N|"
      Top             =   1030
      Width           =   1695
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   5280
      MaxLength       =   12
      TabIndex        =   24
      Text            =   "Precio Caja"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   3600
      MaxLength       =   12
      TabIndex        =   23
      Text            =   "Precio Unidad"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   5615
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6795
      TabIndex        =   9
      Top             =   5615
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6795
      TabIndex        =   10
      Top             =   5615
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   480
      TabIndex        =   19
      Top             =   5450
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
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   630
      Width           =   4005
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   975
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "Cod. Tarifa|N|N|0|999|slista|codlista|000|S|"
      Text            =   "Text1"
      Top             =   975
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1800
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "Cod. Articulo|T|N|||slista|codartic||S|"
      Text            =   "Text1"
      Top             =   600
      Width           =   1710
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
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
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   5880
         TabIndex        =   18
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3960
      Top             =   5640
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   2760
      Top             =   5640
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacTarifasPrecios.frx":000C
      Height          =   2625
      Left            =   480
      TabIndex        =   11
      Top             =   2685
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4630
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
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1500
      ToolTipText     =   "Buscar tarifa"
      Top             =   1005
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1500
      ToolTipText     =   "Buscar art�culo"
      Top             =   645
      Width           =   240
   End
   Begin VB.Label Label5 
      Caption         =   "C�d. Tarifa"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   975
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Cod. Articulo"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   630
      Width           =   975
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacTarifasPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticulos 'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmT As frmFacTarifas  'Form Mantenimiento Tarifas
Attribute frmT.VB_VarHelpID = -1

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer


Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String

Private HaDevueltoDatos As Boolean


Private Sub chkPermiteDto_GotFocus()
'    chkPermiteDto.TabIndex = 2
    ConseguirfocoChk Modo
End Sub

Private Sub chkPermiteDto_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkPermiteDto_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then PonerFoco Text1(2) 'ENTER
    KEYpress KeyAscii
End Sub

Private Sub chkPermiteDto_LostFocus()
'    PonerFoco Text1(5)
'    chkPermiteDto.TabIndex = 20
End Sub

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'Busqueda
            HacerBusqueda
        Case 3 'Insertar
            If DatosOk Then
                If InsertarDesdeForm(Me) Then PosicionarData
            End If
            
        Case 4 'Modificar
           If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    
                    'si se ha modificado el precio actual
                    If CCur(Me.data1.Recordset!precioac) <> CCur(Me.Text1(2).Text) Then
                        'si la tarifa modificada coincide con la de parametros aplicacion
                        If CInt(Me.data1.Recordset!codlista) = vParamAplic.CodTarifa Then
                            'actualizar el PVP del articulo a partir del precio tarifa nuevo
                            'y quitandole el margen de la tarifa correspondiente
                            '---------------------------------------------------------
                            '- bloquear el articulo a modificar
                            If BloquearArticulo Then
                                Screen.MousePointer = vbDefault
                                MsgBox "Se va a actualizar el PVP del art�culo.", vbInformation
                                '- actualizar su PVP
                                Screen.MousePointer = vbHourglass
                                ActualizarPVPArticulo
                                TerminaBloquear
                            End If
                        End If
                    End If
                    
                    PosicionarData
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
    
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
                'PonerOpcionesMenu
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

'Private Sub cmdRegresar_Click()
''Este es el boton Cabecera
'Dim cad As String
'Dim Indicador As String
'
'    'Quitar lineas y volver a la cabecera
'    If Modo = 5 Then 'modo 5: Lineas Articulos x Almacen
'        DataGrid1.ClearFields
'        cad = "(codmovim=" & Val(Text1(0).Text) & ")"
'        If SituarData(Data1, cad, Indicador) Then
'            PonerModo 2
'            lblIndicador.Caption = Indicador
'            Me.Toolbar1.Buttons(9).Enabled = True
'            Me.Toolbar1.Buttons(10).Enabled = True
'        End If
'    End If
'End Sub



Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    
    For kCampo = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    
    
    'La toolbar
    btnPrimero = 14 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'A�adir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(11).Image = 15 'Salir
        .Buttons(14).Image = 6 'Primero
        .Buttons(15).Image = 7 'Anterior
        .Buttons(16).Image = 8 'Siguiente
        .Buttons(17).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "slista" 'Tabla Tarifas Precios de Articulos
    Ordenacion = " ORDER BY codartic"
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE codartic = -1" 'No recupera datos"

    data1.ConnectionString = conn
    data1.RecordSource = CadenaConsulta
    data1.Refresh
    
    PonerModo 0
    CargaGrid (Modo = 2)
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim Sql As String

    On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, Sql, False

    DataGrid1.Columns(0).visible = False 'Cod. Articulo
    DataGrid1.Columns(1).visible = False 'Cod. Lista
    i = 2
    
    'Numero Linea
    DataGrid1.Columns(i).Caption = "Num. Linea"
    DataGrid1.Columns(i).Width = 1200
    
    'Fecha Cambio
    DataGrid1.Columns(i + 1).Caption = "Fecha Cambio"
    DataGrid1.Columns(i + 1).Width = 1500
    
    'Precio Unidad
    DataGrid1.Columns(i + 2).Caption = "Precio Unidad"
    DataGrid1.Columns(i + 2).Width = 2000
    DataGrid1.Columns(i + 2).Alignment = dbgRight
    DataGrid1.Columns(i + 2).NumberFormat = FormatoPrecio & " "
    
    'Precio Caja
    DataGrid1.Columns(i + 3).Caption = "Precio Caja"
    DataGrid1.Columns(i + 3).Width = 2000
    DataGrid1.Columns(i + 3).Alignment = dbgRight
    DataGrid1.Columns(i + 3).NumberFormat = FormatoPrecio & " "
    
    DataGrid1.ScrollBars = dbgAutomatic
       
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
    Next i
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
    Exit Sub
    
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Articulos
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
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
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    Text1(4).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Tarifas
    Text1(1).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0 'Codigo Articulo
            Set frmA = New frmAlmArticulos
            frmA.DatosADevolverBusqueda2 = "@1@" 'Abre en modo Busqueda
            frmA.Show vbModal
            Set frmA = Nothing
        Case 1  'Cod. Tarifa
            Set frmT = New frmFacTarifas
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
    End Select
    PonerFoco Text1(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   indice = 4
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
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
'On Error Resume Next
'
'    If KeyAscii = 13 And Index = 4 Then 'ENTER
'        If modo = 1 Or modo = 3 Or modo = 4 Then 'Modo Insertar o Modificar
''            cmdAceptar.SetFocus
'            PonerFocoBtn Me.cmdAceptar
'        End If
'    Else
'        KEYpress KeyAscii
'    End If
'    If Err.Number <> 0 Then Err.Clear
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim Sql As String
Dim devuelve As String
Dim campo As String, TipoCampo As String
Dim Codigo As String
Dim Tabla As String
Dim Titulo As String
Dim index2 As Integer 'Indice del control Text2
Dim indImgP As Integer 'Indice del control ImgBuscar (Prismaticos)

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    Select Case Index
        Case 0 'Codigo Articulo
            campo = "nomartic"
            TipoCampo = "T"
            Tabla = "sartic"
            Codigo = "codartic"
            Titulo = "Articulos"
            indImgP = Index 'Imagen Prismaticos
            index2 = Index  'Indice para Text2
            
        Case 1 'Codigo Tarifa
            campo = "nomlista"
            TipoCampo = "N"
            Tabla = "starif"
            Codigo = "codlista"
            Titulo = "Tarifas"
            indImgP = Index
            index2 = Index
            Text1(Index).Text = Format(Text1(Index).Text, "000")
                
        Case 2, 3, 5, 6 'Precios Actuales y Nuevos
            'Formato tipo 2: Decimal(10,4)
            If Text1(Index).Text <> "" And Modo <> 1 Then
                If PonerFormatoDecimal(Text1(Index), 2) And (Index = 5 Or Index = 6) Then
                    'fecha de cambio
                    If Text1(4).Text = "" Then Text1(4).Text = Format(Now, "dd/mm/yyyy")
                End If
            End If
            
        Case 4 'Fecha
            If Modo <> 1 And Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            PonerFocoBtn Me.cmdAceptar
    End Select

    If (Index = 0 Or Index = 1) And Modo <> 1 Then
    'Indices 0 al 1 del textbox Text1 son busquedas a otras tablas
    'Comprobar los cambios para actualizar la descripcion de Text2
        If Text1(Index).Text = "" Then
             Text2(index2).Text = ""
             Exit Sub
        Else
            'BD 1: conexion a BD Aritaxi
            If TipoCampo = "N" Then
                If Not PonerFormatoEntero(Text1(Index)) Then
                    Text2(index2).Text = ""
                    Exit Sub
                End If
            End If
            Sql = DevuelveDesdeBD(conAri, campo, Tabla, Codigo, Text1(Index).Text, TipoCampo)
            Text2(index2).Text = Sql
            If Sql = "" Then 'No existe
                devuelve = "No existe el c�digo de " & Titulo & vbCrLf
                devuelve = devuelve & "C�digo: " & Text1(Index).Text
                MsgBox devuelve, vbExclamation
                PonerFoco Text1(Index)
            End If
        End If
    End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            BotonBuscar
        Case 2 'Ver Todos
            BotonVerTodos
        Case 5 'Nuevo
                BotonAnyadir
        Case 6  'Modificar
                If BLOQUEADesdeFormulario(Me) Then BotonModificar
        Case 7 'Eliminar
                BotonEliminar
        Case 10 'Imprimir
'            BotonImprimir
            AbrirListado (28) '28: Informe Tarifas de Articulos
        Case 11  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
    If KeyAscii = 27 And Modo = 1 Then cmdCancelar_Click 'busqueda
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo

    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    For i = 0 To txtAux.Count - 1
        Text1(i).BackColor = vbWhite
    Next i
    
    
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    
     'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not data1.Recordset.EOF Then
        If data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg

    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
    BloquearChecks Me, Modo
           
    '--------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b And Modo <> 4 'Si modificar no activado pq son claves ajenas
    Next i
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = b 'And Modo <> 3 'Si es insertar se modifica
    Next i
    
'    Me.chkPermiteDto.Enabled = (Modo = 1) Or (Modo = 3) Or (Modo = 4)
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Modo", Err.Description
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    '-------------------------------------
    b = (Modo >= 3)
    'Insertar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnnuevo.Enabled = Not b
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnvertodos.Enabled = Not b
End Sub

Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkPermiteDto.Value = 0
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData data1, Index
    PonerCampos
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
Dim Tabla As String
    
    Tabla = "slist1"

    Sql = "SELECT * FROM " & Tabla
    If enlaza Then
        Sql = Sql & " WHERE codartic=" & DBSet(data1.Recordset!codArtic, "T") & " AND codlista=" & data1.Recordset!codlista
    Else
        Sql = Sql & " WHERE codartic = -1"
    End If
    Sql = Sql & " ORDER BY " & Tabla & ".numlinea desc"
    MontaSQLCarga = Sql
End Function


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False

        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Vac�a los TextBox
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    ModoAnterior = Modo 'Para el bot�n Cancelar en Modo Insertar
    PonerModo 3
    
    'Para que si no se ha cargado el Data1 inicialmente, tenga valor cuando situamos el Data
'    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
'    Data1.RecordSource = CadenaConsulta
           
    'Ponemos el grid de lineas enlazando a ningun sitio
    CargaGrid False
    
    PonerFoco Text1(0)
    
End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1 y campo2 es clave primaria, NO se puede modificar
'    If Text1(4).Text = "" Then Text1(4).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(2)
End Sub


Private Sub BotonEliminar()
Dim Sql As String

    'Ciertas comprobaciones
    If data1.Recordset.EOF Then Exit Sub
    
    Sql = "Tarifas de Precios." & vbCrLf
    Sql = Sql & "---------------------------" & vbCrLf & vbCrLf
    
    Sql = Sql & "Va a Eliminar la Tarifa del Articulo:"
    Sql = Sql & vbCrLf & "Cod. Artic :  " & Text1(0).Text
    Sql = Sql & vbCrLf & "Cod. Tarif :  " & Text1(1).Text
    Sql = Sql & vbCrLf & vbCrLf & "�Desea continuar ? "
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub
        NumRegElim = data1.Recordset.AbsolutePosition
        DataGrid1.Enabled = False
        If SituarDataTrasEliminar(data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            CargaGrid False
            PonerModo 0
        End If
    End If
    
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MuestraError Err.Number, "Eliminar Tarifa de Articulo", Err.Description
            'MsgBox Err.Number & " - " & Err.Description, vbExclamation
            data1.Recordset.CancelUpdate
        End If
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String
On Error GoTo FinEliminar
        
        conn.BeginTrans
        Sql = " WHERE  codartic=" & DBSet(data1.Recordset!codArtic, "T") & ""
        Sql = Sql & " AND codlista=" & Val(data1.Recordset!codlista)
        
        'Lineas
        conn.Execute "Delete  from slist1 " & Sql
        
        'Cabeceras
        conn.Execute "Delete  from slista " & Sql
                      
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    'Comprobar que si hay valores nuevos, la fecha de cambio no es nulo
    If Not EsVacio(Text1(5)) Or Not EsVacio(Text1(6)) Then
        b = (Not EsVacio(Text1(4)))
    End If
    If Not b Then
        MsgBox "La Fecha de Cambio debe tener valor.", vbInformation
        Exit Function
    End If
    
    'Comprobar que si no hay valores nuevos no haya fecha de Cambio
    If EsVacio(Text1(5)) And EsVacio(Text1(6)) Then
        b = (EsVacio(Text1(4)))
    End If
    If Not b Then
        MsgBox "No hay valores nuevos para la fecha de cambio", vbInformation
        Exit Function
    End If
    
    'si se modifica el precio actual no hay fecha de cambio ni precio nuevo
    If Modo = 4 Then
        If CCur(Me.data1.Recordset!precioac) <> CCur(Me.Text1(2).Text) Then
            b = EsVacio(Text1(5)) And EsVacio(Text1(6)) And EsVacio(Text1(4))
            If Not b Then
                MsgBox "Si se modifican precios actuales, los precios nuevos y fecha cambio no deben tener valor.", vbInformation
                Exit Function
            End If
        End If
    End If
    
    
    DatosOk = True
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    cad = cad & ParaGrid(Text1(0), 14, "Cod Artic.")
    cad = cad & "Desc. Artic|sartic|nomartic|T||45�"
    cad = cad & ParaGrid(Text1(1), 10, "Cod Tarifa")
    cad = cad & "Desc. Tarifa|starif|nomlista|T||30�"
    
    Tabla = "(" & NombreTabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic" & ")"
    Tabla = Tabla & " LEFT JOIN starif ON " & NombreTabla & ".codlista=starif.codlista"
       
'    tabla = "slista"
    Titulo = "Tarifas de Art�culos"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|2|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
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
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub



Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    data1.RecordSource = CadenaConsulta
    data1.Refresh
    If data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        PonerModo Modo
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
        data1.Recordset.MoveFirst
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

    If data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, data1
    'Poner el nombre del cod. Articulo
    Text2(0).Text = PonerNombreDeCod(Text1(0), 1, "sartic", "nomartic")
    'Poner el nombre del cod. tarifa
    Text2(1).Text = PonerNombreDeCod(Text1(1), 1, "starif", "nomlista")
    
    'Si los campos de precios nuevos son cero mostrar cadena vacia
    If Text1(5).Text <> "" Then
        If Text1(5).Text = 0 Then Text1(5).Text = ""
    End If
    If Text1(6).Text <> "" Then
        If Text1(6).Text = 0 Then Text1(6).Text = ""
    End If
    
    BloquearChecks Me, Modo
    
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub BotonActualizar()
'Actualizar Tarifas de Articulos
Dim Sql As String

    If data1.Recordset.EOF Then
        MsgBox "Ning�na Tarifa para actualizar.", vbExclamation
        Exit Sub
    End If
    
    If Data2 Is Nothing Then Exit Sub
   
    Sql = "Actualizaci�n Tarifas de Art�culos." & vbCrLf
    Sql = Sql & "-----------------------------------------" & vbCrLf & vbCrLf
    
    Sql = Sql & "Va a Actualizar la Tarifa del Articulo:"
    Sql = Sql & vbCrLf & " Cod. Artic. :  " & data1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & " Cod. Tarifa :  " & CStr(Format(data1.Recordset.Fields(1), "000"))
    Sql = Sql & vbCrLf & vbCrLf & " �Desea continuar ? "
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then
        Exit Sub
    End If
    
    NumRegElim = data1.Recordset.AbsolutePosition
    If ActualizarTarifa Then
        SituarDataTrasEliminar data1, NumRegElim
    End If
End Sub


Private Function ActualizarTarifa() As Boolean
Dim Donde As String
Dim Sql As String
Dim bol As Boolean
On Error GoTo EActualizarTarifa
    
   
    'Aqui empieza transaccion
    conn.BeginTrans
    bol = ActualizarLaTarifa(Donde)

EActualizarTarifa:
        If Err.Number <> 0 Then
            Sql = "Actualizar Tarifa." & vbCrLf & "----------------------------" & vbCrLf
            Sql = Sql & Donde
'            If OpcionActualizar = 1 Then
                MuestraError Err.Number, Sql, Err.Description
'            Else
'                SQL = Donde & " -> " & Err.Description
'                SQL = Mid(SQL, 1, 200)
'                InsertaError SQL
'            End If
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            ActualizarTarifa = True
        Else
            conn.RollbackTrans
        End If
End Function


Private Function ActualizarLaTarifa(ByRef ADonde As String) As Boolean

    ActualizarLaTarifa = False
    
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Historico lineas Tarifas de Almacen"
    If Not InsertarLineasHistorico Then Exit Function
    
    'Modificamos en cabeceras de Tarifas
    ADonde = "Modificando datos en cabecera de Tarifas Precios"
    If Not ModificarCabecera Then Exit Function
    
    ActualizarLaTarifa = True
End Function


Private Function ModificarCabecera() As Boolean
'Modifica la tabla de cabeceras de Tarifas
Dim Sql As String
On Error Resume Next

    Sql = "UPDATE slista SET precioac=precionu, precioa1=precion1, fechanue=null, precionu=0, precion1=0"
    Sql = Sql & " WHERE codartic=" & DBSet(data1.Recordset!codArtic, "T") & " AND codlista=" & data1.Recordset!codlista
   
    conn.Execute Sql
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        ModificarCabecera = False
    Else
        ModificarCabecera = True
    End If
End Function


Private Function InsertarLineasHistorico() As Boolean
Dim Sql As String
Dim numF As String
On Error Resume Next

    'Obtenemos la siguiente numero de linea de tarifa
    Sql = "codartic=" & DBSet(data1.Recordset!codArtic, "T") & " AND codlista=" & data1.Recordset!codlista
    numF = SugerirCodigoSiguienteStr("slist1", "numlinea", Sql)

    Sql = "INSERT INTO slist1 (codartic, codlista, numlinea, fechacam, precioac, precioa1)"
    Sql = Sql & " VALUES (" & DBSet(data1.Recordset.Fields(0).Value, "T") & ", " & data1.Recordset.Fields(1).Value & ", "
    Sql = Sql & numF & ", " & DBSet(Text1(4).Text, "F") & ", "
    Sql = Sql & DBSet(data1.Recordset!precioac, "N") & ", " & DBSet(data1.Recordset!precioa1, "N") & ") "
    conn.Execute Sql
        
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarLineasHistorico = False
    Else
        InsertarLineasHistorico = True
    End If
End Function


Private Sub BotonImprimir()

        frmListado.NumCod = Text1(0).Text
        AbrirListado (8) '8: Informe Movimientos Almacen
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = "(codartic=" & DBSet(Text1(0).Text, "T") & " AND codlista=" & Text1(1).Text & ")"
    If SituarDataMULTI(data1, vWhere, Indicador) Then
'    If SituarData2(Data1, Text1(0).Text, Text1(1).Text, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
'                        LimpiarCampos
        PonerModo 0
    End If
End Sub



Private Function ActualizarPVPArticulo() As Boolean
Dim cTar As CTarifaArt
Dim margen As Currency
Dim newPrecio As Currency
Dim Sql As String
    
    On Error GoTo ErrActPVP
    
    Set cTar = New CTarifaArt
    If cTar.LeerDatos(Text1(0).Text, CInt(Text1(1).Text)) Then
        margen = Round2(cTar.MargenComercial / 100, 4)
        newPrecio = Round2((cTar.PrecioActual / (margen + 1)), 4)
        If newPrecio > 0 Then
            Sql = "UPDATE sartic SET preciove=" & DBSet(newPrecio, "N")
            Sql = Sql & " WHERE codartic=" & DBSet(Text1(0).Text, "T")
            conn.Execute Sql
        End If
    End If
    Set cTar = Nothing
    Exit Function
    
ErrActPVP:
    MuestraError Err.Number, "Actualizar precio PVP del articulo.", Err.Description
End Function


Private Function BloquearArticulo() As Boolean
Dim cadWHERE As String
    cadWHERE = "codartic=" & DBSet(Text1(0), "T")
    BloquearArticulo = BloqueaRegistro("sartic", cadWHERE)
End Function


