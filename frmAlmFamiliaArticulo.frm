VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmFamiliaArticulo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Familias de Art�culos"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frmAlmFamiliaArticulo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   8
      Left            =   2060
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   33
      Text            =   "Text2"
      Top             =   4200
      Width           =   3585
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   1365
      MaxLength       =   4
      TabIndex        =   9
      Tag             =   "Centro de coste|T|S|||sfamia|codccost||N|"
      Top             =   4200
      Width           =   630
   End
   Begin VB.Frame Frame3 
      Caption         =   "Compras "
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
      Left            =   120
      TabIndex        =   28
      Top             =   2940
      Width           =   7695
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Cta. Contable compras|T|N|||sfamia|ctacompr||N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Cta.Abono Compras|T|N|||sfamia|abocompr||N|"
         Text            =   "Text1"
         Top             =   675
         Width           =   1125
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   675
         Width           =   3885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   240
         Width           =   3885
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Abono Compras"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   32
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Contable Compras"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   31
         Top             =   270
         Width           =   1815
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   2040
         Picture         =   "frmAlmFamiliaArticulo.frx":000C
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   315
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   3
         Left            =   2040
         Picture         =   "frmAlmFamiliaArticulo.frx":010E
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ventas "
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
      Height          =   1935
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   7695
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Cta. Contable Ventas|T|N|||sfamia|ctaventa||N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Cta. Abono Ventas|T|N|||sfamia|aboventa||N|"
         Text            =   "Text1"
         Top             =   675
         Width           =   1125
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   675
         Width           =   3885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   240
         Width           =   3885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   6
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   1080
         Width           =   3885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   7
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   1485
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Cta. Alternativa Abonos|T|N|||sfamia|abovent1||N|"
         Text            =   "Text1"
         Top             =   1485
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Cta. Alternativa Ventas|T|N|||sfamia|ctavent1||N|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Abono Ventas"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   27
         Top             =   675
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Contable Ventas"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   26
         Top             =   270
         Width           =   1575
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "frmAlmFamiliaArticulo.frx":0210
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   285
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   2040
         Picture         =   "frmAlmFamiliaArticulo.frx":0312
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   705
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   5
         Left            =   2040
         Picture         =   "frmAlmFamiliaArticulo.frx":0414
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   4
         Left            =   2040
         Picture         =   "frmAlmFamiliaArticulo.frx":0516
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Alternativa Ventas"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   1110
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Alternativa Abonos"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   24
         Top             =   1515
         Width           =   1815
      End
   End
   Begin VB.CheckBox chkInstalac 
      Caption         =   "�Es instalaci�n?"
      Height          =   195
      Left            =   6240
      TabIndex        =   2
      Tag             =   "�Es instalaci�n?|N|N|||sfamia|instalac||N|"
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6690
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   2640
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "Denominaci�n familia de Art�culo|T|N|||sfamia|nomfamia||N|"
      Text            =   "Text1"
      Top             =   600
      Width           =   3285
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   600
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "C�digo familia de art�culo|N|N|0|9999|sfamia|codfamia|0000|S|"
      Text            =   "Text"
      Top             =   600
      Width           =   630
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   150
      TabIndex        =   13
      Top             =   4635
      Width           =   2655
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6690
      TabIndex        =   12
      Top             =   4800
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   4800
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   270
      Top             =   4635
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8010
      _ExtentX        =   14129
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
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
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
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5880
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1005
      Picture         =   "frmAlmFamiliaArticulo.frx":0618
      ToolTipText     =   "Buscar centro coste"
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "CCoste"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   34
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Denominaci�n"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cod."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   375
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
Attribute VB_Name = "frmAlmFamiliaArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
Private ModoAnterior As Byte

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean


Private Sub chkInstalac_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkInstalac_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    TratarCtaContable
                    PosicionarData
                End If
            End If
        
        Case 4  'MODIFICAR
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
    Select Case Modo
        Case 1 'Busqueda
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
    End Select
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    'A�adiremos el boton de aceptar y demas objetos para insertar
    ModoAnterior = Modo 'Para el bot�n cancelar en Modo Insertar
    PonerModo 3
    Text1(0).Text = SugerirCodigoSiguienteStr("sfamia", "codfamia")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else 'Modo=1 Busqueda
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
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
        CadenaConsulta = "Select * from " & NombreTabla
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub BotonModificar()
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    
    
    '### a mano
    cad = "�Seguro que desea eliminar la Familia de Art�culo?:" & vbCrLf
    cad = cad & vbCrLf & "Cod. : " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "Desc.: " & Data1.Recordset.Fields(1)

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Data1.Recordset.Delete
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
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Familia de Articulo", Err.Description
    End If
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = Data1.Recordset.Fields(0) & "|"
    cad = cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    If vParamAplic.Descriptores Then Me.Caption = "Categorias Art."
    ' ICONITOS DE LA BARRA
    btnPrimero = 14 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Bot�n Buscar
        .Buttons(2).Image = 2   'Bot�n Todos
        
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        .Buttons(10).Image = 16  ' Imprimir
        .Buttons(11).Image = 15  'Salir
        
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    
    'Pone el Tag del primer bot�n de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sfamia, BD: Aritaxi
    'Si tag>0 abre busqueda en la tabla: Cuentas, BD: Conta
    imgCuentas(0).Tag = "-1"
    Me.imgBuscar(0).Tag = "-1"
        
  
    '## A mano
    NombreTabla = "sfamia"
    Ordenacion = " ORDER BY codfamia"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
       
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE codfamia=-1"
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkInstalac.Value = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
Dim Indice As Byte
    
    If CadenaDevuelta <> "" Then
        If Val(imgCuentas(0).Tag) >= 0 Then
            'Se llama desde un bot�n de busqueda de Cuentas
            'Recuperar solo el campo c�digo y Descripci�n
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            Indice = Val(Me.imgCuentas(0).Tag)
            Text1(Indice + 2).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(Indice + 2).Text = RecuperaValor(CadenaDevuelta, 2)
            
        ElseIf Val(imgBuscar(0).Tag) >= 0 Then
            'Centro de coste
            Text1(8).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(8).Text = RecuperaValor(CadenaDevuelta, 2)
        Else
            'Recupera todo el registro de Banco Propio
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            '   Como la clave principal es unica, con poner el sql apuntando
            '   al valor devuelto sobre la clave ppal es suficiente
            'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            'If CadB <> "" Then CadB = CadB & " AND "
            'CadB = CadB & Aux
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub


Private Sub imgBuscar_Click(Index As Integer)
    ' ---- [19/10/2009] [LAURA]: a�adir campo centro de coste familia
    
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    
    Select Case Index
        Case 0 'Centros de coste de la conta
            Screen.MousePointer = vbHourglass
            Me.imgBuscar(0).Tag = Index
            Set frmB = New frmBuscaGrid
            frmB.vCampos = "Codigo|cabccost|codccost|T||20�Descripci�n|cabccost|nomccost|T||70�"
            frmB.vTabla = "cabccost"
            frmB.vSQL = ""
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vTitulo = "Centros de coste"
            frmB.vselElem = 0
            frmB.vConexionGrid = conConta
            
            frmB.Show vbModal
            Set frmB = Nothing
            imgBuscar(0).Tag = -1
            Screen.MousePointer = vbDefault
            PonerFoco Text1(8)
    End Select
End Sub


Private Sub imgCuentas_Click(Index As Integer)

    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    imgCuentas(0).Tag = Index
    MandaBusquedaPrevia "apudirec='S'"
    imgCuentas(0).Tag = -1
    PonerFoco Text1(Index + 2)
    Screen.MousePointer = vbDefault
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
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Codigo familia
'            If Text1(Index).Text <> "" Then
             If PonerFormatoEntero(Text1(Index)) Then
                'Comprobar si ya existe el cod de familia en la tabla
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
        '#### lo hemos puesto en el evento VALIDATE
'         Case 2, 3, 4, 5 'Cuentas
'            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo)
        '####
        
        ' ---- [19/10/2009] [LAURA]: a�adir campo centro de coste familia
        Case 8: Me.Text2(Index).Text = PonerNombreCCoste(Me.Text1(Index))
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
    
    cadB = ObtenerBusqueda(Me, False)

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim CargaF As Boolean 'Para saber si se carga el frame o no en el BuscaGrid
Dim Conexion As Byte

        'Llamamos a al form
        '##A mano
        cad = ""
        If Val(Me.imgCuentas(0).Tag) >= 0 Then
        'Se llama a Busqueda desde un campo de Cuenta
            '#A MANO: Porque busca en la tabla Cuentas
            'de la base de datos de Contabilidad
            cad = cad & "C�digo|Cuentas|codmacta|T||15�Denominacion|Cuentas|nommacta|T||70�"
            tabla = "Cuentas"
            Titulo = "Cuentas"
            Conexion = conConta    'Conexi�n a BD: Conta
            CargaF = True 'Se puede cargar el frame
        Else
            'Busqueda de una Fam�lia de Art�culo
            cad = cad & ParaGrid(Text1(0), 15, "C�digo")
            cad = cad & ParaGrid(Text1(1), 80, "Denominacion")
            tabla = "sfamia"
            Titulo = "Fam�lia de Art�culos"
            If vParamAplic.Descriptores Then Titulo = "Categorias Art."
            Conexion = conAri    'Conexi�n a BD: Aritaxi
            CargaF = False 'No se carga el frame
        End If
        
        If cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = cad
            frmB.vTabla = tabla
            frmB.vSQL = cadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vTitulo = Titulo
            frmB.vselElem = 1
            frmB.vConexionGrid = Conexion
            frmB.vCargaFrame = CargaF
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If kCampo < 5 Then PonerFoco Text1(kCampo + 1)
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then
                    If Not (Val(Me.imgCuentas(0).Tag) >= 0) Then cmdRegresar_Click
                End If
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco Text1(kCampo)
'                If Modo = 1 Then
'                    MsgBox "No hay ning�n registro en la tabla " & tabla
'                    PonerFoco Text1(0)
'                End If
            End If
        End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ning�n registro en la tabla " & NombreTabla & " para ese criterio de B�squeda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
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
Dim i As Byte
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'poner la descripcion de las cuentas
    For i = 2 To 7
        Text2(i).Text = PonerNombreCuenta(Text1(i), Modo)
    Next i
        
    ' ---- [19/10/2009] [LAURA]: a�adir campo centro de coste familia
    Me.Text2(8).Text = PonerNombreCCoste(Me.Text1(8))
        
        
    BloquearChecks Me, Modo
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    '-------------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    'Poner Boton de Cabecera o Aceptar/Cancelar
    PonerBotonCabecera B Or (Modo = 0)
        
    'Bloquear Registros si modo distinto de Insertar o Modificar
    BloquearText1 Me, Modo
    
    BloquearChecks Me, Modo
        
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n MODO
    PonerOpcionesMenu   'Activar opciones de menu seg�n NIVEL
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
On Error Resume Next

    B = (Modo = 2) Or (Modo = 0) Or (Modo = 1)
    'A�adir
    Toolbar1.Buttons(5).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    mnEliminar.Enabled = B
    
     '---------------------------------
    B = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean

    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function
    
    'Comprobar si ya existe el cod de familia en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then B = False
    End If
    
    DatosOk = B
End Function




Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    If Modo = 3 Or Modo = 4 Then
        Select Case Index
            Case 2, 3, 4, 5, 6, 7 'Cuentas
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(1).Text)
                If Text1(Index).Text <> "" And Text2(Index).Text = "" Then Cancel = True
        End Select
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnBuscar_Click
        Case 2: mnVerTodos_Click
        Case 5  'Nuevo
                mnNuevo_Click
        Case 6  'Modificar
                mnModificar_Click
        Case 7  'Borrar
                mnEliminar_Click
        Case 10 'Imprimir listado
            BotonImprimir
        Case 11: mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(B As Boolean)
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    If B Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(codfamia=" & Text1(0).Text & ")"
    If SituarData(Data1, cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
   
    cadFormula = ""
    cadParam = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 5
        .Titulo = "Listado Familia de Art�culos"
        .NombreRPT = "rAlmFamArtic.rpt"  'Nombre fichero .rpt a Imprimir
        .Show vbModal
    End With
End Sub


Private Sub TratarCtaContable()
Dim i As Integer
Dim CtaCreadas As String
    For i = 2 To 7
        If Text2(i).Text = vbCrearNuevaCta Then
            If InStr(1, CtaCreadas, Text1(i).Text & "|") = 0 Then
                InsertarCuentaCble Text1(i).Text, "", "", Text1(1).Text
                CtaCreadas = CtaCreadas & Text1(i).Text & "|"
            End If
            Text2(i).Text = Text1(1).Text
        End If
    Next i
End Sub
