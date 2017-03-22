VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmConfTipoMov 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Movimiento"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7650
   Icon            =   "frmConfTipoMov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   6000
      MaxLength       =   1
      TabIndex        =   6
      Tag             =   "Tipo de Documento|N|S|0|9|stipom|tipodocu|0|N|"
      Text            =   "C"
      Top             =   4440
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   240
      TabIndex        =   10
      Top             =   5140
      Width           =   2355
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
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1920
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   4920
      MaxLength       =   3
      TabIndex        =   5
      Tag             =   "Letra Serie|T|S|||stipom|letraser||N|"
      Text            =   "L"
      Top             =   4440
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   3480
      MaxLength       =   7
      TabIndex        =   4
      Tag             =   "Contador|N|N|0|9999999|stipom|contador|0000000|N|"
      Text            =   "contado"
      Top             =   4440
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Denominaci�n|T|N|||stipom|nomtipom||N|"
      Text            =   "Descripcion"
      Top             =   4440
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   240
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "C�digo Tipo de Movimiento|T|N|||stipom|codtipom||S|"
      Text            =   "Cod"
      Top             =   4440
      Width           =   800
   End
   Begin VB.ComboBox CboMueveStock 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmConfTipoMov.frx":000C
      Left            =   2520
      List            =   "frmConfTipoMov.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Mueve Stock|N|N|||stipom|muevesto||N|"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   5265
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6195
      TabIndex        =   8
      Top             =   5265
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insertar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2760
      Top             =   5280
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
      Height          =   4410
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   7100
      _ExtentX        =   12515
      _ExtentY        =   7779
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfTipoMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos

Dim Modo As Byte
Dim PrimeraVez As Boolean


Private Sub CboMueveStock_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim i As String
On Error Resume Next

    Select Case Modo
        Case 1 'Insertar
            If DatosOk Then
               If InsertarDesdeForm(Me) Then
                  CargaGrid
                  BotonAnyadir
                End If
            End If
        Case 2  'Modificar
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 3) Then
                    TerminaBloquear
                    i = Data1.Recordset.Fields(0)
                    PonerModo 0
                    CancelaADODC Me.Data1
                    CargaGrid
                    Data1.Recordset.Find (Data1.Recordset.Fields(0).Name & " ='" & i & "'")
                End If
                Me.DataGrid1.SetFocus
                'Data1.Recordset.MoveFirst
'                lblIndicador.Caption = ""
            End If
        End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
    TerminaBloquear
    Select Case Modo
        Case 1 'Insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
        Case 3
            CargaGrid
    End Select
    PonerModo 0
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If (Modo = 0) Then Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    Me.DataGrid1.SetFocus
'    PonerCadenaBusqueda
End Sub


'Private Sub Form_KeyPress(KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub


Private Sub Form_Load()
Dim Sql As String

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Eliminar
        .Buttons(5).Image = 15  'Salir
    End With
        
    DataGrid1.ClearFields
    CadAncho = False
    PrimeraVez = True
    '## A mano
    NombreTabla = "stipom"
    'Ordenacion = " ORDER BY codtipom"
           
    'ASignamos un SQL al DATA1
    Sql = "Select codtipom, nomtipom, If(muevesto=1,""Si"",""No"") AS MovStock, contador, letraser, tipodocu "
    CadenaConsulta = Sql & " from " & NombreTabla
 
    CargaGrid
    CargaCombo
    PonerModo 0
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Insertar
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Eliminar
            mnEliminar_Click
        Case 5 'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub BotonAnyadir()
Dim anc As Single
Dim i As Byte
    
    lblIndicador.Caption = "INSERTAR"
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Me.Data1
      
    If DataGrid1.Row < 0 Then
        anc = 820
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 600
    End If

    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    Me.CboMueveStock.ListIndex = 1
    
    CargaTxtAux anc, 0
    PonerModo 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
Dim i As Integer
Dim anc As Single

    If Data1.Recordset.EOF Then Exit Sub
    If Data1.Recordset.RecordCount < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 290
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 600
    End If
    
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = Format(DataGrid1.Columns(3).Text, "0000000")
    txtAux(3).Text = DataGrid1.Columns(4).Text
    txtAux(4).Text = DataGrid1.Columns(5).Text
    
    Select Case DataGrid1.Columns(2).Value
        Case "Si"
            Me.CboMueveStock.ListIndex = 0
        Case "No"
            Me.CboMueveStock.ListIndex = 1
    End Select
    
    CargaTxtAux anc, 1
    PonerModo 2
    If Not BLOQUEADesdeFormulario(Me) Then
        cmdCancelar_Click
        Exit Sub
    End If
    
    'Como es modificar
    'Si es usuario Administrador
    If vUsu.Nivel = 1 Then PonerFoco txtAux(2)
    'Si es usuario root
    If vUsu.Nivel = 0 Then PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault

End Sub


Private Sub BotonEliminar()
Dim Sql As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    
    '### a mano
    Sql = "�Seguro que desea eliminar el Tipo de Movimiento?"
    Sql = Sql & vbCrLf & "C�digo: " & Data1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Denominaci�n: " & Data1.Recordset.Fields(1)
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        Sql = "Delete from stipom where codtipom='" & Data1.Recordset!codtipom & "'"
        conn.Execute Sql
        CancelaADODC Me.Data1
        CargaGrid ""
        CancelaADODC Me.Data1
        Me.DataGrid1.SetFocus
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Tipo de Movimiento", Err.Description
End Sub


Private Sub CargaGrid(Optional Sql As String)
Dim i As Byte
    
    If Sql <> "" Then
        Sql = CadenaConsulta & " WHERE " & Sql
    Else
        Sql = CadenaConsulta
    End If
    
    CargaGridGnral DataGrid1, Me.Data1, Sql, PrimeraVez
    
    
    i = 0  'C�digo
        DataGrid1.Columns(i).Caption = "C�digo"
        DataGrid1.Columns(i).Width = 600
    
    i = 1  'Nombre Tipo Movimiento
        DataGrid1.Columns(i).Caption = "Denominaci�n"
        DataGrid1.Columns(i).Width = 2200
'        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
            
    i = 2   'Mueve stock: Si o No
        DataGrid1.Columns(i).Caption = "Mueve Stock"
        DataGrid1.Columns(i).Width = 1100
        DataGrid1.Columns(i).Alignment = dbgCenter
'        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
            
    i = 3   'Contador
        DataGrid1.Columns(i).Caption = "Contador"
        DataGrid1.Columns(i).Width = 900
'        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
        DataGrid1.Columns(i).NumberFormat = "0000000"
        DataGrid1.Columns(i).Alignment = dbgCenter
        
    i = 4  'Letra Serie
        DataGrid1.Columns(i).Caption = "Letra Serie"
        DataGrid1.Columns(i).Width = 900
        DataGrid1.Columns(i).Alignment = dbgCenter
'        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
        
    i = 5  'Tipo de Documento
        DataGrid1.Columns(i).Caption = "T.Docum."
        DataGrid1.Columns(i).Width = 760
        DataGrid1.Columns(i).Alignment = dbgCenter
'        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width

   'No permitir cambiar tama�o de columnas
   For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
   Next i
   
   DataGrid1.Enabled = True
   DataGrid1.ScrollBars = dbgAutomatic
   
   PrimeraVez = False
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
            
    DatosOk = False
    b = CompForm(Me, 3)
    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(vModo As Byte)
Dim b As Boolean
Dim i As Byte

    Modo = vModo
    b = (Modo = 0)

    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = Not b
    Next i
    
    Me.CboMueveStock.visible = Not b
    
'    If b Then Me.lblIndicador.Caption = ""
    
    If b And Not Data1.Recordset.EOF Then
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    Else
        Me.lblIndicador.Caption = ""
    End If
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b

    'Si estamo modificar or insert
    If Modo = 2 Then
       txtAux(0).BackColor = &H80000018
    Else
        txtAux(0).BackColor = &H80000005
    End If
    
    Me.DataGrid1.Enabled = (Modo <> 1 And Modo <> 2)
    cmdAceptar.Enabled = (vUsu.Nivel <= 1)

    'Solo usuario root puede Insertar, Modificar y Borrar
    'Usuario administrador solo puede modificar campos "contador", "letra serie", y "Tipo Documento"
    'Modo 2: Modificar,  1: Insertar
    txtAux(0).Enabled = (Modo = 1) And (vUsu.Nivel = 0)
    txtAux(1).Enabled = ((Modo = 1) Or (Modo = 2)) And (vUsu.Nivel = 0)
    Me.CboMueveStock.Enabled = ((Modo = 1) Or (Modo = 2)) And (vUsu.Nivel = 0)
    txtAux(2).Enabled = (((Modo = 1) Or (Modo = 2)) And (vUsu.Nivel = 0)) Or (Modo = 2 And vUsu.Nivel = 1)
    txtAux(3).Enabled = (((Modo = 1) Or (Modo = 2)) And (vUsu.Nivel = 0)) Or (Modo = 2 And vUsu.Nivel = 1)
    txtAux(4).Enabled = (((Modo = 1) Or (Modo = 2)) And (vUsu.Nivel = 0)) Or (Modo = 2 And vUsu.Nivel = 1)
    
    PonerModoOpcionesMenu
    PonerOpcionesMenu
End Sub



Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 1) Or (Modo = 2) '1:insertar, 2: Modificar
    'A�adir
    Me.Toolbar1.Buttons(1).Enabled = Not b
    Me.mnNuevo.Enabled = Not b
    'Modificar
    Me.Toolbar1.Buttons(2).Enabled = Not b
    Me.mnModificar.Enabled = Not b
'    Me.Toolbar1.Buttons(2).Enabled = (Not b) 'Modificar
End Sub

Private Sub CargaCombo()
    'Carga la lista de impresi�n de etiquetas
    CboMueveStock.Clear
    CboMueveStock.AddItem "Si"
    CboMueveStock.ItemData(CboMueveStock.NewIndex) = 1
    
    CboMueveStock.AddItem "No"
    CboMueveStock.ItemData(CboMueveStock.NewIndex) = 0
    
End Sub


Private Sub CargaTxtAux(alto As Single, xModo As Byte)
Dim i As Byte

    DeseleccionaGrid Me.DataGrid1
    
    'Fijamos el ancho
    For i = 0 To txtAux.Count - 1
        txtAux(i).Top = alto
    Next i
    CboMueveStock.Top = alto - 15
    
    txtAux(0).Left = DataGrid1.Left + 330
    txtAux(0).Width = DataGrid1.Columns(0).Width - 30
    
    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 20
    txtAux(1).Width = DataGrid1.Columns(1).Width - 10
    
    CboMueveStock.Left = txtAux(1).Left + txtAux(1).Width + 15
    CboMueveStock.Width = DataGrid1.Columns(2).Width - 30
    
    txtAux(2).Left = CboMueveStock.Left + CboMueveStock.Width + 15
    txtAux(2).Width = DataGrid1.Columns(3).Width - 10
    
    txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 15
    txtAux(3).Width = DataGrid1.Columns(4).Width - 10
    
    txtAux(4).Left = txtAux(3).Left + txtAux(3).Width + 15
    txtAux(4).Width = DataGrid1.Columns(5).Width - 30
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), 3
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    'Quitar espacios en blanco por los lados
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then Exit Sub

    If Index = 2 Then 'Contador
        If Not IsNumeric(txtAux(Index).Text) Then
            MsgBox "Contador tiene que ser num�rico", vbExclamation
            PonerFoco txtAux(Index)
            Exit Sub
        Else
            txtAux(Index).Text = Format(txtAux(Index).Text, "0000000")
        End If
    End If
End Sub



Private Function SepuedeBorrar() As Boolean
Dim Sql As String
Dim cad As String

    SepuedeBorrar = False
    Sql = DevuelveDesdeBD(1, "detamovi", "smoval", "detamovi", Data1.Recordset!codtipom, "T")
    If Sql <> "" Then
        cad = "No se puede eliminar la fila. " & vbCrLf
        MsgBox cad & "Esta vinculada con Movimientos de Art�culos", vbExclamation
        Exit Function
    End If
    SepuedeBorrar = True
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

