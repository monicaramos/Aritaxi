VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacturacionCli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturaci�n por cliente"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   Icon            =   "frmFacturacionCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFacturar 
      Caption         =   "Facturar"
      Height          =   375
      Left            =   7440
      TabIndex        =   16
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   8880
      TabIndex        =   5
      Top             =   6960
      Width           =   1335
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5415
      Left            =   5400
      TabIndex        =   4
      Top             =   1320
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9551
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.TextBox txtSitua 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text5"
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtclien 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Situacion"
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   100
         Left            =   840
         Picture         =   "frmFacturacionCli.frx":000C
         ToolTipText     =   "Buscar actividad"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha Vto"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Factura"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "F. Factura"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Pendiente"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   17
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   19
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Pendiente"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   3960
      TabIndex        =   15
      Top             =   4560
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Albaranes pendientes facturar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Albaranes para facturar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   5400
      TabIndex        =   13
      Top             =   1080
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "Cobros pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1320
   End
   Begin VB.Label lblInd 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   4335
   End
End
Attribute VB_Name = "frmFacturacionCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1

Dim Sql As String
Dim Im As Currency

Private Sub cmdFacturar_Click()
Dim i As Integer

    


    If Me.txtclien.Text = "" Then Exit Sub
    If Me.txtnombre.Text = "" Then Exit Sub
    
    If TreeView1.Nodes.Count = 0 Then Exit Sub
    
    
    'Vere si hay alguno marcado para facturar
    Sql = ""
    For i = 1 To TreeView1.Nodes.Count
        If Not TreeView1.Nodes(i).Parent Is Nothing Then
            If TreeView1.Nodes(i).Checked Then
                If Not TreeView1.Nodes(i).Parent.Checked Then
                    MsgBox "Deberia estar marcado: " & TreeView1.Nodes(i).Parent.Text, vbExclamation
                    TreeView1.Nodes(i).Parent.Checked = True
                    Exit Sub
                End If
                
                Sql = "OK"
                Exit For
            End If
        End If
    Next
    If Sql = "" Then
        MsgBox "Ninguna albar�n marcado para facturar", vbExclamation
        Exit Sub
    End If


    CadenaDesdeOtroForm = ""
    frmListado2.Opcion = 25
    frmListado2.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        'OK Vamos a facturar
        Set miRsAux = Nothing
        Screen.MousePointer = vbHourglass
        HacerFacturacionCliente
        CargarDatos2
        Screen.MousePointer = vbDefault
        
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    lblInd.Caption = ""
    limpiar Me
    Set TreeView1.ImageList = frmPpal.imgListComun
    Set TreeView2.ImageList = frmPpal.imgListComun
    Set ListView1.SmallIcons = frmPpal.imgListComun
End Sub



Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtclien.Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub imgBuscarG_Click(Index As Integer)
    Sql = txtclien.Text
    Set frmCli = New frmFacClientes
    frmCli.DatosADevolverBusqueda = "0|1|"
    frmCli.Show vbModal
    Set frmCli = Nothing
    If txtclien.Text <> Sql Then PonerFoco txtclien
        
End Sub

Private Sub TreeView1_DblClick()
    If TreeView1.Nodes.Count = 0 Then Exit Sub
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If TreeView1.SelectedItem.Parent Is Nothing Then Exit Sub
    
    
      
        frmFacEntAlbaranes.hcoCodMovim = DevuelveNumeroAlbaran(TreeView1.SelectedItem.Text)
        frmFacEntAlbaranes.hcoCodTipoM = Mid(TreeView1.SelectedItem.Text, 1, 3)
        frmFacEntAlbaranes.RecuperarFactu = False
        frmFacEntAlbaranes.Show vbModal
        Set frmFacEntAlbaranes = Nothing
        
        'Vuelvo a cargar los datos
        
        CargarDatos2
  
  

  
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim N As Node
Dim Im As Currency

    If Node.Parent Is Nothing Then
        'Ha checkeado(quitado) uno padre. Todos los hijos haran los mismo
        Set N = Node.Child
        'If Node.Checked = False Then N.Tag = 0
            
        Im = 0
        While Not N Is Nothing
            
            N.Checked = N.Parent.Checked
            If N.Checked Then Im = Im + N.Tag
            Set N = N.Next
            
            
        Wend
        Node.Tag = Im
        PonerCadenaImporte Node, True
    Else
        If Node.Checked Then
            Node.Parent.Tag = Node.Parent.Tag + Node.Tag
        Else
            Node.Parent.Tag = Node.Parent.Tag - Node.Tag
        End If
        PonerCadenaImporte Node.Parent, True
        
        
        
        'Comprobare que si hay marcado alguno el ppal este maracdo y al reves
        Im = 0
        Set N = Node.FirstSibling
        
            
        Im = 0  'Ninguno chekeado
        While Not N Is Nothing
            
            If N.Checked Then Im = Im + 1
            Set N = N.Next
            
            
        Wend
        Node.Parent.Checked = Im > 0
    End If
    
    
                        
End Sub

Private Sub PonerCadenaImporte(ByRef N As Node, Padre As Boolean)
Dim i As Integer
Dim J As Integer
    If Padre Then
        J = 24
    Else
        J = 45
    End If
    i = InStr(1, N.Text, ":")
    If i > 0 Then
        N.Text = Mid(N.Text, 1, i)
        N.Text = N.Text & Right(Space(J) & Format(N.Tag, FormatoImporte), J)
    End If
End Sub

Private Sub TreeView2_DblClick()
    If TreeView2.Nodes.Count = 0 Then Exit Sub
    If TreeView2.SelectedItem Is Nothing Then Exit Sub

    
    
      
        frmFacEntAlbaranes.hcoCodMovim = DevuelveNumeroAlbaran(TreeView2.SelectedItem.Text)
        frmFacEntAlbaranes.hcoCodTipoM = Mid(TreeView2.SelectedItem.Text, 1, 3)
        frmFacEntAlbaranes.RecuperarFactu = False
        frmFacEntAlbaranes.Show vbModal
        Set frmFacEntAlbaranes = Nothing
        
        'Vuelvo a cargar los datos
        
        CargarDatos2
End Sub

Private Sub txtclien_GotFocus()
   ConseguirFoco txtclien, 3
End Sub

Private Sub txtclien_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtclien_LostFocus()


    Sql = ""
    txtclien.Text = Trim(txtclien.Text)
    txtSitua.Text = ""
    txtnombre.Text = ""
    If txtclien.Text <> "" Then
    
        If PonerFormatoEntero(txtclien) Then
            
            Set miRsAux = New ADODB.Recordset
            Sql = PonerCliente
            If Sql = "" Then
                MsgBox "No existe el cliente: " & txtclien.Text, vbExclamation
                txtclien.Text = ""
                PonerFoco txtclien
            Else
                'Cargar DATOS
                CargarDatos2
               
            
            End If

        End If
    End If
    If Sql = "" Then
        Me.ListView1.ListItems.Clear
        Me.TreeView1.Nodes.Clear
        Me.TreeView2.Nodes.Clear
        lblTot(0).Caption = ""
        lblTot(1).Caption = ""
    End If
    
    
End Sub

Private Function PonerCliente() As String
    Set miRsAux = New ADODB.Recordset
    Sql = "Select nomclien,nomsitua,codmacta from sclien,ssitua WHERE sclien.codsitua=ssitua.codsitua AND codclien=" & txtclien.Text
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    If Not miRsAux.EOF Then
        Sql = miRsAux!nomclien
        Me.txtnombre.Text = Sql
        PonerCliente = miRsAux!nomclien
        txtSitua = miRsAux!nomsitua
        txtSitua.Tag = miRsAux!codmacta
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Function

Private Sub CargarDatos2()
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset

    'Cargamos cobros pendientes
    lblInd.Caption = "Vencimientos"
    lblInd.Refresh
    CargarVtos
    
    
    'Cargamos albarananes pendientes de facturar
    lblInd.Caption = "Albaranes pendientes facturar"
    lblInd.Refresh
    CargaAlbaranes
    
    
    lblInd.Caption = "Albaranes sin marca facturar"
    lblInd.Refresh
    CargaAlbaranesSin
    
    
    lblInd.Caption = ""
    Screen.MousePointer = vbDefault
    Set miRsAux = Nothing
    
End Sub


Private Sub CargarVtos()
Dim It As ListItem
Dim Im2 As Currency
Dim Pend As Currency

    ListView1.ListItems.Clear
    If vParamAplic.ContabilidadNueva Then
        Sql = "SELECT cobros.* FROM cobros INNER JOIN formapago ON cobros.codforpa=formapago.codforpa "
        Sql = Sql & " WHERE cobros.codmacta = '" & txtSitua.Tag & "'"
    
    Else
        Sql = "SELECT scobro.* FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
        Sql = Sql & " WHERE scobro.codmacta = '" & txtSitua.Tag & "'"
        'SQL = SQL & " AND fecvenci <= ' " & Format(Now, FormatoFecha) & "' "
        ' SQL = SQL & " AND (sforpa.tipforpa between 0 and 3)
    End If
    Sql = Sql & " ORDER BY fecvenci"
    miRsAux.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Im = 0
    Pend = 0
    While Not miRsAux.EOF
        Im2 = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
        If Im2 <> 0 Then
    
            Set It = ListView1.ListItems.Add()
            It.Text = miRsAux!FecVenci
            It.SmallIcon = 23
            'If miRsAux!FecVenci > Now Then
            If vParamAplic.ContabilidadNueva Then
                It.SubItems(1) = miRsAux!numSerie & Format(miRsAux!NumFactu, "00000")
                It.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
            Else
                It.SubItems(1) = miRsAux!numSerie & Format(miRsAux!Codfaccl, "00000")
                It.SubItems(2) = Format(miRsAux!fecfaccl, "dd/mm/yyyy")
            End If
            
            It.SubItems(3) = Format(Im2, FormatoImporte)
            Im = Im + Im2
            If miRsAux!FecVenci < Now Then Pend = Pend + Im2
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If ListView1.ListItems.Count = 0 Then
        lblTot(1).Caption = ""
        lblTot(0).Caption = ""
    Else
        lblTot(0).Caption = Format(Im, FormatoImporte)
        lblTot(1).Caption = Format(Pend, FormatoImporte)
    End If
End Sub



Private Sub CargaAlbaranes()
Dim Anterior As String
Dim Col As Collection
    TreeView1.Nodes.Clear
    'Todo estara en una cadena    direc|forpa|dtopp|dtogn|   Si cambia algo sera salto factura
    'antClien = 0 'cliente SIEMPRE ES EL MISMO
    'antDirec = 0 'direccion/departamento
    'antForpa = 0 'forma de pago
    'antDtoPP = 0 'dto pronto pago
    'antDtoGn = 0 'dto general
    Sql = "Select *  FROM  scaalb  WHERE (scaalb.fechaalb <= '2010-04-06') AND (scaalb.codclien = " & txtclien.Text
    Sql = Sql & ") AND ( scaalb.codtipom='ALV' ) AND ( scaalb.factursn=1)  and ((scaalb.codtipom,scaalb.numalbar) in (select distinct codtipom,numalbar from slialb))"
    Sql = Sql & " ORDER BY scaalb.tipofact, scaalb.codclien, scaalb.coddirec, codforpa, dtoppago, dtognral "
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Anterior = ""
    NumRegElim = 1
    Set Col = New Collection
    
    While Not miRsAux.EOF
        If miRsAux!TipoFact = 1 Then
            'Factura x albaran
            
            
            'Hay que meter una factura anterior
            If Anterior <> "" Then InsertarLineaFactura Col
                
            'Meto esta
            CadenaAlbaran Col
            InsertarLineaFactura Col
            
            Anterior = ""
        Else
            Sql = CadenaIndentificacionAlbaran
            If Sql <> Anterior Then
                'Ha cambiado algun valor
                If Anterior <> "" Then InsertarLineaFactura Col
                
                
                Anterior = Sql
            End If
            CadenaAlbaran Col 'Meto el albaran en el collection
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Col.Count > 0 Then InsertarLineaFactura Col
End Sub

Private Function CadenaIndentificacionAlbaran() As String
  '  direc|forpa|dtopp|dtogn|
    CadenaIndentificacionAlbaran = Format(DBLet(miRsAux!CodDirec, "N"), "000") & "|" & Format(DBLet(miRsAux!codforpa, "N"), "000") & "|"
    CadenaIndentificacionAlbaran = CadenaIndentificacionAlbaran & Format(miRsAux!DtoPPago * 100, "0000") & "|" & Format(miRsAux!DtoGnral * 100, "0000") & "|"
End Function

Private Sub CadenaAlbaran(ByRef Cole As Collection)
Dim C As String

    C = " codtipom = '" & miRsAux!codtipom & "' AND numalbar"
    C = DevuelveDesdeBD(conAri, "sum(importel)", "slialb", C, miRsAux!NumAlbar)
    
    'Ira codtipomNumalbar sapacioblanco fecha  espacios importe
    Cole.Add miRsAux!codtipom & Format(miRsAux!NumAlbar, "000000") & "  " & Format(miRsAux!FechaAlb, "dd/mm/yyyy") & "|" & C & "|"
    
End Sub

Private Function DevuelveNumeroAlbaran(linea As String) As String
Dim J As Integer
    
    DevuelveNumeroAlbaran = "0"
    
    J = InStr(1, linea, " ")
    If J > 0 Then
        DevuelveNumeroAlbaran = Mid(linea, 1, J - 1)
        DevuelveNumeroAlbaran = Mid(DevuelveNumeroAlbaran, 4) 'los tres primeros son el codtipom
    End If
End Function


Private Sub InsertarLineaFactura(ByRef Cole As Collection)
Dim i As Integer
Dim N As Node
Dim TotalFra As Currency



    If Cole.Count = 0 Then
        'Msgbox
        'No tiene albaranes a facturar? algo raro ha pasado
        
    End If
       

    'Meto el raiz
    Set N = TreeView1.Nodes.Add(, , "FRA" & Format(NumRegElim, "000"), "Fra " & NumRegElim)
    N.Image = 43
    N.Checked = True
    TotalFra = 0
    'Los albaranes que iran
    For i = 1 To Cole.Count
        'El importe
        Sql = RecuperaValor(Cole.item(i), 2)
        Im = CCur(Sql)
        TotalFra = TotalFra + Im
        
        'El importe
        Sql = Right(Space(10) & Format(Im, FormatoImporte), 10)
        Sql = RecuperaValor(Cole.item(i), 1) & Sql
        Set N = TreeView1.Nodes.Add("FRA" & Format(NumRegElim, "000"), tvwChild)
        N.Text = Sql
        N.Image = 44
        N.Checked = True
        N.Tag = Im
        
        
        
    Next
    N.Parent.Text = N.Parent.Text & "   Imp: "
    N.Parent.Tag = TotalFra
    PonerCadenaImporte N.Parent, True
    
    N.Parent.Expanded = True
    NumRegElim = NumRegElim + 1
    Set Cole = Nothing
    Set Cole = New Collection
End Sub



Private Sub CargaAlbaranesSin()
Dim Col As Collection
Dim N As Node

    TreeView2.Nodes.Clear
    
    Sql = "Select *  FROM  scaalb  WHERE (scaalb.fechaalb <= '2010-04-06') AND (scaalb.codclien = " & txtclien.Text
    Sql = Sql & ") AND ( scaalb.codtipom='ALV' ) AND ( scaalb.factursn=0)  and ((scaalb.codtipom,scaalb.numalbar) in (select distinct codtipom,numalbar from slialb))"
    Sql = Sql & " ORDER BY scaalb.tipofact, scaalb.codclien, scaalb.coddirec, codforpa, dtoppago, dtognral "
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set Col = New Collection
        CadenaAlbaran Col
        
        Sql = RecuperaValor(Col.item(1), 2)
        
        'El importe
        Sql = Right(Space(10) & Format(Sql, FormatoImporte), 10)
        Sql = RecuperaValor(Col.item(1), 1) & Sql
        Set N = TreeView2.Nodes.Add()
        N.Text = Sql
        N.Image = 44
            
            
        miRsAux.MoveNext
        Set Col = Nothing
    Wend
    miRsAux.Close
    
End Sub



Private Sub HacerFacturacionCliente()
Dim CadenaSQL As String
Dim i As Integer
    
    Sql = ""
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Parent Is Nothing Then
            'NADA
            
        Else
            If TreeView1.Nodes(i).Checked Then Sql = Sql & ", " & DevuelveNumeroAlbaran(TreeView1.Nodes(i).Text)
   
        End If
    Next i
    
    Sql = Mid(Sql, 3)
    
    CadenaSQL = "scaalb.codtipom = 'ALV' AND scaalb.codclien=" & Me.txtclien.Text & " AND  scaalb.numalbar IN (" & Sql & ")"
    Sql = "SELECT scaalb.*,sclien.nomclien FROM scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien  WHERE " & CadenaSQL
    
    i = Val(RecuperaValor(CadenaDesdeOtroForm, 3))
    
    TraspasoAlbaranesFacturas Sql, CadenaSQL, RecuperaValor(CadenaDesdeOtroForm, 1), RecuperaValor(CadenaDesdeOtroForm, 2), Nothing, Me.lblInd, i = 1, "ALV", "", False
End Sub
