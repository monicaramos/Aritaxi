VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCRMImprimir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresion CRM"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmCRMImprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFecha 
      Height          =   375
      Left            =   840
      Picture         =   "frmCRMImprimir.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cambiar fecha desde"
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   9360
      TabIndex        =   6
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin MSComctlLib.TreeView TV1 
      Height          =   5295
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9340
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "932"
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.TreeView tv2 
      Height          =   5295
      Left            =   5640
      TabIndex        =   8
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9340
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tv3 
      Height          =   5295
      Left            =   8760
      TabIndex        =   10
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9340
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   9720
      Picture         =   "frmCRMImprimir.frx":0596
      ToolTipText     =   "Quitar seleccion"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   10080
      Picture         =   "frmCRMImprimir.frx":06E0
      ToolTipText     =   "seleccionar todos"
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblInd 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Impresion"
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
      Index           =   2
      Left            =   8760
      TabIndex        =   12
      Top             =   810
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Salen en CRM"
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
      Index           =   1
      Left            =   5640
      TabIndex        =   11
      Top             =   810
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Informe"
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
      Left            =   120
      TabIndex        =   9
      Top             =   810
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   960
      Picture         =   "frmCRMImprimir.frx":082A
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
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
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmCRMImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public N 'As Node
Private PrimeraVez As Boolean
Private WithEvents frmC2 As frmFacClientes2
Attribute frmC2.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private GuardarConfig As Boolean

Dim J As Integer
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Donde As String




Private vCRM As cCRM
Private HayAlgunDato As Boolean
Private cadParam2 As String   'Para pasarle los parametros al rpt

Dim DatosGuardados As Collection

'Configuracion en el equipo



Private Sub CargaTreeView()

    'EN EL TAG llevara los valores para la cadparam
    ' parametrovisible|parametrofecha|    el de fecha es optativo
    
    'Losw parametros son:
    'Para las fechas
    ' {pDesdeAlbarabSat} {pDesdeAlbaran} {pDesdeAnyo} {pDesdeAvisos} {pDesdeEmail}
    '{pDesdeLlamada} {pDesdeOferta} {pDesdepedido} {pDesdeReclamas} {pDesdeRepara}
    'Para los visibles
    '{pVisAccionesComer} {pVisAlbaranes} {pVisAlbSat} {pVisAvisos} {pVisCobrPdte}
    '{pVisEmails} {pVisFreq} {pVisLlamadas} {pVisMtos} {pVisOfertas}{pVisPedido}
    '{pVisReclamas} {pVisReparas} {pVisVolVenta}
    Configuracion True
    
    CargaAdmon
    CargaComercial
    CargaSAT
End Sub

Private Sub CargaAdmon()

    'Losw parametros son:
    'Para las fechas
    ' {pDesdeAlbarabSat} {pDesdeAlbaran} {} {pDesdeAvisos} {pDesdeEmail}
    '{pDesdeLlamada} {pDesdeOferta} {pDesdepedido} {pDesdeReclamas} {pDesdeRepara}
    'Para los visibles
    '{pVisAccionesComer} {pVisAlbaranes} {pVisAlbSat} {pVisAvisos} {pVisCobrPdte}
    '{pVisEmails} {pVisFreq} {pVisLlamadas} {pVisMtos} {pVisOfertas}{pVisPedido}
    '{pVisReclamas} {pVisReparas} {}
    
    'Departamento administradcio
    Set N = TV1.Nodes.Add(, , "ADM")
    N.Text = "Datos dpto de administración"
    N.Bold = True
    N.Checked = NodoPadreCheckeado(N.Index)    '
    
    FijarNodo3 N, "ADM", "adm1", True, True, "Volumen facturación"
    N.Tag = "pVisVolVenta|pDesdeAnyo|"

    FijarNodo3 N, "ADM", "adm2", False, True, "Facturas pendientes de cobro"
    N.Tag = "pVisCobrPdte||"
    
    
    FijarNodo3 N, "ADM", "adm3", True, False, "Detalle reclamaciones de cobros efectuadas"
    N.Tag = "pVisReclamas|pDesdeReclamas|"

    FijarNodo3 N, "ADM", "adm4", False, False, "Detalle mantenimiento"
    N.Tag = "pVisMtos||"

End Sub

Private Function NodoPadreCheckeado(indice As Integer) As Boolean
    
    NodoPadreCheckeado = True
    If Not DatosGuardados Is Nothing Then
        If DatosGuardados.Count >= indice Then NodoPadreCheckeado = RecuperaValor(DatosGuardados(indice), 1) = "1"
    End If
End Function
Private Sub CargaComercial()
        'Para las fechas
    ' {pDesdeAlbarabSat} {pDesdeAlbaran} {pDesdeAnyo} {pDesdeAvisos} {pDesdeEmail}
    '{pDesdeLlamada} {pDesdeOferta} {pDesdepedido} {pDesdeReclamas} {pDesdeRepara}
    'Para los visibles
    '{pVisAccionesComer} {pVisAlbaranes} {pVisAlbSat} {pVisAvisos} {pVisCobrPdte}
    '{pVisEmails} {pVisFreq} {pVisLlamadas} {pVisMtos} {pVisOfertas}{pVisPedido}
    '{pVisReclamas} {pVisReparas} {pVisVolVenta}
    
    'Departamento administradcio
    Set N = TV1.Nodes.Add(, , "COM")
    N.Text = "Datos dpto de comercial"
    N.Tag = "||"
    N.Bold = True
    N.Checked = NodoPadreCheckeado(N.Index)
     
    FijarNodo3 N, "COM", "com1", True, False, "Detalle ofertas pendientes"
    N.Tag = "pVisOfertas|pDesdeOferta|"
    
    
    
    FijarNodo3 N, "COM", "com2", True, False, "Detalle pedidos pendientes de entregar"
    N.Tag = "pVisPedido|pDesdepedido|"
    
    FijarNodo3 N, "COM", "com3", True, False, "Detalle albaranes pendientes de facturar"
    N.Tag = "pVisAlbaranes|pDesdeAlbaran|"
    
    'Acciones comerciales. Lo hemos intecalado
    FijarNodo3 N, "COM", "com6", True, False, "Acciones comerciales "
    N.Tag = "pVisAccionesComer|pDesdeAccComer|"
    
    
    FijarNodo3 N, "COM", "com4", True, False, "Detalle llamadas"
    N.Tag = "pVisLlamadas|pDesdeLlamada|"
    
            FijarNodo3 N, "com4", "com41", False, False, "Recibidas"
            FijarNodo3 N, "com4", "com42", False, False, "Realizadas"
    
    FijarNodo3 N, "COM", "com5", True, False, "Detalle correos(eMail)"
    N.Tag = "pVisEmails|pDesdeEmail|"
    
    
            FijarNodo3 N, "com5", "com51", False, False, "Recibidos"
            FijarNodo3 N, "com5", "com52", False, False, "Enviados"
            
End Sub



Private Sub CargaSAT()
    
    If Not vParamAplic.Reparaciones Then Exit Sub
    
        'Para las fechas
    ' {pDesdeAlbarabSat} {pDesdeAlbaran} {pDesdeAnyo} {pDesdeAvisos} {pDesdeEmail}
    '{pDesdeLlamada} {pDesdeOferta} {pDesdepedido} {pDesdeReclamas} {pDesdeRepara}
    'Para los visibles
    '{pVisAccionesComer} {pVisAlbaranes} {pVisAlbSat} {pVisAvisos} {pVisCobrPdte}
    '{pVisEmails} {pVisFreq} {pVisLlamadas} {pVisMtos} {pVisOfertas}{pVisPedido}
    '{pVisReclamas} {pVisReparas} {pVisVolVenta}
    

    'Departamento administradcio
    Set N = TV1.Nodes.Add(, , "SAT")
    N.Text = "Datos dpto de S.A.T."
    N.Bold = True
    N.Checked = NodoPadreCheckeado(N.Index)
    
    
 
 
    FijarNodo3 N, "SAT", "sat1", False, False, "Frecuencias"
    N.Tag = "pVisFreq||"
    
    FijarNodo3 N, "SAT", "sat2", True, False, "Albaranes reparacion pendientes facturar"
    N.Tag = "pVisAlbSat|pDesdeAlbarabSat|"
    
    FijarNodo3 N, "SAT", "sat3", True, False, "Avisos pendientes de cerrar"
    N.Tag = "pVisAvisos|pDesdeAvisos|"
    
    
    FijarNodo3 N, "SAT", "sat4", True, False, "Equipos pendientes de reparar"
    N.Tag = "pVisReparas|pDesdeRepara|"
    

End Sub




Private Sub cmdFecha_Click()
    If TV1.Nodes.Count = 0 Then Exit Sub
    If TV1.SelectedItem Is Nothing Then Exit Sub
    
    If Right(TV1.SelectedItem.Text, 1) <> "]" Then
        MsgBox "NO se le asigna fecha a esta opcion", vbExclamation
    Else
        Sql = ""
        J = InStr(1, TV1.SelectedItem, "[")
        If J = 0 Then
            MsgBox "No se ha encotrado la marca de fecha", vbExclamation
        Else
            Donde = Mid(TV1.SelectedItem.Text, J + 1)
            Donde = Mid(Donde, 1, Len(Donde) - 1)
            If Len(Donde) = 4 Then
                'Es AÑO
                J = 0
                Donde = "01/01/" & Donde
            Else
                'Es fecha
                J = 1
            End If
            Sql = ""
            Set frmC = New frmCal
            frmC.Fecha = CDate(Donde)
            frmC.Show vbModal
            Set frmC = Nothing
            If Sql <> "" Then
                
                            'Solo quiero el año
                If J = 0 Then Sql = Year(Sql)
                
                J = InStr(TV1.SelectedItem.Text, "[")
                If J = 0 Then
                    MsgBox ""
            
                Else
                    'Ha retornado dato
                    GuardarConfig = True
                    Donde = Mid(TV1.SelectedItem.Text, 1, J)
                    Donde = Donde & Sql & "]"
                    TV1.SelectedItem.Text = Donde
                End If
                Donde = ""
                Sql = ""
            
                End If
            End If
        End If
End Sub

Private Sub cmdImprimir_Click()

    'el unico control de errores esta aqui
On Error GoTo EcmdImprimir
    
    
    'A ver si esta configurada
    pPdfRpt = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "46")
    If pPdfRpt = "" Then
        MsgBox "Falta configurar en informes(46)", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'En estas cargaremos los albaranes, ofertas y facturas seleccionadas
    ejecutar "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo, False
    NumRegElim = 0 'contador para tmp con los ofe/ped/alb
    
    
    Set RS = New ADODB.Recordset
    Set vCRM = New cCRM
    HayAlgunDato = False
    cadParam2 = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    ''46'
    

    
    
    GenerarDatosInformes
    
    HayAlgunDato = True
    
    If HayAlgunDato Then
        InsertaDatosBasicos
        LlamarImprimir False
        
        
        ImprimirDocumentosAuxiliares
        
    End If
        
    
EcmdImprimir:
    If Err.Number <> 0 Then MuestraError Err.Number, Donde & vbCrLf & Err.Description
    Set RS = Nothing
    Set vCRM = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        Screen.MousePointer = vbHourglass
        PrimeraVez = False
        CargaDatosAux
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    CargaTreeView
    For J = 1 To TV1.Nodes.Count
        'TV1.Nodes(J).Checked = True
        TV1.Nodes(J).EnsureVisible
    Next J

    GuardarConfig = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GuardarConfig Then Configuracion False
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Sql = CStr(vFecha)
End Sub

Private Sub frmC2_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub Image1_Click()
    Sql = ""
    Set frmC2 = New frmFacClientes2
    frmC2.DatosADevolverBusqueda = "0|1|"
    frmC2.Show vbModal
    Set frmC2 = Nothing
    If Sql <> "" Then
        Me.Text1.Text = RecuperaValor(Sql, 1)
        Me.Text2.Text = RecuperaValor(Sql, 2)
        CargaDatosAux
    End If
End Sub

Private Sub imgCheck_Click(Index As Integer)
    If tv3.Nodes.Count = 0 Then Exit Sub
    For J = 1 To tv3.Nodes.Count
        tv3.Nodes(J).Checked = Index = 1
    Next J
End Sub

Private Sub Text1_GotFocus()
    ConseguirFoco Text1, 3
End Sub

Private Sub Text1_LostFocus()
    Sql = ""
    Text1.Text = Trim(Text1.Text)
    If Text1.Text <> "" Then
        If Not IsNumeric(Text1.Text) Then
            MsgBox "Codigo cliente numérico: " & Text1.Text, vbExclamation
            Text1.Text = ""
            PonerFoco Text1
        Else
            Sql = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1.Text)
            If Sql = "" Then
                MsgBox "no existe cliente: " & Text1.Text, vbExclamation
                PonerFoco Text1
    
            End If
        End If
    End If
    Text2.Text = Sql
    CargaDatosAux
    
End Sub

Private Sub TV1_DblClick()
    If TV1.Nodes.Count = 0 Then Exit Sub
    If TV1.SelectedItem Is Nothing Then Exit Sub
    
    If InStr(TV1.SelectedItem.Text, "[") = 0 Then Exit Sub
    
    cmdFecha_Click
End Sub

Private Sub TV1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim CH As Boolean
    If PrimeraVez Then Exit Sub
    
    If Node.Checked Then
        If Not Node.Parent Is Nothing Then Node.Parent.Checked = True
    End If
    
    CH = Node.Checked
    CheckSubNodo Node, CH, False
    GuardarConfig = True
End Sub


Private Sub CheckSubNodo(ByRef N, Checkar As Boolean, EsElTV2 As Boolean)
Dim no
    
    Set no = N
    no.Checked = Checkar
    If EsElTV2 Then CheckeaTambienEnElTv3 no.Index, Checkar
    Set no = N.Child
    While Not no Is Nothing
        CheckSubNodo no, Checkar, EsElTV2
        Set no = no.Next
    Wend
    
    
    
End Sub


Private Sub CheckeaTambienEnElTv3(indice As Integer, ChK)
    On Error Resume Next
    tv3.Nodes(indice).Checked = ChK
    Err.Clear
End Sub
Private Sub LlamarImprimir(PonerNombrePDF As Boolean)
Dim K As Integer

    With frmImprimir
        .FormulaSeleccion = "{tmpcrmclien.codusu} = " & vUsu.Codigo
        
        'Cuantos parametros envio
        NumRegElim = 0
        J = 2
        Do
           K = InStr(J, cadParam2, "|")
           If K > 0 Then
                NumRegElim = NumRegElim + 1
                J = K + 1
            End If
        Loop Until K = 0
        .OtrosParametros = cadParam2
        .NumeroParametros = NumRegElim

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2018
        .Titulo = "CRM"
        .NombreRPT = pPdfRpt
        .NombrePDF = ""
        'If PonerNombrePDF Then .NombrePDF = pPdfRpt
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub



'Generad Datos
Private Sub InsertaDatosBasicos()
Dim Aux As String

    'Si habian metido algun dato...
    Sql = "insert into `tmpcrmclien` (`codusu`,`codclien`,`saldopdte`,saldototal,`nomactiv`,`nomforpa`) values ("
    Sql = Sql & vUsu.Codigo & "," & Text1.Text & ","
    
    'Saldo pdte (a fecha NOW
    Aux = "Imp"
    ComprobarCobrosCliente2 Text1.Text, Now, Aux
    If Aux = "" Or Aux = "Imp" Then Aux = "0"
    Sql = Sql & DBSet(Aux, "N") & ","
    'saldo totoal A fecha 31/12/2222"
    Aux = "Imp"
    ComprobarCobrosCliente2 Text1.Text, CDate("31/12/2222"), Aux
    If Aux = "" Or Aux = "Imp" Then Aux = "0"
    Sql = Sql & DBSet(Aux, "N") & ","
    
    
    
    Aux = DevuelveDesdeBD(conAri, "nomactiv", "sclien,sactiv", "sclien.codactiv=sactiv.codactiv and codclien", Text1.Text)
    Sql = Sql & DBSet(Aux, "T") & ","
    Aux = DevuelveDesdeBD(conAri, "nomforpa", "sclien,sforpa", "sclien.codforpa=sforpa.codforpa and codclien", Text1.Text)
    Sql = Sql & DBSet(Aux, "T") & ")"
    conn.Execute Sql
End Sub



Private Sub GenerarDatosInformes()

    vCRM.BorrarTemporales
    vCRM.CodClien = CInt(Text1.Text)
    vCRM.codmacta = DevuelveDesdeBD(conAri, "codmacta", "sclien", "codclien", Text1.Text)

    
    
    J = DevuelveIndiceNodo("ADM")
    If Me.TV1.Nodes(J).Checked Then
        GenerarDatosAdmon
    Else
        'PONGO TODOS LOS SUBPARAMETROS A FALSE
        PonerparametrosVisiblesFalse
    End If
    
    
    'Para saber si tiene datos cada secccion
    J = DevuelveIndiceNodo("COM")
    If Me.TV1.Nodes(J).Checked Then
        GenerarDatosComer
    Else
        'PONGO TODOS LOS SUBPARAMETROS A FALSE
        PonerparametrosVisiblesFalse
    End If


    'Para saber si tiene datos cada secccion
    If vParamAplic.Reparaciones Then
        J = DevuelveIndiceNodo("SAT")
        If Me.TV1.Nodes(J).Checked Then
            GenerarDatosSAT
        Else
            'PONGO TODOS LOS SUBPARAMETROS A FALSE
            PonerparametrosVisiblesFalse
        End If
    Else
        cadParam2 = cadParam2 & "pVisFreq=0|pVisAlbSat=0|pVisAvisos=0|pVisReparas=0|"
    End If
    
End Sub


Private Sub PonerparametrosVisiblesFalse()
Dim N As Node
    'en TV1(j) tengo el NODO padre
    'Con lo cual, recorrro todos sus hijos, obteneido la cadena param de visible y poneindola a cero
    Set N = TV1.Nodes(J).Child '
    While Not (N Is Nothing)
        Sql = RecuperaValor(N.Tag, 1)
        If Sql <> "" Then cadParam2 = cadParam2 & Sql & "=0|"
        Set N = N.Next
    Wend
End Sub



Private Function DevuelveIndiceNodo(Clave As String) As Integer
Dim I As Integer
    
    For I = 1 To TV1.Nodes.Count
        If TV1.Nodes(I).Key = Clave Then
            DevuelveIndiceNodo = I
            Exit Function
        End If
    Next
    
    'Si llega aqui generaremos un erro
    Err.Raise 512, , "NO se encuentra NODO : " & Clave
End Function


'COmercia
'---------------------------
Private Sub GenerarDatosComer()
Dim cad As String
Dim Contador As Long
Dim F As Date
    Donde = "Comercial"
    'Volumen facturacion
    J = DevuelveIndiceNodo("com1")
    If HayKprocesarNodo(J, F) Then
        Donde = "Ofertas pendientes"
        
    
    End If
    
    
    J = DevuelveIndiceNodo("com2")
    If HayKprocesarNodo(J, F) Then
        Donde = "Pedidos pendientes"
        
    End If
    
    
    J = DevuelveIndiceNodo("com3")
    If HayKprocesarNodo(J, F) Then
        Donde = "Albaranes pdtes"
        
    End If
    
    
    'Acciones comerciales
    J = DevuelveIndiceNodo("com6")
    If HayKprocesarNodo(J, F) Then
        Donde = "Acciones comerciales"
    End If
    
    
    
    Contador = 0
    
    J = DevuelveIndiceNodo("com4")
    If HayKprocesarNodo(J, F) Then
        Donde = "Llamadas"

        
        'Si no quiere las recibidas
        J = DevuelveIndiceNodo("com41")
        If HayKprocesarNodo(J, F) Then
            'insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,`trabajador`,`adjuntos`) values ( '1','0','','',NULL,NULL,NULL,NULL)
            Sql = "select feholla,usuario,nomllama1,observac,codtraba,nomtraba from sllama,sllama1 "
            Sql = Sql & "  where sllama.codllama1 = sllama1.codllama1"
            Sql = Sql & " and codclien=" & vCRM.CodClien
            Sql = Sql & " AND feholla>=" & DBSet(F, "F")
            
            
            
            
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                NumRegElim = NumRegElim + 1
                Sql = "insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,"
                Sql = Sql & "`trabajador`,`adjuntos`) values ( " & vUsu.Codigo & "," & NumRegElim & ",0,"
                Sql = Sql & DBSet(RS!feholla, "FH") & ","
                'En sllama siempre son RECIBIDAS
                Sql = Sql & "'Recibida',"
                cad = DBLetMemo(RS!observac)
                cad = Replace(cad, vbCrLf, " ")
                Sql = Sql & DBSet(cad, "T", "S") & ","
                'Trabajador
                Sql = Sql & DBSet(RS!NomTraba, "T") & ","
                'En adjuntos guardare el tipop llamada
                Sql = Sql & DBSet(RS!nomllama1, "T") & ")"
                
                conn.Execute Sql
                RS.MoveNext
            Wend
            RS.Close
            'Ha metido algun dato
           ' If NumRegElim > 0 Then comer(4) = True   'tiene datos
            Contador = NumRegElim
        End If
            
        'Si no quiere las realizadas
        J = DevuelveIndiceNodo("com42")
        If HayKprocesarNodo(J, F) Then
            'insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,`trabajador`,`adjuntos`) values ( '1','0','','',NULL,NULL,NULL,NULL)
            Sql = "select fechora ,usuario,nomtraba ,observaciones from"
            Sql = Sql & " scrmacciones left join straba on scrmacciones.codtraba=straba.codtraba "
            Sql = Sql & " WHERE scrmacciones.tipo=1  and codclien= " & vCRM.CodClien
            Sql = Sql & " AND fechora>=" & DBSet(F, "F")
            
            
            
            
            
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                NumRegElim = NumRegElim + 1
                Sql = "insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,"
                Sql = Sql & "`trabajador`,`adjuntos`) values ( " & vUsu.Codigo & "," & NumRegElim & ",0,"
                Sql = Sql & DBSet(RS!fechora, "FH") & ","
                'En sllama siempre son RECIBIDAS
                Sql = Sql & "'Realizada',"
                cad = DBLetMemo(RS!Observaciones)
                cad = Replace(cad, vbCrLf, " ")
                Sql = Sql & DBSet(cad, "T", "S") & ","
                'Trabajador
                Sql = Sql & DBSet(RS!NomTraba, "T") & ","
                'En adjuntos guardare el tipop llamada
                Sql = Sql & "NULL)"
                
                conn.Execute Sql
                RS.MoveNext
            Wend
            RS.Close
            'Ha metido algun dato
            'If NumRegElim > Contador Then comer(4) = True   'tiene datos
            Contador = NumRegElim
        End If
        
    End If
    
    
    
    
    J = DevuelveIndiceNodo("com5")
    If HayKprocesarNodo(J, F) Then
        Donde = "Emails"
        
        
        'Si no quiere las recibidas
        NumRegElim = 0
        J = DevuelveIndiceNodo("com51")
        If TV1.Nodes(J).Checked Then NumRegElim = 1
        
        J = DevuelveIndiceNodo("com51")
        If TV1.Nodes(J).Checked Then NumRegElim = NumRegElim + 2
        
        If NumRegElim > 0 Then
                'Ha selecionado alguno de los dos, o los dos
                
                'insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,`trabajador`,`adjuntos`) values ( '1','0','','',NULL,NULL,NULL,NULL)
                Sql = "select fechahora,enviado,email,asunto,adjuntos from scrmmail"
                Sql = Sql & " WHERE codclien=" & vCRM.CodClien
                 Sql = Sql & " AND fechahora>=" & DBSet(F, "F")
                If NumRegElim = 1 Or NumRegElim = 2 Then
                    cad = "1"
                    If NumRegElim = 2 Then cad = "0"
                    'Ha selecionado solo una de las dos
                    Sql = Sql & " AND enviado = " & cad
                End If
                NumRegElim = Contador
                
            
            
            
                RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                    NumRegElim = NumRegElim + 1
                    Sql = "insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,"
                    Sql = Sql & "`trabajador`,`adjuntos`) values ( " & vUsu.Codigo & "," & NumRegElim & ",1,"  '1.email
                    Sql = Sql & DBSet(RS!fechahora, "FH") & ","
                    'En sllama siempre son RECIBIDAS
                    If Val(RS!Enviado) = 1 Then
                        Sql = Sql & "'Enviado',"
                    Else
                        Sql = Sql & "'Recibido',"
                    End If
                    cad = DBLetMemo(RS!asunto)
                    cad = Replace(cad, vbCrLf, " ")
                    Sql = Sql & DBSet(cad, "T", "S") & ","
                    'Trabajador
                    Sql = Sql & DBSet(RS!email, "T", "S") & ","
                    'En adjuntos guardare el tipop llamada
                    cad = "'*'"
                    If DBLet(RS!adjuntos, "T") = "" Then cad = "NULL"
                    Sql = Sql & cad & ")"
                    
                    conn.Execute Sql
                    RS.MoveNext
                Wend
                RS.Close
                'Ha metido algun dato
                'If NumRegElim > Contador Then comer(5) = True   'tiene datos
                Contador = NumRegElim
        End If
            
        

        
    End If
    
    
End Sub










Private Sub GenerarDatosAdmon()
Dim Impor1 As Currency
Dim Base As Currency
Dim cad As String
Dim F As Date

    Donde = "Administracion"
    'Volumen facturacion
    J = DevuelveIndiceNodo("adm1")
    If HayKprocesarNodo(J, F) Then
        Donde = "Volumen fact."
        
        'Volumen facturacion
        Sql = "select year(fecfactu) anyo,sum(totalfac) totalfac from scafac "
        Sql = Sql & " where codclien=" & Text1.Text & " and codtipom <>'FAZ' and codtipom<>'FRT' and codtipom<>'FRC' "
        Sql = Sql & " AND fecfactu>='" & Format(F, FormatoFecha) & "'"
        'Aqui va lo de ultimos años
        Sql = Sql & " group by 1 order by 1,2"
        
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        
        While Not RS.EOF
            cad = ""
        
            NumRegElim = NumRegElim + 1
            Impor1 = DBLet(RS!TotalFac, "N")
            
            Sql = "insert into `tmpcrmtesor` (`codusu`,`codigo`,`importe`,`anyotxt`,`variacion`)"
            Sql = Sql & " values (" & vUsu.Codigo & "," & NumRegElim & "," & TransformaComasPuntos(CStr(Impor1)) & ",'"
            
            If Val(RS!Anyo) = Year(Now) Then
                'Valor actual.
                Sql = Sql & "actual',"
                'Cambio la base para comprar con el mismo periodo del actual
                
                cad = "codtipom <>'FAZ' and codtipom<>'FRT' and codtipom<>'FRC' and "
                cad = cad & " fecfactu>='" & Year(Now) - 1 & "-01-01' and "
                cad = cad & " fecfactu<='" & Year(Now) - 1 & "-" & Format(Now, "mm-dd") & "' AND codclien "
                cad = DevuelveDesdeBD(conAri, "sum(totalfac)", "scafac", cad, Text1.Text)
                If cad = "" Then cad = "0"
                Base = CCur(cad)
                If NumRegElim > 1 And Base <> 0 Then
                    Impor1 = CStr(((100 * Impor1) / Base) - 100)
                    cad = Format(Impor1, FormatoPorcen) & "% sobre misma fecha año anterior"
                Else
                    cad = ""
                End If
            Else
                'Otro año cualquiera
                 Sql = Sql & RS!Anyo & "',"
                If NumRegElim > 1 And Base <> 0 Then
                    Impor1 = CStr(((100 * Impor1) / Base) - 100)
                    cad = Format(Impor1, FormatoPorcen) & "%"
                End If
                 
            End If
            Base = DBLet(RS!TotalFac, "N")
            Sql = Sql & "'" & cad & "')"
          

            conn.Execute Sql
            RS.MoveNext
        Wend
        RS.Close
        'If NumRegElim > 0 Then admon(1) = True
    
    
    End If
    
    
    J = DevuelveIndiceNodo("adm2")
    If HayKprocesarNodo(J, F) Then
        Donde = "Cobros pendientes"
        'insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,`importe`,`observa`) values ( '1','0','0','','','',NULL,NULL)
        If vParamAplic.ContabilidadNueva Then
            Sql = "SELECT cobros.*,nomforpa FROM cobros INNER JOIN formapago ON cobros.codforpa=formapagos.codforpa "
            Sql = Sql & " WHERE cobros.codmacta = '" & vCRM.codmacta & "'"
            'JUNIO 2010
            'PONGO Toooodos los vtos es decir, comento la linea inferior
            'SQL = SQL & " AND fecvenci <= ' " & Format(Now, FormatoFecha) & "' "
            'SQL = SQL & " AND (sforpa.tipforpa between 0 and 3) ORDER BY fecvenci desc"
            Sql = Sql & "  AND recedocu=0 ORDER BY fecvenci desc"
        
        Else
            Sql = "SELECT scobro.*,nomforpa FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
            Sql = Sql & " WHERE scobro.codmacta = '" & vCRM.codmacta & "'"
            'JUNIO 2010
            'PONGO Toooodos los vtos es decir, comento la linea inferior
            'SQL = SQL & " AND fecvenci <= ' " & Format(Now, FormatoFecha) & "' "
            'SQL = SQL & " AND (sforpa.tipforpa between 0 and 3) ORDER BY fecvenci desc"
            Sql = Sql & "  AND recedocu=0 ORDER BY fecvenci desc"
        End If
        NumRegElim = 0
        RS.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        Base = 0
        Impor1 = 0
        
        While Not RS.EOF
              'trozo copiado d ela funcion de ver cobros pdtes
          If DBLet(RS!Devuelto, "N") = 1 Then
                'SALE SEGURO (si no esta girado otra vez ¿no?
                'Si esta girado otra vez tendra impcobro, con lo cual NO tendra diferencia de importes
                Impor1 = RS!ImpVenci + DBLet(RS!Gastos, "N") - DBLet(RS!impcobro, "N")
                
            Else
                'Si esta recibido NO lo saco
                If Val(RS!recedocu) = 1 Then
                    Impor1 = 0
                Else
                    'NO esta recibido. Si tiene diferencia
                    Impor1 = RS!ImpVenci + DBLet(RS!Gastos, "N") - DBLet(RS!impcobro, "N")
            
                End If
          End If
          If Impor1 <> 0 Then
                NumRegElim = NumRegElim + 1
                Sql = "insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,"
                Sql = Sql & "`importe`,`observa`,forpa) values ( "
                Sql = Sql & vUsu.Codigo & "," & NumRegElim & ",0,'"
                
                If vParamAplic.ContabilidadNueva Then
                    Sql = Sql & RS!numSerie & Format(RS!NumFactu, "000000")
                    If RS!FecVenci > Now Then Sql = Sql & " *"
                    Sql = Sql & "','" & Format(RS!FecFactu, FormatoFecha)
                Else
                    Sql = Sql & RS!numSerie & Format(RS!Codfaccl, "000000")
                    If RS!FecVenci > Now Then Sql = Sql & " *"
                    Sql = Sql & "','" & Format(RS!fecfaccl, FormatoFecha)
                End If
                
                Sql = Sql & "','" & Format(RS!FecVenci, FormatoFecha) & "'," & TransformaComasPuntos(CStr(Impor1)) & ",NULL"
                'Mayo 2010
                'Con forma de pago
                Sql = Sql & ",'" & Format(RS!codforpa, "000") & " - " & DevNombreSQL(RS!nomforpa) & "')"
                conn.Execute Sql
          End If
          RS.MoveNext

            
        
        Wend
        RS.Close
        
        'If NumRegElim > 0 Then admon(2) = True   'tiene datos
         
        
    End If
    
    J = DevuelveIndiceNodo("adm3")
    If HayKprocesarNodo(J, F) Then
        Donde = "Hco reclamas"
        
        If vParamAplic.ContabilidadNueva Then
            Sql = "SELECT reclama.codigo,reclama_facturas.numserie,reclama_facturas.numfactu codfaccl,reclama_facturas.fecfactu fecfaccl,fecreclama,reclama_facturas.impvenci,reclama.codmacta,reclama.observaciones from reclama_facturas inner join reclama on reclama.codigo = reclama_facturas.codigo "
            Sql = Sql & " WHERE codmacta = '" & vCRM.codmacta & "'"
            Sql = Sql & " AND fecreclama >= '" & Format(F, FormatoFecha) & "' "
        Else
        
            Sql = "SELECT codigo,numserie,codfaccl,fecfaccl,fecreclama,impvenci,codmacta,observaciones from shcocob "
            Sql = Sql & " WHERE codmacta = '" & vCRM.codmacta & "'"
            Sql = Sql & " AND fecreclama >= '" & Format(F, FormatoFecha) & "' "
            
        End If
        'SQL = SQL & " AND (sforpa.tipforpa between 0 and 3) ORDER BY fecvenci desc"
        J = CInt(NumRegElim) 'pk puede que haya metidos de cobros. NO reseteo Numregelim
        
        RS.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            NumRegElim = NumRegElim + 1
            Sql = "insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,`importe`,`observa`) values ("
            Sql = Sql & vUsu.Codigo & "," & NumRegElim & ",1,'"
            Sql = Sql & DBLet(RS!numSerie, "T") & Format(DBLet(RS!Codfaccl, "N"), "000000") & "','"
            Sql = Sql & Format(RS!fecfaccl, FormatoFecha) & "','" & Format(RS!fecreclama, FormatoFecha) & "',"
            Sql = Sql & TransformaComasPuntos(RS!ImpVenci) & ",'"
            cad = DBLetMemo(RS!Observaciones)
            cad = Replace(cad, vbCrLf, " ")
            Sql = Sql & DevNombreSQL(cad) & "')"
            conn.Execute Sql
            RS.MoveNext
        Wend
        RS.Close
        
        'Ha metido algun dato
        'If NumRegElim > J Then admon(3) = True   'tiene datos
    End If
    
    
    'Vere si teiene manteinimeots para mostrar/o no en el rpt
    J = DevuelveIndiceNodo("adm4")
    If HayKprocesarNodo(J, F) Then
        Sql = "Select count(*) from scaman where codclien = " & Text1.Text
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        If Not RS.EOF Then NumRegElim = DBLet(RS.Fields(0), "N")
        RS.Close
        'If NumRegElim > 0 Then admon(4) = True
    End If
End Sub


Private Sub GenerarDatosSAT()
Dim cad As String
Dim Contador As Long
Dim F As Date

   

    Donde = "SAT"
    'Volumen facturacion
    J = DevuelveIndiceNodo("sat1")
    If HayKprocesarNodo(J, F) Then
        Donde = "Frecuencias"
        
    
    End If
    
    
    J = DevuelveIndiceNodo("sat2")
    If HayKprocesarNodo(J, F) Then
        Donde = "Albaranes reparacion"
        
    End If
    
    J = DevuelveIndiceNodo("sat3")
    If HayKprocesarNodo(J, F) Then
        Donde = "Avisos pdtes de cerrar"
        
    End If
    
    J = DevuelveIndiceNodo("sat4")
    If HayKprocesarNodo(J, F) Then
        Donde = "Equipos pendientes reparar"
        
    End If
    
End Sub



Private Sub FijarNodo3(ByRef Nod, Padre As String, Clave As String, LlevaFecha As Boolean, Anyo As Boolean, Texto As String)
Dim Aux As String
Dim Fecha As Date
Dim leido As Boolean

    'Primero AÑADO EL NODO
    Set Nod = TV1.Nodes.Add(Padre, tvwChild, Clave)
    Nod.Text = Texto
    
    'Veo si estan leido los datos de preselccion
    leido = False
    If Not DatosGuardados Is Nothing Then
        If DatosGuardados.Count > 0 Then leido = True
    End If
        
    If leido Then
        If Nod.Index > DatosGuardados.Count Then
            leido = False
        End If
    End If
    
    
    If Not leido Then
        Nod.Checked = True
        
    Else
        Nod.Checked = RecuperaValor(DatosGuardados(Nod.Index), 1) = "1"
        'Debug.Print Nod.Text & " " & Nod.Checked
    End If
    
    If LlevaFecha Then
        If Not leido Then
            Fecha = "01/01/2010"
        Else
            Aux = RecuperaValor(DatosGuardados(Nod.Index), 2)
            If Aux = "" Then
                Aux = "01/10/2010"
            Else
                If Not IsDate(Aux) Then Aux = "01/01/2010"
            End If
            Fecha = Aux
            
        End If
        
        Aux = Nod.Text & "   ["
        If Anyo Then
            Aux = Aux & Year(Fecha)
        Else
            Aux = Aux & Format(Fecha, "dd/mm/yyyy")
        End If
        Aux = Aux & "]"
        Nod.Text = Aux
    End If
End Sub



''''''Private Sub FijarNodoConFecha(ByRef Nod, Anyo As Boolean)
''''''Dim Aux As String
''''''Dim Fecha As Date
''''''
''''''    'Leeriamos de datos guardados
''''''    If False Then
''''''
''''''    Else
''''''        Fecha = "01/01/2010"
''''''    End If
''''''
''''''
''''''
''''''
''''''    Aux = Nod.Text & "   ["
''''''    If Anyo Then
''''''        Aux = Aux & Year(Fecha)
''''''    Else
''''''        Aux = Aux & Format(Fecha, "dd/mm/yyyy")
''''''    End If
''''''    Aux = Aux & "]"
''''''    Nod.Text = Aux
''''''End Sub





'Dado un NODO
Private Function HayKprocesarNodo(indice As Integer, ByRef Fecha As Date) As Boolean
Dim I As Integer
Dim Valor As String
Dim TieneFecha As Boolean
Dim CadenaFecha As String
Dim CadenaVisible As String
Dim Aux As String
Dim NodoOfertaPedidoAlbaran As Boolean


    Fecha = CDate("01/01/2007")
    I = InStr(1, TV1.Nodes(indice).Text, "[")
    TieneFecha = I > 0
    
    
    If TieneFecha Then
        Valor = Mid(TV1.Nodes(indice).Text, I + 1)
        Valor = Mid(Valor, 1, Len(Valor) - 1)
    End If
    
    'Sabremos si esta marcado o no
    HayKprocesarNodo = TV1.Nodes(indice).Checked
    
    
    'Si es un NODO padre no leo mas, ya que no hay campos visibles para ellos
    If TV1.Nodes(indice).Parent Is Nothing Then Exit Function
    

    NodoOfertaPedidoAlbaran = False
    If indice = 7 Or indice = 8 Or indice = 9 Then NodoOfertaPedidoAlbaran = True
        
    If NodoOfertaPedidoAlbaran Then
        CadenaVisible = RecuperaValor(TV1.Nodes(indice).Tag, 1)
        If CadenaVisible <> "" Then
            'El nodo esta marcado para imprimir
            If Not CadenaOfePedAlb(indice, Aux) Then
                CadenaVisible = ""  'para qe no imprima

            End If
        End If
        
    Else
        CadenaVisible = RecuperaValor(TV1.Nodes(indice).Tag, 1)
    End If  'para los nodos de ofer,ped alb y el resto
    
    
    If CadenaVisible <> "" Then
        cadParam2 = cadParam2 & CadenaVisible & "=" & Val(Abs(TV1.Nodes(indice).Checked)) & "|"
    Else
       ' MsgBox "No hay campo visible en el rpt", vbInformation
    End If
    CadenaFecha = RecuperaValor(TV1.Nodes(indice).Tag, 2)
    'FECHA
    'Si hay fecha
    If CadenaFecha <> "" Then
        If Len(Valor) = 4 Then
            'Es solo el año
            cadParam2 = cadParam2 & CadenaFecha & "=" & Valor
            Fecha = CDate("01/01/" & Valor)
        Else
            cadParam2 = cadParam2 & CadenaFecha & "=" & "Date(" & Year(Valor) & ", " & Month(Valor) & ", " & Day(Valor) & ")"
            Fecha = CDate(Valor)
        End If
        cadParam2 = cadParam2 & "|"
    Else
        If Valor <> "" Then MsgBox "Hay fecha y no hay campo en el rpt para indicarla", vbInformation
    End If
             
        
    
        
End Function

Private Sub Configuracion(Leer As Boolean)
    Sql = App.Path & "\crmdef.dat"
    If Leer Then
        If Dir(Sql, vbArchive) <> "" Then
            'Lo cargo todo
            If Not ProcFicheroConfig(True) Then Set DatosGuardados = Nothing
        End If
    Else
        ProcFicheroConfig False
    
    End If
End Sub



Private Function ProcFicheroConfig(Leer As Boolean) As Boolean
Dim TieneF As Boolean
Dim I As Integer
Dim Aux As String
Dim NF As Integer

    On Error GoTo eLeerFicheroConfig
    ProcFicheroConfig = False
    NF = FreeFile
    If Leer Then
        Open Sql For Input As #NF
        
        Set DatosGuardados = New Collection
        Sql = ""
        While Not EOF(NF)
            Line Input #NF, Sql
            DatosGuardados.Add Sql
        Wend
        Close #NF
        
    Else
    
        Open Sql For Output As #NF
        For J = 1 To TV1.Nodes.Count
            I = InStr(1, TV1.Nodes(J), "[")
            TieneF = I > 0
            
            Sql = Abs(TV1.Nodes(J).Checked) & "|"
            If TieneF Then
                Aux = Mid(TV1.Nodes(J).Text, I + 1)
                Aux = Mid(Aux, 1, Len(Aux) - 1)
                If Len(Aux) = 4 Then Aux = "01/01/" & Aux
                
            Else
                Aux = ""
            End If
            Sql = Sql & Aux & "|"
            Print #NF, Sql
        Next J
        Close #NF
    End If
    
    ProcFicheroConfig = True
    
    Exit Function
eLeerFicheroConfig:
    MuestraError Err.Number, "LeerFicheroConfig"
    TrataCerrarFichero NF
End Function

Private Sub TrataCerrarFichero(ByRef NFF As Integer)
    On Error Resume Next
    Close #NFF
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaDatosAux()
Dim C As Byte
    C = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    tv2.Nodes.Clear
    tv3.Nodes.Clear
    If Text1.Text <> "" Then
        Set RS = New ADODB.Recordset
        lblInd.Caption = ""
        CargaImpresionAuxiliar
        lblInd.Caption = ""
        Set RS = Nothing
    End If
    Screen.MousePointer = C
End Sub

Private Sub CargaImpresionAuxiliar()
Dim PpalInsertado As Boolean
Dim N

    
        
    '***********************************************************************
    'OFERTAS
    lblInd.Caption = "OFERTAS"
    lblInd.Refresh
    Sql = "Select numofert,fecofert from scapre where codclien =" & Text1.Text & " AND "
    Sql = Sql & DevFecha(7, "fecofert")
    Sql = Sql & " ORDER BY fecofert"
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PpalInsertado = False
    While Not RS.EOF
        If Not PpalInsertado Then
            Set N = tv2.Nodes.Add(, , "OFE")
            N.Text = "OFERTAS"
            N.Bold = True
            N.Checked = True
            
            Set N = tv3.Nodes.Add(, , "OFE")
            N.Text = "OFERTAS"
            N.Bold = True
            N.Checked = True
            PpalInsertado = True
        End If
        
        Sql = Format(RS!NumOfert, "000000") & "  -  " & Format(RS!fecofert, "dd/mm/yyyy")
        Set N = tv2.Nodes.Add("OFE", tvwChild)
        N.Text = Sql
        N.Checked = True
        Set N = tv3.Nodes.Add("OFE", tvwChild)
        N.Text = Sql
        N.Checked = True
        RS.MoveNext
    Wend
    RS.Close
    If PpalInsertado Then
        tv2.Nodes(N.Index).EnsureVisible
        tv3.Nodes(N.Index).EnsureVisible
    End If
    
    
    
    '***********************************************************************
    'PEDIDO
    lblInd.Caption = "PEDIDOS"
    lblInd.Refresh
    Sql = "Select numpedcl,fecpedcl from scaped where codclien =" & Text1.Text & " AND "
    Sql = Sql & DevFecha(8, "fecpedcl")
    Sql = Sql & " ORDER BY fecpedcl"
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PpalInsertado = False
    While Not RS.EOF
        If Not PpalInsertado Then
            Set N = tv2.Nodes.Add(, , "PED")
            N.Text = "PEDIDOS"
            N.Bold = True
            N.Checked = True
            N.ForeColor = &H4000&
            Set N = tv3.Nodes.Add(, , "PED")
            N.Text = "PEDIDOS"
            N.Bold = True
            N.Checked = True
            PpalInsertado = True
            N.ForeColor = &H4000&
        End If
        
        Sql = Format(RS!Numpedcl, "000000") & "  -  " & Format(RS!fecpedcl, "dd/mm/yyyy")
        Set N = tv2.Nodes.Add("PED", tvwChild)
        N.Text = Sql
        N.Checked = True
        Set N = tv3.Nodes.Add("PED", tvwChild)
        N.Text = Sql
        N.Checked = True
        RS.MoveNext
    Wend
    RS.Close
    If PpalInsertado Then
        tv2.Nodes(N.Index).EnsureVisible
        tv3.Nodes(N.Index).EnsureVisible
    End If
    
    
    '***********************************************************************
    'ALBARANES
    lblInd.Caption = "ALBARANES"
    lblInd.Refresh
    Sql = "Select codtipom,numalbar,fechaalb from scaalb where "
    Sql = Sql & DevFecha(9, "fechaalb")
    Sql = Sql & " AND codtipom <>'ALZ' and codtipom<>'ALR' and "
    Sql = Sql & " codClien = " & Text1.Text
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PpalInsertado = False
    While Not RS.EOF
        If Not PpalInsertado Then
            Set N = tv2.Nodes.Add(, , "ALB")
            N.Text = "ALBARANES"
            N.Bold = True
            N.Checked = True
            N.ForeColor = &H80&
            Set N = tv3.Nodes.Add(, , "ALB")
            N.Text = "ALBARANES"
            N.Bold = True
            N.Checked = True
            N.ForeColor = &H80&
            PpalInsertado = True
        End If
        
        Sql = RS!codtipom & Format(RS!NumAlbar, "000000") & "  -  " & Format(RS!FechaAlb, "dd/mm/yy")
        Set N = tv2.Nodes.Add("ALB", tvwChild)
        N.Checked = True
        N.Text = Sql
        Set N = tv3.Nodes.Add("ALB", tvwChild)
        N.Text = Sql
        N.Checked = True
        
        RS.MoveNext
    Wend
    RS.Close
    If PpalInsertado Then
        tv2.Nodes(N.Index).EnsureVisible
        tv3.Nodes(N.Index).EnsureVisible
    End If
    
End Sub


Private Function DevFecha(indice As Integer, CampoBD As String) As String
Dim I As Integer
Dim F As String
    F = CDate("01/01/1900")
    I = InStr(1, TV1.Nodes(indice).Text, "[")
    If I > 0 Then F = Mid(TV1.Nodes(indice), I + 1, 10)
    DevFecha = CampoBD & " >= '" & Format(F, FormatoFecha) & "'"
End Function

Private Sub tv2_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    If PrimeraVez Then Exit Sub
    
    'Pong el nodo en el tv3 chcec(unche
    tv3.Nodes(Node.Index).Checked = Node.Checked
    
    Dim CH As Boolean
    
    If Node.Checked Then
        If Not Node.Parent Is Nothing Then Node.Parent.Checked = True
    End If
    CH = Node.Checked
    CheckSubNodo Node, CH, True
    
    
    Err.Clear
End Sub

Private Sub tv3_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim CH As Boolean
    If PrimeraVez Then Exit Sub
    
    If Node.Checked Then
        If Not Node.Parent Is Nothing Then Node.Parent.Checked = True
    End If
    
    
    CH = Node.Checked
    CheckSubNodo Node, CH, False
    
    
End Sub


Private Function CadenaOfePedAlb(Index As Integer, CadenaSQL_ As String) As Boolean
Dim J As Integer
Dim N As Node
Dim Pad As Node
Dim C2 As String

    CadenaOfePedAlb = False
    CadenaSQL_ = "-1"
    If tv2.Nodes.Count <= 1 Then Exit Function  'si no hay modos, nos piaramos
    
    Set Pad = tv2.Nodes(1)
    
    Select Case Index
    Case 7
        'OFERTAS
        If Pad.Key <> "OFE" Then Exit Function
        Set N = Pad.Child
        CadenaSQL_ = ""
        While Not N Is Nothing
            
            If N.Checked Then
                J = InStr(1, N.Text, "-")
                If J > 0 Then CadenaSQL_ = CadenaSQL_ & ", " & Trim(Mid(N.Text, 1, J - 1))
            End If
            Set N = N.Next
       Wend
 
        
    Case 8
        J = 0
        While J = 0
            If Pad.Key = "PED" Then
                J = 1
            Else
                Set Pad = Pad.Next
                If Pad Is Nothing Then J = 1
            End If
        Wend
        
        If Pad Is Nothing Then Exit Function
        Set N = Pad.Child
        CadenaSQL_ = ""
        While Not N Is Nothing
            If N.Checked Then
                J = InStr(1, N.Text, "-")
                If J > 0 Then CadenaSQL_ = CadenaSQL_ & ", " & Trim(Mid(N.Text, 1, J - 1))
            End If
            Set N = N.Next
        Wend

       
       
       
    Case 9
         'ALBARANES
         J = 0
         While J = 0
             If Pad.Key = "ALB" Then
                 J = 1
             Else
                 Set Pad = Pad.Next
                 If Pad Is Nothing Then J = 1
             End If
         Wend
         
         If Pad Is Nothing Then Exit Function
         Set N = Pad.Child
         CadenaSQL_ = ""
         While Not N Is Nothing
             If N.Checked Then
                J = InStr(1, N.Text, "-")
                If J > 0 Then
                    C2 = Trim(Mid(N.Text, 1, J - 1))
                    CadenaSQL_ = CadenaSQL_ & ", ('" & Mid(C2, 1, 3) & "'," & Mid(C2, 4) & ")"
                End If
             End If
             Set N = N.Next
         Wend

    End Select
    
          'Ninguno seleccionado
       If InStr(1, CadenaSQL_, ",") = 0 Then
            CadenaOfePedAlb = False
            CadenaSQL_ = "-1"
       Else
            CadenaSQL_ = Mid(CadenaSQL_, 2)
            CadenaOfePedAlb = True
       
            InsertarEnTmpsOfePedAlb Index, CadenaSQL_
       
       
       
       
       
       
       
       End If
    
End Function




Private Sub InsertarEnTmpsOfePedAlb(indice As Integer, ByRef Conjunto As String)
Dim C As String
Dim C2 As String
    Select Case indice
    Case 7
        C = "Select * from scapre where numofert in (" & Conjunto & ") ORDER by fecofert asc"
        RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            '                               secuencial   ofe/ped/alb  iden     dpto     vacio
            NumRegElim = NumRegElim + 1
            C = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`nombre1`,`nombre2`,`nombre3`,`importe1`,`fecha1`,`fecha2`)"
            
            'ANTES MAYO2010
'            C = C & " VALUES (" & vUsu.Codigo & "," & NumRegElim & ",1,"
'            'identificador
'            C = C & Format(RS!NumOfert, "000000") & ","
'
            'AHORA
            C = C & " VALUES (" & vUsu.Codigo & "," & RS!NumOfert & ",1,"
            'identificador
            C = C & Format(NumRegElim, "000000") & ","
                        
            If IsNull(RS!CodDirec) Then
                C2 = "NULL"
            Else
                C2 = "'" & RS!CodDirec & "   " & DevNombreSQL(DBLet(RS!nomdirec, "T")) & "'"
            End If
            '               vacio de momento
            C = C & C2 & ",NULL,"
            C2 = DevuelveDesdeBD(conAri, "sum(importel)", "slipre", "numofert", RS!NumOfert, "N")
            If C2 = "" Then C2 = "0"
            C = C & TransformaComasPuntos(C2)
            C = C & "," & DBSet(RS!fecofert, "F") & "," & DBSet(RS!FecEntre, "F") & ")"
            conn.Execute C
            RS.MoveNext
        Wend
        RS.Close
        
    Case 8
        C = "Select * from scaped where numpedcl IN (" & Conjunto & ") ORDER by fecpedcl asc"
        RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            '                               secuencial   ofe/ped/alb  iden     dpto     vacio
            NumRegElim = NumRegElim + 1
            C = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`nombre1`,`nombre2`,`nombre3`,`importe1`,`fecha1`,`fecha2`)"
            C = C & " VALUES (" & vUsu.Codigo & "," & NumRegElim & ",2,"  '2 de pedido
            'identificador
            C = C & Format(RS!Numpedcl, "000000") & ","
            If IsNull(RS!CodDirec) Then
                C2 = "NULL"
            Else
                C2 = "'" & RS!CodDirec & "   " & DBLet(RS!nomdirec, "T") & "'"
            End If
            '               vacio de momento
            C = C & C2 & ",NULL,"
            C2 = DevuelveDesdeBD(conAri, "sum(importel)", "sliped", "numpedcl", RS!Numpedcl, "N")
            If C2 = "" Then C2 = "0"
            C = C & TransformaComasPuntos(C2)
            C = C & "," & DBSet(RS!fecpedcl, "F") & "," & DBSet(RS!FecEntre, "F") & ")"
            conn.Execute C
            RS.MoveNext
        Wend
        RS.Close
    
    Case 9
        C = "Select * from scaalb where (codtipom,numalbar)  IN (" & Conjunto & ") ORDER by fechaalb,codtipom asc"
        RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            '                               secuencial   ofe/ped/alb  iden     dpto     vacio
            NumRegElim = NumRegElim + 1
            C = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`nombre1`,`nombre2`,`nombre3`,`importe1`,`fecha1`,`fecha2`)"
            C = C & " VALUES (" & vUsu.Codigo & "," & NumRegElim & ",3,"  '3 de alb
            'identificador
            C = C & "'" & RS!codtipom & Format(RS!NumAlbar, "000000") & "',"
            If IsNull(RS!CodDirec) Then
                C2 = "NULL"
            Else
                C2 = "'" & RS!CodDirec & "   " & DBLet(RS!nomdirec, "T") & "'"
            End If
            '               vacio de momento
            C = C & C2 & ",NULL,"
            C2 = DevuelveDesdeBD(conAri, "sum(importel)", "slialb", "codtipom = '" & RS!codtipom & "' AND numalbar", RS!NumAlbar, "N")
            If C2 = "" Then C2 = "0"
            C = C & TransformaComasPuntos(C2)
            C = C & "," & DBSet(RS!FechaAlb, "F") & ",NULL)"
            conn.Execute C
            RS.MoveNext
        Wend
        RS.Close
    
    End Select
        
End Sub

Private Sub ImprimirDocumentosAuxiliares()
Dim Cuantos As Integer
Dim N As Node

    If tv3.Nodes.Count = 0 Then Exit Sub
    
    
    Set N = tv3.Nodes(1)
    Sql = ""
    For J = 1 To tv3.Nodes.Count
        If tv3.Nodes(J).Checked Then
            If Not tv3.Nodes(J).Parent Is Nothing Then
                Sql = "OK"   'Si es nodo hijo
                Exit For
            End If
        End If
    Next
    
    If Sql = "" Then
      '  MsgBox "Ningun datos seleccionado", vbExclamation
        J = 0
    Else
        J = 1
        Sql = "Va a imprimir las ofertas/pedidos/albaranes seleccionados" & vbCrLf & vbCrLf
        Sql = Sql & "¿Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then J = 0
    End If
    If J = 0 Then Exit Sub
    
    Set N = tv3.Nodes(1)
    While Not N Is Nothing
        ImprimirReports N
        
        Set N = N.Next
    Wend
    
End Sub


'       0- Ofertas   1-Pedidos   2-Albaranes
Private Sub ImprimirReports(ByRef NodoPadre As Node)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim devuelve As String, campo As String
Dim OpcRPT As Integer
Dim numParam As Byte
Dim cadFormula As String
Dim N As Node
Dim AntiguoTipmov As String

'Dim campo1 As String, campo2 As String, campo3 As String
    
    J = 0
    Set N = NodoPadre.Child
    While Not (N Is Nothing)
        If N.Checked Then J = 1
        Set N = N.Next
    Wend
    
    If J = 0 Then Exit Sub 'No hay ninguno
  
    '===================================================
    '============ PARAMETROS ===========================
    Select Case NodoPadre.Key
        Case "PED"
            indRPT = 7 '7: Pedidos de Clientes
            OpcRPT = 38  'impreison pedidos
            
        Case "OFE"
            indRPT = 5
            OpcRPT = 31
        Case Else
            'NodoPadre .key ="ALB"
            indRPT = 10
            OpcRPT = 45
    End Select
    numParam = 0
    cadParam2 = ""
    If Not PonerParamRPT(indRPT, cadParam2, numParam, Donde, pImprimeDirecto, pPdfRpt) Then Exit Sub
     
    
    
        'Añadimos a los parametros el tipo de IVA que se aplica a ese cliente (para saber si esta exento o no de IVA)
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1.Text, "N")
        If devuelve <> "" Then
            cadParam2 = cadParam2 & "pTipoIVA=" & devuelve & "|"
            numParam = numParam + 1
        End If
        
        'PORTES
        cadParam2 = cadParam2 & "vPortes=""" & vParamAplic.ArtPortes & """|"
        numParam = numParam + 1
    

    cadFormula = ""
    Sql = ""
    AntiguoTipmov = ""
    Set N = NodoPadre.Child
    While Not (N Is Nothing)
        If N.Checked Then
            
            Select Case NodoPadre.Key
            Case "PED"
                If Sql = "" Then Sql = "{scaped.codclien} = " & Text1.Text & " AND {scaped.numpedcl} IN "
                J = InStr(1, N.Text, "-")
                cadFormula = cadFormula & ", " & Trim(Mid(N.Text, 1, J - 1))
                
            Case "OFE"
                'Añado el parametro de carta NO
                If cadFormula = "" Then
                    'Es la 1era vez k entra aqui
                    cadParam2 = cadParam2 & "pCodCarta=0|"
                    numParam = numParam + 1
                    Sql = "{scapre.codclien} = " & Text1.Text & " AND {scapre.numofert} IN "
                End If
                 J = InStr(1, N.Text, "-")
                 cadFormula = cadFormula & ", " & Trim(Mid(N.Text, 1, J - 1))
                 
                 
            Case Else
                If Mid(N.Text, 1, 3) <> AntiguoTipmov Then
                    If AntiguoTipmov <> "" Then Imprime cadFormula, OpcRPT, cadParam2, numParam
                    cadFormula = ""
                    AntiguoTipmov = Mid(N.Text, 1, 3)
                End If
                'ALBARANES
                '{scaalb.codtipom}='ALV' AND ({scaalb.numalbar}=14)
                If cadFormula = "" Then
                    'Es la 1era vez k entra aqui
'--[Monica]arituclo de reciclado pasa a ser reutilizado como articulo de gastos de administracion
'                    'PUNTO VERDE
'                    cadParam2 = cadParam2 & "PuntoVerde=""" & vParamAplic.ArtReciclado & """|"
'                    numParam = numParam + 1
                    
                    'Si se imprimen importes y/o
                    devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", Text1.Text, "N")
                    If devuelve = "" Then devuelve = "0"
                    ' 0 "Todo"
                    ' 1 "Cantidad y Precio"
                    ' 2 "Cantidad"
                    cadParam2 = cadParam2 & "Albarcon=" & devuelve & "|"
                    numParam = numParam + 1
                    
                    Sql = "{scaalb.codclien} = " & Text1.Text & " AND {scaalb.codtipom}= '" & AntiguoTipmov & "' AND {scaalb.numalbar} IN "
                    
                End If
                 J = InStr(1, N.Text, "-")
                 cadFormula = cadFormula & ", " & Trim(Mid(N.Text, 4, J - 4))
                
                
            End Select
            
            
        End If
        Set N = N.Next
        
    Wend
    
    Imprime cadFormula, OpcRPT, cadParam2, numParam
            
            
       
            
    
End Sub





Private Sub Imprime(cadFormula As String, OpcRPT As Integer, cadParam As String, numParam As Byte)
        cadFormula = Mid(cadFormula, 2) 'quito la primera coma
        cadFormula = "[" & cadFormula & "]"
        cadFormula = Sql & cadFormula
    
         With frmImprimir
                    
                    .outTipoDocumento = 0
            '        If DatosEnvioMail <> "" Then
            '            .outTipoDocumento = RecuperaValor(DatosEnvioMail, 1)
            '            .outCodigoCliProv = RecuperaValor(DatosEnvioMail, 2)
            '            .outClaveNombreArchiv = RecuperaValor(DatosEnvioMail, 3)
            '        End If
                    .FormulaSeleccion = cadFormula
                    .OtrosParametros = cadParam2
                    .NumeroParametros = numParam
                    .SoloImprimir = True
                    .EnvioEMail = False
                    .Opcion = OpcRPT
                    .Titulo = "Datos auxiliares desde CRM"
                    If OpcRPT = 31 Then
                        .Titulo = .Titulo & "(OFERTAS)"
                    ElseIf OpcRPT = 38 Then
                        .Titulo = .Titulo & "(PEDIDOS)"
                    Else
                        .Titulo = .Titulo & "(ALBARANES)"
                    End If
                    .NombreRPT = Donde  'tendra el nomrtp
                    'If PonerNombrePDF Then .NombrePDF = cadPDFrpt
                    .ConSubInforme = True
                    .Show vbModal
                End With
                Me.Refresh
                DoEvents
                Screen.MousePointer = vbHourglass
                Espera 0.4
End Sub
