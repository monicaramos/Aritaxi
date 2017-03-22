VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlertas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alertas"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlertas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlertas.frx":6862
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5175
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Numero"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cod"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nombre"
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   9128
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   3
      Appearance      =   1
   End
End
Attribute VB_Name = "frmAlertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Sql As String
Dim F As Date


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon
    
    Set miRsAux = New ADODB.Recordset
    Set TreeView1.ImageList = Me.ImageList1
    Set ListView1.SmallIcons = frmPpal.ImgListPpal
    CargaTreeView
    Set TreeView1.SelectedItem = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaTreeView()
Dim no As Node
    
    TreeView1.Nodes.Clear
    
    'Para cada opcion de alertas vamos viendo si lo ponemos.
    Set no = TreeView1.Nodes.Add(, , "c1", "PEDIDOS CLIENTE")
    no.Image = 1
    no.Tag = 3
    If vParamAplic.avipedcli = 0 Then
        no.ForeColor = RGB(192, 192, 192)
        no.Tag = 1   'Pondremos el icono
        no.Image = 2
    End If
    
    
    Set no = TreeView1.Nodes.Add(, , "c2", "PEDIDOS PROVEEDORES")
    no.Image = 1
    no.Tag = 4
    If vParamAplic.avipedpro = 0 Then
        no.ForeColor = RGB(192, 192, 192)
        no.Image = 2
    End If
    
    
    Set no = TreeView1.Nodes.Add(, , "c3", "ALBARANES CLIENTE")
    no.Image = 1
    no.Tag = 7
    If vParamAplic.avialbcli = 0 Then
        no.ForeColor = RGB(192, 192, 192)
        no.Image = 2
    End If
    
    Set no = TreeView1.Nodes.Add(, , "c4", "ALBARANES PROVEEDORES")
    no.Image = 1
    no.Tag = 10
    If vParamAplic.avipedpro = 0 Then
        no.ForeColor = RGB(192, 192, 192)
        no.Image = 2
    End If
    
    Set no = TreeView1.Nodes.Add(, , "c5", "REPARACIONES")
    no.Image = 1
    no.Tag = 16
    If vParamAplic.avirepara = 0 Then
        no.ForeColor = RGB(192, 192, 192)
        no.Image = 2
    End If
    
    Set no = TreeView1.Nodes.Add(, , "c6", "AVISOS")
    no.Image = 1
    no.Tag = 1
    If vParamAplic.aviavisos = 0 Then
        no.ForeColor = RGB(192, 192, 192)
        no.Image = 2
    End If
    
    Set no = TreeView1.Nodes.Add(, , "c7", "MANTENIMIENTOS")
    no.Image = 1
    no.Tag = 12
    If vParamAplic.avimanteni = 0 Then
        no.ForeColor = RGB(192, 192, 192)
        no.Image = 2
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set miRsAux = Nothing
End Sub



Private Sub CargaListView(NumNod As Integer, LaImagen As Integer)
Dim It As ListItem

    On Error GoTo ECA
    FijaCadenaSQL NumNod
    If Sql = "" Then Exit Sub
    
    'SI no cargamos. SIiiiempre sera el mismo orden para los campos
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        Set It = ListView1.ListItems.Add(, "c" & NumRegElim)
        
        It.Text = miRsAux.Fields(0)
        It.SubItems(1) = Format(miRsAux.Fields(1), "dd/mm/yyyy")
        It.SubItems(2) = miRsAux.Fields(2)
        It.SubItems(3) = miRsAux.Fields(3)
        It.SubItems(4) = miRsAux.Fields(4)
        'IT.SubItems(4) = Format(miRsAux.Fields(4), FormatoImporte)
        
        It.SmallIcon = LaImagen
        miRsAux.MoveNext
        NumRegElim = NumRegElim + 1
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Exit Sub
ECA:
    MuestraError Err.Number, Sql
    Set miRsAux = Nothing
End Sub



Private Sub FijaCadenaSQL(Opcion As Integer)
    
    Sql = ""
    Select Case Opcion
    Case 1
        ' "PEDIDOS CLIENTE")
        'SQl = "select scaped.numpedcl,scaped.fecpedcl,scaped.codclien,scaped.nomclien,sum(importel) "
        'SQl = SQl & " from scaped,sliped WHERE scaped.numpedcl=sliped.numpedcl"
        'WHERE del alerta
        'f = DateAdd("d", -vParamAplic.avipedcli, Now)
        'SQl = SQl & " AND scaped.fecpedcl <= '" & Format(f, FormatoFecha) & "' group by 1 ORDER BY fecpedcl"
    
    Case 2
        '"PEDIDOS PROVEEDORES")
        Sql = "select scappr.numpedpr,scappr.fecpedpr,scappr.codprove,scappr.nomprove,sum(importel)"
        Sql = Sql & " from scappr,slippr WHERE scappr.numpedpr=slippr.numpedpr "
        F = DateAdd("d", -vParamAplic.avipedpro, Now)
        Sql = Sql & " AND scappr.fecpedpr <= '" & Format(F, FormatoFecha) & "' group by 1 ORDER BY fecpedpr"
    
    Case 3
        'Set NO = TreeView1.Nodes.Add(, , "c3", "ALBARANES CLIENTE")
        Sql = "select concat(scaalb.codtipom , scaalb.numalbar),scaalb.fechaalb,scaalb.codclien,scaalb.nomclien,sum(importel)"
        Sql = Sql & " from scaalb,slialb WHERE scaalb.numalbar=slialb.numalbar and scaalb.codtipom=slialb.codtipom"
        F = DateAdd("d", -vParamAplic.avialbcli, Now)
        Sql = Sql & " AND scaalb.fechaalb <= '" & Format(F, FormatoFecha) & "' group by 1 ORDER BY fechaalb"
    
    Case 4
        '"ALBARANES PROVEEDORES"
        Sql = "select  scaalp.numalbar,scaalp.fechaalb,scaalp.codprove,scaalp.nomprove,sum(importel)"
        Sql = Sql & " from scaalp,slialp WHERE scaalp.numalbar=slialp.numalbar and scaalp.fechaalb=slialp.fechaalb"
        Sql = Sql & " and scaalp.codprove=slialp.codprove"
        F = DateAdd("d", -vParamAplic.avialbpro, Now)
        Sql = Sql & " AND scaalp.fechaalb <= '" & Format(F, FormatoFecha) & "' group by 1 ORDER BY fechaalb"
    
    Case 5
        'Set NO = TreeView1.Nodes.Add(, , "c5", "REPARACIONES")
        F = DateAdd("d", -vParamAplic.avirepara, Now)
        Sql = "select scarep.numrepar,fecrepar,scarep.codclien,scarep.nomclien,if(imppresu1 is null,'0.0',imppresu1) from"
        Sql = Sql & " scarep,sclien where scarep.codclien=sclien.codclien  AND motivore is null "
        Sql = Sql & " AND scarep.fecrepar <= '" & Format(F, FormatoFecha) & "' group by 1 ORDER BY fecrepar"
    Case 6
        '"AVISOS"
        'f = DateAdd("d", -vParamAplic.aviavisos, Now)
        'SQl = "select numaviso,fechaavi,codclien,nomclien,'' from scaavi where situacio=0 "
        'SQl = SQl & " AND scaavi.fechaavi <= '" & Format(f, FormatoFecha) & "' group by 1 ORDER BY fechaavi"
    
    Case 7
        'Mantenimientos
        'f = DateAdd("d", -vParamAplic.avimanteni, Now)
        'SQl = "select scaman.nummante,concat(""01/"" , ulmesfac,""/" & Year(f) & """),scaman.codclien,nomclien,"" "" from scaman,sclien"
        'SQl = SQl & " Where scaman.CodClien = sclien.CodClien And ("
        'SQl = SQl & "(tipopago = 0 And ulmesfac < " & Month(f)
        'SQl = SQl & ") Or (tipopago = 1 And ulmesfac < " & Month(f) - 3
        'SQl = SQl & ") Or (tipopago = 2 And ulmesfac <" & Month(f) - 6
        'SQl = SQl & ") Or (tipopago = 3 And ulmesfac = 0))"
    
    
    
    End Select
    
End Sub


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    
    If ListView1.Tag = CStr(Node.Index) Then Exit Sub
    
    ListView1.ListItems.Clear
    ListView1.Tag = Node.Index
    If Node.Image <> 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    CargaListView Node.Index, CInt(Node.Tag)
    Screen.MousePointer = vbDefault
    
End Sub
