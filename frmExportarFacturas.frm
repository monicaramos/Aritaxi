VERSION 5.00
Begin VB.Form frmExportarFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar Facturas"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   Icon            =   "frmExportarFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   3960
      TabIndex        =   6
      Top             =   4140
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   4140
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Height          =   2910
      Left            =   150
      TabIndex        =   8
      Top             =   600
      Width           =   5640
      Begin VB.TextBox txtCodigo 
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
         Index           =   2
         Left            =   2790
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codsocio|||"
         Top             =   1755
         Width           =   1380
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   3
         Left            =   2790
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Código Postal|T|S|||clientes|codsocio|||"
         Top             =   2100
         Width           =   1380
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Incluir los ya traspasados"
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
         Left            =   270
         TabIndex        =   13
         Top             =   2490
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   240
         Width           =   2490
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   1
         Left            =   2790
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1290
         Width           =   1380
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   2790
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   915
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Socio / Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   2
         Left            =   270
         TabIndex        =   16
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1860
         TabIndex        =   15
         Top             =   2100
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1860
         TabIndex        =   14
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   270
         TabIndex        =   12
         Top             =   270
         Width           =   1605
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   1860
         TabIndex        =   11
         Top             =   930
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   1860
         TabIndex        =   10
         Top             =   1290
         Width           =   660
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2520
         ToolTipText     =   "Buscar fecha"
         Top             =   930
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   2520
         ToolTipText     =   "Buscar fecha"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   270
         TabIndex        =   9
         Top             =   720
         Width           =   1665
      End
   End
   Begin VB.Label lblInf 
      Alignment       =   2  'Center
      Caption         =   "Información del proceso"
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
      Left            =   240
      TabIndex        =   7
      Top             =   3615
      Width           =   5295
   End
End
Attribute VB_Name = "frmExportarFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1009


Public Tipo As Byte
    'Tipo:  0 Impresion de facturas

Dim DesdeFecha As Date
Dim Hastafecha As Date
Dim frmVis As frmVisReport

Dim nompath As String
Dim ConDetalle As Byte

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

Private Sub cmdAceptar_Click()
    If Not DatosOk() Then Exit Sub
    
    '-- Cargar facturas  entre las fechas seleccionadas
    Select Case Combo1(1).ListIndex
        Case 0 ' facturas liquidacion socio
            '[Monica]19/02/2018: Entra Cordoba
                '[Monica]19/11/2018: Entra Sevilla
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3 Then
                If MsgBox("¿ Desea imprimir el detalle de servicios ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                    ConDetalle = 0
                Else
                    ConDetalle = 1
                End If
            End If
            'hasta aquí
            CargaFacturasLiq DesdeFecha, Hastafecha
        
        Case 1 ' facturas de cliente
            '[Monica]19/02/2018: Entra Cordoba
            '[Monica]31/03/2014: en el caso de teletaxi pedimos si imprime o no detalle
                '[Monica]19/11/2018: Entra Sevilla
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3 Then
                If MsgBox("¿ Desea imprimir el detalle de servicios ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                    ConDetalle = 0
                Else
                    ConDetalle = 1
                End If
            End If
            'hasta aquí
            
            CargaFacturasCliente DesdeFecha, Hastafecha
            
        Case 2 ' facturas cuotas socios
            CargaFacturasCuotasSocio DesdeFecha, Hastafecha
            
        Case 3 ' facturas de publicidad socios
            CargaFacturasPublicidadSocio DesdeFecha, Hastafecha
        
        Case 4 ' facturas de publicidad clientes
            CargaFacturasPublicidadCliente DesdeFecha, Hastafecha
        
        Case 5 ' facturas de venta socios
            CargaFacturasVentaSocio DesdeFecha, Hastafecha
            
    End Select
    cmdSalir_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function DatosOk() As Boolean
Dim Codigo As String

    'Comprobacion de fechas
    DesdeFecha = CDate(txtcodigo(0).Text)
    Hastafecha = CDate(txtcodigo(1).Text)
    If DesdeFecha > Hastafecha Then
        MsgBox "La fecha desde debe ser menor que la fecha hasta", vbInformation
        Exit Function
    End If
    
    'Comprobacion de configuracion de carpeta destino
    Codigo = ""
    If vParamAplic.PathFacturaE = "" Then
        Codigo = "Falta configurar parametros"
    Else
        If Dir(vParamAplic.PathFacturaE, vbDirectory) = "" Then Codigo = "No existe carpeta"
    End If
    If Codigo <> "" Then
        MsgBox Codigo, vbExclamation
        Exit Function
    End If
    
    DatosOk = True
End Function




Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmppal.Icon

    txtcodigo(0).Text = Date
    txtcodigo(1).Text = Date
    
    CargaCombo
    
    Combo1(1).ListIndex = 0
    
    
    For I = 0 To Me.imgFec.Count - 1
        imgFec(I).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next
    
    

   '[Monica]05/09/2012: la carpeta de destino la tenemos en parámetros
    nompath = vParamAplic.PathFacturaE
        
End Sub


Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object
    
    Set frmC = New frmCal
    frmC.Fecha = Now

    esq = imgFec(Index).Left
    dalt = imgFec(Index).top

    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.top
        Set obj = obj.Container
    Wend
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.top = dalt + imgFec(Index).Parent.top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(0).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    
    PonerFormatoFecha txtcodigo(Index)
    If txtcodigo(Index).Text <> "" Then frmC.Fecha = CDate(txtcodigo(Index).Text)

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtcodigo(CByte(imgFec(0).Tag))
    ' ***************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtcodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub


Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = teclaBuscar Then
'        Select Case Index
'            Case 0: KEYFecha KeyAscii, 0 'fecha desde
'            Case 1: KEYFecha KeyAscii, 1 'fecha hasta
'            Case 2: KEYFecha KeyAscii, 1 'fecha de factura
'        End Select
'    Else
        KEYpress KeyAscii
'    End If

End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

'    'Quitar espacios en blanco por los lados
'    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0, 1 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
    End Select
End Sub


Private Sub CargaFacturasLiq(DFecha As Date, HFecha As Date)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim C1 As String
    Dim C2 As String
    Dim c3 As String
    Dim c4 As String
    Dim F1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim Fr As frmVisReport
    
    Dim Variedad As String
    
    Dim TipoFact1 As Byte
    Dim Gastos As Currency
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadParam As String

    '[Monica]29/02/2012: añadida la insercion en la tabla temporal por el report
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "select sfactusoc.*, stipom.letraser " & _
            " from sfactusoc, stipom, sclien where sfactusoc.fecfactu >= " & DBSet(txtcodigo(0).Text, "F") & _
            " and sfactusoc.fecfactu <= " & DBSet(txtcodigo(1).Text, "F") & _
            " and sfactusoc.codtipom = stipom.codtipom " & _
            " and sfactusoc.codtipom in ('FLI','FRL') "
    
    '[Monica]12/11/2014: faltaba la condicion de solo los socios que tienen facturae
    Sql = Sql & " and sclien.facturae=1 and sfactusoc.codsocio = sclien.codclien "
            
    '[Monica]05/09/2012: añadimos el check de solo los que no estan traspasados
    If Check2.Value = 0 Then Sql = Sql & " and sfactusoc.exportada = 0"
            
    '[Monica]02/05/2016: añadimos desde/hasta cliente
    If txtcodigo(2).Text <> "" Then Sql = Sql & " and sfactusoc.codsocio >= " & DBSet(txtcodigo(2).Text, "N")
    If txtcodigo(3).Text <> "" Then Sql = Sql & " and sfactusoc.codsocio <= " & DBSet(txtcodigo(3).Text, "N")
            
            
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Sql = "insert into tmpinformes (codusu, nombre1, importe1, codigo1, fecha1) "
        Sql = Sql & " select " & vUsu.Codigo & ", sfactusoc.codtipom, sfactusoc.numfactu, sfactusoc.codsocio, sfactusoc.fecfactu from sfactusoc, stipom, sclien where sfactusoc.fecfactu >= " & DBSet(txtcodigo(0).Text, "F") & _
                " and sfactusoc.fecfactu <= " & DBSet(txtcodigo(1).Text, "F") & _
                " and sfactusoc.codtipom = stipom.codtipom " & _
                " and sfactusoc.codtipom in ('FLI','FRL') "
                
        '[Monica]12/11/2014: faltaba la condicion de solo los socios que tienen facturae
        Sql = Sql & " and sclien.facturae=1 and sfactusoc.codsocio = sclien.codclien "
        
        If Check2.Value = 0 Then Sql = Sql & " and sfactusoc.exportada = 0"
                
        conn.Execute Sql
        
        
        Rs.MoveFirst
        While Not Rs.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\docum.pdf"
            
            Set Fr = New frmVisReport
            
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = 1
            
            '[Monica]31/04/2014: si se detallan los servicios en las facturas o no
            '[Monica]19/02/2018: Entra Cordoba
                '[Monica]19/11/2018: Entra Sevilla
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3 Then
                cadParam = cadParam & "pDetalle=" & ConDetalle & "|"
                numParam = numParam + 1
            End If
            
            
            indRPT = 51
            
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt) Then Exit Sub
            '++
            Fr.NumeroParametros = numParam
            Fr.OtrosParametros = cadParam
            Fr.ConSubInforme = True
            Fr.Informe = App.Path & "\Informes\" & pPdfRpt 'nomDocu
            Fr.ExportarPDF = True
            Fr.FormulaSeleccion = "{tmpinformes.codusu} = " & vUsu.Codigo & " and " & _
                                  "{tmpinformes.nombre1} = '" & Rs!codtipom & "' and " & _
                                  "{tmpinformes.importe1} =" & CStr(Rs!NumFactu) & " and " & _
                                  "{tmpinformes.codigo1} =" & DBSet(Rs!codSocio, "N") & " and " & _
                                  "{tmpinformes.fecha1} = Date(" & Format(Rs!FecFactu, "yyyy") & _
                                                        "," & Format(Rs!FecFactu, "mm") & _
                                                        "," & Format(Rs!FecFactu, "dd") & ")"
'            fr.FicheroPDF = FicheroPDF
            Load Fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            FileCopy FicheroPDF, nompath & "\aritaxi_" & Format(Rs!codSocio, "000000") & Rs!LetraSer & "_" & Format(Rs!NumFactu, "0000000") & "_" & Format(Rs!FecFactu, "yyyymmdd") & "_S.pdf"
            
            'actualizamos el pasaridoc de facturas socios
            Sql = "update sfactusoc set exportada = 1 where codtipom = " & DBSet(Rs!codtipom, "T")
            Sql = Sql & " and numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F")
            Sql = Sql & " and codsocio = " & DBSet(Rs!codSocio, "N")
            conn.Execute Sql
            
            Unload Fr
            Set Fr = Nothing
            
            Rs.MoveNext
        Wend
        
        MsgBox "Proceso finalizado", vbInformation
        
    Else
        MsgBox "No hay datos entre esos límites.", vbExclamation
    
    End If
    
    Set Rs = Nothing
    
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas Liquidación Socio"
    End If
End Sub

Private Sub CargaFacturasCliente(DFecha As Date, HFecha As Date)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim C1 As String
    Dim C2 As String
    Dim c3 As String
    Dim c4 As String
    Dim F1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim Fr As frmVisReport
    
    Dim Variedad As String
    
    Dim TipoFact1 As Byte
    Dim Gastos As Currency
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadParam As String
    
    Sql = "select scafaccli.*, stipom.letraser " & _
            " from scafaccli, stipom, scliente where scafaccli.fecfactu >= " & DBSet(txtcodigo(0).Text, "F") & _
            " and scafaccli.fecfactu <= " & DBSet(txtcodigo(1).Text, "F") & _
            " and scafaccli.codtipom in ('FAC','FVC','FRN') " & _
            " and scafaccli.codtipom = stipom.codtipom "
    '[Monica]12/11/2014: faltaba la condicion de solo los clientes que tienen facturae
    Sql = Sql & " and scliente.tasareciclado=1 and scafaccli.codclien = scliente.codclien "

    '[Monica]05/09/2012: añadimos el check de solo los que no estan traspasados
    If Check2.Value = 0 Then Sql = Sql & " and scafaccli.exportada = 0"
            
    '[Monica]02/05/2016: añadimos desde/hasta cliente
    If txtcodigo(2).Text <> "" Then Sql = Sql & " and scafaccli.codclien >= " & DBSet(txtcodigo(2).Text, "N")
    If txtcodigo(3).Text <> "" Then Sql = Sql & " and scafaccli.codclien <= " & DBSet(txtcodigo(3).Text, "N")
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\docum.pdf"
            
            Set Fr = New frmVisReport
            
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = 1
            
            indRPT = 52
            If DBLet(Rs!codtipom) = "FRN" Then indRPT = 54
            
            '[Monica]31/04/2014: si se detallan los servicios en las facturas o no
            '[Monica]19/02/2018: Entra Cordoba
                '[Monica]19/11/2018: Entra Sevilla
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3 Then
                cadParam = cadParam & "pDetalle=" & ConDetalle & "|"
                numParam = numParam + 1
            End If
            
            
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt) Then Exit Sub
            '++
            Fr.NumeroParametros = numParam
            Fr.OtrosParametros = cadParam
            Fr.ConSubInforme = True
            Fr.Informe = App.Path & "\Informes\" & pPdfRpt 'nomDocu
            Fr.ExportarPDF = True
            Fr.FormulaSeleccion = "{scafaccli.codtipom} = '" & Rs!codtipom & "' and " & _
                                  "{scafaccli.numfactu} =" & CStr(Rs!NumFactu) & " and " & _
                                  "{scafaccli.fecfactu} = Date(" & Format(Rs!FecFactu, "yyyy") & _
                                                        "," & Format(Rs!FecFactu, "mm") & _
                                                        "," & Format(Rs!FecFactu, "dd") & ")"
'            fr.FicheroPDF = FicheroPDF
            Load Fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            FileCopy FicheroPDF, nompath & "\aritaxi_" & Rs!LetraSer & "_" & Format(Rs!NumFactu, "0000000") & "_" & Format(Rs!FecFactu, "yyyymmdd") & "_F.pdf"
            
            'actualizamos el pasaridoc de facturas clientes
            Sql = "update scafaccli set exportada = 1 where codtipom = " & DBSet(Rs!codtipom, "T")
            Sql = Sql & " and numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F")
            
            conn.Execute Sql
            
            Unload Fr
            Set Fr = Nothing
            
            Rs.MoveNext
        Wend
        
        MsgBox "Proceso finalizado", vbInformation
        
    Else
        MsgBox "No hay datos entre esos límites.", vbExclamation
    
    End If
    
    Set Rs = Nothing
    
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas Cliente"
    End If
End Sub

Private Sub CargaFacturasCuotasSocio(DFecha As Date, HFecha As Date)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim C1 As String
    Dim C2 As String
    Dim c3 As String
    Dim c4 As String
    Dim F1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim Fr As frmVisReport
    
    Dim Variedad As String
    
    Dim TipoFact1 As Byte
    Dim Gastos As Currency
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadParam As String

    Sql = "select scafac.*, stipom.letraser " & _
            " from scafac, stipom, sclien where scafac.fecfactu >= " & DBSet(txtcodigo(0).Text, "F") & _
            " and scafac.fecfactu <= " & DBSet(txtcodigo(1).Text, "F") & _
            " and scafac.codtipom = stipom.codtipom " & _
            " and scafac.codtipom  in ('FCN','FCE','FRC') "
            
    '[Monica]12/11/2014: faltaba la condicion de solo los socios que tienen facturae
    Sql = Sql & " and sclien.facturae=1 and scafac.codclien = sclien.codclien "
            
            
    '[Monica]05/09/2012: añadimos el check de solo los que no estan traspasados
    If Check2.Value = 0 Then Sql = Sql & " and scafac.exportada = 0"
            
    '[Monica]02/05/2016: añadimos desde/hasta cliente
    If txtcodigo(2).Text <> "" Then Sql = Sql & " and scafac.codclien >= " & DBSet(txtcodigo(2).Text, "N")
    If txtcodigo(3).Text <> "" Then Sql = Sql & " and scafac.codclien <= " & DBSet(txtcodigo(3).Text, "N")
            
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\docum.pdf"
            
            Set Fr = New frmVisReport
            
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = 1
            
            indRPT = 49
            
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt) Then Exit Sub
            '++
            Fr.NumeroParametros = numParam
            Fr.OtrosParametros = cadParam
            Fr.ConSubInforme = True
            Fr.Informe = App.Path & "\Informes\" & pPdfRpt 'nomDocu
            Fr.ExportarPDF = True
            Fr.FormulaSeleccion = "{scafac.codtipom} = '" & Rs!codtipom & "' and " & _
                                  "{scafac.numfactu} =" & CStr(Rs!NumFactu) & " and " & _
                                  "{scafac.fecfactu} = Date(" & Format(Rs!FecFactu, "yyyy") & _
                                                        "," & Format(Rs!FecFactu, "mm") & _
                                                        "," & Format(Rs!FecFactu, "dd") & ")"
'            fr.FicheroPDF = FicheroPDF
            Load Fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            FileCopy FicheroPDF, nompath & "\aritaxi_" & Rs!LetraSer & "_" & Format(Rs!NumFactu, "0000000") & "_" & Format(Rs!FecFactu, "yyyymmdd") & "_C.pdf"
            
            'actualizamos el pasaridoc de facturas socios
            Sql = "update scafac set exportada = 1 where codtipom = " & DBSet(Rs!codtipom, "T")
            Sql = Sql & " and numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F")
            conn.Execute Sql
            
            Unload Fr
            Set Fr = Nothing
            
            Rs.MoveNext
        Wend
        
        MsgBox "Proceso finalizado", vbInformation
        
    Else
        MsgBox "No hay datos entre esos límites.", vbExclamation
    
    End If
    
    Set Rs = Nothing
    
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas Cuotas Socio"
    End If
End Sub

Private Sub CargaFacturasPublicidadSocio(DFecha As Date, HFecha As Date)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim C1 As String
    Dim C2 As String
    Dim c3 As String
    Dim c4 As String
    Dim F1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim Fr As frmVisReport
    
    Dim Variedad As String
    
    Dim TipoFact1 As Byte
    Dim Gastos As Currency
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadParam As String

    Sql = "select sfactusoc.*, stipom.letraser " & _
            " from sfactusoc, stipom, sclien where sfactusoc.fecfactu >= " & DBSet(txtcodigo(0).Text, "F") & _
            " and sfactusoc.fecfactu <= " & DBSet(txtcodigo(1).Text, "F") & _
            " and sfactusoc.codtipom = stipom.codtipom " & _
            " and sfactusoc.codtipom = 'FPS' "
            
    '[Monica]12/11/2014: faltaba la condicion de solo los socios que tienen facturae
    Sql = Sql & " and sclien.facturae=1 and sfactusoc.codsocio = sclien.codclien "
            
            
    '[Monica]05/09/2012: añadimos el check de solo los que no estan traspasados
    If Check2.Value = 0 Then Sql = Sql & " and sfactusoc.exportada = 0"
            
    '[Monica]02/05/2016: añadimos desde/hasta cliente
    If txtcodigo(2).Text <> "" Then Sql = Sql & " and sfactusoc.codsocio >= " & DBSet(txtcodigo(2).Text, "N")
    If txtcodigo(3).Text <> "" Then Sql = Sql & " and sfactusoc.codsocio <= " & DBSet(txtcodigo(3).Text, "N")
            
            
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\docum.pdf"
            
            Set Fr = New frmVisReport
            
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = 1
            
            indRPT = 48
            
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt) Then Exit Sub
            '++
            Fr.NumeroParametros = numParam
            Fr.OtrosParametros = cadParam
            Fr.ConSubInforme = True
            Fr.Informe = App.Path & "\Informes\" & pPdfRpt 'nomDocu
            Fr.ExportarPDF = True
            Fr.FormulaSeleccion = "{sfactusoc.codtipom} = '" & Rs!codtipom & "' and " & _
                                  "{sfactusoc.numfactu} =" & CStr(Rs!NumFactu) & " and " & _
                                  "{sfactusoc.codsocio} =" & DBSet(Rs!codSocio, "N") & " and " & _
                                  "{sfactusoc.fecfactu} = Date(" & Format(Rs!FecFactu, "yyyy") & _
                                                        "," & Format(Rs!FecFactu, "mm") & _
                                                        "," & Format(Rs!FecFactu, "dd") & ")"
'            fr.FicheroPDF = FicheroPDF
            Load Fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            FileCopy FicheroPDF, nompath & "\aritaxi_" & Format(Rs!codSocio, "000000") & Rs!LetraSer & "_" & Format(Rs!NumFactu, "0000000") & "_" & Format(Rs!FecFactu, "yyyymmdd") & "_S" & ".pdf"
            
            'actualizamos el pasaridoc de facturas socios
            Sql = "update sfactusoc set exportada = 1 where codtipom = " & DBSet(Rs!codtipom, "T")
            Sql = Sql & " and numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F")
            Sql = Sql & " and codsocio = " & DBSet(Rs!codSocio, "N")
            conn.Execute Sql
            
            Unload Fr
            Set Fr = Nothing
            
            Rs.MoveNext
        Wend
        
        MsgBox "Proceso finalizado", vbInformation
        
    Else
        MsgBox "No hay datos entre esos límites.", vbExclamation
    
    End If
    
    Set Rs = Nothing
    
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas Publicidad Socio"
    End If
End Sub

Private Sub CargaFacturasPublicidadCliente(DFecha As Date, HFecha As Date)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim C1 As String
    Dim C2 As String
    Dim c3 As String
    Dim c4 As String
    Dim F1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim Fr As frmVisReport
    
    Dim Variedad As String
    
    Dim TipoFact1 As Byte
    Dim Gastos As Currency
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadParam As String
    
    Sql = "select scafaccli.*, stipom.letraser " & _
            " from scafaccli, stipom, scliente where scafaccli.fecfactu >= " & DBSet(txtcodigo(0).Text, "F") & _
            " and scafaccli.fecfactu <= " & DBSet(txtcodigo(1).Text, "F") & _
            " and scafaccli.codtipom in ('FPC','FRP') " & _
            " and scafaccli.codtipom = stipom.codtipom "
    
    '[Monica]12/11/2014: faltaba la condicion de solo los clientes que tienen facturae
    Sql = Sql & " and scliente.tasareciclado=1 and scafaccli.codclien = scliente.codclien "
            
    
    '[Monica]05/09/2012: añadimos el check de solo los que no estan traspasados
    If Check2.Value = 0 Then Sql = Sql & " and scafaccli.exportada = 0"
            
    '[Monica]02/05/2016: añadimos desde/hasta cliente
    If txtcodigo(2).Text <> "" Then Sql = Sql & " and scafaccli.codclien >= " & DBSet(txtcodigo(2).Text, "N")
    If txtcodigo(3).Text <> "" Then Sql = Sql & " and scafaccli.codclien <= " & DBSet(txtcodigo(3).Text, "N")
            
            
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\docum.pdf"
            
            Set Fr = New frmVisReport
            
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = 1
            
            indRPT = 47
            
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt) Then Exit Sub
            '++
            Fr.NumeroParametros = numParam
            Fr.OtrosParametros = cadParam
            Fr.ConSubInforme = True
            Fr.Informe = App.Path & "\Informes\" & pPdfRpt 'nomDocu
            Fr.ExportarPDF = True
            Fr.FormulaSeleccion = "{scafaccli.codtipom} = '" & Rs!codtipom & "' and " & _
                                  "{scafaccli.numfactu} =" & CStr(Rs!NumFactu) & " and " & _
                                  "{scafaccli.fecfactu} = Date(" & Format(Rs!FecFactu, "yyyy") & _
                                                        "," & Format(Rs!FecFactu, "mm") & _
                                                        "," & Format(Rs!FecFactu, "dd") & ")"
'            fr.FicheroPDF = FicheroPDF
            Load Fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            FileCopy FicheroPDF, nompath & "\aritaxi_" & Rs!LetraSer & "_" & Format(Rs!NumFactu, "0000000") & "_" & Format(Rs!FecFactu, "yyyymmdd") & "_F.pdf"
            
            'actualizamos el pasaridoc de facturas clientes
            Sql = "update scafaccli set exportada = 1 where codtipom = " & DBSet(Rs!codtipom, "T")
            Sql = Sql & " and numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F")
            
            conn.Execute Sql
            
            Unload Fr
            Set Fr = Nothing
            
            Rs.MoveNext
        Wend
        
        MsgBox "Proceso finalizado", vbInformation
        
    Else
        MsgBox "No hay datos entre esos límites.", vbExclamation
    
    End If
    
    Set Rs = Nothing
    
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas Liquidación Socio"
    End If
End Sub

Private Sub CargaFacturasVentaSocio(DFecha As Date, HFecha As Date)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim C1 As String
    Dim C2 As String
    Dim c3 As String
    Dim c4 As String
    Dim F1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim Fr As frmVisReport
    
    Dim Variedad As String
    
    Dim TipoFact1 As Byte
    Dim Gastos As Currency
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadParam As String

    Sql = "select scafac.*, stipom.letraser " & _
            " from scafac, stipom, sclien where scafac.fecfactu >= " & DBSet(txtcodigo(0).Text, "F") & _
            " and scafac.fecfactu <= " & DBSet(txtcodigo(1).Text, "F") & _
            " and scafac.codtipom = stipom.codtipom " & _
            " and scafac.codtipom  in ('FAV','FRT') "
            
    '[Monica]12/11/2014: faltaba la condicion de solo los clientes que tienen facturae
    Sql = Sql & " and sclien.facturae=1 and scafac.codclien = sclien.codclien "
            
    '[Monica]05/09/2012: añadimos el check de solo los que no estan traspasados
    If Check2.Value = 0 Then Sql = Sql & " and scafac.exportada = 0"
            
    '[Monica]02/05/2016: añadimos desde/hasta cliente
    If txtcodigo(2).Text <> "" Then Sql = Sql & " and scafac.codclien >= " & DBSet(txtcodigo(2).Text, "N")
    If txtcodigo(3).Text <> "" Then Sql = Sql & " and scafac.codclien <= " & DBSet(txtcodigo(3).Text, "N")
            
            
            
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\docum.pdf"
            
            Set Fr = New frmVisReport
            
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = 1
            
            indRPT = 12
            
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt) Then Exit Sub
            '++
            Fr.NumeroParametros = numParam
            Fr.OtrosParametros = cadParam
            Fr.ConSubInforme = True
            Fr.Informe = App.Path & "\Informes\" & pPdfRpt 'nomDocu
            Fr.ExportarPDF = True
            Fr.FormulaSeleccion = "{scafac.codtipom} = '" & Rs!codtipom & "' and " & _
                                  "{scafac.numfactu} =" & CStr(Rs!NumFactu) & " and " & _
                                  "{scafac.fecfactu} = Date(" & Format(Rs!FecFactu, "yyyy") & _
                                                        "," & Format(Rs!FecFactu, "mm") & _
                                                        "," & Format(Rs!FecFactu, "dd") & ")"
'            fr.FicheroPDF = FicheroPDF
            Load Fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            FileCopy FicheroPDF, nompath & "\aritaxi_" & Rs!LetraSer & "_" & Format(Rs!NumFactu, "0000000") & "_" & Format(Rs!FecFactu, "yyyymmdd") & "_C.pdf"
            
            'actualizamos el pasaridoc de facturas socios
            Sql = "update scafac set exportada = 1 where codtipom = " & DBSet(Rs!codtipom, "T")
            Sql = Sql & " and numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F")
            conn.Execute Sql
            
            Unload Fr
            Set Fr = Nothing
            
            Rs.MoveNext
        Wend
        
        MsgBox "Proceso finalizado", vbInformation
        
    Else
        MsgBox "No hay datos entre esos límites.", vbExclamation
    
    End If
    
    Set Rs = Nothing
    
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas Venta Socio"
    End If
End Sub


Private Sub CargaCombo()
Dim ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I
    
    Combo1(1).Clear
    
    Combo1(1).AddItem "Liquidación Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Factura Clientes"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Facturas Cuotas Socios"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    '[Monica]05/02/2015: desmarcamos los otros tipos de facturas
    Combo1(1).AddItem "Facturas Publicidad Socios"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 3
    Combo1(1).AddItem "Facturas Publicidad Clientes"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 4
    Combo1(1).AddItem "Facturas Venta Socios"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 5
    
    
End Sub




Private Function IntentaMatar(FicheroPDF As String) As Boolean
Dim I As Integer

    On Error Resume Next
    I = 1
    IntentaMatar = False
    Do
        If Dir(FicheroPDF, vbArchive) <> "" Then
            Kill FicheroPDF
            If Err.Number <> 0 Then
                Err.Clear
                I = I + 1
            Else
                IntentaMatar = True
                I = 6
            End If
        Else
            IntentaMatar = True
            I = 6
        End If
    Loop Until I < 5 Or IntentaMatar = True
    
    
End Function






