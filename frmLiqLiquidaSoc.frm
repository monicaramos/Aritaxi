VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLiqLiquidaSoc 
   Caption         =   "Informes"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk_FactHco 
      Caption         =   "Facturación sobre Hco de llamadas"
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
      Left            =   270
      TabIndex        =   31
      Top             =   6000
      Width           =   4035
   End
   Begin VB.TextBox txtcodigo 
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
      Index           =   5
      Left            =   1410
      TabIndex        =   7
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
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
      Index           =   5
      Left            =   2280
      TabIndex        =   28
      Top             =   4920
      Width           =   3705
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Imprimir Factura"
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   270
      TabIndex        =   27
      Top             =   5760
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Imprimir Resumen"
      Enabled         =   0   'False
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
      Index           =   1
      Left            =   270
      TabIndex        =   26
      Top             =   5460
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Frame FrameProgress 
      Height          =   915
      Left            =   210
      TabIndex        =   22
      Top             =   6420
      Visible         =   0   'False
      Width           =   5835
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgess 
         Caption         =   "Facturando:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   135
         Width           =   4215
      End
      Begin VB.Label lblProgess 
         Caption         =   "Iniciando el proceso ..."
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   350
         Width           =   4335
      End
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   1410
      TabIndex        =   4
      Top             =   3150
      Width           =   1005
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   1410
      TabIndex        =   5
      Top             =   3570
      Width           =   735
   End
   Begin VB.TextBox txtnombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
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
      Left            =   2190
      TabIndex        =   19
      Top             =   3570
      Width           =   3825
   End
   Begin VB.TextBox txtnombre 
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
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1275
      Width           =   3765
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   1410
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
      Top             =   1275
      Width           =   855
   End
   Begin VB.TextBox txtnombre 
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
      Index           =   1
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1680
      Width           =   3765
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   1410
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtcodigo 
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
      Index           =   85
      Left            =   1410
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2430
      Width           =   1245
   End
   Begin VB.TextBox txtcodigo 
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
      Index           =   86
      Left            =   3915
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2445
      Width           =   1215
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
      Left            =   3750
      TabIndex        =   8
      Top             =   5460
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   4980
      TabIndex        =   9
      Top             =   5460
      Width           =   1035
   End
   Begin VB.TextBox txtcodigo 
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
      Index           =   4
      Left            =   1410
      TabIndex        =   6
      Top             =   4230
      Width           =   4605
   End
   Begin VB.Label Label9 
      Caption         =   "Cuenta Pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   29
      Top             =   4650
      Width           =   1815
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Index           =   4
      Left            =   1140
      Tag             =   "-1"
      ToolTipText     =   "Buscar cuenta"
      Top             =   4980
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   2850
      Width           =   1725
   End
   Begin VB.Label Label4 
      Caption         =   "F.Pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   1140
      ToolTipText     =   "Buscar fecha"
      Top             =   3150
      Width           =   240
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Index           =   2
      Left            =   1140
      Tag             =   "-1"
      ToolTipText     =   "Buscar forma de pago"
      Top             =   3600
      Width           =   240
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Index           =   0
      Left            =   1170
      Top             =   1275
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "V Socio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   930
      Width           =   750
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   3
      Left            =   480
      TabIndex        =   17
      Top             =   1275
      Width           =   600
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Index           =   1
      Left            =   1170
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   4
      Left            =   480
      TabIndex        =   16
      Top             =   1680
      Width           =   570
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   23
      Left            =   1140
      Top             =   2430
      Width           =   240
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   7
      Left            =   240
      TabIndex        =   13
      Top             =   2070
      Width           =   630
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   7
      Left            =   450
      TabIndex        =   12
      Top             =   2460
      Width           =   600
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   8
      Left            =   3000
      TabIndex        =   11
      Top             =   2490
      Width           =   570
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   24
      Left            =   3630
      Top             =   2460
      Width           =   240
   End
   Begin VB.Label Label10 
      Caption         =   "Liquidación Socios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   300
      Width           =   5655
   End
   Begin VB.Label Label6 
      Caption         =   "Concepto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   3960
      Width           =   1065
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Index           =   3
      Left            =   1140
      Tag             =   "-1"
      ToolTipText     =   "Ver observaciones"
      Top             =   4260
      Width           =   240
   End
End
Attribute VB_Name = "frmLiqLiquidaSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tabla As String
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim codtipom As String
Dim cadSelect As String
Dim indCodigo As Long
Dim cadNombreRPT As String
Dim cadTitulo As String
Dim ConSubInforme As Boolean
Dim conSubRPT As Boolean

Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmMtoV As frmGesVSocio ' V socios
Attribute frmMtoV.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago ' formas de pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios
Attribute frmMtoBancosPro.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes ' socios con liquidacion en efectivo
Attribute frmMens.VB_VarHelpID = -1


' Importes para la Grabacion de Cabecera de Facturas de Socio
Dim TotalFac As Currency
Dim TotalLiq As Currency
Dim BaseImpo As Currency
Dim BaseReten As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim vPorcIva As String
Dim PorceIVA As Currency

Dim tipoMov As String
Dim codSocio As String
Dim SociosContado As String

Dim kCampo As Integer

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Sub chk_FactHco_Click()
    If chk_FactHco.Value = 1 Then
        Tabla = "shilla"
    Else
        Tabla = "sfactsoctr"
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim Codigo As String
Dim FecFac As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean

Dim SqlUve As String

    If Not DatosOk Then Exit Sub
    
    InicializarVbles
    
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 1
   
    'Desde/Hasta numero de V
    '---------------------------------------------
    If txtcodigo(0).Text <> "" Or txtcodigo(1).Text <> "" Then
        Codigo = "{" & Tabla & ".numeruve}"
        If Not PonerDesdeHasta(Codigo, "N", 0, 1, "pDHUve=""") Then Exit Sub
    End If
    
    '[Monica]25/10/2012: seleccionamos que socios han de ser en efectivo
    SqlUve = ""
    If CLng(ComprobarCero(txtcodigo(0).Text)) <> 0 Then SqlUve = SqlUve & " and numeruve >= " & DBSet(txtcodigo(0).Text, "N")
    If CLng(ComprobarCero(txtcodigo(1).Text)) <> 0 Then SqlUve = SqlUve & " and numeruve <= " & DBSet(txtcodigo(1).Text, "N")
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtcodigo(85).Text <> "" Or txtcodigo(86).Text <> "" Then
        If Tabla = "shilla" Then
            Codigo = "{" & Tabla & ".fecha}"
        Else
            Codigo = "{" & Tabla & ".fecfactu}"
        End If
        If Not PonerDesdeHasta(Codigo, "F", 85, 86, "pDHFecha=""") Then Exit Sub
    End If
    
'[Monica]10/09/2014: partimos de la shilla
    ' que no este liquidado
    If Tabla = "shilla" Then
        If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".liquidadosocio} = 0") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".liquidadosocio} = 0") Then Exit Sub
    Else
        If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".facturado} = 0") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".facturado} = 0") Then Exit Sub
    End If



'[Monica]10/09/2014: partimos de la shilla, añado esta condicion
    If Tabla = "shilla" Then
        If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".impcompr} <> 0") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".impcompr} <> 0") Then Exit Sub
        
        '[Monica]19/09/2014: añadida esta condicion por teletaxi
        ' solo servicios de credito
        If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".tipservi} = 1") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".tipservi} = 1") Then Exit Sub
    
        '[Monica]13/11/2014: añadida la condicion de solo los servicios que esten validados para Teletaxi
        If vParamAplic.Cooperativa = 0 Then
            ' solo servicios validados
            If Not AnyadirAFormula(cadFormula, "{" & Tabla & ".validado} = 1") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".validado} = 1") Then Exit Sub
        End If
    End If


    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    
    '[Monica]25/10/2012: cuando es teletaxi vemos que socios deben ser facturados en efectivo
    If vParamAplic.Cooperativa = 0 Then
    
        '[Monica]31/03/2014: si no hay socios de contado no sacar ventana, añadida la condicion
        Dim SqlVer As String
        '[Monica]10/09/2014: antes era sobre rfactsoctr, ahora sobre la shilla
        If Tabla = "shilla" Then
            SqlVer = "select count(*) from shilla where " & cadSelect & " and shilla.codsocio in (select codclien from sclien where escontado = 1)"
        Else
            SqlVer = "select count(*) from sfactsoctr where " & cadSelect & " and sfactsoctr.codsocio in (select codclien from sclien where escontado = 1)"
        End If
        If TotalRegistros(SqlVer) <> 0 Then
    
            Set frmMens = New frmMensajes
            
            frmMens.OpcionMensaje = 25
            '[Monica]10/09/2014: antes era sobre rfactsoctr, ahora sobre la shilla
            frmMens.cadWHERE2 = Tabla
            If Tabla = "shilla" Then
                frmMens.cadWHERE = cadSelect & " and shilla.codsocio in (select codclien from sclien where escontado = 1)"
            Else
                frmMens.cadWHERE = cadSelect & " and sfactsoctr.codsocio in (select codclien from sclien where escontado = 1)"
            End If
            frmMens.Show vbModal
        
            Set frmMens = Nothing
            
            If SociosContado = "Cancelado" Then
                Unload Me
                Exit Sub
            End If
        
        End If
        
    End If

    ' proceso de liquidacion a socios
    If Tabla = "shilla" Then
        If ProcesoLiquidacionSocioNew(cadSelect, txtcodigo(3).Text) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(1).Value Then
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Liquidación"
            
                cadFormula = ""
                cadParam = cadParam & "pFecFac= """ & txtcodigo(3).Text & """|"
                numParam = numParam + 1
                cadParam = cadParam & "pTitulo= ""Resumen Liquidaciones Socio""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtcodigo(3).Text)
                cadAux = "{sfactusoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                ConSubInforme = False
                
                LlamarImprimir False
            End If
            
            '[Monica]20/11/2014: en teletaxi no dejamos imprimir las facturas que lo hagan en reimpresion de facturas
            If vParamAplic.Cooperativa = 0 Then Check1(0).Value = 0
            
            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA LIQUIDACION SOCIO
            If Me.Check1(0).Value Then
                cadFormula = ""
                cadSelect = ""
                cadAux = "({sfactusoc.codtipom} = 'FLI')"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                'Nº Socio
                If txtcodigo(0).Text <> "" Then
                    cadAux = "{sfactusoc.numeruve} >= " & txtcodigo(0).Text
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                End If
                If txtcodigo(1).Text <> "" Then
                    cadAux = "{sfactusoc.numeruve} <= " & txtcodigo(1).Text
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                End If
                 
                'Fecha de Factura
                FecFac = CDate(txtcodigo(3).Text)
                cadAux = "{sfactusoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{sfactusoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                               
                indRPT = 51 'Impresion de facturas de liquidacion a socios
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, False, "") Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresión de Facturas Liquidación a Socios"
                ConSubInforme = True
                
                conSubRPT = ConSubInforme
                
                '[Monica]02/04/2012: faltaba esta condicion para que no saque otras facturas realizadas anteriormente
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                
                '[Monica]31/03/2014: en el caso de teletaxi solo imprimimos las facturas de los socios que no tengan facturacion electronica
                If vParamAplic.Cooperativa = 0 Then
                    'preguntamos si quiere imprimirlo o no con los servicios
                    If MsgBox("¿ Desea imprimir el detalle de servicios ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                        cadParam = cadParam & "pDetalle=0|"
                    Else
                        cadParam = cadParam & "pDetalle=1|"
                    End If
                    numParam = numParam + 1
                    
                    If Not AnyadirAFormula(cadFormula, "{sclien.facturae} = 0") Then Exit Sub
                End If
                
                LlamarImprimir False
            End If
        
            cmdCancelar_Click
            
        End If
    Else
        If ProcesoLiquidacionSocio(cadSelect, txtcodigo(3).Text) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(1).Value Then
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Liquidación"
            
                cadFormula = ""
                cadParam = cadParam & "pFecFac= """ & txtcodigo(3).Text & """|"
                numParam = numParam + 1
                cadParam = cadParam & "pTitulo= ""Resumen Liquidaciones Socio""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtcodigo(3).Text)
                cadAux = "{sfactusoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                ConSubInforme = False
                
                LlamarImprimir False
            End If
            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA LIQUIDACION SOCIO
            If Me.Check1(0).Value Then
                cadFormula = ""
                cadSelect = ""
                cadAux = "({sfactusoc.codtipom} = 'FLI')"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                'Nº Socio
                If txtcodigo(0).Text <> "" Then
                    cadAux = "{sfactusoc.numeruve} >= " & txtcodigo(0).Text
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                End If
                If txtcodigo(1).Text <> "" Then
                    cadAux = "{sfactusoc.numeruve} <= " & txtcodigo(1).Text
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                End If
                 
                'Fecha de Factura
                FecFac = CDate(txtcodigo(3).Text)
                cadAux = "{sfactusoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{sfactusoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                               
                indRPT = 51 'Impresion de facturas de liquidacion a socios
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, False, "") Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresión de Facturas Liquidación a Socios"
                ConSubInforme = True
                
                conSubRPT = ConSubInforme
                
                '[Monica]02/04/2012: faltaba esta condicion para que no saque otras facturas realizadas anteriormente
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                
                '[Monica]31/03/2014: en el caso de teletaxi solo imprimimos las facturas de los socios que no tengan facturacion electronica
                If vParamAplic.Cooperativa = 0 Then
                    'preguntamos si quiere imprimirlo o no con los servicios
                    If MsgBox("¿ Desea imprimir el detalle de servicios ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                        cadParam = cadParam & "pDetalle=0|"
                    Else
                        cadParam = cadParam & "pDetalle=1|"
                    End If
                    numParam = numParam + 1
                    
                    If Not AnyadirAFormula(cadFormula, "{sclien.facturae} = 0") Then Exit Sub
                End If
                
                LlamarImprimir False
            End If
        
            cmdCancelar_Click
            
        End If
    
    
    End If

End Sub

Private Sub LlamarImprimir(duplicado As Boolean)
        
    With frmImprimir
        .Titulo = cadTitulo
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        'El nombre es el del documento
        .NombreRPT = cadNombreRPT
        .Opcion = 101
        .ConSubInforme = ConSubInforme
        .Show vbModal
    End With
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    cadFormula = ""
    numParam = 0
    cadParam = ""


    PonerFoco txtcodigo(0)

End Sub

Private Sub Form_Load()

    'Icono del form
    Me.Icon = frmPpal.Icon
    
    Me.chk_FactHco.Value = 1
   
    For kCampo = 0 To Me.imgBuscarOfer.Count - 1
        Me.imgBuscarOfer(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    
    For kCampo = 23 To 24
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    Me.imgFecha(0).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    
    
    If vParamAplic.Cooperativa = 0 Then
        If Me.chk_FactHco.Value = 1 Then
            '[Monica]10/09/2014
            Tabla = "shilla"
        Else
            Tabla = "sfactsoctr"
        End If
    Else
        Tabla = "shilla"
    End If

    ' para las impresiones
    Check1(0).Value = 1
    Check1(1).Value = 1

End Sub

Private Function DatosOk() As Boolean
Dim encontrado As String
Dim Codigo As String

    DatosOk = True
    If txtcodigo(2).Text = "" Then
        MsgBox "Debe introducir obligatoriamente la forma de pago.", vbExclamation
        DatosOk = False
        Exit Function
    End If
    If txtcodigo(3).Text = "" Then
        MsgBox "Debe introducir obligatoriamente la fecha de liquidación.", vbExclamation
        DatosOk = False
        Exit Function
    End If
    'concepto
    If txtcodigo(4).Text = "" Then
        MsgBox "Es necesario introducir el concepto de la factura.", vbExclamation
        PonerFoco txtcodigo(4)
        DatosOk = False
        Exit Function
    End If
    'banco
    If txtcodigo(5).Text = "" Then
        MsgBox "Es necesario introducir la cuenta de banco.", vbExclamation
        PonerFoco txtcodigo(5)
        DatosOk = False
        Exit Function
    Else
        encontrado = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", txtcodigo(5).Text, "N")
        'comprobar que la cuenta prevista de cobro tiene valor
        If encontrado = "" Then
            MsgBox "La cta. prevista de cobro debe tener valor.", vbExclamation
            DatosOk = False
            Exit Function
        End If
    End If

End Function



Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtcodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(2).Text = RecuperaValor(CadenaSeleccion, 1)
    txtnombre(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    SociosContado = ""
    If CadenaSeleccion <> "" Then
        If CadenaSeleccion = "Cancelado" Then
            SociosContado = "Cancelado"
        Else
            SociosContado = "codsocio in (" & CadenaSeleccion & ")"
        End If
    End If
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    txtcodigo(5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtnombre(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoV_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 0, 1 ' Vsocio
            indCodigo = Index
            
            Set frmMtoV = New frmGesVSocio
            frmMtoV.DeConsulta = True
            frmMtoV.DatosADevolverBusqueda = "0|1|2|"
            frmMtoV.Show vbModal
            Set frmMtoV = Nothing
        
        Case 2 ' forma de pago
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0|1|"
            frmFP.Show vbModal
            Set frmFP = Nothing
    
        Case 3 ' concepto
            CadenaDesdeOtroForm = txtcodigo(4).Text
            frmFacClienteObser.Modificar = True
            frmFacClienteObser.Text1 = CadenaDesdeOtroForm
            frmFacClienteObser.Show vbModal
            If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then txtcodigo(4).Text = Mid(CadenaDesdeOtroForm, 3)
            CadenaDesdeOtroForm = ""
            PonerFoco txtcodigo(4)
        
        
        Case 4 ' banco propio
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
    
    End Select
    PonerFoco txtcodigo(indCodigo)

End Sub

Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 23, 24 'fechas de factura
            indCodigo = Index + 62
        Case 0
            indCodigo = 3
   End Select
   
   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtcodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)
End Sub

Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    If txtcodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtcodigo(indD).Text
        If txtnombre(indD).Text <> "" Then Cad = Cad & " - " & txtnombre(indD).Text
    End If
    If txtcodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtcodigo(indH).Text
        If txtnombre(indH).Text <> "" Then Cad = Cad & " - " & txtnombre(indH).Text
    End If
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Tabla As String
Dim codCampo As String, NomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean


    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    Select Case Index
        Case 85, 86  'FECHA Desde Hasta
            PonerFormatoFecha txtcodigo(Index)
            
        Case 0, 1 'V Socio
            PonerFormatoEntero txtcodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sclien", "nomclien", "numeruve", "N")
            
        Case 2 ' forma de pago
            PonerFormatoEntero txtcodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sforpa", "nomforpa", "codforpa", "N")
            
        Case 3 ' fecha de liquidacion
            PonerFormatoFecha txtcodigo(Index)
            
        Case 5 ' banco propio
            PonerFormatoEntero txtcodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sbanpr", "nombanpr", "codbanpr", "N")
            
        
        
    End Select
End Sub



Private Function ProcesoLiquidacionSocio(cadWHERE As String, FecFac As String) As Boolean
'Desde Historico de llamadas Genera las Facturas correspondientes
Dim RSalb As ADODB.Recordset 'Ordenados por: codsocio
Dim b As Boolean
Dim Sql As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim antSocio As Long
Dim actSocio As Long
Dim antV As Long
Dim actV As Long

'Concatenamos todas las facturas generadas para listarlas en el informe
Dim ListFactu As String
Dim vSocio As CSocio
Dim Inc As Integer
Dim condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura

Dim nTotal As Long
Dim SQL2 As String
Dim NumFactu As Long
Dim devuelve As Long
Dim Existe As Boolean

Dim vFacSoc As CFacturaSoc
Dim MensError As String
Dim FormatSocio As String
Dim FPagContado As String
Dim BancoContado As String

    On Error GoTo ETraspasoAlbFac

    ProcesoLiquidacionSocio = False

    SQL2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL2
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
    
    conn.BeginTrans
    ConnConta.BeginTrans

    tipoMov = "FLI"

    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("LIQSOC") 'facturas de venta
    If Not BloqueoManual("LIQSOC", "1") Then
        MsgBox "No se puede liquidar. Hay otro usuario realizando el proceso.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los registros de llamadas que vamos a facturar (cabeceras y lineas)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    If vParamAplic.Cooperativa = 0 Then
        If Not BloqueaRegistro("sfactsoctr", cadWHERE) Then
            Screen.MousePointer = vbDefault
            'comprobamos que no haya nadie facturando
            DesBloqueoManual ("LIQSOC")
            Exit Function
        End If
    Else
        If Not BloqueaRegistro("shilla", cadWHERE) Then
            Screen.MousePointer = vbDefault
            'comprobamos que no haya nadie facturando
            DesBloqueoManual ("LIQSOC")
            Exit Function
        End If
    End If
    
    If vParamAplic.Cooperativa = 0 Then
        Sql = "select numeruve, sum(if(importe is null,0,importe)) from sfactsoctr where " & cadWHERE & " group by numeruve having sum(if(importe is null,0,importe)) <> 0 "
    Else
        Sql = "select numeruve, sum(if(impcompr is null,0,impcompr)) from shilla where " & cadWHERE & " group by numeruve having sum(if(impcompr is null,0,impcompr)) <> 0 "
    End If
    nTotal = TotalRegistrosConsulta(Sql)
    PB1.Max = nTotal
    
    FrameProgress.visible = True
    
    'EMPEZAMOS LA FACTURA
    If vParamAplic.Cooperativa = 0 Then
        Sql = "select numeruve, sum(numserv) servicios, sum(if(importe is null,0,importe)) importe from sfactsoctr where " & cadWHERE & " group by numeruve, concepto having sum(if(importe is null,0,importe)) <> 0 "
        Sql = Sql & " ORDER BY sfactsoctr.numeruve"
    Else
        Sql = "select numeruve, count(*) servicios, sum(if(impcompr is null,0,impcompr)) importe from shilla where " & cadWHERE & " group by numeruve having sum(if(impcompr is null,0,impcompr)) <> 0 "
        Sql = Sql & " ORDER BY shilla.numeruve"
    End If
    Set RSalb = New ADODB.Recordset
    RSalb.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    b = True
    While Not RSalb.EOF And b
        codSocio = DevuelveValor("select codclien from sclien where numeruve = " & DBLet(RSalb!NumerUve, "N"))
        
        IncrementarProgresNew PB1, 1
        
        Set vSocio = New CSocio
        If vSocio.LeerDatos(codSocio) Then
            NumFactu = vSocio.ConseguirContador(tipoMov)
            If NumFactu = -1 Then b = False
            Do
                NumFactu = vSocio.ConseguirContador(tipoMov)
                Sql = "select numfactu from rfactusoc where codtipom = " & DBSet(tipoMov, "T") & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFac, "F") & " and codsocio = " & DBSet(vSocio.Codigo, "N")
                devuelve = DevuelveValor(Sql) 'DevuelveDesdeBDNew(cAgro, "rfacttra", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> 0 Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vSocio.IncrementarContador (tipoMov)
                    NumFactu = vSocio.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            vPorcIva = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", vParamAplic.IVA_REA, "N")
            
            PorceIVA = 0
            If vPorcIva <> "" Then PorceIVA = CCur(vPorcIva)
            
            TotalFac = DBLet(RSalb.Fields(2).Value, "N")
            BaseImpo = Round2(TotalFac / (1 + (PorceIVA / 100)), 2)
            BaseReten = TotalFac
            ImpoIva = TotalFac - BaseImpo
            ImpoReten = Round2(TotalFac * vParamAplic.PorReten / 100, 2)
            TotalLiq = TotalFac - ImpoReten
            
'            txtcodigo(4).Text = RSalb!Concepto
            
            'insertar cabecera de factura
            b = InsertCabeceraFactSocio(tipoMov, CStr(NumFactu), FecFac, vSocio.Codigo, RSalb!NumerUve, RSalb!Servicios)
            
            '[Monica]31/03/2014: en el caso de teletaxi insertamos los servicios
            If vParamAplic.Cooperativa = 0 Then
                If b Then InsertServiciosFactSocio tipoMov, CStr(NumFactu), FecFac, vSocio.Codigo, RSalb!NumerUve, cadWHERE
            End If
            MensError = ""
            If b Then
                Set vFacSoc = New CFacturaSoc
                '[Monica]22/11/2013: iban
                vFacSoc.CCC_Iban = vSocio.Iban
                vFacSoc.CCC_Entidad = vSocio.Banco
                vFacSoc.CCC_Oficina = vSocio.Sucursal
                vFacSoc.CCC_CC = vSocio.DigControl
                vFacSoc.CCC_CTa = vSocio.CuentaBan
                
                If vParamAplic.Cooperativa = 0 Then
                    If EsSocioContado(codSocio, SociosContado) Then
                        FPagContado = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", "nomforpa", "EFECTIVO", "T")
                        If FPagContado = "" Then
                            vFacSoc.ForPago = 0
                        Else
                            vFacSoc.ForPago = CInt(FPagContado)
                        End If
                    Else
                        vFacSoc.ForPago = txtcodigo(2).Text
                    End If
                Else
                    vFacSoc.ForPago = txtcodigo(2).Text
                End If
                
                vFacSoc.NumFactu = NumFactu
                vFacSoc.FecFactu = FecFac
                vFacSoc.TotalFac = TotalLiq
                vFacSoc.ImpRet2 = ImpoReten
                vFacSoc.Socio = vSocio.Codigo
                vFacSoc.tipoMov = "FLI"
                
                vFacSoc.CtaSocio = vSocio.CtaSocioLiq
                
                '[Monica]25/10/2012: socios contado ponemos la cuenta prevista como la de caja
                If vParamAplic.Cooperativa = 0 Then
                    If EsSocioContado(codSocio, SociosContado) Then
                        BancoContado = DevuelveDesdeBDNew(conAri, "sbanpr", "codbanpr", "nombanpr", "CAJA", "T")
                        If BancoContado = "" Then
                            vFacSoc.CuentaPrev = ""
                        Else
                            vFacSoc.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", BancoContado, "N")
                        End If
                    Else
                        vFacSoc.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", txtcodigo(5).Text, "N")
                    End If
                
                Else
                    vFacSoc.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", txtcodigo(5).Text, "N")
                End If
                
                If b Then b = InsertarRetencion(RSalb!NumerUve, vFacSoc)
                
                If b Then b = vFacSoc.InsertarEnTesoreria(MensError)
                
                Set vFacSoc = Nothing
            End If

            If b And vParamAplic.Cooperativa <> 0 Then b = InsertLineasFactSocio(tipoMov, CStr(NumFactu), FecFac, vSocio.Codigo, RSalb!NumerUve, cadWHERE)
            
            If b And vParamAplic.Cooperativa <> 0 Then b = ActualizarLlamadas(RSalb!NumerUve, cadWHERE)

            If b And vParamAplic.Cooperativa = 0 Then b = ActualizarLlamadas2(RSalb!NumerUve, cadWHERE)

            '[Monica]25/10/2012: socio contado los desmarcados como contado
            If b And vParamAplic.Cooperativa = 0 Then b = ActualizarSociosContado(codSocio)
            
            If b Then b = InsertResumen(tipoMov, CStr(NumFactu), vSocio.Codigo, FecFac)
            
            If b Then b = vSocio.IncrementarContador(tipoMov)
            
        Else
            b = False
            MsgBox "No existe el código de socio para la Uve " & RSalb!NumerUve, vbExclamation
        End If
        
        Set vSocio = Nothing
        
        RSalb.MoveNext
    Wend
    RSalb.Close
    Set RSalb = Nothing
    
ETraspasoAlbFac:
    If Err.Number <> 0 Or Not b Then
        If Err.Number <> 0 Or MensError <> "" Then MuestraError Err.Number, "Liquidación Socio:", Err.Description & " " & MensError
        conn.RollbackTrans
        ConnConta.RollbackTrans
        ProcesoLiquidacionSocio = False
    Else
        conn.CommitTrans
        ConnConta.CommitTrans
        ProcesoLiquidacionSocio = True
    End If
    DesBloqueoManual ("LIQSOC")
    TerminaBloquear
End Function

Private Function InsertCabeceraFactSocio(tipoMov As String, NumFactu As String, FecFac As String, Socio As String, Uve As String, Serv As String) As Boolean
Dim Sql As String
Dim MensError As String
Dim FPagContado As String
    On Error GoTo eInsertCabe
    
    MensError = ""
    InsertCabeceraFactSocio = False
    
    Sql = "insert into sfactusoc (codtipom,codsocio,numfactu,fecfactu,concepto,importel,baseiva1,impoiva1, "
    Sql = Sql & "codiiva1,porciva1,basereten,porcreten,totalfac,impreten,intconta,codforpa,numeruve, numserv) values ("
    Sql = Sql & DBSet(tipoMov, "T") & "," & DBSet(Socio, "N") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & ","
    Sql = Sql & DBSet(txtcodigo(4).Text, "T") & "," & DBSet(TotalFac, "N") & "," & DBSet(BaseImpo, "N") & "," & DBSet(ImpoIva, "N") & ","
    Sql = Sql & vParamAplic.IVA_REA & "," & DBSet(PorceIVA, "N") & "," & DBSet(BaseReten, "N") & "," & DBSet(vParamAplic.PorReten, "N") & ","
    
    If vParamAplic.Cooperativa = 0 Then
        If EsSocioContado(Socio, SociosContado) <> 0 Then
            FPagContado = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", "nomforpa", "EFECTIVO", "T")
            If FPagContado = "" Then FPagContado = "0"
            Sql = Sql & DBSet(TotalLiq, "N") & "," & DBSet(ImpoReten, "N") & ",0," & DBSet(FPagContado, "N") & "," & DBSet(Uve, "N") & ","
        Else
            Sql = Sql & DBSet(TotalLiq, "N") & "," & DBSet(ImpoReten, "N") & ",0," & DBSet(txtcodigo(2).Text, "N") & "," & DBSet(Uve, "N") & ","
        End If
    Else
        Sql = Sql & DBSet(TotalLiq, "N") & "," & DBSet(ImpoReten, "N") & ",0," & DBSet(txtcodigo(2).Text, "N") & "," & DBSet(Uve, "N") & ","
    End If
    Sql = Sql & DBSet(Serv, "N") & ")"
    
    conn.Execute Sql
    
    InsertCabeceraFactSocio = True
    
    Exit Function

eInsertCabe:
    MensError = "Error en la inserción en Cabecera de Factura " & NumFactu & " del Socio " & Socio
    MuestraError Err.Number, MensError
End Function

Private Function EsSocioContado(Socio As String, CadSocios As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from sclien where codclien = " & DBSet(Socio, "N") & " and " & Replace(CadSocios, "codsocio", "codclien")
    
    EsSocioContado = (TotalRegistros(Sql) <> 0)

End Function
'Insertar Linea de factura (llamadas)
Private Function InsertLineasFactSocio(tipoMov As String, NumFactu As String, FecFac As String, Socio As String, NumerUve As String, cadWHERE As String) As Boolean
Dim Precio As Currency
Dim Sql As String
Dim SQL2 As String
Dim SqlValues As String
Dim linea As Long
Dim MensError As String
Dim RS As ADODB.Recordset
    
    On Error GoTo eInsertLinea
    
    InsertLineasFactSocio = False
    
    MensError = ""
    
    Sql = "insert into sfactusoc_serv (codtipom,codsocio,numfactu,fecfactu,numlinea,fecha,hora,numeruve,"
    Sql = Sql & "codclien,nomclien,dirllama,numllama,puerllama,ciudadre,telefono,impventa,idservic, observac2) values "
    
    SQL2 = "select * from shilla where numeruve = " & DBSet(NumerUve, "N")
    SQL2 = SQL2 & " and " & cadWHERE
    SQL2 = SQL2 & " order by fecha, hora "
    
    SqlValues = ""
    linea = 0
    
    Set RS = New ADODB.Recordset
    RS.Open SQL2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        linea = linea + 1
    
        SqlValues = SqlValues & "(" & DBSet(tipoMov, "T") & "," & DBSet(Socio, "N") & "," & DBSet(NumFactu, "N") & ","
        SqlValues = SqlValues & DBSet(FecFac, "F") & "," & DBSet(linea, "N") & "," & DBSet(RS!Fecha, "F") & ","
        SqlValues = SqlValues & DBSet(RS!hora, "H") & "," & DBSet(RS!NumerUve, "N") & "," & DBSet(RS!CodClien, "N") & ","
        SqlValues = SqlValues & DBSet(RS!nomclien, "T") & "," & DBSet(RS!dirllama, "T") & "," & DBSet(RS!numllama, "T") & ","
        SqlValues = SqlValues & DBSet(RS!puerllama, "T") & "," & DBSet(RS!ciudadre, "T") & "," & DBSet(RS!Telefono, "T") & ","
        SqlValues = SqlValues & DBSet(RS!impcompr, "N") & "," & DBSet(RS!idservic, "T") & "," & DBSet(RS!observac2, "T") & "),"
        
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If linea <> 0 Then
        'quitamos la ultima coma
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        
        conn.Execute Sql & SqlValues
    End If
    
    InsertLineasFactSocio = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de servicios de la factura del socio " & Socio
        MuestraError Err.Number, MensError, Err.descripc
    End If
End Function


'Insertar Resumen
Private Function InsertResumen(Tipo As String, NumFactu As String, Socio As String, FecFac As String) As Boolean
Dim MensError As String
Dim Sql As String
    
    On Error GoTo eInsertResumen
    
    MensError = ""
    InsertResumen = False
    
                                        ' codtipom, numfactu, codsocio, fecfactu
    Sql = "insert into tmpinformes (codusu, nombre1, importe1, codigo1, fecha1) values ( " & vUsu.Codigo
    Sql = Sql & ",'" & Tipo & "'," & DBSet(NumFactu, "N") & "," & DBSet(Socio, "N") & "," & DBSet(FecFac, "F") & ")"
    
    conn.Execute Sql
    
    InsertResumen = True
    
    Exit Function

eInsertResumen:
    MensError = "Error en la inserción de la factura " & NumFactu & " en el Resumen "
    MuestraError Err.Number, MensError
End Function

Private Function ListaFacturasGeneradas(Tipo As String) As String
Dim Sql As String
Dim rs1 As ADODB.Recordset
Dim Cad As String
    
    On Error GoTo eFacturasGeneradas

    ListaFacturasGeneradas = ""

    Sql = "select nombre1, importe1, codigo1 from tmpinformes where codusu = " & vUsu.Codigo
    Sql = Sql & " and nombre1 = " & DBSet(Trim(Tipo), "T")
    
    Set rs1 = New ADODB.Recordset
    rs1.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Cad = ""
    While Not rs1.EOF
        Cad = Cad & "(" & DBSet(rs1.Fields(0).Value, "T") & "," & DBLet(rs1.Fields(2).Value, "N") & ","
        Cad = Cad & DBLet(rs1.Fields(1).Value, "N") & "," & DBSet(txtcodigo(3).Text, "F") & "),"
    
        rs1.MoveNext
    Wend
    Set rs1 = Nothing
    
    'si hay facturas quitamos la ultima coma
    If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
    
    ListaFacturasGeneradas = Cad
    Exit Function
    
eFacturasGeneradas:
    MuestraError Err.Number, "Cadena de Facturas Generadas", Err.Description
End Function



'Actualizar llamadas
Private Function ActualizarLlamadas(Uve As String, cadWHERE As String) As Boolean
Dim Precio As Currency
Dim Sql As String
Dim SQL2 As String
Dim SqlValues As String
Dim linea As Long
Dim MensError As String

    On Error GoTo eInsertLinea
    
    ActualizarLlamadas = False
    
    MensError = ""
    
    Sql = "update shilla set liquidadosocio = 1 "
    If vParamAplic.Cooperativa = 0 Then
          Sql = Sql & ", abonados = 1 "
    End If
    Sql = Sql & " where " & cadWHERE & " and numeruve = " & DBSet(Uve, "N")
    
    
    conn.Execute Sql
    
    ActualizarLlamadas = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la actualización de servicios de la factura del socio NºV " & Uve
        MuestraError Err.Number, MensError, Err.descripc
    End If
End Function

Private Function ActualizarLlamadas2(Uve As String, cadWHERE As String) As Boolean
Dim Precio As Currency
Dim Sql As String
Dim SQL2 As String
Dim SqlValues As String
Dim linea As Long
Dim MensError As String

    On Error GoTo eInsertLinea
    
    ActualizarLlamadas2 = False
    
    MensError = ""
    
    Sql = "update sfactsoctr set facturado = 1 where " & cadWHERE & " and numeruve = " & DBSet(Uve, "N")
    
    conn.Execute Sql
    
    ActualizarLlamadas2 = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la actualización de servicios de la factura del socio NºV " & Uve
        MuestraError Err.Number, MensError, Err.descripc
    End If
End Function

Private Function InsertarRetencion(Uve As String, ByRef vFac As CFacturaSoc) As Boolean
Dim Precio As Currency
Dim Sql As String
Dim SQL2 As String
Dim SqlValues As String
Dim linea As Long
Dim MensError As String

    On Error GoTo eInsertLinea
    
    InsertarRetencion = False
    
    MensError = ""
    
    Sql = "insert into sreten (codsocio, numeruve, fecfactu, numfactu, impreten, tiporeten) values ("
    Sql = Sql & DBSet(vFac.Socio, "N") & "," & DBSet(Uve, "N") & "," & DBSet(vFac.FecFactu, "F") & ","
    Sql = Sql & DBSet(vFac.NumFactu, "N") & "," & DBSet(vFac.ImpRet2, "N") & ",0)"
    
    conn.Execute Sql
    
    InsertarRetencion = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de retencion de la factura del socio NºV " & Uve
        MuestraError Err.Number, MensError, Err.descripc
    End If
End Function


Private Function InsertarEnTesoreria(tipoMov As String, NumFactu As String, FecFac As String, Socio As String) As Boolean
'Guarda datos de Tesoreria en tablas: aritaxi.svenci y en conta.scobros
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim RsFact As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single
Dim ForPago As String
Dim CtaProve As String
Dim FormatSocio As String
Dim CuentaPrev As String

Dim MenError As String

'[Monica]22/11/2013
Dim mCCC_Iban As String

Dim mCCC_Entidad As String
Dim mCCC_Oficina As String
Dim mCCC_CC As String
Dim mCCC_CTa As String

Dim vSocio As CSocio

Dim i As Byte

    On Error GoTo EInsertarTesoreria

'    b = False
    InsertarEnTesoreria = False
    CadValues2 = ""

    
    Sql = "select * from sfactusoc where codtipom = " & DBSet(tipoMov, "T") & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
    Sql = Sql & " and codsocio = " & DBSet(Socio, "N")
    
    Set RsFact = New ADODB.Recordset
    RsFact.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RsFact.EOF Then
        MenError = MenError & "No existe el registro de factura"
        Set RsFact = Nothing
        Exit Function
    End If
    
    ForPago = DBLet(RsFact!codforpa, "N")
    

    'Obtener el Nº de Vencimientos de la forma de pago
    Sql = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago de la tabla de parametros: spara1
            Sql = " SELECT  diapago1, diapago2, diapago3,mesnogir "
            Sql = Sql & " FROM spara1 "
            Sql = Sql & " WHERE codigo=1"
            Set RS = New ADODB.Recordset
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not RS.EOF Then
               ' se construye como en el caso de publicidad
               CtaProve = ""
               
               Socio = RsFact!codSocio
               Set vSocio = New CSocio
               
               If vSocio.LeerDatos(Socio) Then
                   FormatSocio = String((vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior), "0")
                   CtaProve = Trim(vParamAplic.Raiz_Cta_Soc_Liqui & Format(Socio, FormatSocio))
                
                   'vamos creando la cadena para insertar en spagosp de la CONTA
                   CadValuesAux2 = "("
                   If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & DBSet(SerieFraPro, "T") & ","
                   CadValuesAux2 = CadValuesAux2 & "'" & CtaProve & "', " & DBSet(RsFact!NumFactu, "T") & ", '" & Format(RsFact!FecFactu, FormatoFecha) & "', "
                  
                  'Primer Vencimiento
                  '------------------------------------------------------------
                  i = 1
                  'FECHA VTO
                  FecVenci = CDate(RsFact!FecFactu)
                  '=== Modificado: Laura 23/01/2007
    '              FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                  FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                  '==================================
                  'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                  FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
    
                  'Comprobar si  tiene mes a no girar
                  FecVenci1 = FecVenci
                  If DBSet(RS!mesnogir, "N") <> 0 Then
                      FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(RS!mesnogir, "N"), DBSet(0, "N"), RS!DiaPago1, RS!DiaPago2, RS!DiaPago3)
                  End If
                 
                  CadValues2 = CadValuesAux2 & i
                  CadValues2 = CadValues2 & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                  'IMPORTE del Vencimiento
                  If rsVenci!numerove = 1 Then
                        ImpVenci = RsFact!TotalFac
                  Else
                        ImpVenci = Round(RsFact!TotalFac / rsVenci!numerove, 2)
                        'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                        If ImpVenci * rsVenci!numerove <> RsFact!TotalFac Then
                            ImpVenci = Round(ImpVenci + (RsFact!TotalFac - ImpVenci * rsVenci!numerove), 2)
                        End If
                  End If
                  
                  CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", txtcodigo(5).Text, "N")
                  
                  
                  CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CuentaPrev, "T") & ","
                  
                  
                  If Not vParamAplic.ContabilidadNueva Then
                        'David. Para que ponga la cuenta bancaria (SI LA tiene)
                        CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                        CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & ","
                  End If
    
                  'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
                  Sql = "Factura num.: " & RsFact!NumFactu & "-" & Format(RsFact!FecFactu, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                  Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
                  
                  If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & "," & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Domicilio, "T") & "," & DBSet(vSocio.Poblacion, "T") & ","
                        CadValues2 = CadValues2 & DBSet(vSocio.CPostal, "T") & "," & DBSet(vSocio.Provincia, "T") & "," & DBSet(vSocio.NIF, "T") & ",'ES',"
                        CadValues2 = CadValues2 & DBSet(vSocio.Iban, "T") & ")"
                  Else
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & "," & DBSet(mCCC_Iban, "T", "S") & ")"
                        Else
                            CadValues2 = CadValues2 & ")"
                        End If
                  End If
     
                  'Resto Vencimientos
                  '--------------------------------------------------------------------
                  For i = 2 To rsVenci!numerove
                     'FECHA Resto Vencimientos
                      '==== Modificado: Laura 23/01/2007
                      'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                      FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                      '==================================================
                      'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                      FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
    
                      'Comprobar si tiene mes a no girar
                      FecVenci1 = FecVenci
                      If DBSet(RS!mesnogir, "N") <> 0 Then
                            FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(RS!mesnogir, "N"), DBSet(0, "N"), RS!DiaPago1, RS!DiaPago2, RS!DiaPago3)
                      End If
    
                      CadValues2 = CadValues2 & ", " & CadValuesAux2 & i & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
    
                      'IMPORTE Resto de Vendimientos
                      ImpVenci = Round(RS!TotalFac / rsVenci!numerove, 2)
    
                      CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "',"
                      
                      
                      'David. Para que ponga la cuenta bancaria (SI LA tiene)
                      If Not vParamAplic.ContabilidadNueva Then
                            CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                            CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & ","
                      End If
                      
                      Sql = "Factura num.: " & RsFact!NumFactu & "-" & Format(RsFact!FecFactu, "dd/mm/yyyy")
                      CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                      Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
                      CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
                      
                      If vParamAplic.ContabilidadNueva Then
                            CadValues2 = CadValues2 & "," & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Domicilio, "T") & "," & DBSet(vSocio.Poblacion, "T") & ","
                            CadValues2 = CadValues2 & DBSet(vSocio.CPostal, "T") & "," & DBSet(vSocio.Provincia, "T") & "," & DBSet(vSocio.NIF, "T") & ",'ES',"
                            CadValues2 = CadValues2 & DBSet(vSocio.Iban, "T") & ")"
                      Else
                            '[Monica]22/11/2013: tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & "," & DBSet(mCCC_Iban, "T", "S") & ")"
                            Else
                                CadValues2 = CadValues2 & ")"
                            End If
                      End If
                  Next i
                End If
                
                Set vSocio = Nothing
            
            End If
        End If
        RS.Close
        Set RS = Nothing
    End If
    
    rsVenci.Close
    Set rsVenci = Nothing
    
    'Grabar tabla spagop de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        'antes de grabar en la spago comprobar que existe en conta.sforpa la
        'forma de pago de la factura. Sino existe insertarla

        'vemos si existe en la conta
        If vParamAplic.ContabilidadNueva Then
            CadValuesAux2 = DevuelveDesdeBDNew(conConta, "formapago", "codforpa", "codforpa", ForPago, "N")
        Else
            CadValuesAux2 = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", ForPago, "N")
        End If
        'si no existe la forma de pago en conta, insertamos la de aritaxi
        If CadValuesAux2 = "" Then
        
'++

            Dim Sql8 As String
            Dim RS8 As ADODB.Recordset

            Sql8 = "select * from sforpa where codfopa = " & DBSet(ForPago, "N")
            Set RS8 = New ADODB.Recordset
            RS8.Open Sql8, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS8.EOF Then
                'insertamos e sforpa de la CONTA
                If vParamAplic.ContabilidadNueva Then
                    Sql8 = "INSERT INTO formapago(codforpa,nomforpa,tipforpa,numerove,primerve,restoven)"
                Else
                    Sql8 = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                End If
                Sql8 = Sql8 & " VALUES(" & ForPago & ", " & DBSet(RS!nomforpa, "T") & ", " & DBSet(RS!tipforpa, "N")
                If vParamAplic.ContabilidadNueva Then
                    Sql8 = Sql8 & "," & DBSet(RS!numerove, "N") & "," & DBSet(RS!primerve, "N") & "," & DBSet(RS!restoven, "N") & ")"
                Else
                    Sql8 = Sql8 & ")"
                End If
                ConnConta.Execute Sql8
            End If
            RS8.Close
            Set RS8 = Nothing
        
'++
        
        
        End If

        'Insertamos en la tabla spagop de la CONTA
        'SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1) "
        'David. Cuenta bancaria y descripcion textos
        If vParamAplic.ContabilidadNueva Then
            Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb," ') "
            Sql = Sql & "nomprove,domprove,pobprove,cpprove,proprove,nifprove,codpais,iban)"
        Else
            Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
            '[Monica]22/11/2013: tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                Sql = Sql & ",iban)"
            Else
                Sql = Sql & ")"
            End If
        End If
        Sql = Sql & " VALUES " & CadValues2
        ConnConta.Execute Sql
    End If

    b = True
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
        MuestraError Err.Number, MenError
    End If
    InsertarEnTesoreria = b
End Function



Private Function ActualizarSociosContado(Socio As String) As Boolean
Dim Precio As Currency
Dim Sql As String
Dim SQL2 As String
Dim SqlValues As String
Dim linea As Long
Dim MensError As String

    On Error GoTo eInsertLinea
    
    ActualizarSociosContado = False
    
    MensError = ""
    
    Sql = "update sclien set escontado = 0 where codclien = " & DBSet(Socio, "N")
    
    conn.Execute Sql
    
    ActualizarSociosContado = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la actualización de socios contado "
        MuestraError Err.Number, MensError, Err.descripc
    End If
End Function


Private Function InsertServiciosFactSocio(tipoMov As String, NumFactu As String, FecFac As String, Socio As String, Uve As String, cWhere As String) As Boolean
Dim Sql As String
Dim MensError As String
Dim FPagContado As String

    On Error GoTo eInsertServiciosFactSocio
    
    MensError = ""
    InsertServiciosFactSocio = False
    
    If Tabla = "shilla" Then
        Sql = "insert into sfactusoc_serv (codtipom,codsocio,numfactu,fecfactu,numlinea,fecha,hora,numeruve,codclien,nomclien,dirllama,"
        Sql = Sql & " impventa,idservic,observac2,matricul, codusuar, destino, codautor, licencia, fecfinal, horfinal) "  '[Monica]03/10/2014: insertamos el destino
        Sql = Sql & " select " & DBSet(tipoMov, "T") & "," & DBSet(Socio, "N") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & ","
        Sql = Sql & " @rownum:=@rownum+1 AS rownum, fecha, hora, numeruve, shilla.codclien, scliente.nomclien, concat(dirllama,' ',numllama) , impcompr, idservic, '', matricul, codusuar, destino, codautor, licencia, fecfinal, horfinal " '[Monica]03/10/2014: insertamos el destino
        Sql = Sql & " from shilla left join scliente on shilla.codclien = scliente.codclien, (SELECT @rownum:=0) r "
    '[Monica]10/09/2014: cambiamos ahora los servicios son de la shilla
    '    SQL = SQL & " where (numeruve, fecfactu) in (select numeruve, fecfactu from sfactsoctr "
        Sql = Sql & " where " & cWhere
        Sql = Sql & " and shilla.numeruve = " & DBSet(Uve, "N") & " and shilla.codsocio = " & DBSet(Socio, "N")
        Sql = Sql & " order by fecha, hora "
    '    SQL = SQL & ")"
    
    Else
        Sql = "insert into sfactusoc_serv (codtipom,codsocio,numfactu,fecfactu,numlinea,fecha,hora,numeruve,codclien,nomclien,dirllama,"
        Sql = Sql & " impventa,idservic,observac2,matricul) "
        Sql = Sql & " select " & DBSet(tipoMov, "T") & "," & DBSet(Socio, "N") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & ","
        Sql = Sql & " @rownum:=@rownum+1 AS rownum, fecha, hora, numeruve, sfactsoctr_serv.codclien, scliente.nomclien, origen , importe, nroservicio, destino, matricul "
        Sql = Sql & " from sfactsoctr_serv left join scliente on sfactsoctr_serv.codclien = scliente.codclien, (SELECT @rownum:=0) r "
    '[Monica]10/09/2014: cambiamos ahora los servicios son de la shilla
        Sql = Sql & " where (numeruve, fecfactu) in (select numeruve, fecfactu from sfactsoctr "
        Sql = Sql & " where " & cWhere
        Sql = Sql & " and sfactsoctr_serv.numeruve = " & DBSet(Uve, "N") & " and sfactsoctr_serv.codsocio = " & DBSet(Socio, "N")
        Sql = Sql & " order by fecha, hora "
        Sql = Sql & ")"
    
    
    End If
    
    conn.Execute Sql
    
    InsertServiciosFactSocio = True
    
    Exit Function

eInsertServiciosFactSocio:
    MensError = "Error en la inserción de Servicios de Factura " & NumFactu & " del Socio " & Socio
    MuestraError Err.Number, MensError
End Function

'###############################################################################################
'###################
'###################    NUEVO PROCESO DE LIQUIDACION DE SOCIOS PARA TELETAXI, PARTIMOS DE LA SHILLA
'###################
'###############################################################################################

Private Function ProcesoLiquidacionSocioNew(cadWHERE As String, FecFac As String) As Boolean
'Desde Historico de llamadas Genera las Facturas correspondientes
Dim RSalb As ADODB.Recordset 'Ordenados por: codsocio
Dim b As Boolean
Dim Sql As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim antSocio As Long
Dim actSocio As Long
Dim antV As Long
Dim actV As Long

'Concatenamos todas las facturas generadas para listarlas en el informe
Dim ListFactu As String
Dim vSocio As CSocio
Dim Inc As Integer
Dim condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura

Dim nTotal As Long
Dim SQL2 As String
Dim NumFactu As Long
Dim devuelve As Long
Dim Existe As Boolean

Dim vFacSoc As CFacturaSoc
Dim MensError As String
Dim FormatSocio As String
Dim FPagContado As String
Dim BancoContado As String

    On Error GoTo ETraspasoAlbFac

    ProcesoLiquidacionSocioNew = False

    SQL2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL2
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
    
    conn.BeginTrans
    ConnConta.BeginTrans

    tipoMov = "FLI"

    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("LIQSOC") 'facturas de venta
    If Not BloqueoManual("LIQSOC", "1") Then
        MsgBox "No se puede liquidar. Hay otro usuario realizando el proceso.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los registros de llamadas que vamos a facturar (cabeceras y lineas)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    If vParamAplic.Cooperativa = 0 Then
        If Not BloqueaRegistro("shilla", cadWHERE) Then
            Screen.MousePointer = vbDefault
            'comprobamos que no haya nadie facturando
            DesBloqueoManual ("LIQSOC")
            Exit Function
        End If
    Else
        If Not BloqueaRegistro("shilla", cadWHERE) Then
            Screen.MousePointer = vbDefault
            'comprobamos que no haya nadie facturando
            DesBloqueoManual ("LIQSOC")
            Exit Function
        End If
    End If
'[Monica]10/09/2014: partimos en ambos casos de la shilla
'    If vParamAplic.Cooperativa = 0 Then
'        SQL = "select numeruve, sum(if(importe is null,0,importe)) from sfactsoctr where " & cadwhere & " group by numeruve having sum(if(importe is null,0,importe)) <> 0 "
'    Else
        Sql = "select numeruve, sum(if(impcompr is null,0,impcompr)) from shilla where " & cadWHERE & " group by numeruve having sum(if(impcompr is null,0,impcompr)) <> 0 "
'    End If
    nTotal = TotalRegistrosConsulta(Sql)
    PB1.Max = nTotal
    
    FrameProgress.visible = True
    
    'EMPEZAMOS LA FACTURA
'[Monica]10/09/2014: partimos en ambos casos de la shilla
'    If vParamAplic.Cooperativa = 0 Then
'        SQL = "select numeruve, sum(numserv) servicios, sum(if(importe is null,0,importe)) importe from sfactsoctr where " & cadwhere & " group by numeruve, concepto having sum(if(importe is null,0,importe)) <> 0 "
'        SQL = SQL & " ORDER BY sfactsoctr.numeruve"
'    Else
        Sql = "select numeruve, count(*) servicios, sum(if(impcompr is null,0,impcompr)) importe from shilla where " & cadWHERE & " group by numeruve having sum(if(impcompr is null,0,impcompr)) <> 0 "
        Sql = Sql & " ORDER BY shilla.numeruve"
'    End If
    Set RSalb = New ADODB.Recordset
    RSalb.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    b = True
    While Not RSalb.EOF And b
        codSocio = DevuelveValor("select codclien from sclien where numeruve = " & DBLet(RSalb!NumerUve, "N"))
        
        IncrementarProgresNew PB1, 1
        
        Set vSocio = New CSocio
        If vSocio.LeerDatos(codSocio) Then
            NumFactu = vSocio.ConseguirContador(tipoMov)
            If NumFactu = -1 Then b = False
            Do
                NumFactu = vSocio.ConseguirContador(tipoMov)
                Sql = "select numfactu from rfactusoc where codtipom = " & DBSet(tipoMov, "T") & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFac, "F") & " and codsocio = " & DBSet(vSocio.Codigo, "N")
                devuelve = DevuelveValor(Sql) 'DevuelveDesdeBDNew(cAgro, "rfacttra", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> 0 Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vSocio.IncrementarContador (tipoMov)
                    NumFactu = vSocio.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            vPorcIva = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", vParamAplic.IVA_REA, "N")
            
            PorceIVA = 0
            If vPorcIva <> "" Then PorceIVA = CCur(vPorcIva)
            
            TotalFac = DBLet(RSalb.Fields(2).Value, "N")
            BaseImpo = Round2(TotalFac / (1 + (PorceIVA / 100)), 2)
            BaseReten = TotalFac
            ImpoIva = TotalFac - BaseImpo
            ImpoReten = Round2(TotalFac * vParamAplic.PorReten / 100, 2)
            TotalLiq = TotalFac - ImpoReten
            
'            txtcodigo(4).Text = RSalb!Concepto
            
            'insertar cabecera de factura
            b = InsertCabeceraFactSocio(tipoMov, CStr(NumFactu), FecFac, vSocio.Codigo, RSalb!NumerUve, RSalb!Servicios)
            
            '[Monica]31/03/2014: en el caso de teletaxi insertamos los servicios
            If vParamAplic.Cooperativa = 0 Then
                If b Then InsertServiciosFactSocio tipoMov, CStr(NumFactu), FecFac, vSocio.Codigo, RSalb!NumerUve, cadWHERE
            End If
            MensError = ""
            If b Then
                Set vFacSoc = New CFacturaSoc
                '[Monica]22/11/2013: iban
                vFacSoc.CCC_Iban = vSocio.Iban
                vFacSoc.CCC_Entidad = vSocio.Banco
                vFacSoc.CCC_Oficina = vSocio.Sucursal
                vFacSoc.CCC_CC = vSocio.DigControl
                vFacSoc.CCC_CTa = vSocio.CuentaBan
                
                If vParamAplic.Cooperativa = 0 Then
                    If EsSocioContado(codSocio, SociosContado) Then
                        FPagContado = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", "nomforpa", "EFECTIVO", "T")
                        If FPagContado = "" Then
                            vFacSoc.ForPago = 0
                        Else
                            vFacSoc.ForPago = CInt(FPagContado)
                        End If
                    Else
                        vFacSoc.ForPago = txtcodigo(2).Text
                    End If
                Else
                    vFacSoc.ForPago = txtcodigo(2).Text
                End If
                
                vFacSoc.NumFactu = NumFactu
                vFacSoc.FecFactu = FecFac
                vFacSoc.TotalFac = TotalLiq
                vFacSoc.ImpRet2 = ImpoReten
                vFacSoc.Socio = vSocio.Codigo
                vFacSoc.tipoMov = "FLI"
                
                vFacSoc.CtaSocio = vSocio.CtaSocioLiq
                
                '[Monica]11/05/2017
                vFacSoc.NombreSocio = vSocio.Nombre
                vFacSoc.DomicilioSocio = vSocio.Domicilio
                vFacSoc.CPostalSocio = vSocio.CPostal
                vFacSoc.PoblacionSocio = vSocio.Poblacion
                vFacSoc.ProvinciaSocio = vSocio.Provincia
                vFacSoc.nifSocio = vSocio.NIF
                
                
                '[Monica]25/10/2012: socios contado ponemos la cuenta prevista como la de caja
                If vParamAplic.Cooperativa = 0 Then
                    If EsSocioContado(codSocio, SociosContado) Then
                        BancoContado = DevuelveDesdeBDNew(conAri, "sbanpr", "codbanpr", "nombanpr", "CAJA", "T")
                        If BancoContado = "" Then
                            vFacSoc.CuentaPrev = ""
                        Else
                            vFacSoc.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", BancoContado, "N")
                        End If
                    Else
                        vFacSoc.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", txtcodigo(5).Text, "N")
                    End If
                
                Else
                    vFacSoc.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", txtcodigo(5).Text, "N")
                End If
                
                If b Then b = InsertarRetencion(RSalb!NumerUve, vFacSoc)
                
                If b Then b = vFacSoc.InsertarEnTesoreria(MensError)
                
                Set vFacSoc = Nothing
            End If

            If b And vParamAplic.Cooperativa <> 0 Then b = InsertLineasFactSocio(tipoMov, CStr(NumFactu), FecFac, vSocio.Codigo, RSalb!NumerUve, cadWHERE)
            
            If b And vParamAplic.Cooperativa <> 0 Then b = ActualizarLlamadas(RSalb!NumerUve, cadWHERE)
    
            '[Monica]10/09/2014: en ambos casos se actualiza la shilla
            If b And vParamAplic.Cooperativa = 0 Then b = ActualizarLlamadas(RSalb!NumerUve, cadWHERE)

            '[Monica]25/10/2012: socio contado los desmarcados como contado
            If b And vParamAplic.Cooperativa = 0 Then b = ActualizarSociosContado(codSocio)
            
            If b Then b = InsertResumen(tipoMov, CStr(NumFactu), vSocio.Codigo, FecFac)
            
            If b Then b = vSocio.IncrementarContador(tipoMov)
            
        Else
            b = False
            MsgBox "No existe el código de socio para la Uve " & RSalb!NumerUve, vbExclamation
        End If
        
        Set vSocio = Nothing
        
        RSalb.MoveNext
    Wend
    RSalb.Close
    Set RSalb = Nothing
    
ETraspasoAlbFac:
    If Err.Number <> 0 Or Not b Then
        If Err.Number <> 0 Or MensError <> "" Then MuestraError Err.Number, "Liquidación Socio:", Err.Description & " " & MensError
        conn.RollbackTrans
        ConnConta.RollbackTrans
        ProcesoLiquidacionSocioNew = False
    Else
        conn.CommitTrans
        ConnConta.CommitTrans
        ProcesoLiquidacionSocioNew = True
    End If
    DesBloqueoManual ("LIQSOC")
    TerminaBloquear
End Function




'Private Function ProcesoLiquidacionSocioVIP(cadWHERE As String, FecFac As String) As Boolean
''Desde Historico de llamadas Genera las Facturas correspondientes
'Dim RSalb As ADODB.Recordset 'Ordenados por: codsocio
'Dim b As Boolean
'Dim Sql As String
'
''Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
'Dim antSocio As Long
'Dim actSocio As Long
'Dim antV As Long
'Dim actV As Long
'
''Concatenamos todas las facturas generadas para listarlas en el informe
'Dim ListFactu As String
'Dim vSocio As CSocio
'Dim Inc As Integer
'Dim condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura
'
'Dim nTotal As Long
'Dim SQL2 As String
'Dim NumFactu As Long
'Dim devuelve As Long
'Dim Existe As Boolean
'
'Dim vFacSoc As CFacturaSoc
'Dim MensError As String
'Dim FormatSocio As String
'
'    On Error GoTo EProcesoLiquidacionSocioVIP
'
'    ProcesoLiquidacionSocioVIP = False
'
'    SQL2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
'    conn.Execute SQL2
'
'    If cadWHERE <> "" Then
'        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
'        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
'        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
'    End If
'
'    conn.BeginTrans
'    ConnConta.BeginTrans
'
'    tipoMov = "FLI"
'
'    'comprobamos que no haya nadie facturando
'    DesBloqueoManual ("LIQSOC") 'facturas de venta
'    If Not BloqueoManual("LIQSOC", "1") Then
'        MsgBox "No se puede liquidar. Hay otro usuario realizando el proceso.", vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    'Bloqueamos todos los registros de llamadas que vamos a facturar (cabeceras y lineas)
'    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
'    If Not BloqueaRegistro("shilla", cadWHERE) Then
'        Screen.MousePointer = vbDefault
'        'comprobamos que no haya nadie facturando
'        DesBloqueoManual ("LIQSOC")
'        Exit Function
'    End If
'
'    Sql = "select numeruve,  sum(if(impcompr is null, 0, impcompr)) from shilla where " & cadWHERE & " group by numeruve having sum(if(impcompr is null, 0, impcompr)) <> 0 "
'    nTotal = TotalRegistrosConsulta(Sql)
'    Pb1.Max = nTotal
'
'    FrameProgress.visible = True
'
'    'EMPEZAMOS LA FACTURA
'    Sql = "select numeruve, count(*) servicios, sum(if(impcompr is null, 0, impcompr)) importe from shilla where " & cadWHERE & " group by numeruve having sum(if(impcompr is null, 0, impcompr)) <> 0 "
'    Sql = Sql & " ORDER BY shilla.numeruve"
'    Set RSalb = New ADODB.Recordset
'    RSalb.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    b = True
'    While Not RSalb.EOF And b
'        codSocio = DevuelveValor("select codclien from sclien where numeruve = " & DBLet(RSalb!NumerUve, "N"))
'
'        IncrementarProgresNew Pb1, 1
'
'        Set vSocio = New CSocio
'        If vSocio.LeerDatos(codSocio) Then
'            NumFactu = vSocio.ConseguirContador(tipoMov)
'            If NumFactu = -1 Then b = False
'            Do
'                NumFactu = vSocio.ConseguirContador(tipoMov)
'                Sql = "select numfactu from rfactusoc where codtipom = " & DBSet(tipoMov, "T") & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFac, "F") & " and codsocio = " & DBSet(vSocio.Codigo, "N")
'                devuelve = DevuelveValor(Sql) 'DevuelveDesdeBDNew(cAgro, "rfacttra", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
'                If devuelve <> 0 Then
'                    'Ya existe el contador incrementarlo
'                    Existe = True
'                    vSocio.IncrementarContador (tipoMov)
'                    NumFactu = vSocio.ConseguirContador(tipoMov)
'                Else
'                    Existe = False
'                End If
'            Loop Until Not Existe
'
'            vPorcIva = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", vParamAplic.IVA_REA, "N")
'
'            PorceIva = 0
'            If vPorcIva <> "" Then PorceIva = CCur(vPorcIva)
'
'            TotalFac = DBLet(RSalb.Fields(2).Value, "N")
'            BaseImpo = Round2(TotalFac / (1 + (PorceIva / 100)), 2)
'            BaseReten = TotalFac
'            ImpoIva = TotalFac - BaseImpo
'            ImpoReten = Round2(TotalFac * vParamAplic.PorReten / 100, 2)
'            TotalLiq = TotalFac - ImpoReten
'
'
'            'insertar cabecera de factura
'            b = InsertCabeceraFactSocio(tipoMov, CStr(NumFactu), FecFac, vSocio.Codigo, RSalb!NumerUve, RSalb!Servicios)
'            MensError = ""
'            If b Then
'                Set vFacSoc = New CFacturaSoc
'                vFacSoc.CCC_Entidad = vSocio.Banco
'                vFacSoc.CCC_Oficina = vSocio.Sucursal
'                vFacSoc.CCC_CC = vSocio.DigControl
'                vFacSoc.CCC_CTa = vSocio.CuentaBan
'                vFacSoc.ForPago = txtcodigo(2).Text
'                vFacSoc.NumFactu = NumFactu
'                vFacSoc.FecFactu = FecFac
'                vFacSoc.TotalFac = TotalLiq
'                vFacSoc.ImpRet2 = ImpoReten
'                vFacSoc.Socio = vSocio.Codigo
'
'                vFacSoc.CtaSocio = vSocio.CtaSocioLiq
'                vFacSoc.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", txtcodigo(5).Text, "N")
'
'                If b Then b = InsertarRetencion(RSalb!NumerUve, vFacSoc)
'
'                If b Then b = vFacSoc.InsertarEnTesoreria(MensError)
'
'                Set vFacSoc = Nothing
'            End If
'
'            If b Then b = InsertLineasFactSocio(tipoMov, CStr(NumFactu), FecFac, vSocio.Codigo, RSalb!NumerUve, cadWHERE)
'
'            If b Then b = ActualizarLlamadas(RSalb!NumerUve, cadWHERE)
''            If b Then b = ActualizarLlamadas2(RSalb!NumerUve, cadWHERE)
'
'            If b Then b = InsertResumen(tipoMov, CStr(NumFactu), vSocio.Codigo)
'
'            If b Then b = vSocio.IncrementarContador(tipoMov)
'
'        Else
'            b = False
'            MsgBox "No existe el código de socio para la Uve " & RSalb!NumerUve, vbExclamation
'        End If
'
'        Set vSocio = Nothing
'
'        RSalb.MoveNext
'    Wend
'    RSalb.Close
'    Set RSalb = Nothing
'
'EProcesoLiquidacionSocioVIP:
'    If Err.Number <> 0 Or Not b Then
'        If Err.Number <> 0 Or MensError <> "" Then MuestraError Err.Number, "Liquidación Socio:", Err.Description & " " & MensError
'        conn.RollbackTrans
'        ConnConta.RollbackTrans
'        ProcesoLiquidacionSocioVIP = False
'    Else
'        conn.CommitTrans
'        ConnConta.CommitTrans
'        ProcesoLiquidacionSocioVIP = True
'    End If
'    DesBloqueoManual ("LIQSOC")
'    TerminaBloquear
'End Function



