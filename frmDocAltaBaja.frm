VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDocAltaBaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6270
   Icon            =   "frmDocAltaBaja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCalidades 
      Height          =   7860
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6165
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   3720
         MaxLength       =   12
         TabIndex        =   3
         Top             =   2340
         Width           =   1305
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2835
         Left            =   90
         TabIndex        =   25
         Top             =   4230
         Width           =   5955
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   2250
            Width           =   3105
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Cta.Dif.negativas|T|S|||sparam|ctanegtat|||"
            Top             =   2250
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1890
            Width           =   3105
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta.Dif.negativas|T|S|||sparam|ctanegtat|||"
            Top             =   1890
            Width           =   585
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   2010
            MaxLength       =   6
            TabIndex        =   5
            Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
            Top             =   750
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   750
            Width           =   3075
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Cta.Dif.negativas|T|S|||sparam|ctanegtat|||"
            Top             =   1530
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1530
            Width           =   3105
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1140
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   1140
            Width           =   3075
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   390
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   390
            Width           =   3075
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Reserva"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   37
            Top             =   1890
            Width           =   1410
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1710
            Picture         =   "frmDocAltaBaja.frx":000C
            ToolTipText     =   "Buscar forma pago"
            Top             =   2250
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Aportac."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   1
            Left            =   270
            TabIndex        =   35
            Top             =   2250
            Width           =   1410
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1710
            Picture         =   "frmDocAltaBaja.frx":010E
            ToolTipText     =   "Buscar forma pago"
            Top             =   1890
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1710
            Picture         =   "frmDocAltaBaja.frx":0210
            ToolTipText     =   "Buscar concepto"
            Top             =   1530
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1710
            Picture         =   "frmDocAltaBaja.frx":0312
            ToolTipText     =   "Buscar concepto"
            Top             =   1140
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1710
            Picture         =   "frmDocAltaBaja.frx":0414
            ToolTipText     =   "Buscar diario"
            Top             =   390
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Pago"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   33
            Top             =   780
            Width           =   1065
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1710
            Picture         =   "frmDocAltaBaja.frx":0516
            ToolTipText     =   "Buscar cuenta"
            Top             =   750
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concep.Reserva"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   24
            Left            =   270
            TabIndex        =   31
            Top             =   1170
            Width           =   1560
         End
         Begin VB.Label Label1 
            Caption         =   "Concep.Aportac."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   25
            Left            =   270
            TabIndex        =   30
            Top             =   1530
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "Número Diario "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   29
            Top             =   420
            Width           =   1350
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1380
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ficha Socio"
         Height          =   330
         Index           =   4
         Left            =   3420
         TabIndex        =   15
         Top             =   3900
         Width           =   2310
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Notificación Banco"
         Height          =   330
         Index           =   3
         Left            =   3420
         TabIndex        =   14
         Top             =   3510
         Width           =   2310
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Documento Admisión"
         Height          =   330
         Index           =   2
         Left            =   3420
         TabIndex        =   13
         Top             =   3120
         Width           =   2310
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Realizar Contabilización"
         Height          =   330
         Index           =   1
         Left            =   330
         TabIndex        =   12
         Top             =   3090
         Width           =   2310
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Num vehiculo|N|N|||shilla|numeruve|000000|S|"
         Top             =   900
         Width           =   795
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   900
         Width           =   3735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   3720
         MaxLength       =   12
         TabIndex        =   2
         Top             =   1830
         Width           =   1305
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   11
         Top             =   7320
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   3930
         TabIndex        =   10
         Top             =   7320
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Height          =   440
         Left            =   7860
         Picture         =   "frmDocAltaBaja.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command7 
         Height          =   440
         Left            =   7860
         Picture         =   "frmDocAltaBaja.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Importe Aportación Capital Social titulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   38
         Top             =   2370
         Width           =   3345
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmDocAltaBaja.frx":0C2C
         ToolTipText     =   "Buscar fecha"
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   330
         TabIndex        =   24
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Impresión"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   3420
         TabIndex        =   23
         Top             =   2820
         Width           =   870
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Importe Reserva Legal Obligatoria gto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   7
         Left            =   330
         TabIndex        =   22
         Top             =   1860
         Width           =   3255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "V Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   21
         Top             =   930
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmDocAltaBaja.frx":0CB7
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Documentos Alta Socio"
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
         Left            =   360
         TabIndex        =   19
         Top             =   330
         Width           =   5025
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5895
      Top             =   5265
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDocAltaBaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados BASICOS ====
    '=============================
    '
    
Public NumCod As String 'Para indicar nro de uve de socio

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'calendario fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmMtoV As frmGesVSocio ' V socios
Attribute frmMtoV.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios ' banco propio (de pago)
Attribute frmMtoBancosPro.VB_VarHelpID = -1
Private WithEvents frmConce As frmConceConta 'conceptos de contabilidad
Attribute frmConce.VB_VarHelpID = -1
Private WithEvents frmTDia As frmDiaConta 'diarios de contabilidad
Attribute frmTDia.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Formas de Pago en menu Facturacion
Attribute frmFP.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim Indice As Integer

Dim NumCampo As String

Dim ImpGasto As Currency
Dim ImpTitulo As Currency

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub Check1_Click(Index As Integer)
    If Index = 1 Then
        Frame2.Enabled = Check1(1).Value
        If Check1(1).Value = 0 Then
            txtcodigo(1).Text = ""
            txtcodigo(2).Text = ""
            txtcodigo(4).Text = ""
            txtcodigo(6).Text = ""
        End If
    End If
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
    If Not DatosOk Then Exit Sub
    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    ' Numero de uve del socio
    cadParam = cadParam & "pCodigo=" & NumCod & "|"
    numParam = numParam + 1
    
    'alta de socios
    cadTitulo = "Documentos Alta de Socios"
    
    ImpGasto = 0
    If txtcodigo(5).Text <> "" Then
        ImpGasto = txtcodigo(5).Text 'TransformaComasPuntos(ImporteSinFormato(txtCodigo(5).Text))
        cadParam = cadParam & "pGasto=" & TransformaComasPuntos(ImporteSinFormato(txtcodigo(5).Text)) & "|"
        numParam = numParam + 1
    End If
    
    ImpTitulo = 0
    If txtcodigo(9).Text <> "" Then 'check1(0).Value Then
        ImpTitulo = txtcodigo(9).Text 'TransformaComasPuntos(ImporteSinFormato(txtCodigo(9).Text))
        cadParam = cadParam & "pTitulo=" & TransformaComasPuntos(ImporteSinFormato(txtcodigo(9).Text)) & "|"
        numParam = numParam + 1
    End If
    
'    cadParam = cadParam & "pBanco=" & Check1(0).Value & "|"
'    numParam = numParam + 1
    
    ' Impresion de documento de alta
    cadParam = cadParam & "pDocum=" & Check1(2).Value & "|"
    numParam = numParam + 1
    
    ' Impresion de la notificacion banco
    cadParam = cadParam & "pNotific=" & Check1(3).Value & "|"
    numParam = numParam + 1
    
    ' Impresion de la ficha del socio
    cadParam = cadParam & "pFicha=" & Check1(4).Value & "|"
    numParam = numParam + 1
    
    
    cadParam = cadParam & "pFecha=""" & txtcodigo(3).Text & """|"
    numParam = numParam + 1
    
    
    If Not AnyadirAFormula(cadFormula, "{sclien.numeruve} = " & NumCod) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "sclien.numeruve = " & NumCod) Then Exit Sub
    
    
    'Nombre fichero .rpt a Imprimir
    indRPT = 53 ' alta socios
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, False, "") Then Exit Sub
    
    frmImprimir.NombreRPT = nomDocu
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme("sclien", cadSelect) Then
        If Check1(2).Value Or Check1(3).Value Or Check1(4).Value Then LlamarImprimir
        
        If Check1(1).Value Then
            If MsgBox("¿ Desea continuar con la contabilización ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                If InsertarAsiento Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (2)
                Else
                    MsgBox "No se ha realizado el proceso. Llame a Ariadna.", vbExclamation
                End If
            Else
                cmdCancel_Click (2)
            End If
        Else
            cmdCancel_Click (2)
        End If
    End If

End Sub


Private Function InsertarAsiento() As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Lineas As String
Dim i As Integer

Dim Mc As Contadores
Dim cad As String
Dim CADENA As String
Dim mCtaSocio As String
Dim mCtaBanco As String
Dim codSocio As Long
Dim CadValues As String
Dim Obs As String

Dim ImporteD As Currency
Dim ImporteH As Currency
Dim ImpTotal As Currency

Dim cadMen As String
Dim b As Boolean

Dim ampliacion1 As String
Dim ampliacion2 As String

Dim Documento As String


    On Error GoTo eInsertarAsiento
    
    InsertarAsiento = False
    
    ConnConta.BeginTrans
    
    Set Mc = New Contadores
    
    If Mc.ConseguirContador("0", (CDate(txtcodigo(3).Text) <= CDate(vEmpresa.FechaFin)), True) = 0 Then
    
        Obs = "Alta de Socios "
        i = 0
    
        codSocio = DevuelveValor("select codclien from sclien where numeruve = " & DBSet(txtcodigo(0).Text, "N"))
        Documento = "SOC-" & Format(codSocio, "000000")
    
        'Insertar en la conta Cabecera Asiento
        b = InsertarCabAsientoDia(txtcodigo(2).Text, Mc.Contador, CDate(txtcodigo(3).Text), Obs, cadMen)
        cadMen = "Insertando Cab. Asiento: " & cadMen
        
        If b Then
            
            CADENA = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
            mCtaSocio = vParamAplic.Raiz_CtaAltaSoc & Format(codSocio, CADENA)
            
            cad = ""
            CadValues = ""
            ' si hay titulo
            
            ampliacion1 = Trim(DevuelveDesdeBDNew(conConta, "conceptos", "nomconce", "codconce", txtcodigo(4).Text, "N"))
            ampliacion2 = Trim(DevuelveDesdeBDNew(conConta, "conceptos", "nomconce", "codconce", txtcodigo(6).Text, "N"))
            
            ImporteD = 0
            ImporteH = 0
            
            ImpTotal = 0
          
            If ImpTitulo <> 0 Then  'Check1(0).Value Then
                ImpTotal = ImpTotal + ImpTitulo  ' vParamAplic.ImpTituloAlta
                
                ' apunte al debe
                i = i + 1
                
                cad = DBSet(txtcodigo(2).Text, "N") & "," & DBSet(txtcodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(i, "N") & "," & DBSet(mCtaSocio, "T") & ",'" & Documento & "',"
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If ImpTitulo > 0 Then
                    ' importe al debe en positivo
                    cad = cad & DBSet(txtcodigo(6).Text, "N") & "," & DBSet(ampliacion2, "T") & "," & DBSet(ImpTitulo, "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(vParamAplic.CtaTituloAlta, "T") & ",'CONTAB',0"
                
                    ImporteD = ImporteD + ImpTitulo
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    cad = cad & DBSet(txtcodigo(6).Text, "N") & "," & DBSet(ampliacion2, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet((ImpTitulo * -1), "N") & "," & DBSet(vParamAplic.CtaTituloAlta, "T") & "," & ValorNulo & ",'CONTAB',0"
                
                    ImporteH = ImporteH + (CCur(ImpTitulo) * (-1))
                End If
                
                cad = "(" & cad & "),"
            
                CadValues = CadValues & cad
                
                ' apunte al haber
                i = i + 1
                cad = DBSet(txtcodigo(2).Text, "N") & "," & DBSet(txtcodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(i, "N") & "," & DBSet(vParamAplic.CtaTituloAlta, "T") & ",'" & Documento & "',"
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If ImpTitulo > 0 Then
                    ' importe al haber en positivo
                    cad = cad & DBSet(txtcodigo(6).Text, "N") & "," & DBSet(ampliacion2, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet((ImpTitulo), "N") & "," & ValorNulo & "," & DBSet(mCtaSocio, "T") & ",'CONTAB',0"
                    ImporteH = ImporteH + (CCur(ImpTitulo))
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    cad = cad & DBSet(txtcodigo(6).Text, "N") & "," & DBSet(ampliacion2, "T") & "," & DBSet(ImpTitulo * (-1), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(mCtaSocio, "T") & ",'CONTAB',0"
                    
                    ImporteD = ImporteD + (CCur(ImpTitulo) * (-1))
                End If
                
                cad = "(" & cad & "),"
            
                CadValues = CadValues & cad
                
            End If
        
            
            If ImpGasto <> 0 Then
                ImpTotal = ImpTotal + ImpGasto
                
                ' apunte al debe
                i = i + 1
                
                cad = DBSet(txtcodigo(2).Text, "N") & "," & DBSet(txtcodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(i, "N") & "," & DBSet(mCtaSocio, "T") & ",'" & Documento & "',"
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If ImpGasto > 0 Then
                    ' importe al debe en positivo
                    cad = cad & DBSet(txtcodigo(4).Text, "N") & "," & DBSet(ampliacion1, "T") & "," & DBSet(ImpGasto, "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(vParamAplic.CtaGastoAlta, "T") & ",'CONTAB',0"
                
                    ImporteD = ImporteD + ImpGasto
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    cad = cad & DBSet(txtcodigo(4).Text, "N") & "," & DBSet(ampliacion1, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(ImpGasto * (-1), "N") & "," & ValorNulo & "," & DBSet(vParamAplic.CtaGastoAlta, "T") & ",'CONTAB',0"
                
                    ImporteH = ImporteH + (ImpGasto * (-1))
                End If
                
                cad = "(" & cad & "),"
            
                CadValues = CadValues & cad
                
                ' apunte al haber
                i = i + 1
                cad = DBSet(txtcodigo(2).Text, "N") & "," & DBSet(txtcodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(i, "N") & "," & DBSet(vParamAplic.CtaGastoAlta, "T") & ",'" & Documento & "',"
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If ImpGasto > 0 Then
                    ' importe al haber en positivo
                    cad = cad & DBSet(txtcodigo(4).Text, "N") & "," & DBSet(ampliacion1, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(ImpGasto, "N") & "," & ValorNulo & "," & DBSet(mCtaSocio, "T") & ",'CONTAB',0"
                    
                    ImporteH = ImporteH + (ImpGasto)
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    cad = cad & DBSet(txtcodigo(4).Text, "N") & "," & DBSet(ampliacion1, "T") & "," & DBSet(ImpGasto * (-1), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(mCtaSocio, "T") & ",'CONTAB',0"
                    ImporteD = ImporteD + (ImpGasto * (-1))
                End If
                
                cad = "(" & cad & "),"
            
                CadValues = CadValues & cad
            End If
        
'            ' total sobre la cuenta del banco
'            If ImpTotal <> 0 Then
'                ' apunte al debe
'                i = i + 1
'
'                mCtaBanco = DevuelveValor("select codmacta from sbanpr where codbanpr = " & DBSet(txtCodigo(1).Text, "N"))
'
'                cad = DBSet(txtCodigo(2).Text, "N") & "," & DBSet(txtCodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
'                cad = cad & DBSet(i, "N") & "," & DBSet(mCtaBanco, "T") & "," & ValorNulo & ","
'
'                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
'                If ImpTotal > 0 Then
'                    ' importe al debe en positivo
'                    cad = cad & DBSet(txtCodigo(4).Text, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpTotal, "N") & ","
'                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(mCtaSocio, "T") & ",'CONTAB',0"
'
'                    ImporteD = ImporteD + ImpTotal
'                Else
'                    ' importe al haber en positivo, cambiamos el signo
'                    cad = cad & DBSet(txtCodigo(6).Text, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
'                    cad = cad & DBSet(ImpTotal * (-1), "N") & "," & ValorNulo & "," & DBSet(mCtaSocio, "T") & ",'CONTAB',0"
'
'                    ImporteH = ImporteH + (ImpTotal * (-1))
'                End If
'
'                cad = "(" & cad & "),"
'
'                cadValues = cadValues & cad
'
'                ' apunte al haber
'                i = i + 1
'                cad = DBSet(txtCodigo(2).Text, "N") & "," & DBSet(txtCodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
'                cad = cad & DBSet(i, "N") & "," & DBSet(mCtaSocio, "T") & "," & ValorNulo & ","
'
'                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
'                If ImpTotal > 0 Then
'                    ' importe al haber en positivo
'                    cad = cad & DBSet(txtCodigo(6).Text, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
'                    cad = cad & DBSet(ImpTotal, "N") & "," & ValorNulo & "," & DBSet(mCtaBanco, "T") & ",'CONTAB',0"
'
'                    ImporteH = ImporteH + (ImpTotal)
'                Else
'                    ' importe al debe en positivo, cambiamos el signo
'                    cad = cad & DBSet(txtCodigo(6).Text, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpTotal * (-1), "N") & ","
'                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(mCtaBanco, "T") & ",'CONTAB',0"
'                    ImporteD = ImporteD + (ImpTotal * (-1))
'                End If
'
'                cad = "(" & cad & "),"
'
'                cadValues = cadValues & cad
'
'            End If
        
            cad = Mid(CadValues, 1, Len(CadValues) - 1)
        
            b = InsertarLinAsientoDia(cad, cadMen)
            cadMen = "Insertando Lineas Asiento: "
        End If
    
        If b And ImpGasto <> 0 Then b = InsertarEnTesoreria(0, CStr(codSocio), ImpGasto, "")
        If b And ImpTitulo <> 0 Then b = InsertarEnTesoreria(1, CStr(codSocio), ImpTitulo, "")
    Else
        b = False
    End If
    
    Set Mc = Nothing
    
    InsertarAsiento = b
    If b Then
        ConnConta.CommitTrans
        Exit Function
    Else
        ConnConta.RollbackTrans
        Exit Function
    End If
    
eInsertarAsiento:
    MuestraError Err.Number, "Insertar Asiento en Contabilidad", Err.Description
    ConnConta.RollbackTrans
End Function

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Importe As Currency
Dim mCtaSocio As String
Dim mCtaBanco As String
Dim codSocio As Long
Dim CADENA As String
Dim Sql As String
Dim ImporteGasto As Currency
Dim ImporteTitulo As Currency

Dim CtaPrev As String

    b = True
    
    If txtcodigo(3).Text = "" Then
        MsgBox "Debe introducir una fecha. Revise.", vbExclamation
        b = False
    End If
    
    If b And txtcodigo(0).Text = "" Then
        MsgBox "Debe introducir un Nro. de uve de socio. Revise.", vbExclamation
        b = False
    End If
    
    If b And Check1(1).Value Then ' hay contabilizacion
        If txtcodigo(1).Text = "" Then ' si no hay banco propio
            MsgBox "Debe introducir el banco propio para la contabilización. Revise.", vbExclamation
            b = False
        End If
        If b Then
            If txtcodigo(2).Text = "" Then ' si no hay diario
                MsgBox "Debe introducir el diario para la contabilización. Revise.", vbExclamation
                b = False
            End If
        End If
        If b Then
            If txtcodigo(4).Text = "" Then ' si no hay concepto al debe
                MsgBox "Debe introducir el concepto al debe para la contabilización. Revise.", vbExclamation
                b = False
            End If
        End If
        If b Then
            If txtcodigo(6).Text = "" Then ' si no hay concepto al debe
                MsgBox "Debe introducir el concepto al haber para la contabilización. Revise.", vbExclamation
                b = False
            End If
        End If
        ' comprobamos que las distintas cuentas existe
        ' cuenta del socio
        If b Then
            codSocio = DevuelveValor("select codclien from sclien where numeruve = " & DBSet(txtcodigo(0).Text, "N"))
            CADENA = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
            mCtaSocio = vParamAplic.Raiz_CtaAltaSoc & Format(codSocio, CADENA)
        
            Sql = DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", mCtaSocio, "T")
            If Sql = "" Then
                MsgBox "No existe la cuenta contable " & mCtaSocio & " en contabilidad. Revise.", vbExclamation
                b = False
            End If
        End If
        ' cuenta de titulo
        If b Then
            ' si hay titulo
            If txtcodigo(9).Text <> "" Then 'Check1(0).Value Then
                ImporteTitulo = CCur(ImporteSinFormato(txtcodigo(9).Text))
                If vParamAplic.CtaTituloAlta = "" Then
                    MsgBox "No existe la cuenta contable de título en parámetros. Revise.", vbExclamation
                    b = False
                Else
                    Sql = DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.CtaTituloAlta, "T")
                    If Sql = "" Then
                        MsgBox "No existe la cuenta contable de aportación " & vParamAplic.CtaTituloAlta & " en contabilidad. Revise.", vbExclamation
                        b = False
                    End If
                End If
            End If
        End If
        ' cuenta de gasto
        If b Then
            If txtcodigo(5).Text <> "" Then
                ImporteGasto = CCur(ImporteSinFormato(txtcodigo(5).Text))
                If ImporteGasto <> 0 Then
                    If vParamAplic.CtaGastoAlta = "" Then
                        MsgBox "No existe la cuenta contable de gasto en parámetros. Revise.", vbExclamation
                        b = False
                    Else
                        Sql = DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.CtaGastoAlta, "T")
                        If Sql = "" Then
                            MsgBox "No existe la cuenta contable de reserva legal " & vParamAplic.CtaGastoAlta & " en contabilidad. Revise.", vbExclamation
                            b = False
                        End If
                    End If
                End If
            End If
        End If
        ' cuenta de banco
        If b Then
            mCtaBanco = DevuelveValor("select codmacta from sbanpr where codbanpr = " & DBSet(txtcodigo(1).Text, "N"))
            If mCtaBanco = "" Then
                MsgBox "El banco no tiene asignada una cuenta contable. Revise.", vbExclamation
                b = False
            Else
                Sql = DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", mCtaBanco, "T")
                If Sql = "" Then
                    MsgBox "No existe la cuenta contable del pago " & mCtaBanco & " en contabilidad. Revise.", vbExclamation
                    b = False
                End If
            End If
        End If
    End If
    
    If b Then
        ' si hay contabilizacion o documento de admision
        If Check1(1).Value Or Check1(2).Value Then
            ImporteTitulo = 0
            If txtcodigo(9).Text <> "" Then 'Check1(0).Value Then
                ImporteTitulo = CCur(ImporteSinFormato(txtcodigo(9).Text))
            End If
            ImporteGasto = 0
            If txtcodigo(5).Text <> "" Then
                ImporteGasto = CCur(ImporteSinFormato(txtcodigo(5).Text))
            End If
            Importe = ImporteTitulo + ImporteGasto
            If Importe = 0 Then
                MsgBox "Debe introducir un importe. Revise.", vbExclamation
                b = False
            End If
        End If
    End If
    
    If b Then
        If Check1(1).Value Then
            If ImporteGasto <> 0 Then
                If txtcodigo(7).Text = "" Then
                    MsgBox "Debe introducir una forma de pago de Reserva Legal Obligatoria. Revise.", vbExclamation
                    b = False
                End If
            End If
            If b And ImporteTitulo <> 0 Then
                If txtcodigo(8).Text = "" Then
                    MsgBox "Debe introducir una forma de pago de Aportacion Capital Social. Revise.", vbExclamation
                    b = False
                End If
            End If
            
        End If
    End If

    DatosOk = b
    
End Function


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
Dim i As Integer

    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtcodigo(2)
        For i = 1 To Check1.Count
            Check1(i).Value = 1
        Next i
        ' Numero de uve
        txtcodigo(0).Text = NumCod
        PonerFormatoEntero txtcodigo(0)
        txtNombre(0).Text = PonerNombreDeCod(txtcodigo(0), conAri, "sclien", "nomclien", "numeruve", "N")
        ' fecha de documento
        txtcodigo(3).Text = Format(Now, "dd/mm/yyyy")
        txtcodigo(5).Text = Format(vParamAplic.ImpGastoAlta, "###,###,##0.00")
        txtcodigo(9).Text = Format(vParamAplic.ImpTituloAlta, "###,###,##0.00")
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    '###Descomentar
'    CommitConexion
    FrameCalidadesVisible True, H, W
    
    Frame2.Enabled = True
    
    indFrame = 2
    Tabla = "sclien"
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtcodigo(Indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(Indice).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(Indice).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    NumCampo = CadenaSeleccion
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub frmMtoV_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub frmTDia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmConce_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0 'SOCIOS
            AbrirFrmSocios (Index)
        
        Case 1 ' cuenta de pago, banco propio
            indCodigo = 1
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
        
        Case 2 ' TIPOS DE DIARIO
            AbrirFrmDiario (Index)
        
        Case 3, 4 'CONCEPTOS CONTABLES
            AbrirFrmConceptos (Index)
        
        Case 5, 6 ' FORMAS DE PAGO
            Indice = Index + 2
            indCodigo = Indice
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 7 ' banco propio
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
        
        
    End Select
    PonerFoco txtcodigo(indCodigo)
End Sub

Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgFec_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 0
            indCodigo = 3
   End Select
   
   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtcodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)
   '*******************************
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtcodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub imgFecha_Click(Index As Integer)

End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
'ahora
'    If KeyAscii = teclaBuscar Then
'        Select Case Index
'            Case 2: KEYBusqueda KeyAscii, 0 ' socio receptor de transmision
'            Case 1: KEYBusqueda KeyAscii, 1 'banco de pago
'        End Select
'    Else
        KEYpress KeyAscii
'    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'V Socio
            PonerFormatoEntero txtcodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sclien", "nomclien", "numeruve", "N")
            
        Case 1 ' banco propio
            PonerFormatoEntero txtcodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sbanpr", "nombanpr", "codbanpr", "N")
        
        Case 5, 9 ' Importe
            If txtcodigo(Index).Text <> "" Then PonerFormatoDecimal txtcodigo(Index), 1
            
            
        Case 3 ' Fecha del documento
            PonerFormatoFecha txtcodigo(Index)
            
            
        Case 2 ' NUMERO DE DIARIO
            If txtcodigo(Index).Text <> "" Then
                txtNombre(Index).Text = ""
                txtNombre(Index).Text = DevuelveDesdeBDNew(conConta, "tiposdiario", "desdiari", "numdiari", txtcodigo(Index).Text, "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "Número de Diario no existe en la contabilidad. Reintroduzca.", vbExclamation
'                    PonerFoco txtcodigo(Index)
                End If
            End If
        
        Case 4, 6 'CONCEPTOS
            If txtcodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreConcepto(txtcodigo(Index))
            If txtNombre(Index).Text = "" Then
                MsgBox "Número de Concepto no existe en la contabilidad. Reintroduzca.", vbExclamation
'                PonerFoco txtcodigo(Index)
            End If
            
            
        Case 7, 8 ' formas de pago
            PonerFormatoEntero txtcodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sforpa", "nomforpa", "codforpa", "N")
            
    End Select
End Sub


Private Sub FrameCalidadesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameCalidades.visible = visible
    If visible = True Then
        Me.FrameCalidades.Top = -90
        Me.FrameCalidades.Left = 0
        Me.FrameCalidades.Height = 7860
        Me.FrameCalidades.Width = 6165
        W = Me.FrameCalidades.Width
        H = Me.FrameCalidades.Height
    End If
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .ConSubInforme = True
'        .NombreRPT = cadNombreRPT
        .Opcion = OpcionListado
        .Show vbModal
    End With
End Sub


Private Sub AbrirFrmSocios(Indice As Integer)
    Set frmMtoV = New frmGesVSocio
    frmMtoV.DeConsulta = True
    frmMtoV.DatosADevolverBusqueda = "0|1|2|"
    frmMtoV.Show vbModal
    Set frmMtoV = Nothing
End Sub

Private Sub AbrirFrmDiario(Indice As Integer)
    indCodigo = 2
    Set frmTDia = New frmDiaConta
    frmTDia.DatosADevolverBusqueda = "0|1|"
    frmTDia.CodigoActual = txtcodigo(indCodigo)
    frmTDia.Show vbModal
    Set frmTDia = Nothing
End Sub

Private Sub AbrirFrmConceptos(Indice As Integer)
    Select Case Indice
        Case 3
            indCodigo = 4
        Case 4
            indCodigo = 6
    End Select
    Set frmConce = New frmConceConta
    frmConce.DatosADevolverBusqueda = "0|1|"
    frmConce.CodigoActual = txtcodigo(indCodigo)
    frmConce.Show vbModal
    Set frmConce = Nothing
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = OpcionListado
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub

Private Function InsertarCabAsientoDia(Diario As String, Asiento As String, Fecha As String, Obs As String, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim RS As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    
    cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
    cad = cad & "''," & ValorNulo & "," & DBSet(Obs, "T")
    cad = "(" & cad & ")"

    'Insertar en la contabilidad
    Sql = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) "
    Sql = Sql & " VALUES " & cad
    ConnConta.Execute Sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabAsientoDia = False
        cadErr = Err.Description
    Else
        InsertarCabAsientoDia = True
    End If
End Function



Private Function InsertarLinAsientoDia(cad As String, cadErr As String) As Boolean

Dim RS As ADODB.Recordset
Dim Aux As String
Dim Sql As String
Dim i As Byte
Dim TotImp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

 
    Sql = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
    Sql = Sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
    Sql = Sql & " VALUES " & cad
    
    ConnConta.Execute Sql

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinAsientoDia = False
        cadErr = Err.Description
    Else
        InsertarLinAsientoDia = True
    End If
End Function



Private Function InsertarEnTesoreria(Tipo As Byte, Socio As String, Importe As Currency, MenError As String) As Boolean
'Tipo 0 = Aportacion Capital social (gasto)
'     1 = Reserva Legal Obligatoria (titulo)

'Guarda datos de Tesoreria en tablas: aritaxi.svenci y en conta.scobros
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String, codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAuxConta As String 'para insertar en conta.scobro
Dim CadValues3 As String
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single 'importe para insertar en la svenci
Dim ImpVenci2 As Single 'importe para insertar en conta.scobro
Dim i As Byte
Dim TotalFactura2 As Currency   'Por si acaso lleva aportacion al terminal
'1 Julio 2009. Los graba en scobro
Dim CadenaDatosFiscales As String
Dim ForPago As String

Dim CADENA As String
Dim vSocio As CSocio
Dim vTextosCSB As String

Dim TipForPago As String
Dim CuentaPrev As String

Dim LEtra As String

    On Error GoTo EInsertarTesoreria

    b = False
    InsertarEnTesoreria = False

    vTextosCSB = "NULL,NULL,NULL"
    
    CadValues3 = ""
    CadValues = ""
    CadValues2 = ""

    'campo para insertar en conta.scobro de Tesoreria
    If Tipo = 0 Then
        textcsb33 = "Cuota de ingreso" '"Reserva Legal Obligatoria"
        ForPago = CCur(txtcodigo(7).Text)
        LEtra = "'V'"
    Else
        textcsb33 = "Aportación Capital Social"
        ForPago = CCur(txtcodigo(8).Text)
        LEtra = "'W'"
    End If
    
    Set vSocio = New CSocio
    If vSocio.LeerDatos(Socio) Then
        CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", txtcodigo(1).Text, "N")

    
    
        'Datos fiscales en scobro     Julio 2009
        'nomclien,domclien,pobclien, cpclien,proclien
        CadenaDatosFiscales = DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Domicilio, "T") & "," & DBSet(vSocio.Poblacion, "T")
        CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(vSocio.CPostal, "T") & "," & DBSet(vSocio.Provincia, "T")
        
        
        'Obtener el Nº de Vencimientos de la forma de pago
        Sql = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
        Set rsVenci = New ADODB.Recordset
        rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not rsVenci.EOF Then
            If rsVenci!numerove > 0 And CCur(Importe) <> 0 Then
            
                'Comporbamos si el importe es <>0
            
                'Obtener los dias de pago del cliente : el socio no tiene dias de pago
                Sql = " SELECT  0 diapago1, 0 diapago2, 0 diapago3, 0 mesnogir, 0 diavtoat, '' codmacta "
                Sql = Sql & " FROM sclien "
                Sql = Sql & " WHERE codclien=" & Socio
                Set RS = New ADODB.Recordset
                RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
                CADENA = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
                codmacta = vParamAplic.Raiz_CtaAltaSoc & Format(Socio, CADENA)
    
    '            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
                
                If Not RS.EOF Then
                    cadValuesAux = "(" & LEtra & ", " & Socio & ", '" & Format(txtcodigo(3).Text, FormatoFecha) & "', "
                    CadValuesAuxConta = "(" & LEtra & ", " & Socio & ", '" & Format(txtcodigo(3).Text, FormatoFecha) & "', "
                    '                    Añadire a la cadena fija esta los valores de textcsb41,txcs
                    CadValuesAuxConta = CadValuesAuxConta & vTextosCSB & ","
                    '-------- Primer Vencimiento
                    i = 1
                    'FECHA VTO
                    FecVenci = CDate(txtcodigo(3).Text)
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                    FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                    '===
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    TipForPago = DevuelveDesdeBDNew(conConta, "sforpa", "tipforpa", "codforpa", CStr(ForPago), "N")
                    If CCur(TipForPago) <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If CInt(DBLet(RS!mesnogir, "N")) <> 0 Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(RS!mesnogir, "N"), DBLet(RS!DiaVtoAt, "N"), DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                    End If
                    
                    'Comprobar si cliente tiene dia de vencimiento atrasado
                    CadValues = cadValuesAux & i & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValuesAuxConta & i & ", "
                    CadValues2 = CadValues2 & DBSet(codmacta, "T") & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                    'IMPORTE del Vencimiento
                    TotalFactura2 = Importe
                    If rsVenci!numerove = 1 Then
                        ImpVenci = TotalFactura2
                        ImpVenci2 = TotalFactura2
                    Else

                        ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                        ImpVenci2 = Round2((TotalFactura2) / rsVenci!numerove, 2)
                        'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                        If ImpVenci * rsVenci!numerove <> TotalFactura2 Then
                            ImpVenci = Round(ImpVenci + (TotalFactura2 - ImpVenci * rsVenci!numerove), 2)
                        End If
                        'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                        If (ImpVenci2 * rsVenci!numerove) <> TotalFactura2 Then
                            ImpVenci2 = Round(ImpVenci2 + (TotalFactura2 - (ImpVenci2 * rsVenci!numerove)), 2)
                        End If
                    End If
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', " & DBSet(vSocio.Banco, "N") & ", " & DBSet(vSocio.Sucursal, "N") & ", " & DBSet(vSocio.DigControl, "T") & ", " & DBSet(vSocio.CuentaBan, "T") & ", '" & textcsb33 & "', " & DBSet(vParamAplic.PorDefecto_Agente, "N")
                    'departamento y transfer
                    CadValues2 = CadValues2 & "," & DBSet("", "N", "S") & ",NULL"
                    '1 Julio 2009
                    ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                     CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                    
                    '[Monica]22/11/2013: tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                       CadValues2 = CadValues2 & "," & DBSet(vSocio.Iban, "T", "S") & ")"
                    Else
                       CadValues2 = CadValues2 & ")"
                    End If
                     
                     
                     
                    
                    'Resto Vencimientos
                    '--------------------------------------------------------------------
                    For i = 2 To rsVenci!numerove
                       'FECHA Resto Vencimientos
                        '=== Laura 23/01/2007
                        'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                        FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                        '===
                        'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                        If TipForPago <> 0 Then
                            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                        Else
                            FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                        End If
                        'Comprobar si cliente tiene mes a no girar
                        FecVenci1 = FecVenci
                        If DBLet(RS!mesnogir, "N") <> "0" Then
                            FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(RS!mesnogir, "N"), DBLet(RS!DiaVtoAt, "N"), DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                        End If

                        CadValues = CadValues & ", " & cadValuesAux & i & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                        CadValues2 = CadValues2 & ", " & CadValuesAuxConta & i & ", " & DBSet(codmacta, "T") & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "

                        'IMPORTE Resto de Vendimientos
                        ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                        ImpVenci2 = Round2((TotalFactura2) / rsVenci!numerove, 2)
                        CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & DBSet(vSocio.Banco, "N") & ", " & DBSet(vSocio.Sucursal, "N") & ", " & DBSet(vSocio.DigControl, "T") & ", " & DBSet(vSocio.CuentaBan, "T") & ", '" & textcsb33 & "', " & DBSet(vParamAplic.PorDefecto_Agente, "N") & ", "
                        CadValues2 = CadValues2 & DBSet("", "N", "S") & ",NULL"
                        '1 Julio 2009
                        ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                        CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & "," & DBSet(vSocio.Iban, "T", "S") & ")"
                        Else
                           CadValues2 = CadValues2 & ")"
                        End If
                        

                    Next i
                    
                End If
                RS.Close
            Else
                'totalfac =0 and numerovtos >=1
                b = True
            End If
            
            Set RS = Nothing
        End If
        rsVenci.Close
        Set rsVenci = Nothing
        
        'Grabar tabla scobro de la CONTABILIDAD
        '-------------------------------------------------
        If CadValues2 <> "" Then
            '01/09/06
    '        If (NumTicket = "") Or (NumTicket <> "" And TipForPago <> 0) Then
                If CuentaPrev <> "" Then
                    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
                    'forma de pago de la factura. Sino existe insertarla
                    'vemos si existe en la conta
                    b = InsertarFormaPagoEnConta(CStr(ForPago), MenError)
                    
                    If b Then
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl,text41csb ,text42csb, text43csb, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur,"
                        Sql = Sql & "digcontr, cuentaba,text33csb,agente,departamento,transfer "
                        'JULIO 2009
                        'nomclien,domclien,pobclien, cpclien,proclien
                        Sql = Sql & ",nomclien,domclien,pobclien, cpclien,proclien" ')"
                        
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           Sql = Sql & ",iban)"
                        Else
                           Sql = Sql & ")"
                        End If
                       
                        Sql = Sql & " VALUES " & CadValues2
                        ConnConta.Execute Sql
                    End If
                Else
                    'DAVID ####
                    'ENERO 2008
                    'Si no inserto en tesoreria por que ctaprvsta="" entonces dejo continuar
                    b = True
            End If
        End If
    End If
    
    Set vSocio = Nothing
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Insertar en Tesoreria: " & vbCrLf & Err.Description
    End If
    InsertarEnTesoreria = b
End Function



Private Function InsertarFormaPagoEnConta(nForPa As String, cadErr As String) As Boolean
Dim cadAux As String
Dim cadAux2 As String
Dim Sql As String

    On Error GoTo ErrInsForpa
    InsertarFormaPagoEnConta = False
    
    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
    'forma de pago de la factura. Sino existe insertarla
    
    'vemos si existe en la conta
    cadAux = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", nForPa, "N")
    'si no existe la forma de pago en conta, insertamos la de aritaxi
    If cadAux = "" Then
        cadAux2 = "tipforpa"
        cadAux = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", nForPa, "N", cadAux2)
        If cadAux <> "" Then
            'insertamos e sforpa de la CONTA
            Sql = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
            Sql = Sql & " VALUES(" & nForPa & ", " & DBSet(cadAux, "T") & ", " & cadAux2 & ")"
            ConnConta.Execute Sql
            InsertarFormaPagoEnConta = True
        Else
            InsertarFormaPagoEnConta = False
        End If
    Else
        InsertarFormaPagoEnConta = True
    End If
    
    
    Exit Function
    
ErrInsForpa:
    InsertarFormaPagoEnConta = False
    cadErr = "Insertar forma de pago en Contablilidad: " & vbCrLf & Err.Description
End Function



