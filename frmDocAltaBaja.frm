VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDocAltaBaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6930
   Icon            =   "frmDocAltaBaja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCalidades 
      Height          =   8310
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6855
      Begin VB.CheckBox Check1 
         Caption         =   "Certificado aportaci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3420
         TabIndex        =   39
         Top             =   4275
         Width           =   2850
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
         Index           =   9
         Left            =   4770
         MaxLength       =   12
         TabIndex        =   3
         Top             =   2340
         Width           =   1305
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilizaci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2835
         Left            =   90
         TabIndex        =   25
         Top             =   4725
         Width           =   6555
         Begin VB.TextBox txtNombre 
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
            Index           =   8
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   2250
            Width           =   3315
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
            Index           =   8
            Left            =   2310
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Cta.Dif.negativas|T|S|||sparam|ctanegtat|||"
            Top             =   2250
            Width           =   795
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   7
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1890
            Width           =   3315
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
            Index           =   7
            Left            =   2310
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta.Dif.negativas|T|S|||sparam|ctanegtat|||"
            Top             =   1890
            Width           =   795
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   2  'Center
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
            MaxLength       =   6
            TabIndex        =   5
            Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
            Top             =   750
            Width           =   795
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   1
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   750
            Width           =   3285
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
            Index           =   6
            Left            =   2310
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Cta.Dif.negativas|T|S|||sparam|ctanegtat|||"
            Top             =   1530
            Width           =   795
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   6
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1530
            Width           =   3315
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
            Index           =   4
            Left            =   2310
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1140
            Width           =   795
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   4
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   1140
            Width           =   3285
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
            Index           =   2
            Left            =   2310
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   390
            Width           =   795
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   390
            Width           =   3285
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Reserva"
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
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   1890
            Width           =   1830
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   2010
            ToolTipText     =   "Buscar forma pago"
            Top             =   2250
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Aportac."
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
            Index           =   1
            Left            =   150
            TabIndex        =   35
            Top             =   2250
            Width           =   1980
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   2010
            ToolTipText     =   "Buscar forma pago"
            Top             =   1890
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   2010
            ToolTipText     =   "Buscar concepto"
            Top             =   1530
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   2010
            ToolTipText     =   "Buscar concepto"
            Top             =   1140
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   2010
            ToolTipText     =   "Buscar diario"
            Top             =   390
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
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
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   780
            Width           =   1380
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   2010
            ToolTipText     =   "Buscar cuenta"
            Top             =   750
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concep.Reserva"
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
            Index           =   24
            Left            =   120
            TabIndex        =   31
            Top             =   1170
            Width           =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "Concep.Aportac."
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
            Height          =   285
            Index           =   25
            Left            =   120
            TabIndex        =   30
            Top             =   1530
            Width           =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "N�mero Diario "
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
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   420
            Width           =   1860
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1380
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ficha Socio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   3420
         TabIndex        =   15
         Top             =   3900
         Width           =   2310
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Notificaci�n Banco"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   3420
         TabIndex        =   14
         Top             =   3510
         Width           =   2310
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Documento Admisi�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   3420
         TabIndex        =   13
         Top             =   3120
         Width           =   2460
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Realizar Contabilizaci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   330
         TabIndex        =   12
         Top             =   3090
         Width           =   2670
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
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
         Left            =   1500
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Num vehiculo|N|N|||shilla|numeruve|000000|S|"
         Top             =   900
         Width           =   795
      End
      Begin VB.TextBox txtNombre 
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
         TabIndex        =   20
         Top             =   900
         Width           =   3735
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
         Index           =   5
         Left            =   4770
         MaxLength       =   12
         TabIndex        =   2
         Top             =   1830
         Width           =   1305
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   2
         Left            =   5490
         TabIndex        =   11
         Top             =   7725
         Width           =   1135
      End
      Begin VB.CommandButton CmdAceptar 
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
         Index           =   2
         Left            =   4230
         TabIndex        =   10
         Top             =   7725
         Width           =   1135
      End
      Begin VB.CommandButton Command8 
         Height          =   440
         Left            =   7860
         Picture         =   "frmDocAltaBaja.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command7 
         Height          =   440
         Left            =   7860
         Picture         =   "frmDocAltaBaja.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Importe Aportaci�n Capital Social titulo"
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
         Index           =   1
         Left            =   150
         TabIndex        =   38
         Top             =   2370
         Width           =   4230
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1200
         ToolTipText     =   "Buscar fecha"
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Impresi�n"
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
         Index           =   0
         Left            =   3420
         TabIndex        =   23
         Top             =   2820
         Width           =   1110
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Importe Reserva Legal Obligatoria gto"
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
         Left            =   150
         TabIndex        =   22
         Top             =   1860
         Width           =   4170
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
         Left            =   180
         TabIndex        =   21
         Top             =   930
         Width           =   750
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1200
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
         Left            =   180
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

Private Const IdPrograma = 206

Public OpcionListado As Byte
    '==== Listados BASICOS ====
    '=============================
    '
    
Public NumCod As String 'Para indicar nro de uve de socio
Public Socio As String ' para indicar codigo de socio para el certificado

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

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
Private WithEvents frmTDia As frmBasico2 'diarios de contabilidad
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
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim indice As Integer

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
            txtCodigo(1).Text = ""
            txtCodigo(2).Text = ""
            txtCodigo(4).Text = ""
            txtCodigo(6).Text = ""
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
    
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    ' Numero de uve del socio
    cadParam = cadParam & "pCodigo=" & NumCod & "|"
    numParam = numParam + 1
    
    'alta de socios
    cadTitulo = "Documentos Alta de Socios"
    
    ImpGasto = 0
    If txtCodigo(5).Text <> "" Then
        ImpGasto = txtCodigo(5).Text 'TransformaComasPuntos(ImporteSinFormato(txtCodigo(5).Text))
        cadParam = cadParam & "pGasto=" & TransformaComasPuntos(ImporteSinFormato(txtCodigo(5).Text)) & "|"
        numParam = numParam + 1
    End If
    
    ImpTitulo = 0
    If txtCodigo(9).Text <> "" Then 'check1(0).Value Then
        ImpTitulo = txtCodigo(9).Text 'TransformaComasPuntos(ImporteSinFormato(txtCodigo(9).Text))
        cadParam = cadParam & "pTitulo=" & TransformaComasPuntos(ImporteSinFormato(txtCodigo(9).Text)) & "|"
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
    
    
    cadParam = cadParam & "pFecha=""" & txtCodigo(3).Text & """|"
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
        
        '[Monica]15/03/2019:
        If Check1(0).Value Then
            cadTitulo = "Certificado de Aportaci�n"
        
        
            indRPT = 58 ' certificado aportaciones
            
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, False, "") Then Exit Sub
            
            cadFormula = "{sclien.codclien} = " & Socio
            
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
            cadParam = cadParam & "pFecha=""" & txtCodigo(3).Text & """|"
            numParam = numParam + 1
            cadParam = cadParam & "pImporte=" & vParamAplic.ImpTituloAlta & "|"
            numParam = numParam + 1
            cadParam = cadParam & "pPresidente=""" & vParamAplic.AporPresidente & """|"
            numParam = numParam + 1
            cadParam = cadParam & "pSecretario=""" & vParamAplic.AporSecretario & """|"
            numParam = numParam + 1
            cadParam = cadParam & "pSerie=""" & vParamAplic.AporSerie & """|"
            numParam = numParam + 1
            cadParam = cadParam & "pNroTitulo=" & vParamAplic.AporNumero & "|"
            numParam = numParam + 1
            
            frmImprimir.NombreRPT = nomDocu
            
            LlamarImprimir
            
            If MsgBox("� Impresion correcta para actualizar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                If ActualizarTablas Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                End If
            End If
            
        End If
        
        If Check1(1).Value Then
            If MsgBox("� Desea continuar con la contabilizaci�n ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
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


Private Function ActualizarTablas() As Boolean
Dim Sql As String
Dim numF As Long
Dim b As Boolean
    
    On Error GoTo eActualizarTablas
    
    b = False
    
    conn.BeginTrans
    
    Sql = "select coalesce(max(sclien_aportaciones.numlinea),0) + 1  from sclien_aportaciones where codsocio= " & DBSet(Socio, "N")
    numF = DevuelveValor(Sql)
    
    Sql = "insert into sclien_aportaciones (codsocio,numlinea,nroapor,serieapor,importe,fecha) values ("
    Sql = Sql & DBSet(Socio, "N") & "," & DBSet(numF, "N") & "," & DBSet(vParamAplic.AporNumero + 1, "N") & "," & DBSet(vParamAplic.AporSerie, "T") & ","
    Sql = Sql & DBSet(vParamAplic.ImpTituloAlta, "N") & "," & DBSet(txtCodigo(3).Text, "F") & ")"
    
    conn.Execute Sql
    
    vParamAplic.AporNumero = vParamAplic.AporNumero + 1
    b = (vParamAplic.Modificar(1) = 1)
    
eActualizarTablas:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        ActualizarTablas = False
    Else
        conn.CommitTrans
        ActualizarTablas = True
    End If
End Function

Private Function InsertarAsiento() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Lineas As String
Dim i As Integer

Dim Mc As Contadores
Dim Cad As String
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
    
    If Mc.ConseguirContador("0", (CDate(txtCodigo(3).Text) <= CDate(vEmpresa.FechaFin)), True) = 0 Then
    
        Obs = "Alta de Socios "
        i = 0
    
        codSocio = DevuelveValor("select codclien from sclien where numeruve = " & DBSet(txtCodigo(0).Text, "N"))
        Documento = "SOC-" & Format(codSocio, "000000")
    
        'Insertar en la conta Cabecera Asiento
        b = InsertarCabAsientoDia(txtCodigo(2).Text, Mc.Contador, CDate(txtCodigo(3).Text), Obs, cadMen)
        cadMen = "Insertando Cab. Asiento: " & cadMen
        
        If b Then
            
            CADENA = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
            mCtaSocio = vParamAplic.Raiz_CtaAltaSoc & Format(codSocio, CADENA)
            
            Cad = ""
            CadValues = ""
            ' si hay titulo
            
            ampliacion1 = Trim(DevuelveDesdeBDNew(conConta, "conceptos", "nomconce", "codconce", txtCodigo(4).Text, "N"))
            ampliacion2 = Trim(DevuelveDesdeBDNew(conConta, "conceptos", "nomconce", "codconce", txtCodigo(6).Text, "N"))
            
            ImporteD = 0
            ImporteH = 0
            
            ImpTotal = 0
          
            If ImpTitulo <> 0 Then  'Check1(0).Value Then
                ImpTotal = ImpTotal + ImpTitulo  ' vParamAplic.ImpTituloAlta
                
                ' apunte al debe
                i = i + 1
                
                Cad = DBSet(txtCodigo(2).Text, "N") & "," & DBSet(txtCodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                Cad = Cad & DBSet(i, "N") & "," & DBSet(mCtaSocio, "T") & ",'" & Documento & "',"
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If ImpTitulo > 0 Then
                    ' importe al debe en positivo
                    Cad = Cad & DBSet(txtCodigo(6).Text, "N") & "," & DBSet(ampliacion2, "T") & "," & DBSet(ImpTitulo, "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(vParamAplic.CtaTituloAlta, "T") & ",'CONTAB',0"
                
                    ImporteD = ImporteD + ImpTitulo
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    Cad = Cad & DBSet(txtCodigo(6).Text, "N") & "," & DBSet(ampliacion2, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet((ImpTitulo * -1), "N") & "," & DBSet(vParamAplic.CtaTituloAlta, "T") & "," & ValorNulo & ",'CONTAB',0"
                
                    ImporteH = ImporteH + (CCur(ImpTitulo) * (-1))
                End If
                
                Cad = "(" & Cad & "),"
            
                CadValues = CadValues & Cad
                
                ' apunte al haber
                i = i + 1
                Cad = DBSet(txtCodigo(2).Text, "N") & "," & DBSet(txtCodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                Cad = Cad & DBSet(i, "N") & "," & DBSet(vParamAplic.CtaTituloAlta, "T") & ",'" & Documento & "',"
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If ImpTitulo > 0 Then
                    ' importe al haber en positivo
                    Cad = Cad & DBSet(txtCodigo(6).Text, "N") & "," & DBSet(ampliacion2, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet((ImpTitulo), "N") & "," & ValorNulo & "," & DBSet(mCtaSocio, "T") & ",'CONTAB',0"
                    ImporteH = ImporteH + (CCur(ImpTitulo))
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    Cad = Cad & DBSet(txtCodigo(6).Text, "N") & "," & DBSet(ampliacion2, "T") & "," & DBSet(ImpTitulo * (-1), "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(mCtaSocio, "T") & ",'CONTAB',0"
                    
                    ImporteD = ImporteD + (CCur(ImpTitulo) * (-1))
                End If
                
                Cad = "(" & Cad & "),"
            
                CadValues = CadValues & Cad
                
            End If
        
            
            If ImpGasto <> 0 Then
                ImpTotal = ImpTotal + ImpGasto
                
                ' apunte al debe
                i = i + 1
                
                Cad = DBSet(txtCodigo(2).Text, "N") & "," & DBSet(txtCodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                Cad = Cad & DBSet(i, "N") & "," & DBSet(mCtaSocio, "T") & ",'" & Documento & "',"
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If ImpGasto > 0 Then
                    ' importe al debe en positivo
                    Cad = Cad & DBSet(txtCodigo(4).Text, "N") & "," & DBSet(ampliacion1, "T") & "," & DBSet(ImpGasto, "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(vParamAplic.CtaGastoAlta, "T") & ",'CONTAB',0"
                
                    ImporteD = ImporteD + ImpGasto
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    Cad = Cad & DBSet(txtCodigo(4).Text, "N") & "," & DBSet(ampliacion1, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet(ImpGasto * (-1), "N") & "," & ValorNulo & "," & DBSet(vParamAplic.CtaGastoAlta, "T") & ",'CONTAB',0"
                
                    ImporteH = ImporteH + (ImpGasto * (-1))
                End If
                
                Cad = "(" & Cad & "),"
            
                CadValues = CadValues & Cad
                
                ' apunte al haber
                i = i + 1
                Cad = DBSet(txtCodigo(2).Text, "N") & "," & DBSet(txtCodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                Cad = Cad & DBSet(i, "N") & "," & DBSet(vParamAplic.CtaGastoAlta, "T") & ",'" & Documento & "',"
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If ImpGasto > 0 Then
                    ' importe al haber en positivo
                    Cad = Cad & DBSet(txtCodigo(4).Text, "N") & "," & DBSet(ampliacion1, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet(ImpGasto, "N") & "," & ValorNulo & "," & DBSet(mCtaSocio, "T") & ",'CONTAB',0"
                    
                    ImporteH = ImporteH + (ImpGasto)
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    Cad = Cad & DBSet(txtCodigo(4).Text, "N") & "," & DBSet(ampliacion1, "T") & "," & DBSet(ImpGasto * (-1), "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(mCtaSocio, "T") & ",'CONTAB',0"
                    ImporteD = ImporteD + (ImpGasto * (-1))
                End If
                
                Cad = "(" & Cad & "),"
            
                CadValues = CadValues & Cad
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
        
            Cad = Mid(CadValues, 1, Len(CadValues) - 1)
        
            b = InsertarLinAsientoDia(Cad, cadMen)
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
    
    If txtCodigo(3).Text = "" Then
        MsgBox "Debe introducir una fecha. Revise.", vbExclamation
        b = False
    End If
    
    If b And txtCodigo(0).Text = "" Then
        MsgBox "Debe introducir un Nro. de uve de socio. Revise.", vbExclamation
        b = False
    End If
    
    If b And Check1(1).Value Then ' hay contabilizacion
        If txtCodigo(1).Text = "" Then ' si no hay banco propio
            MsgBox "Debe introducir el banco propio para la contabilizaci�n. Revise.", vbExclamation
            b = False
        End If
        If b Then
            If txtCodigo(2).Text = "" Then ' si no hay diario
                MsgBox "Debe introducir el diario para la contabilizaci�n. Revise.", vbExclamation
                b = False
            End If
        End If
        If b Then
            If txtCodigo(4).Text = "" Then ' si no hay concepto al debe
                MsgBox "Debe introducir el concepto al debe para la contabilizaci�n. Revise.", vbExclamation
                b = False
            End If
        End If
        If b Then
            If txtCodigo(6).Text = "" Then ' si no hay concepto al debe
                MsgBox "Debe introducir el concepto al haber para la contabilizaci�n. Revise.", vbExclamation
                b = False
            End If
        End If
        ' comprobamos que las distintas cuentas existe
        ' cuenta del socio
        If b Then
            codSocio = DevuelveValor("select codclien from sclien where numeruve = " & DBSet(txtCodigo(0).Text, "N"))
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
            If txtCodigo(9).Text <> "" Then 'Check1(0).Value Then
                ImporteTitulo = CCur(ImporteSinFormato(txtCodigo(9).Text))
                If vParamAplic.CtaTituloAlta = "" Then
                    MsgBox "No existe la cuenta contable de t�tulo en par�metros. Revise.", vbExclamation
                    b = False
                Else
                    Sql = DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.CtaTituloAlta, "T")
                    If Sql = "" Then
                        MsgBox "No existe la cuenta contable de aportaci�n " & vParamAplic.CtaTituloAlta & " en contabilidad. Revise.", vbExclamation
                        b = False
                    End If
                End If
            End If
        End If
        ' cuenta de gasto
        If b Then
            If txtCodigo(5).Text <> "" Then
                ImporteGasto = CCur(ImporteSinFormato(txtCodigo(5).Text))
                If ImporteGasto <> 0 Then
                    If vParamAplic.CtaGastoAlta = "" Then
                        MsgBox "No existe la cuenta contable de gasto en par�metros. Revise.", vbExclamation
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
            mCtaBanco = DevuelveValor("select codmacta from sbanpr where codbanpr = " & DBSet(txtCodigo(1).Text, "N"))
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
            '[MOnica]22/03/2019: solo en el caso de contabilizacion
        If Check1(1).Value Then 'Or Check1(2).Value Then
            ImporteTitulo = 0
            If txtCodigo(9).Text <> "" Then 'Check1(0).Value Then
                ImporteTitulo = CCur(ImporteSinFormato(txtCodigo(9).Text))
            End If
            ImporteGasto = 0
            If txtCodigo(5).Text <> "" Then
                ImporteGasto = CCur(ImporteSinFormato(txtCodigo(5).Text))
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
                If txtCodigo(7).Text = "" Then
                    MsgBox "Debe introducir una forma de pago de Reserva Legal Obligatoria. Revise.", vbExclamation
                    b = False
                End If
            End If
            If b And ImporteTitulo <> 0 Then
                If txtCodigo(8).Text = "" Then
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
        PonerFoco txtCodigo(2)
        For i = 0 To Check1.Count - 1
            Check1(i).Value = 1
        Next i
        ' Numero de uve
        txtCodigo(0).Text = NumCod
        PonerFormatoEntero txtCodigo(0)
        txtnombre(0).Text = PonerNombreDeCod(txtCodigo(0), conAri, "sclien", "nomclien", "numeruve", "N")
        ' fecha de documento
        txtCodigo(3).Text = Format(Now, "dd/mm/yyyy")
        txtCodigo(5).Text = Format(vParamAplic.ImpGastoAlta, "###,###,##0.00")
        txtCodigo(9).Text = Format(vParamAplic.ImpTituloAlta, "###,###,##0.00")
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer, i As Integer
Dim List As Collection
    
    'Icono del formulario
    Me.Icon = frmppal.Icon

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
    
    For i = 0 To Me.imgBuscar.Count - 1
        imgBuscar(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    For i = 0 To Me.imgFec.Count - 1
        imgFec(i).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next
    
    
End Sub


Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indice).Text = RecuperaValor(CadenaSeleccion, 1)
    txtnombre(indice).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    NumCampo = CadenaSeleccion
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub frmMtoV_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub frmTDia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmConce_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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
            indice = Index + 2
            indCodigo = indice
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
    PonerFoco txtCodigo(indCodigo)
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
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
   '*******************************
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub imgFecha_Click(Index As Integer)

End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
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

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'V Socio
            PonerFormatoEntero txtCodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sclien", "nomclien", "numeruve", "N")
            
        Case 1 ' banco propio
            PonerFormatoEntero txtCodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sbanpr", "nombanpr", "codbanpr", "N")
        
        Case 5, 9 ' Importe
            If txtCodigo(Index).Text <> "" Then PonerFormatoDecimal txtCodigo(Index), 1
            
            
        Case 3 ' Fecha del documento
            PonerFormatoFecha txtCodigo(Index)
            
            
        Case 2 ' NUMERO DE DIARIO
            If txtCodigo(Index).Text <> "" Then
                txtnombre(Index).Text = ""
                txtnombre(Index).Text = DevuelveDesdeBDNew(conConta, "tiposdiario", "desdiari", "numdiari", txtCodigo(Index).Text, "N")
                If txtnombre(Index).Text = "" Then
                    MsgBox "N�mero de Diario no existe en la contabilidad. Reintroduzca.", vbExclamation
'                    PonerFoco txtcodigo(Index)
                End If
            End If
        
        Case 4, 6 'CONCEPTOS
            If txtCodigo(Index).Text <> "" Then txtnombre(Index).Text = PonerNombreConcepto(txtCodigo(Index))
            If txtnombre(Index).Text = "" Then
                MsgBox "N�mero de Concepto no existe en la contabilidad. Reintroduzca.", vbExclamation
'                PonerFoco txtcodigo(Index)
            End If
            
            
        Case 7, 8 ' formas de pago
            PonerFormatoEntero txtCodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sforpa", "nomforpa", "codforpa", "N")
            
    End Select
End Sub


Private Sub FrameCalidadesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameCalidades.visible = visible
    If visible = True Then
        Me.FrameCalidades.top = -90
        Me.FrameCalidades.Left = 0
        Me.FrameCalidades.Height = 8385
        Me.FrameCalidades.Width = 6855
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


Private Sub AbrirFrmSocios(indice As Integer)
    Set frmMtoV = New frmGesVSocio
    frmMtoV.DeConsulta = True
    frmMtoV.DatosADevolverBusqueda = "0|1|2|"
    frmMtoV.Show vbModal
    Set frmMtoV = Nothing
End Sub

Private Sub AbrirFrmDiario(indice As Integer)

    indCodigo = 2
    
    Set frmTDia = New frmBasico2
    
    AyudaDiarios frmTDia
    
    Set frmTDia = Nothing
End Sub

Private Sub AbrirFrmConceptos(indice As Integer)
    Select Case indice
        Case 3
            indCodigo = 4
        Case 4
            indCodigo = 6
    End Select
    Set frmConce = New frmConceConta
    frmConce.DatosADevolverBusqueda = "0|1"
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
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    If vParamAplic.ContabilidadNueva Then
        Cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
        Cad = Cad & DBSet(Obs, "T") & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARITAXI'"
        Cad = "(" & Cad & ")"
        
        'Insertar en la contabilidad
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) "
        Sql = Sql & " VALUES " & Cad
    Else
        Cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
        Cad = Cad & "''," & ValorNulo & "," & DBSet(Obs, "T")
        Cad = "(" & Cad & ")"
    
        'Insertar en la contabilidad
        Sql = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) "
        Sql = Sql & " VALUES " & Cad
        ConnConta.Execute Sql
    End If
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabAsientoDia = False
        cadErr = Err.Description
    Else
        InsertarCabAsientoDia = True
    End If
End Function



Private Function InsertarLinAsientoDia(Cad As String, cadErr As String) As Boolean

Dim Rs As ADODB.Recordset
Dim Aux As String
Dim Sql As String
Dim i As Byte
Dim TotImp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    If vParamAplic.ContabilidadNueva Then
        Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        Sql = Sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
        Sql = Sql & " VALUES " & Cad
    Else
        Sql = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        Sql = Sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
        Sql = Sql & " VALUES " & Cad
    End If
    
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
Dim Rs As ADODB.Recordset
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
Dim vvIban As String

    On Error GoTo EInsertarTesoreria

    b = False
    InsertarEnTesoreria = False


    If vParamAplic.ContabilidadNueva Then
        vTextosCSB = "NULL"
    Else
        vTextosCSB = "NULL,NULL,NULL"
    End If
    
    CadValues3 = ""
    CadValues = ""
    CadValues2 = ""

    'campo para insertar en conta.scobro de Tesoreria
    If Tipo = 0 Then
        textcsb33 = "Cuota de ingreso" '"Reserva Legal Obligatoria"
        ForPago = CCur(txtCodigo(7).Text)
        LEtra = "'V'"
    Else
        textcsb33 = "Aportaci�n Capital Social"
        ForPago = CCur(txtCodigo(8).Text)
        LEtra = "'W'"
    End If
    
    Set vSocio = New CSocio
    If vSocio.LeerDatos(Socio) Then
        CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", txtCodigo(1).Text, "N")

    
    
        'Datos fiscales en scobro     Julio 2009
        'nomclien,domclien,pobclien, cpclien,proclien
        CadenaDatosFiscales = DBSet(vSocio.NOMBRE, "T") & "," & DBSet(vSocio.Domicilio, "T") & "," & DBSet(vSocio.Poblacion, "T")
        CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(vSocio.CPostal, "T") & "," & DBSet(vSocio.Provincia, "T")
        
        If vParamAplic.ContabilidadNueva Then
            CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(vSocio.NIF, "T") & ",'ES'"
        End If
        
        'Obtener el N� de Vencimientos de la forma de pago
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
                Set Rs = New ADODB.Recordset
                Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
                CADENA = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
                codmacta = vParamAplic.Raiz_CtaAltaSoc & Format(Socio, CADENA)
    
    '            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
                
                If Not Rs.EOF Then
                    cadValuesAux = "(" & LEtra & ", " & Socio & ", '" & Format(txtCodigo(3).Text, FormatoFecha) & "', "
                    CadValuesAuxConta = "(" & LEtra & ", " & Socio & ", '" & Format(txtCodigo(3).Text, FormatoFecha) & "', "
                    '                    A�adire a la cadena fija esta los valores de textcsb41,txcs
                    CadValuesAuxConta = CadValuesAuxConta & vTextosCSB & ","
                    '-------- Primer Vencimiento
                    i = 1
                    'FECHA VTO
                    FecVenci = CDate(txtCodigo(3).Text)
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                    FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                    '===
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    TipForPago = DevuelveDesdeBDNew(conConta, "sforpa", "tipforpa", "codforpa", CStr(ForPago), "N")
                    If CCur(TipForPago) <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If CInt(DBLet(Rs!mesnogir, "N")) <> 0 Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
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
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "',  '" & textcsb33 & "', " & DBSet(vParamAplic.PorDefecto_Agente, "N")
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', " & DBSet(vSocio.Banco, "N") & ", " & DBSet(vSocio.Sucursal, "N") & ", " & DBSet(vSocio.DigControl, "T") & ", " & DBSet(vSocio.CuentaBan, "T") & ", '" & textcsb33 & "', " & DBSet(vParamAplic.PorDefecto_Agente, "N")
                    End If
                    'departamento y transfer
                    CadValues2 = CadValues2 & "," & DBSet("", "N", "S") & ",NULL"
                    '1 Julio 2009
                    ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                     CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                    
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(vSocio.Iban, "") & MiFormat(vSocio.Banco, "0000") & MiFormat(vSocio.Sucursal, "0000") & MiFormat(vSocio.DigControl, "00") & MiFormat(vSocio.CuentaBan, "0000000000")
                    
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                    
                    Else
                    
                        '[Monica]22/11/2013: tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & "," & DBSet(vSocio.Iban, "T", "S") & ")"
                        Else
                           CadValues2 = CadValues2 & ")"
                        End If
                         
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
                            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                        Else
                            FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                        End If
                        'Comprobar si cliente tiene mes a no girar
                        FecVenci1 = FecVenci
                        If DBLet(Rs!mesnogir, "N") <> "0" Then
                            FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                        End If

                        CadValues = CadValues & ", " & cadValuesAux & i & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                        CadValues2 = CadValues2 & ", " & CadValuesAuxConta & i & ", " & DBSet(codmacta, "T") & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "

                        'IMPORTE Resto de Vendimientos
                        ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                        ImpVenci2 = Round2((TotalFactura2) / rsVenci!numerove, 2)
                        CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                        If vParamAplic.ContabilidadNueva Then
                            CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", '" & textcsb33 & "', " & DBSet(vParamAplic.PorDefecto_Agente, "N") & ", "
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & DBSet(vSocio.Banco, "N") & ", " & DBSet(vSocio.Sucursal, "N") & ", " & DBSet(vSocio.DigControl, "T") & ", " & DBSet(vSocio.CuentaBan, "T") & ", '" & textcsb33 & "', " & DBSet(vParamAplic.PorDefecto_Agente, "N") & ", "
                        End If
                        CadValues2 = CadValues2 & DBSet("", "N", "S") & ",NULL"
                        '1 Julio 2009
                        ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                        CadValues2 = CadValues2 & "," & CadenaDatosFiscales '& ")"
                        
                        If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(vSocio.Iban, "") & MiFormat(vSocio.Banco, "0000") & MiFormat(vSocio.Sucursal, "0000") & MiFormat(vSocio.DigControl, "00") & MiFormat(vSocio.CuentaBan, "0000000000")
                        
                            CadValues2 = CadValues2 & "," & DBSet(vvIban, "T", "S") & ")"
                        
                        Else
                            '[Monica]22/11/2013: tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                               CadValues2 = CadValues2 & "," & DBSet(vSocio.Iban, "T", "S") & ")"
                            Else
                               CadValues2 = CadValues2 & ")"
                            End If
                        End If
                    Next i
                    
                End If
                Rs.Close
            Else
                'totalfac =0 and numerovtos >=1
                b = True
            End If
            
            Set Rs = Nothing
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
                        If vParamAplic.ContabilidadNueva Then
                            'Insertamos en la tabla scobro de la CONTA
                            Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu,text41csb, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, "
                            Sql = Sql & "text33csb,agente,departamento,transfer "
                            'JULIO 2009
                            'nomclien,domclien,pobclien, cpclien,proclien
                            Sql = Sql & ",nomclien,domclien,pobclien, cpclien,proclien, nifclien,codpais, iban)" ')"
                        Else
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
Dim Rs As ADODB.Recordset
    On Error GoTo ErrInsForpa
    InsertarFormaPagoEnConta = False
    
    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
    'forma de pago de la factura. Sino existe insertarla
    
    'vemos si existe en la conta
    If vParamAplic.ContabilidadNueva Then
        cadAux = DevuelveDesdeBDNew(conConta, "formapago", "codforpa", "codforpa", nForPa, "N")
    Else
        cadAux = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", nForPa, "N")
    End If
    'si no existe la forma de pago en conta, insertamos la de aritaxi
    If cadAux = "" Then
        Sql = "select * from sforpa where codforpa = " & DBSet(nForPa, "N")
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            'insertamos e sforpa de la CONTA
            If vParamAplic.ContabilidadNueva Then
                Sql = "INSERT INTO formapago(codforpa,nomforpa,tipforpa,numerove,primerve,restoven)"
            Else
                Sql = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
            End If
            Sql = Sql & " VALUES(" & nForPa & ", " & DBSet(Rs!nomforpa, "T") & ", " & DBSet(Rs!tipforpa, "N")
            If vParamAplic.ContabilidadNueva Then
                Sql = Sql & "," & DBSet(Rs!numerove, "N") & "," & DBSet(Rs!primerve, "N") & "," & DBSet(Rs!restoven, "N") & ")"
            Else
                Sql = Sql & ")"
            End If
            ConnConta.Execute Sql
            InsertarFormaPagoEnConta = True
        Else
            InsertarFormaPagoEnConta = False
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        InsertarFormaPagoEnConta = True
    End If
    
    
    Exit Function
    
ErrInsForpa:
    InsertarFormaPagoEnConta = False
    cadErr = "Insertar forma de pago en Contablilidad: " & vbCrLf & Err.Description
End Function



