VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Varios"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameEstadisticasConsultas 
      Height          =   3855
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdListConsultaPedido 
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
         Left            =   4200
         TabIndex        =   28
         Top             =   3120
         Width           =   1135
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   4260
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   2520
         Width           =   1245
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   1560
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   2520
         Width           =   1245
      End
      Begin VB.TextBox txtArt 
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
         Left            =   1560
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtArtD 
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
         Index           =   3
         Left            =   3000
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   2
         Left            =   5400
         TabIndex        =   29
         Top             =   3120
         Width           =   1135
      End
      Begin VB.TextBox txtArt 
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
         Left            =   1560
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtArtD 
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
         Index           =   2
         Left            =   3000
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Image imgFecha 
         Height          =   255
         Index           =   1
         Left            =   3930
         Top             =   2520
         Width           =   255
      End
      Begin VB.Image imgFecha 
         Height          =   255
         Index           =   0
         Left            =   1200
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   38
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   37
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   2160
         Width           =   660
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   3
         Left            =   1200
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   35
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Estadísticas consultas artículo / cliente"
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
         Index           =   2
         Left            =   360
         TabIndex        =   33
         Top             =   240
         Width           =   6195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   960
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   2
         Left            =   1200
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   31
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame FrameDHArticulo 
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdEliminarArticulos 
         Caption         =   "Buscar"
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
         Left            =   4140
         TabIndex        =   2
         Top             =   2040
         Width           =   1135
      End
      Begin VB.TextBox txtArtD 
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
         Left            =   2880
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtArt 
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
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1560
         Width           =   1545
      End
      Begin VB.TextBox txtArtD 
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
         Left            =   2880
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtArt 
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
         Left            =   1320
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1200
         Width           =   1545
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   1
         Left            =   5460
         TabIndex        =   3
         Top             =   2040
         Width           =   1135
      End
      Begin VB.Label lblElim 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   1
         Left            =   1080
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   0
         Left            =   1080
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   36
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   960
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Eliminar artículos"
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
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame FrameImpresionFacturasDirectas 
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   0
         Left            =   4200
         TabIndex        =   9
         Top             =   1680
         Width           =   1135
      End
      Begin VB.Label lblImpr 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label lblImpr 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Impresión facturas"
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
         Index           =   8
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame FrameFormaEnvio 
      Height          =   3495
      Left            =   120
      TabIndex        =   39
      Top             =   0
      Width           =   6135
      Begin VB.ListBox ListEnvio 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   5655
      End
      Begin VB.CommandButton cmdFormaEnvio 
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
         Left            =   4740
         TabIndex        =   41
         Top             =   3000
         Width           =   1135
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Forma de envio"
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
         Index           =   3
         Left            =   1680
         TabIndex        =   40
         Top             =   360
         Width           =   2835
      End
   End
   Begin VB.Frame FrameListArticulos 
      Height          =   6855
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   6855
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAccionListview 
         Caption         =   "Eliminar"
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
         Left            =   4230
         TabIndex        =   4
         Top             =   6330
         Width           =   1135
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
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
         Left            =   5460
         TabIndex        =   5
         Top             =   6330
         Width           =   1135
      End
      Begin VB.Label lblElim 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   22
         Top             =   6480
         Width           =   2415
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmVarios.frx":0000
         ToolTipText     =   "Quitar al haber"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmVarios.frx":014A
         ToolTipText     =   "Puntear al haber"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Eliminar artículos"
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
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame FrameSelClien 
      Height          =   6975
      Left            =   0
      TabIndex        =   43
      Top             =   -30
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdClientes 
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
         Left            =   4320
         TabIndex        =   47
         Top             =   6450
         Width           =   1135
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
         Index           =   4
         Left            =   5550
         TabIndex        =   46
         Top             =   6450
         Width           =   1135
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5655
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6703
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   720
         Picture         =   "frmVarios.frx":0294
         ToolTipText     =   "Puntear al haber"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   240
         Picture         =   "frmVarios.frx":03DE
         ToolTipText     =   "Quitar al haber"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Seleccionar clientes"
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
         Index           =   4
         Left            =   360
         TabIndex        =   44
         Top             =   240
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frmVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    ' 0.-   Impresion de facturas directas (tipo 4tonda)
    ' 1.-   Eliminar articulos masiva
    ' 2.-   Estadisticas consultas (archivo-facturacion-pedidos-consulta precio/cliente

    ' 3.-   Eleccion del metodo de envio para los albaranes

    ' 4.-   Ver clientes para añadir acciones comerciales

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticulos
Attribute frmA.VB_VarHelpID = -1


Private Cad As String
Private SePuedeCerrar As Boolean   'Puede llevar DoEvents
Private PrimeraVez As Boolean


Private Sub cmdAccionListview_Click()
Dim T1 As Single

    CadenaDesdeOtroForm = ""
    For NumRegElim = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(NumRegElim).Checked Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "1"
    Next
    
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Seleccione algun artículo para eliminar", vbInformation
    Else
        CadenaDesdeOtroForm = Len(CadenaDesdeOtroForm)
        CadenaDesdeOtroForm = "Va a eliminar " & CadenaDesdeOtroForm & " artículo(s).   ¿Continuar?"
        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbNo Then CadenaDesdeOtroForm = ""
    End If
    If CadenaDesdeOtroForm = "" Then Exit Sub
    
    
    
    
    
    'AHora eliminamos
    'Y el log de acciones
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    
    
    '-----------------------------------------------------------------------------
    
    Screen.MousePointer = vbHourglass
    lblElim(1).Caption = ""
    For NumRegElim = ListView1.ListItems.Count To 1 Step -1
        If ListView1.ListItems(NumRegElim).Checked Then
            T1 = Timer
            ListView1.ListItems(NumRegElim).EnsureVisible
            conn.BeginTrans
            If EliminarArticulo(ListView1.ListItems(NumRegElim).Text, lblElim(1)) Then
                LOG.Insertar 7, vUsu, ListView1.ListItems(NumRegElim).Text & " " & ListView1.ListItems(NumRegElim).SubItems(1)
                conn.CommitTrans
                'QUitamos del nodo
                ListView1.ListItems.Remove ListView1.ListItems(NumRegElim).Index
                T1 = 1.5 - (Timer - T1)
                If T1 > 0 Then Espera T1
                
            Else
                'NO se ha podido eliminar
                conn.RollbackTrans
                ListView1.ListItems(NumRegElim).Bold = True
                ListView1.ListItems(NumRegElim).ForeColor = vbRed
                ListView1.ListItems(NumRegElim).Checked = False
            End If
        End If
    Next
    lblElim(1).Caption = ""
    Screen.MousePointer = vbDefault
    Set LOG = Nothing
    If ListView1.ListItems.Count = 0 Then
        SePuedeCerrar = True
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    
    If Opcion = 0 Then
        'Esta haciendo cosas. Preguntar si cerrar
        If Not SePuedeCerrar Then
            If MsgBox("Seguro que desea finalizar el proceso?", vbQuestion + vbYesNo) = vbYes Then SePuedeCerrar = True
            Exit Sub
        End If
        
    End If
    
    SePuedeCerrar = True
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    SePuedeCerrar = True
    Unload Me
End Sub

Private Sub cmdClientes_Click()
        Cad = ""
        For NumRegElim = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(NumRegElim).Checked Then Cad = Cad & ", " & CStr(Val(ListView2.ListItems(NumRegElim).Text))
        Next NumRegElim
        If Cad = "" Then
            MsgBox "Seleccione algun dato", vbExclamation
            Exit Sub
        End If
        CadenaDesdeOtroForm = Mid(Cad, 2) 'le quito la primera coma
        Unload Me
End Sub

Private Sub cmdEliminarArticulos_Click()
Dim Sql As String
Dim It As ListItem

    '
    lblElim(0).Caption = "Cargando datos"
    lblElim(0).Refresh
    
    'Eliminamos los datos de tmpnseries
    conn.Execute "DELETE FROM tmpnseries where codusu = " & vUsu.Codigo
    
    
    'Cargamos tmpnseries con los articulos del desde hasta
    Sql = ""
    If Me.txtArt(0).Text <> "" Then Sql = Sql & " codartic >=" & DBSet(txtArt(0).Text, "T")
    If Me.txtArt(1).Text <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & " codartic <=" & DBSet(txtArt(1).Text, "T")
    End If
    If Sql <> "" Then Sql = " WHERE " & Sql
    Sql = " SELECT " & vUsu.Codigo & ",codartic,0,0 FROM sartic " & Sql
    Sql = "insert into `tmpnseries` (`codusu`,`codartic`,`numlinealb`,`numlinea`) " & Sql
    conn.Execute Sql
    
    
    Set miRsAux = New ADODB.Recordset
    
    'Eliminamos de tmpnseries los articulos que seguro estan en
    ' alba, fact....
    EliminandoArticulos_Paso1
    
    
    'Ya tengo los articulos. Vere cuales puedo borrar
    lblElim(0).Caption = "Obteniendo registros"
    lblElim(0).Refresh
    
    Sql = "Select tmpnseries.codartic,nomartic from tmpnseries,sartic where codusu = " & vUsu.Codigo
    Sql = Sql & " AND tmpnseries.codartic=sartic.codartic"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        lblElim(0).Caption = ""
        MsgBox "No existen registros", vbExclamation
        miRsAux.Close
        Set miRsAux = Nothing
        Exit Sub
    End If
    
    'Ajustamos los tamaños para cargar el LISTVIEW
    CargaColumnas
    NumRegElim = (Screen.Width - FrameListArticulos.Width - 420) \ 2
    Me.Left = NumRegElim
    NumRegElim = (Screen.Height - FrameListArticulos.Height - 360) \ 2
    Me.Top = NumRegElim
    Me.FrameDHArticulo.visible = False
    PonerFrameVisible Me.FrameListArticulos
    Me.lblTitulo(1).Caption = "Eliminar artículos"
    DoEvents
    
    'Vamos cargando los registros
    While Not miRsAux.EOF
        Set It = ListView1.ListItems.Add()
        It.Text = miRsAux!codArtic
        It.SubItems(1) = miRsAux!NomArtic
        It.Checked = True
        'Sig
        miRsAux.MoveNext
    Wend
End Sub

Private Sub cmdFormaEnvio_Click()
Dim I As Integer

    If ListEnvio.ListIndex < 0 Then Exit Sub
    Cad = ListEnvio.List(ListEnvio.ListIndex)
    I = InStrRev(Cad, "(")
    Cad = Trim(Mid(Cad, I + 1))
    I = InStrRev(Cad, ")")
    Cad = Mid(Cad, 1, I - 1) 'quitamos el ultmio parentesis
    CadenaDesdeOtroForm = Cad
    
    I = InStrRev(ListEnvio.List(ListEnvio.ListIndex), "(")
    Cad = Mid(ListEnvio.List(ListEnvio.ListIndex), 1, I - 1)  'quito el precio kilo
    
    I = Val(Mid(Cad, 1, 10))
    
    Cad = Trim(Mid(Cad, 11))
    
    CadenaDesdeOtroForm = I & "|" & Cad & "|" & CadenaDesdeOtroForm & "|"
    
    'Desde kilo
    Cad = ListEnvio.List(ListEnvio.ListIndex)
    I = InStrRev(ListEnvio.List(ListEnvio.ListIndex), "Desde :")
    Cad = Mid(Cad, I + 7)
    Cad = Trim(Mid(Cad, 1, Len(Cad) - 2)) 'Le kito kg
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Cad & "|"
    SePuedeCerrar = True
    Unload Me
End Sub

Private Sub cmdListConsultaPedido_Click()
Dim Aux As String


    Cad = ""
    Aux = CadenaDesdeHastaBD(txtArt(2).Text, txtArt(3).Text, "codartic", "T")
    If Aux <> "" Then Cad = Aux
    
    'La fecha
    Aux = CadenaDesdeHastaBD(txtFecha(0).Text, txtFecha(1).Text, "DiaHora", "FH")
    If Aux <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & Aux
    End If
        
    If Not HayRegParaInforme("sconsulta", Cad) Then Exit Sub
    
    
    'Para el informe
    Cad = ""
    Aux = CadenaDesdeHasta(txtArt(2).Text, txtArt(3).Text, "{sconsulta.codartic}", "T")
    If Aux <> "" Then Cad = Aux
    
    'La fecha
    Aux = CadenaDesdeHasta(txtFecha(0).Text, txtFecha(1).Text, "{sconsulta.DiaHora}", "FH")
    If Aux <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & Aux
    End If
    
    
    
    
    With frmImprimir
        .FormulaSeleccion = Cad
        .OtrosParametros = ""
        .NumeroParametros = 0

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2002
        .Titulo = "Estadistica consultas pedidos"
        .NombreRPT = "rFacConsuPrecioArt.rpt"
        .ConSubInforme = False
        .Show vbModal
    End With
    
    
    
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        Select Case Opcion
        Case 0
            'Se pone a imprimir las facturas
            HacerImpresionFacturas
            
        Case 3
            ListEnvio.SetFocus
        Case 4
            CargaClientes
        End Select
    End If
End Sub

Private Sub CargarIconos()
Dim I As Image


    For Each I In Me.imgArticulo
         I.Picture = frmPpal.imgListComun.ListImages(19).Picture
         I.ToolTipText = "Articulo"
    Next
    For Each I In Me.imgFecha
         I.Picture = frmPpal.imgListComun.ListImages(23).Picture
         I.ToolTipText = "fecha"
    Next
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    PrimeraVez = True
    
    limpiar Me
    CargarIconos
    FrameListArticulos.visible = False
    FrameDHArticulo.visible = False
    Me.FrameImpresionFacturasDirectas.visible = False
    Me.FrameEstadisticasConsultas.visible = False
    FrameFormaEnvio.visible = False
    FrameSelClien.visible = False
    SePuedeCerrar = True
    Select Case Opcion
    Case 0
        PonerFrameVisible Me.FrameImpresionFacturasDirectas
    Case 1
        PonerFrameVisible FrameDHArticulo
    Case 2
        PonerFrameVisible Me.FrameEstadisticasConsultas
    Case 3
        'Metodo de envio
        'En cadena deotro form llevo las lineas para seelccionar una de ellas
        SePuedeCerrar = False
        PonerFrameVisible FrameFormaEnvio
        CargaFormasEnvioPosibles
    
    Case 4
        PonerFrameVisible Me.FrameSelClien
        
        
    End Select
    
    If Opcion <> 3 Then cmdCancelar(Opcion).Cancel = True
    
End Sub



Private Sub PonerFrameVisible(Fr As Frame)
    Fr.visible = True
    Fr.Top = 0
    Fr.Left = 120
    Me.Height = Fr.Height + 480
    Me.Width = Fr.Width + 320
End Sub




Private Sub HacerImpresionFacturas()
Dim I As Integer
Dim Fin As Boolean
    SePuedeCerrar = False
    
    Me.lblImpr(0).Caption = "Leyendo datos"
    lblImpr(0).Refresh
    Espera 0.25
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open "Select count(*) from scafac WHERE " & CadenaDesdeOtroForm, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    If NumRegElim = 0 Then Exit Sub
    
    CadenaDesdeOtroForm = "Select codtipom, numfactu, fecfactu, nomclien from scafac where " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " ORDER BY fecfactu,numfactu"
    
    miRsAux.Open CadenaDesdeOtroForm, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    Fin = False
    While Not Fin
        I = I + 1
        Me.lblImpr(1).Caption = "Fac. " & Format(miRsAux!NumFactu, "00000") & " de " & Format(miRsAux!FecFactu, "dd/mm/yyyy") & "     " & Mid(miRsAux!nomclien, 1, 20)
        lblImpr(1).Refresh
        Me.lblImpr(0).Caption = "Registro: " & I & "   de   " & NumRegElim
        lblImpr(0).Refresh
    
        'IMprimimos la factura
        ReImprimirDirectoFact " scafac.codtipom ='" & miRsAux!codtipom & "' AND scafac.numfactu = " & miRsAux!NumFactu
    
        DoEvents
        If SePuedeCerrar Then
            Fin = True  'Han pulsado cancelar
        Else
            'Siguiente
            miRsAux.MoveNext
            Fin = miRsAux.EOF
        End If
        If I Mod 50 = 25 Then Me.Refresh
            
        
    Wend
    If miRsAux.EOF Then
        'Significa que ha acabado toda la impresion. Con lo cual
        'pongo CadenaDesdeOtroForm="" para que el form de reimpresion lo cierre
        CadenaDesdeOtroForm = ""
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    SePuedeCerrar = True
    Unload Me  'Y cierro
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not SePuedeCerrar Then Cancel = 1
    
    
End Sub


Private Sub imgSel_Click(Index As Integer)

End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    Cad = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgArticulo_Click(Index As Integer)
        Cad = ""
        Set frmA = New frmAlmArticulos
        frmA.DeConsulta = True
        frmA.DatosADevolverBusqueda2 = "@1@"
        frmA.Show vbModal
        Set frmA = Nothing
        If Cad <> "" Then
            Me.txtArt(Index).Text = RecuperaValor(Cad, 1)
            Me.txtArtD(Index).Text = RecuperaValor(Cad, 2)
        End If
End Sub

Private Sub imgCheck_Click(Index As Integer)
    If Index < 2 Then
        'LISTVIEW 1
        For NumRegElim = 1 To ListView1.ListItems.Count
            ListView1.ListItems(NumRegElim).Checked = Index = 1
        Next NumRegElim
        
    Else
        For NumRegElim = 1 To ListView2.ListItems.Count
            ListView2.ListItems(NumRegElim).Checked = Index = 3
        Next NumRegElim
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Cad = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    If Cad <> "" Then txtFecha(Index).Text = Cad
End Sub

Private Sub txtArt_GotFocus(Index As Integer)
 PonerFoco txtArt(Index)
End Sub

Private Sub txtArt_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtArt_LostFocus(Index As Integer)
Dim C As String

    txtArt(Index).Text = Trim(txtArt(Index).Text)
    If txtArt(Index).Text = "" Then
        C = ""
    Else
        C = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", txtArt(Index).Text, "T")
        If C = "" Then
            'El articulo no existe. SI fuera obligado ponerlo es aqui donde habria que poner el ocdigo
            
        End If
    End If
    txtArtD(Index).Text = C
End Sub



Private Sub EliminandoArticulos_Paso1()
Dim C As String
Dim Sql As String
Dim Aux As String
Dim NT As Integer
Dim J As Byte

    If Me.txtArt(0).Text <> "" Then Sql = Sql & " codartic >=" & DBSet(txtArt(0).Text, "T")
    If Me.txtArt(1).Text <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & " codartic <=" & DBSet(txtArt(1).Text, "T")
    End If
    If Sql <> "" Then Sql = " WHERE " & Sql
    
     
    'El stock
    lblElim(0).Caption = "Almacenes"
    lblElim(0).Refresh
    C = "select codartic,sum(canstock) from salmac " & Sql & " group by codartic having sum(canstock) <> 0"
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
         conn.Execute "DELETE FROM tmpnseries where codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux.Fields(0), "T")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    
    For J = 0 To 2
        DevuelveTablasBorre J, C, Aux, NT
        For NumRegElim = 1 To NT
            
            lblElim(0).Caption = RecuperaValor(Aux, CInt(NumRegElim)) & "   -"
            If J = 0 Then
                lblElim(0).Caption = lblElim(0).Caption & "Clientes"
            ElseIf J = 1 Then
                lblElim(0).Caption = lblElim(0).Caption & "Prove"
            Else
                lblElim(0).Caption = lblElim(0).Caption & "Varios"
            End If
            lblElim(0).Refresh
            
            
            miRsAux.Open "Select codartic from " & RecuperaValor(C, CInt(NumRegElim)) & Sql & " GROUP BY codartic", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                 conn.Execute "DELETE FROM tmpnseries where codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux.Fields(0), "T")
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Me.Refresh
        Next NumRegElim
    Next J
    
End Sub


Private Sub CargaColumnas()
Dim clmX As ColumnHeader

    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    
    Select Case Opcion

    Case 1
        Me.ListView1.Checkboxes = True
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Código"
        clmX.Width = 2200
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Descripción"
        clmX.Width = 3600
        
    End Select
    Me.FrameListArticulos.ZOrder 1  'QUe lo traiga al frente
End Sub


 

Private Sub txtFecha_GotFocus(Index As Integer)
    PonerFoco txtFecha(Index)
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        Cad = txtFecha(Index).Text
        If Not EsFechaOK(Cad) Then
            MsgBox "Fecha incorrecta: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        Else
            txtFecha(Index).Text = Cad
        End If
    End If
End Sub


'En cadenadesdeotroform llevo las lformas posibles. Se trata de ir poniendolas en el list
Private Sub CargaFormasEnvioPosibles()
Dim I As Integer
    
    
    While CadenaDesdeOtroForm <> ""
        I = InStr(1, CadenaDesdeOtroForm, "|")
        If I = 0 Then
            CadenaDesdeOtroForm = ""
            Cad = ""
        Else
            Cad = Mid(CadenaDesdeOtroForm, 1, I)
            CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, I + 1)
            
            Cad = Replace(Cad, "<", "|")
                        
        End If
        If Cad <> "" Then
            
            I = RecuperaValor(Cad, 1)
            
            Cad = Format(I, "0000") & "      " & RecuperaValor(Cad, 2) & "    (" & RecuperaValor(Cad, 3) & ")    Desde :" & RecuperaValor(Cad, 4) & " Kg"
            ListEnvio.AddItem Cad
            
        End If
    Wend
    If ListEnvio.ListCount > 0 Then ListEnvio.Selected(0) = True
End Sub



Private Sub CargaClientes()
Dim It
    On Error GoTo ECargaClientes
    Set miRsAux = New ADODB.Recordset
    
    
    
    miRsAux.Open "select sclien.codclien,nomclien from sclien " & CadenaDesdeOtroForm, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = ""
    While Not miRsAux.EOF
        Set It = ListView2.ListItems.Add()
        It.Text = Format(miRsAux!CodClien, "0000")
        It.SubItems(1) = miRsAux!nomclien
        It.Checked = True
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Caption = ListView2.ListItems.Count
ECargaClientes:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
    
End Sub
