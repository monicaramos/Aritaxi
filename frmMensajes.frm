VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajes"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14160
   Icon            =   "frmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFrasPteContabilizar 
      Height          =   5790
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   13660
      Begin VB.CommandButton cmdCerrarFras 
         Caption         =   "Continuar"
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
         Index           =   5
         Left            =   12060
         TabIndex        =   74
         Top             =   5280
         Width           =   1215
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
         Index           =   0
         ItemData        =   "frmMensajes.frx":000C
         Left            =   240
         List            =   "frmMensajes.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Tag             =   "Tipo de cliente|N|N|0|2|ssocio|tipsocio|||"
         Top             =   240
         Width           =   3180
      End
      Begin MSComctlLib.ListView ListView22 
         Height          =   4545
         Left            =   240
         TabIndex        =   75
         Top             =   630
         Width           =   13085
         _ExtentX        =   23072
         _ExtentY        =   8017
         View            =   3
         LabelWrap       =   -1  'True
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Facturas Pendientes de Contabilizar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   375
         Left            =   4920
         TabIndex        =   76
         Top             =   300
         Width           =   8355
      End
   End
   Begin VB.Frame FrameNSeries 
      Height          =   5000
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdSelTodos 
         Caption         =   "&Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   22
         Top             =   4320
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmdDeselTodos 
         Caption         =   "&Ninguno"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   21
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
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
         Left            =   5040
         TabIndex        =   20
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarNSeries 
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
         Left            =   3960
         TabIndex        =   12
         Top             =   4320
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas en el sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   5295
      End
   End
   Begin VB.Frame FrameEmail 
      Height          =   6975
      Left            =   3600
      TabIndex        =   43
      Top             =   0
      Width           =   8055
      Begin VB.TextBox txmemail 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "Text2"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton cmdEmail 
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
         Height          =   405
         Left            =   6720
         TabIndex        =   53
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txmemail 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3585
         Index           =   3
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Text            =   "frmMensajes.frx":002C
         Top             =   2760
         Width           =   7335
      End
      Begin VB.TextBox txmemail 
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
         Height          =   315
         Index           =   2
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "Text2"
         Top             =   2160
         Width           =   4815
      End
      Begin VB.TextBox txmemail 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Text2"
         Top             =   1560
         Width           =   7335
      End
      Begin VB.TextBox txmemail 
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   960
         Width           =   7335
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   5280
         TabIndex        =   55
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Mensaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   52
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Adjuntos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   50
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Asunto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   48
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   47
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Email CRM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   15
         Left            =   960
         TabIndex        =   46
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrameCobrosPtes 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton cmdCancelarCobros 
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
         Left            =   7080
         TabIndex        =   25
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtParam 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "frmMensajes.frx":0032
         Top             =   120
         Width           =   6615
      End
      Begin VB.CommandButton cmdAceptarCobros 
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
         Left            =   5400
         TabIndex        =   1
         Top             =   4440
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "¿Desea continuar?"
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
         Left            =   480
         TabIndex        =   26
         Top             =   4440
         Width           =   7215
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame FrameServicios 
      Height          =   6615
      Left            =   0
      TabIndex        =   56
      Top             =   -60
      Width           =   10395
      Begin MSComctlLib.ListView ListView5 
         Height          =   4665
         Left            =   150
         TabIndex        =   59
         Top             =   750
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Precio"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Familia"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.CommandButton cmdCanServ 
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
         Index           =   0
         Left            =   8910
         TabIndex        =   58
         Top             =   5700
         Width           =   1215
      End
      Begin VB.CommandButton cmdAcepServ 
         Caption         =   "Continuar"
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
         Left            =   7470
         TabIndex        =   57
         Top             =   5700
         Width           =   1215
      End
      Begin VB.Image imgSelServ 
         Height          =   240
         Index           =   5
         Left            =   180
         Picture         =   "frmMensajes.frx":0038
         Top             =   5550
         Width           =   240
      End
      Begin VB.Image imgSelServ 
         Height          =   240
         Index           =   4
         Left            =   540
         Picture         =   "frmMensajes.frx":0182
         Top             =   5550
         Width           =   240
      End
   End
   Begin VB.Frame FrameEtiqEstant 
      Height          =   7455
      Left            =   0
      TabIndex        =   31
      Top             =   -120
      Width           =   8535
      Begin VB.CommandButton cmdEtiqEstan 
         Caption         =   "Imprimir"
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
         Left            =   5520
         TabIndex        =   34
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdEtiqEstan 
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
         Index           =   0
         Left            =   6960
         TabIndex        =   33
         Top             =   6960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   6495
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   11456
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Precio"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Familia"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmMensajes.frx":02CC
         Top             =   6960
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmMensajes.frx":0416
         Top             =   6960
         Width           =   240
      End
   End
   Begin VB.Frame FrameCorreccionPrecios 
      Height          =   6375
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   12975
      Begin VB.ComboBox cmbActualizarTar 
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
         ItemData        =   "frmMensajes.frx":0560
         Left            =   7800
         List            =   "frmMensajes.frx":0562
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   5960
         Width           =   2175
      End
      Begin VB.CommandButton cmdCorrecotrPrecios 
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
         Left            =   11760
         TabIndex        =   38
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdCorrecotrPrecios 
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
         Index           =   0
         Left            =   10560
         TabIndex        =   37
         Top             =   5880
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   5175
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   12660
         _ExtentX        =   22331
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Denominación"
            Object.Width           =   5715
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "U.P.Compra"
            Object.Width           =   2188
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "% M"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "PVP"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "%T"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "P.Tarifa"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "PVP Correcto"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Tarifa correc."
            Object.Width           =   2011
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Actualizar:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   42
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   11760
         Picture         =   "frmMensajes.frx":0564
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   12360
         Picture         =   "frmMensajes.frx":06AE
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblIndicadorCorregir 
         Caption         =   "Label3"
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
         Left            =   240
         TabIndex        =   40
         Top             =   5880
         Width           =   5055
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Frame FrameErrores 
      Height          =   5535
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
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
         Left            =   7080
         TabIndex        =   29
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   28
         Text            =   "frmMensajes.frx":07F8
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame FrameComponentes 
      Height          =   3975
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdAceptarComp 
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
         Left            =   3240
         TabIndex        =   19
         Top             =   3240
         Width           =   975
      End
      Begin VB.Frame FrameComponentes2 
         Caption         =   "Mostrar Equipos del :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2175
         Left            =   960
         TabIndex        =   15
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton OptCompXClien 
            Caption         =   "Cliente"
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
            Left            =   360
            TabIndex        =   18
            Top             =   1440
            Width           =   2535
         End
         Begin VB.OptionButton OptCompXDpto 
            Caption         =   "Departamento"
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
            Left            =   360
            TabIndex        =   17
            Top             =   960
            Width           =   2535
         End
         Begin VB.OptionButton OptCompXMant 
            Caption         =   "Mantenimiento"
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
            Left            =   360
            TabIndex        =   16
            Top             =   480
            Width           =   2055
         End
      End
   End
   Begin VB.Frame FrameAcercaDe 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6375
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C/ Uruguay 11, Despacho 101"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3000
         TabIndex        =   10
         Top             =   2640
         Width           =   2610
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "46007 - VALENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3000
         TabIndex        =   9
         Top             =   2925
         Width           =   2535
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tfno:  902 88 88 78  -  96 380 55 79"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   2685
         TabIndex        =   8
         Top             =   3195
         Width           =   3165
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax: 963 42 09 38"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3480
         TabIndex        =   7
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   540
         Left            =   240
         Top             =   2640
         Width           =   2160
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
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
         Left            =   -120
         TabIndex        =   6
         Top             =   1260
         Width           =   4155
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   81.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1725
         Left            =   4260
         TabIndex        =   5
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ARIGES"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   1080
         TabIndex        =   4
         Top             =   300
         Width           =   3495
      End
   End
   Begin VB.Frame FrameBloqueoEmpresas 
      Height          =   7455
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CommandButton cmdBloqEmpre 
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
         Index           =   0
         Left            =   8400
         TabIndex        =   67
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton cmdBloqEmpre 
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
         Index           =   1
         Left            =   9840
         TabIndex        =   65
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   ">"
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   64
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   "<"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   63
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   ">>"
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   62
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   "<<"
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   61
         Top             =   3480
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   5775
         Index           =   0
         Left            =   225
         TabIndex        =   66
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empresa"
            Object.Width           =   5644
         EndProperty
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   5775
         Index           =   1
         Left            =   6240
         TabIndex        =   68
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empresa"
            Object.Width           =   5644
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Bloqueo de empresas por usuario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   495
         Index           =   2
         Left            =   2880
         TabIndex        =   71
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label41 
         Caption         =   "Permitidas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   70
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Bloqueadas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   10050
         TabIndex        =   69
         Top             =   600
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================== VBLES PUBLICAS ================================

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionMensaje As Byte
'======================================
'==== FACTURACION =====================
' 1 .- Mensaje de Cobros Pendientes
' 2 .- Mensaje de No hay suficiente Stock para pasar de Pedido a Albaran
' 3 .- Mensaje Acerca de...
' 4 .- Listado de los Nº de Serie de un Articulo
' 5 .- Seleccionar tipo de Componente a Mostrar en Mant. de Nº de Series
' 6 .- Mostrar Prefacturacion de Albaranes
' 7 .- Mostrar Prefacturacion Mantenimientos
' 8 .- Mostrar lista clientes para seleccionar los que queremos imprimir (Etiquetas)
' 9 .- Mostrar lista Proveedores para seleccionar los que queremos imprimir (Etiquetas)
'10 .- Mostrar lista de Errores de las facturas NO contabilizadas
'11 .- Mostrar lista lineas de factura a Rectificar para seleccionar las q queremos traer al Albaran de FAct. Rectificativa
'12 .- Mostrar Albaranes del Rango que no se van a Facturar. (Facturar Albaranes Venta)

'13 .- Mostrar Errores
'14 .- Mostrar Empresas existentes en el sistema



'15 .- Mostrar lista de articulos para imprimir etiquetas estanteria
'16 .- Lista de articulos para corregir importes
'17 .- Etiquetas clientes. LO MISMO QUE EL 8 pero hecho por david
'18 .- Mantenimientos. paso ejercicio siguiente a actual

'20 .- IGual que el 16. Pero los importes son de los articulos que tienen componentes

'21 .- Ver un mensaje enlazado desde el outlook para el CRM

'22 .- Servicios de un cliente para seleccionar para la facturacion

'23 .- Mostrar lista Socios para seleccionar los que queremos imprimir (Etiquetas)

'24 .- Mostrar lista de situaciones para seleccionar las que queremos imprimir (Etiquetas)

'25 .- Mostrar los socios que estan marcados para liquidarlos como contado

'30 .- Bloqueo de empresas por usuarios.

'31 .- facturas no contabilizadas

Public cadWHERE As String 'Cadena para pasarle la WHERE de la SELECT de los cobros pendientes o de Pedido(para comp. stock)
                          'o CodArtic para seleccionar los Nº Series
                          'para cargar el ListView
                          
Public cadWHERE2 As String

Public vCampos As String 'Articulo y cantidad Empipados para Nº de Series
                         'Tambien para pasar el nombre de la tabla de lineas (sliped, slirep,...)
                         'Dependiendo desde donde llamemos, de Pedidos o Reparaciones
Public CADENA As String

'====================== VBLES LOCALES ================================

Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form
Dim PrimeraVez As Boolean

'Para los Nº de Serie
Dim TotalArray As Integer
Dim codArtic() As String
Dim Cantidad() As Integer


Public Parametros As String

Dim I As Integer
Dim Sql As String
Dim Rs As Recordset
Dim ItmX As ListItem
Dim Errores As String
Dim NE As Integer
Dim Ok As Integer
Dim vAnt As Integer

Private Sub cmdAcepServ_Click(Index As Integer)
Dim cad As String

    cad = ""
    For TotalArray = 1 To ListView5.ListItems.Count
        If ListView5.ListItems(TotalArray).Checked Then
            cad = cad & "(" & DBSet(ListView5.ListItems(TotalArray).Text, "F") & "," & DBSet(ListView5.ListItems(TotalArray).SubItems(1), "H") & "," & DBSet(ListView5.ListItems(TotalArray).SubItems(2), "N") & "),"
        End If
    Next TotalArray
    If Len(cad) > 0 Then cad = Mid(cad, 1, Len(cad) - 1) ' quitamos la ultima coma
    RaiseEvent DatoSeleccionado(cad)
    PulsadoSalir = True
    Unload Me
    
End Sub

Private Sub cmdAceptarCobros_Click()
    If OpcionMensaje = 12 Then vCampos = "1"
    Unload Me
End Sub


Private Sub cmdAceptarComp_Click()
'Boton Aceptar de Componentes del Mant. de Nº de Series en Reparaciones
Dim H As Integer, W As Integer

    ponerFrameComponentesVisible False, H, W
    PonerFrameCobrosPtesVisible True, H, W
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Me.OptCompXMant.Value Then
        'Mostrar Resumen de los Nº de Serie del Mantenimiento
        Me.Caption = "Equipos del Mantenimiento"
        CargarListaComponentes (1)
    ElseIf Me.OptCompXDpto.Value Then
        'Mostrar Resumen de los Nº de Serie del Departamento
        Me.Caption = "Equipos del Departamento"
        CargarListaComponentes (2)
    ElseIf Me.OptCompXClien.Value Then
        'Mostrar Resumen de los Nº de Serie del Cliente
        Me.Caption = "Equipos del Cliente"
        CargarListaComponentes (3)
    End If
    PonerFocoBtn Me.cmdAceptarCobros
End Sub


Private Sub cmdAceptarNSeries_Click()
Dim I As Integer, J As Byte
Dim Seleccionados As Integer
Dim cad As String, Sql As String
Dim Articulo As String
Dim Rs As ADODB.Recordset
Dim C1 As String * 10, C2 As String * 10, c3 As String * 10


    If OpcionMensaje = 4 Then
        'Comprobar que se han seleccionado el nº correcto de  Nº de Serie para cada Articulo
        Seleccionados = 0
        Articulo = ""
      
        'Si se ha seleccionado la cantidad correcta de Nº de series, empiparlos y
        'devolverlos al form de Albaranes(facturacion)
        cad = ""
        For J = 0 To TotalArray
            Articulo = codArtic(J)
            cad = cad & Articulo & "|"
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    If Articulo = ListView2.ListItems(I).ListSubItems(1).Text Then
                        If Seleccionados < Abs(Cantidad(J)) Then
                            Seleccionados = Seleccionados + 1
                            cad = cad & ListView2.ListItems(I).Text & "|"
                        End If
                   'cad = cad & Data1.Recordset.Fields(1) & "|"
                    End If
                End If
            Next I
            If Seleccionados < Abs(Cantidad(J)) Then
                'Comprobar que si tiene Nºs de serie de ese articulos cargados seleccione los
                'que corresponden
                Sql = "SELECT count(sserie.numserie)"
                Sql = Sql & " FROM sserie " 'INNER JOIN sartic ON sserie.codartic=sartic.codartic "
                Sql = Sql & " WHERE sserie.codartic=" & DBSet(Articulo, "T")
                Sql = Sql & " AND (isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='') "
                Sql = Sql & " ORDER BY sserie.codartic, numserie "
                Set Rs = New ADODB.Recordset
                Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Rs.Fields(0).Value >= Abs(Cantidad(J)) - Seleccionados Then
                    MsgBox "Debe seleccionar " & Cantidad(J) & " Nº Series para el articulo " & codArtic(J), vbExclamation
                    Exit Sub
                Else
                    'No hay Nº Serie y Pedirlos
                End If
                Rs.Close
                Set Rs = Nothing
            
            End If
            cad = cad & "·"
            Seleccionados = 0
        Next J
      
    ElseIf OpcionMensaje = 8 Or OpcionMensaje = 9 Or OpcionMensaje = 17 Or OpcionMensaje = 23 Or OpcionMensaje = 24 Then
        'concatenar todos los clientes seleccionados para imprimir etiquetas
        If OpcionMensaje = 17 Then
            
            '----------------------------------------------------------------
            cad = "insert into tmpnlotes (codusu,numalbar,fechaalb,codalmac,codprove) values ("
            ' ---- [30/10/2009] (LAURA) : agrupar por cliente y departamento
'            cad = cad & vUsu.Codigo & ",1,'2005-04-12',1,"
            cad = cad & vUsu.Codigo & ",1,'2005-04-12',"
            
            
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    ' ---- [30/10/2009] (LAURA) : agrupar por cliente y departamento
'                    conn.Execute cad & (ListView2.ListItems(I).Text) & ")"
                    conn.Execute cad & DBSet(ListView2.ListItems(I).ListSubItems(3).Text, "N", "S") & "," & (ListView2.ListItems(I).Text) & ")"
                    
                    NumRegElim = NumRegElim + 1
                End If
            Next I
            
            
            '----------------------------------------------------------------
            
        Else
            cad = ""
            NumRegElim = 0
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    NumRegElim = NumRegElim + 1
                    cad = cad & Val(ListView2.ListItems(I).Text) & ","
                     'cad = cad & Data1.Recordset.Fields(1) & "|"
                End If
            Next I
            If NumRegElim > 1000 Then
                MsgBox "Maximo número de etiquetas: 1000 (" & NumRegElim & ")", vbExclamation
                NumRegElim = 0
                cad = ""
                Exit Sub
            End If
            NumRegElim = 0
            If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
        End If
    ElseIf OpcionMensaje = 11 Or OpcionMensaje = 111 Then
    'Lineas Factura a rectificar
        'cad = "(" & cadWHERE & ")"
        cad = ""
        C1 = ""
        C2 = ""
        c3 = ""
        Sql = ""
        For I = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(I).Checked Then
                If Sql = "" Then
                    C1 = DBSet(ListView2.ListItems(I), "T", "N")
                    C2 = ListView2.ListItems(I).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                    cad = "(codtipoa=" & Trim(C1) & " and numalbar=" & Val(C2) & " and numlinea IN (" & ListView2.ListItems(I).ListSubItems(2)

                Else
                    If Trim(DBSet(ListView2.ListItems(I), "T", "N")) = Trim(C1) And Trim(ListView2.ListItems(I).ListSubItems(1)) = Trim(C2) Then
                    'es el mismo albaran y concatenamos lineas
                        cad = "," & ListView2.ListItems(I).ListSubItems(2)

                    Else
                        If cad <> "" Then Sql = Sql & ")) "
                        C1 = DBSet(ListView2.ListItems(I), "T", "N")
                        C2 = ListView2.ListItems(I).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                        cad = " or (codtipoa=" & Trim(C1) & " and numalbar=" & Val(C2) & " and numlinea IN (" & ListView2.ListItems(I).ListSubItems(2)
                        
'                       cad=cad &
                    End If
                End If
                Sql = Sql & cad
'                If cad <> "" Then cad = cad & " OR "
'                cad = cad & "(codtipoa=" & DBSet(ListView2.ListItems(i), "T", "N") & " and numalbar=" & Val(ListView2.ListItems(i).ListSubItems(1)) & " and numlinea=" & ListView2.ListItems(i).ListSubItems(2) & ")"
            Else
'                cad = ""
            End If
        Next I
        If cad <> "" Then
            Sql = Sql & "))"
            cad = "(" & cadWHERE & ") AND (" & Sql & ")"
        End If
'        If cad <> "" Then cad = "(" & cadWHERE & ") AND (" & cad & ")"
    ElseIf OpcionMensaje = 14 Then
        cad = RegresarCargaEmpresas
    ElseIf OpcionMensaje = 25 Then
            cad = ""
            NumRegElim = 0
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    NumRegElim = NumRegElim + 1
                    cad = cad & Val(ListView2.ListItems(I).Text) & ","
                     'cad = cad & Data1.Recordset.Fields(1) & "|"
                End If
            Next I
            If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
    End If
    
    
    
     'Actualizar la tabla sseries asignando los valores correspondientes a los
      'campos: codclien, coddirec, tieneman, codtipom, numalbar, fechavta, numline1
      'y Salir (Volver a Mto Albaranes Clientes (Facturacion)
      PulsadoSalir = True
      'RaiseEvent CargarNumSeries
      RaiseEvent DatoSeleccionado(cad)
      Unload Me
End Sub


Private Sub cmdBlEmp_Click(Index As Integer)

    Select Case Index
    Case 0, 1
        'Index Me dira que listview
        For Ok = ListView6(Index).ListItems.Count To 1 Step -1
            If ListView6(Index).ListItems(Ok).Selected Then
                I = ListView6(Index).ListItems(Ok).Index
                PasarUnaEmpresaBloqueada Index = 0, I
            End If
        Next Ok
    Case Else
        If Index = 2 Then
            Ok = 0
        Else
            Ok = 1
        End If
        For NumRegElim = ListView6(Ok).ListItems.Count To 1 Step -1
            PasarUnaEmpresaBloqueada Ok = 0, ListView6(Ok).ListItems(NumRegElim).Index
        Next NumRegElim
        Ok = 0
    End Select
End Sub

Private Sub PasarUnaEmpresaBloqueada(ABLoquedas As Boolean, indice As Integer)
Dim Origen As Integer
Dim Destino As Integer
Dim It
    If ABLoquedas Then
        Origen = 0
        Destino = 1
        NE = 2
    Else
        Origen = 1
        Destino = 0
        NE = 1 'icono
    End If
    
    Sql = ListView6(Origen).ListItems(indice).Key
    Set It = ListView6(Destino).ListItems.Add(, Sql)
    It.SmallIcon = NE
    It.Text = ListView6(Origen).ListItems(indice).Text
    It.SubItems(1) = ListView6(Origen).ListItems(indice).SubItems(1)

    'Borramos en origen
    ListView6(Origen).ListItems.Remove indice
End Sub

Private Sub cmdBloqEmpre_Click(Index As Integer)
    If Index = 0 Then
        Sql = "DELETE FROM usuarios.usuarioempresasaritaxi WHERE codusu =" & Parametros
        conn.Execute Sql
        Sql = ""
        For I = 1 To ListView6(1).ListItems.Count
            Sql = Sql & ", (" & Parametros & "," & Val(Mid(ListView6(1).ListItems(I).Key, 2)) & ")"
        Next I
        If Sql <> "" Then
            'Quitmos la primera coma
            Sql = Mid(Sql, 2)
            Sql = "INSERT INTO usuarios.usuarioempresasaritaxi(codusu,codempre) VALUES " & Sql
            If Not EjecutaSQL(Sql) Then MsgBox "Se han producido errores insertando datos", vbExclamation
        End If
    End If
    Unload Me
End Sub

Private Function EjecutaSQL(ByRef Sql As String) As Boolean
    EjecutaSQL = False
    On Error Resume Next
    conn.Execute Sql
    If Err.Number <> 0 Then
        Err.Clear
    Else
        EjecutaSQL = True
    End If
End Function


Private Sub cmdCancelar_Click()
    If OpcionMensaje = 4 Then
        MsgBox "Debe introducir los nº de serie necesarios para el Albaran.", vbInformation
        Exit Sub
    End If
    
    If OpcionMensaje = 25 Then
        PulsadoSalir = True
        RaiseEvent DatoSeleccionado("Cancelado")
        Unload Me
        Exit Sub
    End If
    
    PulsadoSalir = True
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdCancelarCobros_Click()
    vCampos = "0"
    Unload Me
End Sub

Private Sub cmdCanServ_Click(Index As Integer)
    RaiseEvent DatoSeleccionado("Salir")
    PulsadoSalir = True
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdCorrecotrPrecios_Click(Index As Integer)
    
    If Index = 0 Then
        
        If Not ActualizarPrecios Then Exit Sub
        
    End If
    Unload Me
End Sub

Private Function ActualizarPrecios() As Boolean
Dim Sql As String
    
    
    
        
        ActualizarPrecios = False
        'Compruebo si ha seleccionado algun articulo de los de precio ultima compra=0
        cadWHERE2 = ""
        Sql = ""
        For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag = "" Then
                    Sql = Sql & "M"
                Else
                    cadWHERE2 = cadWHERE2 & "M"
                End If
            End If
        Next
    
        If Sql <> "" Then
            MsgBox "No puede actualizar los articulos cuyo precio ultima compra sea 0", vbExclamation
            Exit Function
        End If
    
        If cadWHERE2 = "" Then
            MsgBox "Seleccione algun articulo para actualizar", vbExclamation
            Exit Function
        End If
    
        'Llegado aqui todo correcto. Hacemos la pregunta de actualizar y a correr
        Sql = "artículo"
        If Len(cadWHERE2) > 1 Then Sql = Sql & "s"
        Sql = "Va a actualizar los precios de " & Len(cadWHERE2) & " " & Sql & vbCrLf & vbCrLf & "¿Desea continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNo) <> vbYes Then Exit Function
        
        
        'Aqui esta el proceso de actualizacion de articulos
        Me.lblIndicadorCorregir.Caption = "Actualización precios"
        Me.Refresh
        Espera 0.5
        
       'Para el LOG
       Sql = cadWHERE & vbCrLf
       For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag <> "" Then Sql = Sql & ListView4.ListItems(TotalArray).Text & "|"
            End If
        Next
        Sql = Mid(Sql, 1, 237)
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        LOG.Insertar 4, vUsu, "Correccion precios: " & vbCrLf & Sql
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        
        
        
        
        
        
        
        
        
        For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag <> "" Then
                    
                    'lo metemos en transaccion. Si queremos vamos
                    Me.lblIndicadorCorregir.Caption = ListView4.ListItems(TotalArray).Text
                    Me.lblIndicadorCorregir.Refresh
                    
                                        
                    conn.BeginTrans
                    If ActualizaPrecios(TotalArray) Then
                        conn.CommitTrans
                    Else
                        conn.RollbackTrans
                    End If
                    
                    
                End If
            End If
        Next
    
    
        ActualizarPrecios = True
End Function


Private Function ActualizaPrecios(NumeroItem As Integer) As Boolean

On Error GoTo EActualizaPrecios
    ActualizaPrecios = False
    With ListView4.ListItems(NumeroItem)
        If OpcionMensaje = 16 Then
            'ACtualizador de precio normal
            If Me.cmbActualizarTar.ListIndex <> 2 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(7))))
                cadWHERE2 = "UPDATE sartic set preciove=" & cadWHERE2 & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "'"
                conn.Execute cadWHERE2
            End If
            If Me.cmbActualizarTar.ListIndex <> 1 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(8))))
                cadWHERE2 = "UPDATE slista set precioac=" & cadWHERE2 & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "' AND codlista =" & vCampos
                conn.Execute cadWHERE2
            End If
            
        Else
            'Precio articulos componentes
            '----------------------------
            vCampos = ""
            If Me.cmbActualizarTar.ListIndex <> 2 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(7))))
                vCampos = " preciove = " & cadWHERE2
            End If
            If Me.cmbActualizarTar.ListIndex <> 1 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(8))))
                If vCampos <> "" Then vCampos = vCampos & ","
                vCampos = vCampos & " preciouc = " & cadWHERE2
            End If
            cadWHERE2 = "UPDATE sartic set " & vCampos & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "'"
            conn.Execute cadWHERE2
            
            
                        

            
        End If
        
    End With
        
    ActualizaPrecios = True
    Exit Function
EActualizaPrecios:
    MuestraError Err.Number, ListView4.ListItems(NumeroItem).Text
End Function


Private Sub cmdDeselTodos_Click()
Dim I As Integer

    For I = 1 To ListView2.ListItems.Count
        ListView2.ListItems(I).Checked = False
    Next I
End Sub

Private Sub cmdEmail_Click()
    Unload Me
End Sub

Private Sub cmdEtiqEstan_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    If Index = 1 Then
        GenerarEtiquetasEstanterias
    Else
        NumRegElim = 0
    End If
    
    Screen.MousePointer = vbDefault
    Unload Me
End Sub


Private Sub GenerarEtiquetasEstanterias()
    
        'Cargo la tabla temporal con los datos que qeuremos imprimir
        cadWHERE2 = "insert into `tmpnseries` (`codusu`,`codartic`,`numlinea`,`numlinealb`) VALUES "
        cadWHERE = ""
        For NumRegElim = 1 To ListView3.ListItems.Count
            '                                                En el tag YA esta grabado
            If ListView3.ListItems(NumRegElim).Checked Then
                cadWHERE = cadWHERE & ",(" & vUsu.Codigo & "," & ListView3.ListItems(NumRegElim).Tag & ",0)"
                If (NumRegElim Mod 25) = 0 Then
                    conn.Execute cadWHERE2 & Mid(cadWHERE, 2) & ";"
                    cadWHERE = ""
                    DoEvents
                End If
            End If
        Next NumRegElim
        If cadWHERE <> "" Then conn.Execute cadWHERE2 & Mid(cadWHERE, 2) & ";"

End Sub

Private Sub cmdSelTodos_Click()
    Dim I As Integer

    For I = 1 To ListView2.ListItems.Count
        ListView2.ListItems(I).Checked = True
    Next I
End Sub




Private Sub Combo1_Click(Index As Integer)
   Select Case Index
        Case 0
            If vAnt <> Combo1(0).ListIndex Then CargarFacturasPendientesContabilizar
            vAnt = Combo1(0).ListIndex
    End Select
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
    Select Case Index
        Case 0
            vAnt = Combo1(0).ListIndex
    End Select
End Sub




Private Sub Form_Activate()
Dim Ok As Boolean

    
    Select Case OpcionMensaje
        Case 4 'Mostrar Nº Series
            If PrimeraVez Then
                PrimeraVez = False
                Me.Refresh
                Screen.MousePointer = vbHourglass
                Ok = ObtenerTamanyosArray
                If Ok Then Ok = SeparaCampos
                If Not Ok Then
                    'Error en SQL
                    'Salimos
                    Unload Me
                    Exit Sub
                End If
                CargarListaNSeries
            End If
            
        Case 8, 9, 17, 23, 24, 25 'Etiquetas de clientes/Proveedores/socios
            CargarListaClientes
'        Case 10 'Errores al contabilizar facturas
'            CargarListaErrContab
        Case 11, 111 'Lineas Factura a rectificar
            CargarListaLinFactu
            
        Case 14 'Mostrar Empresas del sistema
            CargarListaEmpresas
            
        Case 15
            'Etiquetas estanteria
            CargarArticulosEstanteria
            
        Case 16, 20
            'Articulos para corregir
            If OpcionMensaje = 16 Then
                CargarArticulosCorreccionPrecio
            Else
                CargaPVPPreciosArticulosConComponentes
            End If
            If Me.ListView4.ListItems.Count = 0 Then
                MsgBox "Ningún dato para mostrar", vbExclamation
                Unload Me
            End If
        Case 21
            CargarEmail
        Case 22
            CargarServiciosAFacturar
            
            
        Case 30 ' bloqueo de empresas
            cargaempresasbloquedas
        
        
        Case 31 ' facturas pendientes de contabilizar
            CargarFacturasPendientesContabilizar
            
            Combo1(0).ListIndex = 0
        
        
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim cad As String
On Error Resume Next

    'Icono del formulario
    Me.Icon = frmppal.Icon



    Me.FrameCobrosPtes.visible = False
    Me.FrameAcercaDe.visible = False
    Me.FrameNSeries.visible = False
    Me.FrameComponentes.visible = False
    Me.FrameComponentes2.visible = False
    Me.FrameErrores.visible = False
    FrameEtiqEstant.visible = False
    FrameCorreccionPrecios.visible = False
    FrameEMail.visible = False
    FrameServicios.visible = False
    
    FrameFrasPteContabilizar.visible = False
    
    PulsadoSalir = True
    PrimeraVez = True
    
    Select Case OpcionMensaje
        Case 1 'Mensaje de Cobros Pendientes
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaCobrosPtes
            Me.Caption = "Cobros Pendientes"
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 2 'Mensaje de no hay suficiente Stock
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaArtSinStock (vCampos)
            Me.Caption = "Artículos sin stock suficiente"
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 3 'Mensaje ACERCA DE
            CargaImagen
            Me.Caption = "Acerca de ....."
            PonerFrameAcercaDeVisible True, H, W
            Me.lblVersion.Caption = "Versión:  " & App.Major & "." & App.Minor & "." & App.Revision & " "
        
        Case 4 'Listado Nº Series Articulo
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Nº Serie"
            Me.Label7(1).Caption = "Seleccione los Nº de serie para el Albaran."
            Me.Label7(1).FontSize = 12
            PulsadoSalir = False
            
        Case 5 'Seleccionar tipo de Componente que queremos mostrar en Resumen
                'En mant. de Nº Series de Reparacion
            ponerFrameComponentesVisible True, H, W
            Me.Caption = "Componentes"
            Me.OptCompXMant.Value = True
            PonerFocoBtn Me.cmdAceptarComp
        
        Case 6 'Mostrar Prefacturacion de Albaranes
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaPreFacturar
            Me.Caption = "Prefacturación Albaranes"
            cad = RecuperaValor(vCampos, 1)
            If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
            Me.txtParam.Text = cad
            cad = RecuperaValor(vCampos, 2)
            If cad <> "" Then
                cad = Mid(cad, 1, Len(cad) - 1)
                If Trim(Me.txtParam.Text) <> "" Then
                    txtParam.Text = Me.txtParam.Text & vbCrLf & cad
                Else
                    txtParam.Text = cad
                End If
            End If
            cad = RecuperaValor(vCampos, 3)
            If cad <> "" Then
                cad = Mid(cad, 1, Len(cad) - 1)
                If Trim(Me.txtParam.Text) <> "" Then
                    txtParam.Text = Me.txtParam.Text & vbCrLf & cad
                Else
                    txtParam.Text = cad
                End If
            End If
            
            PonerFocoBtn Me.cmdAceptarComp
            
        Case 8, 17 'Etiquetas de Clientes
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Clientes"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            
        Case 9 'Etiquetas de Proveedores
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Proveedores"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
        
        Case 23 'Etiquetas de Socios
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Socios"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
        
        Case 24 'Situaciones de socios
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Situaciones"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
        
        
        Case 10 'Errores al contabilizar facturas
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaErrContab
            Me.Caption = "Facturas NO contabilizadas: "
            PonerFocoBtn Me.cmdAceptarCobros
        
        Case 11, 111 'Lineas Factura a Rectificar
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Lineas Factura a Rectificar"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            Me.cmdAceptarNSeries.Left = Me.cmdAceptarNSeries.Left + 1000
            Me.cmdCancelar.Left = Me.cmdCancelar.Left + 1000
        
        Case 12 'Mensaje Albaranes que no se van a Facturar
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaAlbaranes
            Me.Caption = "Facturación Albaranes"
            Me.Label1(0).Caption = "Existen Albaranes que NO se van a Facturar:"
            Me.Label1(0).top = 260
            Me.Label1(0).Left = 480
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 13 'Muestra Errores
            H = 6000
            W = 8800
            PonerFrameVisible Me.FrameErrores, True, H, W
            Me.Text1.Text = vCampos
            Me.Caption = "Errores"
        
        Case 14 'Muestra Empresas del sistema
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Selección"
            CargarListaEmpresas
        Case 15
            H = FrameEtiqEstant.Height
            W = FrameEtiqEstant.Width
            PonerFrameVisible FrameEtiqEstant, True, H, W
            
        Case 16, 20
            
            
            Caption = "Corrección precios"
            H = FrameCorreccionPrecios.Height
            W = FrameCorreccionPrecios.Width
            PonerFrameVisible FrameCorreccionPrecios, True, H, W
            Me.cmdCorrecotrPrecios(1).Cancel = True
            lblIndicadorCorregir.Caption = ""
            CargaComboActualizarPrecios
            If OpcionMensaje = 20 Then
                ListView4.ColumnHeaders(9).Text = " PUC correc."
                Label2(0).Caption = " Corrección de precios de articulos con componentes"
            Else
                ListView4.ColumnHeaders(9).Text = "Tarifa correc."
                Label2(0).Caption = " Corrección de errores y actualización de tarifas"
            End If
            
           
        Case 21
            'Ver email
            limpiar Me
            H = FrameEMail.Height
            W = FrameEMail.Width
            PonerFrameVisible FrameEMail, True, H, W
            If cadWHERE2 = "0" Then
                Caption = "Enviados"
                Label5(0).Caption = "Para"
            Else
                Label5(0).Caption = "De"
                Caption = "Recibidos"
            End If
            cmdEmail.Cancel = True
            PonerFocoBtn Me.cmdEmail
            
        Case 22 ' seleccioon de servicios del cliente que se van a facturar
            H = FrameServicios.Height
            W = FrameServicios.Width
            PonerFrameVisible Me.FrameServicios, True, H, W
            Me.Caption = "Servicios del Cliente"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
        
        
        Case 25 'Socios para liquidar como contado
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Socios de Contado"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
        
        Case 30 ' bloqueo de empresas
            Me.FrameBloqueoEmpresas.visible = True
            Caption = "Bloqueo empresas"
            W = Me.FrameBloqueoEmpresas.Width
            H = Me.FrameBloqueoEmpresas.Height + 300
            'Como cuando venga por esta opcion, viene llamado desde el manteusu
            Me.ListView6(0).SmallIcons = frmMantenusu.ImageList1
            Me.ListView6(1).SmallIcons = frmMantenusu.ImageList1
            Me.cmdBloqEmpre(1).Cancel = True
        
        Case 31 ' 31-facturas de pendientes de contabilizar
            H = Me.FrameFrasPteContabilizar.Height
            W = FrameFrasPteContabilizar.Width
            PonerFrameVisible FrameFrasPteContabilizar, True, H, W
        
            CargarCombo
        
        
    End Select
    'Me.cmdCancel(indFrame).Cancel = True
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerFrameCobrosPtesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 4600
        
    Select Case OpcionMensaje
        Case 1
            H = 5000
            W = 8600
            Me.Label1(0).Caption = "CLIENTE: " & vCampos
        Case 2
            W = 8800
            Me.cmdAceptarCobros.top = 4000
            Me.cmdAceptarCobros.Left = 4200
        Case 5 'Componentes
            W = 6000
            H = 5000
            Me.cmdAceptarCobros.Left = 4000

        Case 6, 7 'Prefacturar Albaranes
            W = 7000
            H = 6000
            Me.cmdAceptarCobros.top = 5400
            Me.cmdAceptarCobros.Left = 4600

        Case 10, 12 'Errores al contabilizar facturas
            H = 6000
            W = 8400
            Me.cmdAceptarCobros.top = 5300
            Me.cmdAceptarCobros.Left = 4900
            If OpcionMensaje = 12 Then
                Me.cmdCancelarCobros.top = 5300
                Me.cmdCancelarCobros.Left = 4600
                Me.cmdAceptarCobros.Left = 3300
                Me.Label1(1).top = 4800
                Me.Label1(1).Left = 3400
                Me.cmdAceptarCobros.Caption = "&SI"
                Me.cmdCancelarCobros.Caption = "&NO"
            End If
    End Select
            
    PonerFrameVisible Me.FrameCobrosPtes, visible, H, W

    If visible = True Then
        Me.txtParam.visible = (OpcionMensaje = 6 Or OpcionMensaje = 7)
        Me.Label1(0).visible = (OpcionMensaje = 1) Or (OpcionMensaje = 5) Or (OpcionMensaje = 12)
        Me.cmdCancelarCobros.visible = (OpcionMensaje = 12)
        Me.Label1(1).visible = (OpcionMensaje = 12)
    End If
End Sub


Private Sub PonerFrameAcercaDeVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame ACERCA DE visible y Ajustado al Formulario

    Me.FrameAcercaDe.visible = visible
    If visible = True Then
        'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
        Me.FrameAcercaDe.top = -90
        Me.FrameAcercaDe.Left = 0
        Me.FrameAcercaDe.Height = 4555
        Me.FrameAcercaDe.Width = 6600
        
        W = Me.FrameAcercaDe.Width
        H = Me.FrameAcercaDe.Height
    End If
End Sub


Private Sub PonerFrameNSeriesVisible(visible As Boolean, H As Integer, W As Integer)
'Pone el Frame de Nº Serie Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 5000
   
    If OpcionMensaje = 11 Then 'Lineas Factura a Rectificar
        W = 10900
    ElseIf OpcionMensaje = 14 Then
        W = 6500
        Me.Label7(1).visible = True
    ElseIf OpcionMensaje = 17 Then
        W = 10500
        Me.Label7(1).visible = False
    Else
        W = 8500
        Me.Label7(1).visible = False
    End If
    PonerFrameVisible Me.FrameNSeries, visible, H, W
End Sub


Private Sub ponerFrameComponentesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Componentes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

'    Me.FrameComponentes.visible = visible
    Me.FrameComponentes2.visible = visible
    
    H = 4000
    W = 5300
    PonerFrameVisible Me.FrameComponentes, visible, H, W
        
    If visible = True Then
        'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
        If vParamAplic.Departamento Then
            Me.OptCompXDpto.Caption = "Departamento"
        Else
            Me.OptCompXDpto.Caption = "Dirección"
        End If
    End If
End Sub


Private Sub CargarListaCobrosPtes()
'Muestra la lista Detallada de cobros en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String
Dim Impor As Currency
Dim Borrame As Currency

    If vParamAplic.ContabilidadNueva Then
        Sql = "SELECT numserie, numfactu, fecfactu, fecvenci, impvenci, impcobro ,gastos"
        Sql = Sql & " FROM cobros INNER JOIN formapago ON cobros.codforpa=formapago.codforpa "
        Sql = Sql & cadWHERE
    Else
        Sql = "SELECT numserie, codfaccl, fecfaccl, fecvenci, impvenci, impcobro ,gastos"
        Sql = Sql & " FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
        Sql = Sql & cadWHERE
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    ListView1.top = 900
    ListView1.Height = 3250
    ListView1.Width = 8100
    ListView1.Left = 160
    
    'Los encabezados
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Serie", 600
    ListView1.ColumnHeaders.Add , , "Nº Factura", 1000, 1
    ListView1.ColumnHeaders.Add , , "Fecha Factura", 1200, 2
    ListView1.ColumnHeaders.Add , , "Fecha Venci.", 1200, 2
    ListView1.ColumnHeaders.Add , , "Imp. Venci.(€)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Imp. Cobro(€)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Pte. Cobro(€)", 1250, 1
   ' Borrame = 0
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = Rs.Fields(0).Value 'Nº Serie
        ItmX.SubItems(1) = Rs.Fields(1).Value 'Nº Factura
        ItmX.SubItems(2) = Rs.Fields(2).Value 'Fecha Factura
        ItmX.SubItems(3) = Rs.Fields(3).Value 'Fecha Vencimiento
        ItmX.SubItems(4) = Rs.Fields(4).Value 'Importe Vencido
        ItmX.SubItems(5) = DBLet(Rs.Fields(5).Value, "N") 'Importe Cobrado
        'ItmX.SubItems(6) = RS.Fields(4).Value + DBLet(RS!gastos, "N") - DBLet(RS.Fields(5).Value, "N") 'Pendiente de cobro
        Impor = Rs.Fields(4).Value + DBLet(Rs!Gastos, "N") - DBLet(Rs.Fields(5).Value, "N") 'Pendiente de cobro
        ItmX.SubItems(6) = Impor
        If ItmX.SubItems(6) > 0 Then
            ItmX.ListSubItems(6).ForeColor = vbRed
           ' Borrame = Borrame + Impor
        End If
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub


Private Sub CargarListaArtSinStock(NomTabla As String)
'Muestra la lista Detallada de Articulos que no tienen stock suficiente en un ListView
Dim Rs As ADODB.Recordset
Dim Sql As String
    
    Sql = "SELECT " & NomTabla & ".codalmac," & NomTabla & ".codartic, " & NomTabla & ".nomartic, salmac.canstock as canstock, SUM(cantidad) as cantidad, canstock-SUM(cantidad) as disp ,conjunto "
    Sql = Sql & " FROM ((" & NomTabla & " INNER JOIN sartic ON " & NomTabla & ".codartic=sartic.codartic) INNER JOIN sfamia ON sartic.codfamia=sfamia.codfamia) "
    Sql = Sql & " INNER JOIN salmac ON " & NomTabla & ".codalmac=salmac.codalmac and " & NomTabla & ".codartic=salmac.codartic "
    Sql = Sql & cadWHERE 'Where numpedcl = 2 And sfamia.instalac = 0
    Sql = Sql & " GROUP by " & NomTabla & ".codalmac, " & NomTabla & ".codartic "
    

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    Me.ListView1.top = 500
     
    'Los encabezados
    ListView1.Width = 8400
    ListView1.Height = 3150
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Alm.", 500
    ListView1.ColumnHeaders.Add , , "Articulo", 1800, 2
    ListView1.ColumnHeaders.Add , , "Dec. Artic", 3300
    ListView1.ColumnHeaders.Add , , "Stock", 950, 2
    ListView1.ColumnHeaders.Add , , "Cantidad", 900, 2
    ListView1.ColumnHeaders.Add , , "No Disp.", 900, 2
    
    While Not Rs.EOF
        CargaItemStock Rs, ""
        'Si no tiene produccion miraremos si es conjunto
        If Not vParamAplic.Produccion Then
            If Rs!Conjunto = 1 Then
                Sql = Rs!codAlmac & "|" & Rs!codArtic & "|" & Rs!Cantidad & "|"
                CargaStockConjuntos Sql
            End If
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    Set Rs = Nothing

    
    
End Sub
    
Private Sub CargaStockConjuntos(linea As String)
    
        
        Set miRsAux = New ADODB.Recordset
            'Deberiamos cargar los elementos que tiene subconjuntos
            cadWHERE2 = "SELECT " & RecuperaValor(linea, 1) & ",sarti1.codarti1,nomartic,"
            cadWHERE2 = cadWHERE2 & " sarti1.cantidad * " & TransformaComasPuntos(RecuperaValor(linea, 3)) & " as cantidad,"
            cadWHERE2 = cadWHERE2 & " salmac.canstock as canstock,  canstock-(sarti1.cantidad * " & TransformaComasPuntos(RecuperaValor(linea, 3))
            cadWHERE2 = cadWHERE2 & ") as disp From sarti1, salmac, sartic"
            cadWHERE2 = cadWHERE2 & " Where sarti1.codarti1 = salmac.codArtic And sarti1.codarti1 = sartic.codArtic"
            cadWHERE2 = cadWHERE2 & " and sarti1.codartic='" & DevNombreSQL(RecuperaValor(linea, 2)) & "'"
            
            miRsAux.Open cadWHERE2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                CargaItemStock miRsAux, " * "
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        
        cadWHERE2 = ""
    Set miRsAux = Nothing
End Sub
 
    
Private Sub CargaItemStock(ByRef R As ADODB.Recordset, ByRef TxtAñadido As String)
Dim ItmX As ListItem
     If R!disp < 0 Then
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(R.Fields(0).Value, "000") 'Cod Almacen
            If TxtAñadido <> "" Then TxtAñadido = "[" & TxtAñadido & "]"
            ItmX.SubItems(1) = TxtAñadido & " " & R.Fields(1).Value 'Cod Artic
            ItmX.SubItems(2) = R.Fields(2).Value 'Nom Artic
            ItmX.SubItems(3) = R.Fields(3).Value 'Stock
            ItmX.SubItems(4) = R.Fields(4).Value 'Cantidad
            ItmX.SubItems(5) = R.Fields(5).Value 'No Disp
    End If
End Sub


Private Sub CargarListaNSeries()
'Carga las lista con todos los Nº de serie encontrados en la tabla:sserie
'para el articulo pasado como parametro en la cadwhere: "codartic='00012'"
'y que esten disponibles: numfactu y numalbar no tengan valor
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String
Dim cadLista As String
Dim Dif As Single

    On Error GoTo ECargarLista

    If cadWHERE2 = "" Then
        'Mostramos los nº serie libres para seleccionar la cantidad
        Sql = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
        Sql = Sql & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
        Sql = Sql & cadWHERE 'Where codartic='000012'
        'seleccionamos los que no esten asignados a ninguna factura ni albaran
        Sql = Sql & " AND ((isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='')) "
        Sql = Sql & " ORDER BY sserie.codartic, numserie "
        
    Else 'venimos de modificar la cantidad y seleccionamos los ya asignados
        If InStr(1, cadWHERE2, "|") > 0 Then
            Dif = CSng(RecuperaValor(cadWHERE2, 1))
            cadWHERE2 = RecuperaValor(cadWHERE2, 2)
        
            'seleccionamos nº serie del albaran que modificamos
            Sql = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
            Sql = Sql & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
            Sql = Sql & cadWHERE2
                
            
            If Dif < 0 Then
                'Si la diferencia de cantidad es < 0, mostrar en la lista los nº serie que
                'tiene la linea de albaran asignado con todos marcados y desmarcar el que no queremos
                
            Else
                'si la diferencia de cantidad es > 0, mostrar en la lista los nº de serie que
                'ya tenia asignados la linea del albaran más los libres para seleccionar los que añadimos de mas
                cadLista = ""
                Set Rs = New ADODB.Recordset
                Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs.EOF
                    cadLista = cadLista & ", " & Rs!numSerie
                    Rs.MoveNext
                Wend
                Rs.Close
                Set Rs = Nothing
                
                'mostrar tambien los nº serie sin asignar
                Sql = Sql & " OR (" & Replace(cadWHERE, "WHERE", "") & " and (numalbar=''or isnull(numalbar)))"
            End If
        Else
            'viene de una factura rectificativa, seleccionamos los nº de serie de
            'esa factura y marcamos los que queremos quitar
            Sql = cadWHERE2
        End If
    End If
    

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Los encabezados
    ListView2.Width = 7400
    Me.ListView2.Height = 3100
    Me.ListView2.Left = 650
    ListView2.ColumnHeaders.Clear
    
    ListView2.ColumnHeaders.Add , , "Nº Serie", 1800
    ListView2.ColumnHeaders.Add , , "Articulo", 1800
    ListView2.ColumnHeaders.Add , , "Desc. Artic", 3650
        
    If Rs.EOF Then Unload Me
    
    While Not Rs.EOF
         Set ItmX = ListView2.ListItems.Add
         ItmX.Text = Rs.Fields(0).Value 'num serie
         If Dif < 0 Then
            ItmX.Checked = True
         ElseIf Dif > 0 Then
            If InStr(1, cadLista, CStr(Rs!numSerie)) > 0 Then
                ItmX.Checked = True
            Else
                ItmX.Checked = False
            End If
         Else
            ItmX.Checked = False
         End If
         ItmX.SubItems(1) = Rs.Fields(1).Value 'Desc Artic
         ItmX.SubItems(2) = Rs.Fields(2).Value 'Nom Artic
         Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
ECargarLista:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar Nº Series", Err.Description
End Sub


Private Sub CargarListaComponentes(opt As Byte)
'Muestra la lista Detallada de cobros en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String
Dim Codigo As String, cadCodigo As String

    Select Case opt
        Case 1 'Mantenimiento
            Codigo = RecuperaValor(vCampos, 1)
            If Codigo = "" Then
                cadCodigo = " isnull(nummante) "
            Else
                cadCodigo = " nummante=" & DBSet(Codigo, "T")
            End If
            Sql = ObtenerSQLcomponentes(cadWHERE & " and " & cadCodigo)
            Me.Label1(0).Caption = "Mantenimiento: " & Codigo
            
        Case 2 'Departamento
            Codigo = RecuperaValor(vCampos, 2)
            If Codigo = "" Then
                cadCodigo = "isnull(coddirec)"
            Else
                cadCodigo = " coddirec=" & Codigo
            End If
            Sql = ObtenerSQLcomponentes(cadWHERE & " and " & cadCodigo)
            If vParamAplic.Departamento Then
                Me.Caption = "Equipos del Departamento"
                Me.Label1(0).Caption = " Departamento: " & RecuperaValor(vCampos, 3)
            Else
                Me.Caption = "Equipos de la Dirección"
                Me.Label1(0).Caption = " Dirección: " & Codigo & " " & RecuperaValor(vCampos, 3)
            End If
        
        Case 3 'Cliente
            Sql = ObtenerSQLcomponentes(cadWHERE)
            Me.Caption = "Equipos del Cliente"
            Me.Label1(0).Caption = "Cliente: " & RecuperaValor(vCampos, 4)
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Los encabezados
    ListView1.top = 800
    ListView1.Left = 280
    ListView1.Width = 4900
    ListView1.Height = 3250
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "TA", 760
    ListView1.ColumnHeaders.Add , , "Tipo Articulo", 2800
    ListView1.ColumnHeaders.Add , , "Cantidad", 1280, 2
    
    If Not Rs.EOF Then
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Rs.Fields(0).Value 'TA
            ItmX.SubItems(1) = Rs.Fields(1).Value 'Tipo Articulo
            ItmX.SubItems(2) = Rs.Fields(2).Value 'Cantidad
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
End Sub


Private Sub CargarListaPreFacturar()
'Muestra la lista Detallada de Albaranes a Factura en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList
    
    Sql = "CREATE TEMPORARY TABLE tmp ( "
    Sql = Sql & "codforpa SMALLINT(3) UNSIGNED  DEFAULT '0' NOT NULL, "
    Sql = Sql & "numalbar MEDIUMINT(7) UNSIGNED  DEFAULT '0' NOT NULL, "
    Sql = Sql & "dtoppago DECIMAL(4,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    Sql = Sql & "dtopgnral DECIMAL(4,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    Sql = Sql & "importe DECIMAL(12,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    Sql = Sql & "bruto DECIMAL(12,2) UNSIGNED  DEFAULT '0.0' NOT NULL) "
    conn.Execute Sql
     
'     SQL = "LOCK TABLES scaalb READ, slialb READ;"
'     Conn.Execute SQL
     
    Sql = "SELECT scaalb.codforpa, scaalb.numalbar, dtoppago, dtognral, round(sum(importel),2) as importe, round(sum(importel),2) - round(((round(sum(importel),2)*dtoppago)/100),2) - round(((round(sum(importel),2)*dtognral)/100),2) as bruto "
    Sql = Sql & " FROM (scaalb INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar) "
    Sql = Sql & " WHERE " & cadWHERE
    Sql = Sql & " GROUP BY scaalb.numalbar "
    Sql = Sql & " ORDER BY scaalb.codforpa, scaalb.numalbar "

    Sql = " INSERT INTO tmp " & Sql
    conn.Execute Sql
     
    Sql = " SELECT tmp.codforpa, sforpa.nomforpa, sum(tmp.bruto) as bruto"
    Sql = Sql & " FROM tmp, sforpa WHERE tmp.codforpa=sforpa.codforpa "
    Sql = Sql & " GROUP BY tmp.codforpa "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ListView1.Height = 3850
        ListView1.Width = 5400
        ListView1.Left = 500
        ListView1.top = 1200
    '    ListView1.GridLines = False
    
        'Los encabezados
        ListView1.ColumnHeaders.Clear
        
        ListView1.ColumnHeaders.Add , , " Forma de Pago", 3300
        ListView1.ColumnHeaders.Add , , "Base Imp.(€)", 2020, 1
     
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(Rs!codforpa.Value, "000") & "  " & Rs!nomforpa.Value
            
            ItmX.SubItems(1) = Rs!bruto
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
    'Borrar la tabla temporal
    Sql = " DROP TABLE IF EXISTS tmp;"
    conn.Execute Sql

ECargarList:
    If Err.Number <> 0 Then
         'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmp;"
        conn.Execute Sql
'        SQL = "UNLOCK TABLES "
'        Conn.Execute SQL
    End If
End Sub


Private Sub CargarListaClientes()
'Carga las lista con todos los clientes seleccionados en la tabla:sclien
'para imprimir etiquetas, pasando como parametro la cadwhere
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String, Men As String

    On Error GoTo ECargarLista

    Select Case OpcionMensaje
    Case 8
        'CLIENTES
        Sql = "SELECT codclien,nomclien,nifclien "
        Sql = Sql & "FROM scliente "
        If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
        Sql = Sql & " ORDER BY codclien "
        Men = "Cliente"
    Case 9
        'PROVEEDORES
        Sql = "SELECT codprove,nomprove,nifprove "
        Sql = Sql & "FROM sprove "
        If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
        Sql = Sql & " ORDER BY codprove "
        Men = "Proveedor"
    Case 17
        'CLIENTES MANTENIMIENTO
        Sql = cadWHERE
        Men = "Cliente"
    Case 23
        'Socios
        Sql = "SELECT codclien,nomclien,nifclien "
        Sql = Sql & "FROM sclien "
        If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
        Sql = Sql & " ORDER BY codclien "
        Men = "Socios"
    Case 24
        ' situaciones de socios
        Sql = "SELECT codsitua,nomsitua "
        Sql = Sql & "FROM ssitua "
        If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
        Sql = Sql & " ORDER BY codsitua "
        Men = "Situaciones"
        
    Case 25
        ' socios para liquidar de efectivo
        If cadWHERE2 = "shilla" Then
            Sql = "SELECT distinct sclien.codclien,sclien.numeruve,sclien.nomclien,sclien.nifclien "
            Sql = Sql & "FROM sclien INNER JOIN shilla on sclien.codclien = shilla.codsocio and sclien.numeruve = shilla.numeruve "
        Else
            Sql = "SELECT sclien.codclien,sclien.numeruve,nomclien,nifclien "
            Sql = Sql & "FROM sclien inner join sfactsoctr on sclien.codclien = sfactsoctr.codsocio and sclien.numeruve = sfactsoctr.numeruve "
        End If
        
        If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
        Sql = Sql & " ORDER BY sclien.codclien "
        Men = "Socios"
        
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'Los encabezados
        ListView2.Width = 9400
        ListView2.top = 500
        ListView2.Height = 3620
        ListView2.ColumnHeaders.Clear
        
        ListView2.ColumnHeaders.Add , , Men, 1050
        
        If OpcionMensaje = 25 Then
            ListView2.Width = 7400
            
            ListView2.ColumnHeaders.Add , , "Uve", 1050
        End If
        
        ListView2.ColumnHeaders.Add , , "Nombre", 4000
        
        If OpcionMensaje <> 24 Then
            ListView2.ColumnHeaders.Add , , "NIF", 1050
        End If
        
        If OpcionMensaje = 17 Then
            ListView2.Left = 500
            ListView2.ColumnHeaders.Add , , "Dpto", 550
            If vParamAplic.Departamento Then
                ListView2.ColumnHeaders.Add , , "Departamento", 2500
            Else
                ListView2.ColumnHeaders.Add , , "Direccion", 2500
            End If
        End If
        
        While Not Rs.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = Format(Rs.Fields(0).Value, "000000") 'cod clien/prove
             '[Monica]31/05/2012: por defecto todos tienen que estar marcados antes false
             ItmX.Checked = True
             ItmX.SubItems(1) = Rs.Fields(1).Value 'Nom clien/prove
             If OpcionMensaje <> 24 Then
                ItmX.SubItems(2) = Rs.Fields(2).Value 'NIF clien/prove
             End If
'             If OpcionMensaje <> 25 Then
'                ItmX.SubItems(3) = RS.Fields(3).Value 'NIF clien/prove
'             End If
             
             If OpcionMensaje = 17 Then
                ItmX.SubItems(3) = DBLet(Rs.Fields(3).Value, "T") 'cod dpto
                ItmX.SubItems(4) = DBLet(Rs.Fields(4).Value, "T") 'nom dpto
             End If
             
             If OpcionMensaje = 25 Then
                ItmX.SubItems(3) = Rs.Fields(3).Value 'NIF clien/prove
             End If
            
             Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar " & Men, Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub


Private Sub CargarListaVarios()
'Carga las lista con todos los clientes seleccionados en la tabla:sclien
'para imprimir etiquetas, pasando como parametro la cadwhere
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String, Men As String

    On Error GoTo ECargarLista

    Select Case OpcionMensaje
    Case 24
        ' situaciones de socios
        Sql = "SELECT codsitua,nomsitua "
        Sql = Sql & "FROM ssitua "
        If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
        Sql = Sql & " ORDER BY codsitua "
        Men = "Situaciones"
        
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'Los encabezados
        ListView2.Width = 9400
        ListView2.top = 500
        ListView2.Height = 3620
        ListView2.ColumnHeaders.Clear
        
        ListView2.ColumnHeaders.Add , , Men, 1050
        ListView2.ColumnHeaders.Add , , "Nombre", 4000
        ListView2.ColumnHeaders.Add , , "NIF", 1050
        
        If OpcionMensaje = 17 Then
            ListView2.Left = 500
            ListView2.ColumnHeaders.Add , , "Dpto", 550
            If vParamAplic.Departamento Then
                ListView2.ColumnHeaders.Add , , "Departamento", 2500
            Else
                ListView2.ColumnHeaders.Add , , "Direccion", 2500
            End If
        End If
        
        While Not Rs.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = Format(Rs.Fields(0).Value, "000000") 'cod clien/prove
             ItmX.Checked = False
             ItmX.SubItems(1) = Rs.Fields(1).Value 'Nom clien/prove
             ItmX.SubItems(2) = Rs.Fields(2).Value 'NIF clien/prove
             
             If OpcionMensaje = 17 Then
                ItmX.SubItems(3) = DBLet(Rs.Fields(3).Value, "T") 'cod dpto
                ItmX.SubItems(4) = DBLet(Rs.Fields(4).Value, "T") 'nom dpto
             End If
            
             Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar " & Men, Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub





Private Sub CargarListaErrContab()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    Sql = " SELECT  * "
    Sql = Sql & " FROM tmpErrFac "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ListView1.Height = 4500
        ListView1.Width = 7400
        ListView1.Left = 500
        ListView1.top = 500

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        If Rs.Fields(0).Name = "codprove" Then
            'Facturas de Compra
             ListView1.ColumnHeaders.Add , , "Prove.", 700
        Else 'Facturas de Venta
            ListView1.ColumnHeaders.Add , , "Tipo", 600
        End If
        ListView1.ColumnHeaders.Add , , "Factura", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Add , , "Error", 4620
    
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Format(Rs!NumFactu, "0000000")
            ItmX.SubItems(2) = Rs!FecFactu
            ItmX.SubItems(3) = Rs!Error
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub


Private Sub CargarListaLinFactu()
'Carga las lista con todas las lineas de la factura que estamos rectificando
'seleccionamos las que nos queremos llevar al Albaran de rectificacion
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarLista

    Sql = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre"
    Sql = Sql & " FROM slifac "
    If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
    Sql = Sql & " ORDER BY codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea "
    
    If OpcionMensaje = 111 Then Sql = Replace(Sql, "slifac", "slifaccli")
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        
        ListView2.top = 500
        ListView2.Left = 380
        ListView2.Width = 10100
        ListView2.Height = 3620
        
        'Los encabezados
        ListView2.ColumnHeaders.Clear
    
        ListView2.ColumnHeaders.Add , , "T.Alb", 660
        ListView2.ColumnHeaders.Add , , "Nº Alb", 840
        ListView2.ColumnHeaders.Add , , "Lin.", 450
        ListView2.ColumnHeaders.Item(3).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Alm", 460
        ListView2.ColumnHeaders.Add , , "Artic", 1380
        ListView2.ColumnHeaders.Add , , "Desc. Artic.", 2500
        ListView2.ColumnHeaders.Add , , "Cant.", 600
        ListView2.ColumnHeaders.Item(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Precio", 960
        ListView2.ColumnHeaders.Item(8).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 1", 600
        ListView2.ColumnHeaders.Item(9).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 2", 600
        ListView2.ColumnHeaders.Item(10).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Importe", 950
        ListView2.ColumnHeaders.Item(11).Alignment = lvwColumnRight
    
        While Not Rs.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = Rs!codtipoa 'cod tipo alb
             ItmX.Checked = False
             ItmX.SubItems(1) = Format(Rs!NumAlbar, "0000000") 'Nº Albaran
             ItmX.SubItems(2) = Rs!numlinea 'linea Albaran
             ItmX.SubItems(3) = Format(Rs!codAlmac, "000") 'cod almacen
             ItmX.SubItems(4) = Rs!codArtic 'Cod Articulo
             ItmX.SubItems(5) = Rs!NomArtic 'Nombre del Articulo
             ItmX.SubItems(6) = Rs!Cantidad
             ItmX.SubItems(7) = Format(Rs!precioar, FormatoPrecio)
             ItmX.SubItems(8) = Rs!dtoline1
             ItmX.SubItems(9) = Rs!dtoline2
             ItmX.SubItems(10) = Format(Rs!ImporteL, FormatoImporte)
             Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
    'si aparece la barra de desplazamiento ajustar el ancho
    If Me.ListView2.ListItems.Count > 11 Then
        Me.ListView2.ColumnHeaders(5).Width = 1200 'codartic
        Me.ListView2.ColumnHeaders(8).Width = 920  'precio
    End If
   
    
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Lineas Factura", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub




Private Sub CargarListaAlbaranes()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    Sql = cadWHERE 'cadwhere ya le pasamos toda la SQL
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        ListView1.Height = 3900
        ListView1.Width = 7200
        ListView1.Left = 500
        ListView1.top = 700

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        ListView1.ColumnHeaders.Add , , "Tipo", 700
        ListView1.ColumnHeaders.Add , , "Nº Albaran", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Item(3).Alignment = lvwColumnCenter
        ListView1.ColumnHeaders.Add , , "Cod. Cli.", 900
        ListView1.ColumnHeaders.Add , , "Cliente", 3400
    
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Format(Rs!NumAlbar, "0000000")
            ItmX.SubItems(2) = Rs!FechaAlb
            ItmX.SubItems(3) = Format(Rs!CodClien, "000000")
            ItmX.SubItems(4) = Rs!nomclien
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub



Private Sub CargarListaEmpresas()
'Carga las lista con todas las empresas que hay en el sistema
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String
Dim I As Integer

Dim Prohibidas As String

    On Error GoTo ECargarLista

    VerEmresasProhibidas Prohibidas
    
    Sql = "Select * from usuarios.empresasaritaxi order by codempre"
    Set ListView2.SmallIcons = frmppal.ImageListB
    ListView2.Width = 5000
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Empresa", 4900
    ListView2.HideColumnHeaders = True
    ListView2.GridLines = False
    ListView2.ListItems.Clear
    
    Set Rs = New ADODB.Recordset
    I = -1
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Sql = "|" & Rs!codempre & "|"
        If InStr(1, Prohibidas, Sql) = 0 Then
            Set ItmX = ListView2.ListItems.Add(, , Rs!nomempre, , 5)
            ItmX.Tag = Rs!codempre
            If ItmX.Tag = vEmpresa.codempre Then
                ItmX.Checked = True
                I = ItmX.Index
            End If
            ItmX.ToolTipText = Rs!AriTaxi
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    If I > 0 Then Set ListView2.SelectedItem = ListView2.ListItems(I)

    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargando datos empresas", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub


Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)
Dim Sql As String
Dim Rs As ADODB.Recordset

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    Sql = "Select codempre from usuarios.usuarioempresasaritaxi WHERE codusu = " & (vUsu.Codigo Mod 1000)
    Sql = Sql & " order by codempre"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
          VarProhibidas = VarProhibidas & Rs!codempre & "|"
          Rs.MoveNext
    Wend
    Rs.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte técnico"
    Set Rs = Nothing
End Sub



Private Sub CargaImagen()
On Error Resume Next
    Image2.Picture = LoadPicture(App.Path & "\minilogo.bmp")
    'Image1.Picture = LoadPicture(App.path & "\fondon.gif")
    Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If PulsadoSalir = False Then Cancel = 1
End Sub



Private Function ObtenerTamanyosArray() As Boolean
'Para el frame de los Nº de Serie de los Articulos
'En cada indice pone en CodArtic(i) el codigo del articulo
'y en Cantidad(i) la cantidad solicitada de cada codartic
Dim I As Integer, J As Integer

    ObtenerTamanyosArray = False
    'Primero a los campos de la tabla
    TotalArray = -1
    J = 0
    Do
        I = J + 1
        J = InStr(I, vCampos, "·")
        If J > 0 Then TotalArray = TotalArray + 1
    Loop Until J = 0
    
    If TotalArray < 0 Then Exit Function
    
    'Las redimensionaremos
    ReDim codArtic(TotalArray)
    ReDim Cantidad(TotalArray)
    
    ObtenerTamanyosArray = True
End Function


Private Function SeparaCampos() As Boolean
'Para el frame de los Nº de Serie de los Articulos
Dim Grupo As String
Dim I As Integer
Dim J As Integer
Dim C As Integer 'Contador dentro del array

    SeparaCampos = False
    I = 0
    C = 0
    Do
        J = I + 1
        I = InStr(J, vCampos, "·")
        If I > 0 Then
            Grupo = Mid(vCampos, J, I - J)
            'Y en la martriz
            InsertaGrupo Grupo, C
            C = C + 1
        End If
    Loop Until I = 0
    SeparaCampos = True
End Function


Private Sub InsertaGrupo(Grupo As String, Contador As Integer)
Dim J As Integer
Dim cad As String

    J = 0
    cad = ""
    
    'Cod Artic
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        J = 1
    End If
    codArtic(Contador) = cad
    
    'Cantidad
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        cad = Grupo
        Grupo = ""
    End If
    Cantidad(Contador) = cad
End Sub





Private Sub imgCheck_Click(Index As Integer)
Dim b As Boolean
    If Index < 2 Then
        'En el listview3
        b = Index = 1
        For TotalArray = 1 To ListView3.ListItems.Count
            ListView3.ListItems(TotalArray).Checked = b
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
        
    Else
        'En el listview4
        b = Index = 3
        For TotalArray = 1 To ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Tag <> "" Then
                ListView4.ListItems(TotalArray).Checked = b
            Else
                ListView4.ListItems(TotalArray).Checked = False
            End If
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
    End If
End Sub



Private Sub imgSelServ_Click(Index As Integer)
    If Index = 4 Then
        For TotalArray = 1 To ListView5.ListItems.Count
            ListView5.ListItems(TotalArray).Checked = True
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
    Else
        For TotalArray = 1 To ListView5.ListItems.Count
            ListView5.ListItems(TotalArray).Checked = False
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
    
    
    End If
    
End Sub

Private Sub OptCompXClien_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXDpto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXMant_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub



Private Function RegresarCargaEmpresas() As String
Dim Sql As String
Dim Parametros As String
Dim I As Integer

    CadenaDesdeOtroForm = ""
    
        Sql = ""
        Parametros = ""
        For I = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(I).Checked Then
                Sql = Sql & Me.ListView2.ListItems(I).Text & "|"
                Parametros = Parametros & "1" 'Contador
            End If
        Next I
        CadenaDesdeOtroForm = Len(Parametros) & "|" & Sql
        'Vemos las conta
        Sql = ""
        For I = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(I).Checked Then
                Sql = Sql & Me.ListView2.ListItems(I).Tag & "|"
            End If
        Next I
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Sql
    
    
        RegresarCargaEmpresas = CadenaDesdeOtroForm

End Function



Private Sub CargarArticulosEstanteria()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim It As ListItem

    Sql = "select sartic.codartic,nomartic,preciove,codigiva,nomfamia from sartic,sfamia where sartic.codfamia=sfamia.codfamia"
    If cadWHERE <> "" Then Sql = Sql & " AND " & cadWHERE
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalArray = 0
    While Not Rs.EOF
        Set It = ListView3.ListItems.Add
        
        'Ponemos el codigo de articulo y el TIPO de IVA
        It.Tag = "'" & DevNombreSQL(Rs!codArtic) & "'," & Rs!Codigiva
        It.Text = Rs!NomArtic
        It.SubItems(1) = Format(Rs!preciove, cadWHERE2)
        It.SubItems(2) = Rs!nomfamia
        It.Checked = True
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    
    
End Sub




Private Sub CargarArticulosCorreccionPrecio()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim It As ListItem
Dim margen As Currency
Dim MargenT As Currency
Dim ImpPVP As Currency
Dim ImpTar As Currency
Dim Aux As Currency
Dim decimales As Long
Dim precioUC As Currency
Dim SoloImporteMenor As Boolean
Dim SobreUPC As Boolean

    'El amrgen a aplicar
    'Si la tarifa es sobre el PVP es el articulo
    'si es sobre UPC entonces es sobre el de la tarifa

    lblIndicadorCorregir = "LEYENDO BD"
    lblIndicadorCorregir.Refresh
    

    
    'Si NUMREGELIM=1 entonces esta marcada la opcion(check) de solo importes menores
    If NumRegElim = 1 Then SoloImporteMenor = True
    
    
    
    'Comprobamos la tarifa donde se aplica, si sobre PVP o sobre ultima compra (%tarifa)
    Sql = DevuelveDesdeBD(conAri, "opcionINC", "starif", "codlista", vCampos)
    SobreUPC = Val(Sql) = 1
            
    
    TotalArray = InStr(1, cadWHERE2, ",")
    Sql = Mid(cadWHERE2, TotalArray + 1)
    decimales = Len(Sql)
    'Formato
    cadWHERE2 = "#,##0." & Mid(cadWHERE2, TotalArray + 1)
    
    'Sql
    Sql = " SELECT sartic.nomartic,slista.codartic,sartic.preciove,sartic.preciouc,"
    Sql = Sql & "slista.precioac, slista.codlista, starif.nomlista,"
    Sql = Sql & "sartic.margecom as margenArt,starif.margecom as margetar"
    Sql = Sql & " FROM   (slista INNER JOIN sartic ON slista.codartic=sartic.codartic)"
    Sql = Sql & " INNER JOIN starif  ON slista.codlista=starif.codlista WHERE "

    Sql = Sql & cadWHERE '& " AND "
    ''SQL = SQL & " sartic.preciove <> sartic.preciouc + round((sartic.preciouc * if(isnull(sartic.margecom),0,sartic.margecom))/100," & Decimales & ")"
    
    Sql = Sql & " ORDER BY slista.codartic"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

    '
  
    TotalArray = 0
    
    While Not Rs.EOF
        'Calculo los importes
        lblIndicadorCorregir.Caption = Rs!codArtic
        lblIndicadorCorregir.Refresh
        
        margen = DBLet(Rs!margenart, "N") / 100
        MargenT = DBLet(Rs!margetar, "N") / 100
        precioUC = DBLet(Rs!precioUC, "N")
        
        Aux = margen * precioUC
        ImpPVP = Round2(precioUC + Aux, decimales)
        
        'El de la tarifa
        If SobreUPC Then
            Aux = MargenT * precioUC
            ImpTar = Round2(precioUC + Aux, CLng(decimales))
        Else
        
            Aux = MargenT * ImpPVP
            ImpTar = Round2(ImpPVP + Aux, CLng(decimales))
        End If
        Aux = Round2(Rs!preciove, decimales)
        
        Sql = ""
        

        If SoloImporteMenor Then
            If Aux >= ImpPVP Then
                'El primero esta bien
                'Veamos el segundo. En la tarifa
                Aux = Round2(Rs!precioac, decimales)
                If Aux < ImpTar Then Sql = "M"
            Else
                Sql = "M"
            End If
        
        
        Else
            If Aux = ImpPVP Then
                'El primero esta bien
                'Veamos el segundo. En la tarifa
                Aux = Round2(Rs!precioac, decimales)
                If Aux <> ImpTar Then Sql = "M"
            Else
                Sql = "M"
            End If
        End If
        
        If Sql <> "" Then
            Set It = ListView4.ListItems.Add
            It.Tag = DevNombreSQL(Rs!codArtic)
            It.ToolTipText = It.Tag
            It.Text = It.Tag
            It.SubItems(1) = Rs!NomArtic
            Aux = Round2(precioUC, decimales)
            It.SubItems(2) = Format(Aux, cadWHERE2)
            
            It.SubItems(3) = Format(margen * 100, FormatoPorcen)
            Aux = Round2(Rs!preciove, decimales)
            It.SubItems(4) = Format(Aux, cadWHERE2)
            
            It.SubItems(5) = Format(MargenT * 100, FormatoPorcen)
            Aux = Round2(Rs!precioac, decimales)
            It.SubItems(6) = Format(Aux, cadWHERE2)
            

            It.SubItems(7) = Format(ImpPVP, cadWHERE2)
            It.SubItems(8) = Format(ImpTar, cadWHERE2)
            
            
            
            If precioUC = 0 Then
                'Precio ultima compra =0
                'NOOOOO se puede actualizar la tarifa
                It.Tag = "" 'para no actualizar
                It.Checked = False
                It.Bold = True
                It.ForeColor = vbRed
            Else
                
            End If
            It.Checked = False
        End If
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            Me.Refresh
            DoEvents
        End If
    Wend
    Rs.Close
    cmbActualizarTar.ListIndex = 0
    lblIndicadorCorregir.Caption = ""
End Sub

Private Sub txtMante_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub CargaPVPPreciosArticulosConComponentes()
Dim decimales As Byte
Dim Sql As String
Dim Impor As Currency
Dim IA As Currency
Dim PC As Currency
Dim PCC As Currency

    Set miRsAux = New ADODB.Recordset
    
    lblIndicadorCorregir = "LEYENDO BD"
    lblIndicadorCorregir.Refresh
    
    
    'Fomato importe
    TotalArray = InStr(1, cadWHERE2, ",")
    Sql = Mid(cadWHERE2, TotalArray + 1)
    decimales = Len(Sql)
    'Formato
    cadWHERE2 = "#,##0." & Mid(cadWHERE2, TotalArray + 1)
    
    
    'Tres columna svamos a ponerlas a tamaño 0
    ListView4.ColumnHeaders(6).Width = 0
    ListView4.ColumnHeaders(7).Width = 0
    
    Sql = "select sarti1.*,s1.nomartic,s1.preciove pre2,s1.margecom,s1.preciouc,"
    Sql = Sql & " sarti1.cantidad,s2.preciove, s2.preciouc coste"
    Sql = Sql & " from sarti1,sartic as s1,sartic as s2 where sarti1.codartic=s1.codartic and sarti1.codarti1=s2.codartic"
    'Si lleva WHERE
    If cadWHERE <> "" Then
        vCampos = Replace(cadWHERE, "sartic.", "s1.")
        Sql = Sql & " AND " & vCampos
        vCampos = ""
    End If
    
    Sql = Sql & " ORDER BY sarti1.codartic"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Sql = ""

    While Not miRsAux.EOF
        If Sql <> miRsAux!codArtic Then
            'Nuevo articulo
            lblIndicadorCorregir = miRsAux!codArtic
            lblIndicadorCorregir.Refresh
            If Sql <> "" Then
                'Si precioventa distionto   o pcompra distionto
                If IA <> Impor Or PC <> PCC Then
                    vCampos = vCampos & Format(IA, cadWHERE2) & "|" & Format(Impor, cadWHERE2) & "|"
                    vCampos = vCampos & Format(PC, cadWHERE2) & "|" & Format(PCC, cadWHERE2) & "|"
                    InsertarItemARticuloConjunto vCampos
                End If
                    
                
            End If
            Sql = miRsAux!codArtic
            vCampos = miRsAux!codArtic & "|" & miRsAux!NomArtic & "|"
            PC = DBLet(miRsAux!precioUC, "N")
            vCampos = vCampos & Format(PC, cadWHERE2)
            vCampos = vCampos & "|" & Format(DBLet(miRsAux!margecom, "N"), FormatoPorcen) & "|"
            
            IA = miRsAux!pre2
            PCC = 0 'precio compra calculado
            Impor = 0
        End If
        Impor = Impor + Round2((miRsAux!Cantidad * miRsAux!preciove), CLng(decimales))
        PCC = PCC + Round2((miRsAux!Cantidad * DBLet(miRsAux!coste, "N")), CLng(decimales))
        miRsAux.MoveNext
    Wend
    If Sql <> "" Then
        If IA <> Impor Or PC <> PCC Then
            vCampos = vCampos & Format(IA, cadWHERE2) & "|" & Format(Impor, cadWHERE2) & "|"
            vCampos = vCampos & Format(PC, cadWHERE2) & "|" & Format(PCC, cadWHERE2) & "|"
            InsertarItemARticuloConjunto vCampos
        End If
    End If
    miRsAux.Close
    lblIndicadorCorregir = ""
End Sub



Private Sub InsertarItemARticuloConjunto(Datos As String)
Dim It As ListItem

        Set It = ListView4.ListItems.Add
        It.Tag = RecuperaValor(Datos, 1)
        It.ToolTipText = It.Tag
        It.Text = It.Tag
        It.SubItems(1) = RecuperaValor(Datos, 2)  'nomartic
    
        It.SubItems(2) = RecuperaValor(Datos, 3)  'precio UC del articulo
        It.SubItems(3) = RecuperaValor(Datos, 4)  ' Margen
        
        It.SubItems(4) = RecuperaValor(Datos, 5)  'PVP articulo
        It.SubItems(7) = RecuperaValor(Datos, 6)  'PVP calculado
        It.SubItems(8) = RecuperaValor(Datos, 8)  'PUC calculado
        
            
End Sub

Private Sub CargaComboActualizarPrecios()
    cmbActualizarTar.Clear
    
    If OpcionMensaje = 16 Then
        'ART Y TARIFAS
        cmbActualizarTar.Tag = "Artículos y tarifas|Solo artículo|Solo tarifas|"
    Else
        cmbActualizarTar.Tag = "PVP y PUC|Solo PVP|Solo PUC|"
    End If
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 1)
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 2)
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 3)
    cmbActualizarTar.Tag = ""
    cmbActualizarTar.ListIndex = 0
End Sub



Private Sub CargarEmail()
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from scrmmail WHERE " & cadWHERE, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Me.txmemail(0).Text = miRsAux!email
        
        Me.txmemail(4).Text = miRsAux!FechaHora
        Me.txmemail(1).Text = DBLet(miRsAux!asunto, "T")
        Me.txmemail(2).Text = DBLet(miRsAux!adjuntos, "T")
        Me.txmemail(3).Text = DBLet(miRsAux!cuerpo, "T")
    
    
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub




Private Sub CargarServiciosAFacturar()
'Carga las lista con todas las lineas de la factura que estamos rectificando
'seleccionamos las que nos queremos llevar al Albaran de rectificacion
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarLista

    Sql = "SELECT fecha, hora, shilla.codsocio, sclien.nomclien, idservic, codusuar, impventa "
    Sql = Sql & " FROM shilla INNER JOIN sclien ON shilla.codsocio = sclien.codclien "
    Sql = Sql & " where facturadocliente = 0 "
    If cadWHERE <> "" Then Sql = Sql & " AND " & cadWHERE
    Sql = Sql & " ORDER BY fecha, hora, codsocio "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        
'        ListView5.Top = 500
'        ListView5.Left = 380
'        ListView5.Width = 10100
'        ListView5.Height = 3620
        
        'Los encabezados
        ListView5.ColumnHeaders.Clear
    
        ListView5.ColumnHeaders.Add , , "Fecha", 1200
        ListView5.ColumnHeaders.Add , , "Hora", 840
        ListView5.ColumnHeaders.Add , , "Socio", 750
'        ListView5.ColumnHeaders.item(3).Alignment = lvwColumnRight
        ListView5.ColumnHeaders.Add , , "Nombre", 2360
        ListView5.ColumnHeaders.Add , , "Id.Servic", 1200
        ListView5.ColumnHeaders.Add , , "Usuario", 2500
        ListView5.ColumnHeaders.Add , , "Importe", 1000
        ListView5.ColumnHeaders.Item(7).Alignment = lvwColumnRight
    
        While Not Rs.EOF
             Set ItmX = ListView5.ListItems.Add
             ItmX.Text = Format(Rs!Fecha, "dd/mm/yyyy") 'fecha
             ItmX.Checked = False
             ItmX.SubItems(1) = Format(Rs!hora, "hh:mm:ss") 'hora
             ItmX.SubItems(2) = Format(Rs!codSocio, "000000") 'socio
             ItmX.SubItems(3) = DBLet(Rs!nomclien, "T") 'nombre del socio
             ItmX.SubItems(4) = DBLet(Rs!idservic, "T") 'indentificacion
             ItmX.SubItems(5) = DBLet(Rs!codusuar, "T") 'Usuario
             ItmX.SubItems(6) = Format(Rs!impventa, FormatoImporte)
             Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Servicios del Cliente", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub

Private Sub cargaempresasbloquedas()
Dim It As ListItem
    On Error GoTo Ecargaempresasbloquedas
    Set Rs = New ADODB.Recordset
    Sql = "select empresasaritaxi.codempre,nomempre,nomresum,usuarioempresasaritaxi.codempre bloqueada from usuarios.empresasaritaxi left join usuarios.usuarioempresasaritaxi on "
    Sql = Sql & " empresasaritaxi.codempre = usuarioempresasaritaxi.codempre And (usuarioempresasaritaxi.codusu = " & Parametros & " Or codusu Is Null)"
    '[Monica] solo ariconta
    Sql = Sql & " WHERE aritaxi like 'aritaxi%' "
    Sql = Sql & " ORDER BY empresasaritaxi.codempre"
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Errores = Format(Rs!codempre, "00000")
        Sql = "C" & Errores
        
        If IsNull(Rs!bloqueada) Then
            'Va al list de la derecha
            Set It = ListView6(0).ListItems.Add(, Sql)
            It.SmallIcon = 1
        Else
            Set It = ListView6(1).ListItems.Add(, Sql)
            It.SmallIcon = 2
        End If
        It.Text = Errores
        It.SubItems(1) = Rs!nomempre
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    Errores = ""
    Exit Sub
Ecargaempresasbloquedas:
    MuestraError Err.Number, Err.Description
    Me.cmdBloqEmpre(0).Enabled = False
    Errores = ""
    Set Rs = Nothing
End Sub

Private Sub CargarFacturasPendientesContabilizar()
Dim Sql As String
Dim SQL2 As String
Dim Rs As ADODB.Recordset
Dim It As ListItem

    Sql = CADENA
    
    Select Case Combo1(0).ListIndex
        Case 0 'todos
        
        Case 1 ' facturas venta
            SQL2 = " and codigo1 = 0"
        Case 2 ' facturas cuotas extraordinarias
            SQL2 = " and codigo1 = 1"
        Case 3 ' facturas cuotas
            SQL2 = " and codigo1 = 2"
        Case 4 ' fras rectificativas cuotas
            SQL2 = " and codigo1 = 3"
        Case 5 ' fras servicios cliente
            SQL2 = " and codigo1 = 4"
        Case 6 ' fras publicidad cliente
            SQL2 = " and codigo1 = 5"
        Case 7 ' fras rectificativas cliente
            SQL2 = " and codigo1 = 6"
        Case 8 ' fras rectificativas publicidad
            SQL2 = " and codigo1 = 7"
        Case 9 ' fras ventas cliente
            SQL2 = " and codigo1 = 8"
        Case 10 ' fras liquidacion
            SQL2 = " and codigo1 = 9"
        Case 11 ' fras publicidad socio
            SQL2 = " and codigo1 = 10"
        Case 12 ' fras rectificativas liquidacion
            SQL2 = " and codigo1 = 11"
            
    End Select
    
    Sql = Sql & SQL2 & " order by 7,6 "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView22.ColumnHeaders.Clear

    ListView22.ColumnHeaders.Add , , "Tipo Factura", 3000 '2600
    ListView22.ColumnHeaders.Add , , "Fecha", 1400
    ListView22.ColumnHeaders.Add , , "Factura", 1500, 0
    ListView22.ColumnHeaders.Add , , "Nombre", 4900, 0 '3400, 0
    ListView22.ColumnHeaders.Add , , "Importe", 2000, 1 '1800, 1
    
    ListView22.ListItems.Clear
    
    ListView22.SmallIcons = frmppal.ImgListPpal
    
    TotalArray = 0
    While Not Rs.EOF
        Set It = ListView22.ListItems.Add
            
        It.Text = DBLet(Rs!nombre1, "T")
        It.SubItems(1) = DBLet(Rs!fecha1, "F")
        It.SubItems(2) = DBLet(Rs!Nombre2, "T")
        It.SubItems(3) = DBLet(Rs!nombre3, "T")
        It.SubItems(4) = Format(DBLet(Rs!Importe1, "N"), "###,###,##0.00")
        
        If vEmpresa.TieneSII Then
            If DBLet(Rs!fecha1, "F") < DateAdd("d", vEmpresa.SIIDiasAviso * (-1), Now) Then
                It.ForeColor = vbRed
                It.ListSubItems.Item(1).ForeColor = vbRed
                It.ListSubItems.Item(2).ForeColor = vbRed
                It.ListSubItems.Item(3).ForeColor = vbRed
                It.ListSubItems.Item(4).ForeColor = vbRed
            End If
        End If
        
        Select Case DBLet(Rs!Codigo1, "N")
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 ' clientes
                It.SmallIcon = 8
            Case 10, 11, 12 ' proveedor
                It.SmallIcon = 36
        End Select
        
        ListView22.Refresh
        
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    
End Sub

Private Sub CargarCombo()
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    Combo1(0).Clear
    
    Combo1(0).AddItem "Todas"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Ventas Socios"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Cuotas extraordinarias"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Cuotas Socios"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    Combo1(0).AddItem "Rectificativas Cuotas"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
    
    Combo1(0).AddItem "Servicios Cliente"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 5
    Combo1(0).AddItem "Publicidad Cliente"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 6
    Combo1(0).AddItem "Rectificativas Servicios"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 7
    Combo1(0).AddItem "Rectif.Publicidad Cliente"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 8
    Combo1(0).AddItem "Ventas Cliente"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 9

    Combo1(0).AddItem "Liquidacion Socio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 10
    Combo1(0).AddItem "Publicidad Socio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 11
    Combo1(0).AddItem "Rectificativas Liquidacion"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 12

End Sub


Private Sub cmdCerrarFras_Click(Index As Integer)
    Unload Me
End Sub

