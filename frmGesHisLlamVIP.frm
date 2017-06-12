VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGesHisLlamVIP 
   Caption         =   "Histórico de Llamadas VIP"
   ClientHeight    =   9900
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   90
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   91
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3720
      TabIndex        =   88
      Top             =   90
      Width           =   885
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   150
         TabIndex        =   89
         Top             =   180
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Informe de Servicios"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4650
      TabIndex        =   86
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   87
         Top             =   180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8190
      TabIndex        =   85
      Top             =   300
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Caption         =   "LOCALIZACION DEL SERVICIO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   3375
      Left            =   120
      TabIndex        =   40
      Top             =   900
      Width           =   10675
      Begin VB.CheckBox Check1 
         Caption         =   "Liquidado Socio"
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
         Index           =   3
         Left            =   5850
         TabIndex        =   15
         Tag             =   "Liquidado Socio|N|N|0|1|shilla|liquidadosocio|||"
         Top             =   2880
         Width           =   1875
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Facturado Cliente"
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
         Left            =   7740
         TabIndex        =   16
         Tag             =   "Facturado Cliente|N|N|0|1|shilla|facturadocliente|||"
         Top             =   2880
         Width           =   2085
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
         Height          =   360
         Index           =   16
         Left            =   2130
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Nombre|T|S|||shilla|nomclien|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1050
         Width           =   3615
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
         Height          =   315
         Index           =   15
         Left            =   7500
         MaxLength       =   14
         TabIndex        =   13
         Tag             =   "Autorización|T|S|||shilla|codautor|||"
         Text            =   "Text"
         Top             =   1380
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Index           =   13
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Codigo cliente|N|S|||shilla|codclien|000000||"
         Text            =   "999999"
         Top             =   1050
         Width           =   810
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
         Height          =   360
         Index           =   12
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Telefono|T|S|||shilla|telefono|||"
         Text            =   "1234567890"
         Top             =   1410
         Width           =   1335
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
         Left            =   7500
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "Tipo servicio|N|S|0|1|shilla|tipservi|0||"
         Top             =   1740
         Width           =   1335
      End
      Begin VB.TextBox Text2 
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
         Left            =   3540
         TabIndex        =   52
         Text            =   "Text2"
         Top             =   510
         Width           =   3765
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
         Height          =   360
         Index           =   10
         Left            =   8910
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Matricula|T|S|||shilla|matricul|||"
         Text            =   "Text"
         Top             =   510
         Width           =   1305
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
         Height          =   315
         Index           =   9
         Left            =   7500
         MaxLength       =   7
         TabIndex        =   12
         Tag             =   "Identificacion|T|S|||shilla|idservic|||"
         Text            =   "Text"
         Top             =   1020
         Width           =   1335
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
         Height          =   360
         Index           =   7
         Left            =   1530
         MaxLength       =   8
         TabIndex        =   1
         Tag             =   "Hora|H|N|||shilla|hora|hh:mm:ss|S|"
         Text            =   "99:99:99"
         Top             =   510
         Width           =   1005
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
         Height          =   360
         Index           =   6
         Left            =   7500
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Licencia|T|S|||shilla|licencia|||"
         Text            =   "Text"
         Top             =   510
         Width           =   1305
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
         Height          =   360
         Index           =   5
         Left            =   1290
         MaxLength       =   35
         TabIndex        =   11
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2880
         Width           =   4425
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
         Height          =   360
         Index           =   4
         Left            =   1290
         MaxLength       =   30
         TabIndex        =   10
         Tag             =   "Población|T|S|||shilla|ciudadre|||"
         Text            =   "ABCDEFGHIJKLMNÑOPQRSTUVWXYZABC"
         Top             =   2520
         Width           =   4425
      End
      Begin VB.TextBox Text1 
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
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "HHHHHH"
         Top             =   2160
         Width           =   915
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
         Height          =   360
         Index           =   2
         Left            =   1290
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "Domicilio|T|S|||shilla|dirllama|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1800
         Width           =   4455
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
         Height          =   360
         Index           =   1
         Left            =   210
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Fecha|F|N|||shilla|fecha|dd/mm/yyyy|S|"
         Text            =   "99/99/9999"
         Top             =   510
         Width           =   1275
      End
      Begin VB.TextBox Text1 
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
         Left            =   2550
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Codigo vehiculo|N|N|||shilla|numeruve|000000|S|"
         Text            =   "Text"
         Top             =   510
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Autorización:"
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
         Index           =   16
         Left            =   5820
         TabIndex        =   57
         Top             =   1410
         Width           =   1365
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1020
         ToolTipText     =   "Buscar fecha"
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   3210
         Tag             =   "-1"
         ToolTipText     =   "Buscar Socio"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   14
         Left            =   210
         TabIndex        =   56
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
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
         Index           =   13
         Left            =   210
         TabIndex        =   55
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Hora"
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
         Left            =   1530
         TabIndex        =   54
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   53
         Top             =   285
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1020
         Tag             =   "-1"
         ToolTipText     =   "Buscar cliente"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Matricula"
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
         Index           =   12
         Left            =   8910
         TabIndex        =   51
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Identificación:"
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
         Index           =   11
         Left            =   5820
         TabIndex        =   50
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
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
         Index           =   10
         Left            =   2550
         TabIndex        =   49
         Top             =   285
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de servicio:"
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
         Index           =   9
         Left            =   5820
         TabIndex        =   48
         Top             =   1770
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Licencia"
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
         Index           =   8
         Left            =   7500
         TabIndex        =   47
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
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
         Index           =   7
         Left            =   210
         TabIndex        =   46
         Top             =   2910
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
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
         Index           =   6
         Left            =   210
         TabIndex        =   45
         Top             =   2550
         Width           =   1005
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   780
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   2190
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "CP"
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
         Left            =   210
         TabIndex        =   44
         Top             =   2190
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
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
         Left            =   210
         TabIndex        =   43
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   18
      Left            =   750
      MaxLength       =   10
      TabIndex        =   84
      Tag             =   "Codsocio|N|N|||shilla|codsocio|||"
      Text            =   "ABCDEFGHIJ"
      Top             =   1020
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   17
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   83
      Tag             =   "Puerta|T|S|||shilla|puerllama|||"
      Text            =   "ABCDEFGHIJ"
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Frame Frame6 
      Height          =   1395
      Left            =   150
      TabIndex        =   69
      Top             =   7530
      Width           =   10635
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
         Height          =   360
         Index           =   19
         Left            =   2190
         MaxLength       =   200
         TabIndex        =   36
         Tag             =   "Observaciones 2|T|S|||shilla|observa2|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   960
         Width           =   8235
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
         Height          =   360
         Index           =   22
         Left            =   2190
         MaxLength       =   200
         TabIndex        =   35
         Tag             =   "Observaciones Cliente|T|S|||shilla|observa1|||"
         Text            =   $"frmGesHisLlamVIP.frx":0000
         Top             =   600
         Width           =   8235
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
         Height          =   360
         Index           =   23
         Left            =   2190
         MaxLength       =   60
         TabIndex        =   34
         Tag             =   "Observaciones|T|S|||shilla|observac2|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   240
         Width           =   8235
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1920
         Tag             =   "-1"
         ToolTipText     =   "Ver Observaciones"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1920
         Tag             =   "-1"
         ToolTipText     =   "Ver Observaciones"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1920
         Tag             =   "-1"
         ToolTipText     =   "Ver Observaciones"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones II"
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
         Index           =   35
         Left            =   120
         TabIndex        =   72
         Top             =   960
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones Cli"
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
         Index           =   34
         Left            =   120
         TabIndex        =   71
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
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
         Index           =   33
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "IMPORTES FACTURADOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   3135
      Left            =   5790
      TabIndex        =   59
      Top             =   4320
      Width           =   4995
      Begin VB.TextBox Text1 
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
         Index           =   35
         Left            =   3420
         TabIndex        =   31
         Tag             =   "Imp.TX|N|S|||shilla|impespera|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1770
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Index           =   34
         Left            =   3420
         TabIndex        =   27
         Tag             =   "Imp.TX|N|S|||shilla|impdistanci|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   690
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   3420
         TabIndex        =   80
         Text            =   "Text2"
         Top             =   1770
         Width           =   1155
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   3420
         TabIndex        =   79
         Text            =   "Text2"
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox Text1 
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
         Index           =   33
         Left            =   1470
         TabIndex        =   30
         Tag             =   "Tpo.Espera|N|S|||shilla|tpoespera|###,##0||"
         Text            =   "Text"
         Top             =   1770
         Width           =   975
      End
      Begin VB.TextBox Text1 
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
         Index           =   27
         Left            =   3420
         TabIndex        =   29
         Tag             =   "Imp.Peaje|N|S|||shilla|imppeaje|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1410
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Index           =   31
         Left            =   3420
         TabIndex        =   28
         Tag             =   "Suplemento|N|S|||shilla|suplemen|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1050
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Index           =   32
         Left            =   1470
         TabIndex        =   26
         Tag             =   "Distancia|N|S|||shilla|distanci|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   690
         Width           =   975
      End
      Begin VB.TextBox Text1 
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
         Index           =   30
         Left            =   3420
         TabIndex        =   33
         Tag             =   "Imp.Venta|N|S|||shilla|impventa|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   2580
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Index           =   29
         Left            =   3420
         TabIndex        =   32
         Tag             =   "Imp.Compra|N|S|||shilla|impcompr|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   2220
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Index           =   28
         Left            =   3420
         TabIndex        =   25
         Tag             =   "Imp.TX|N|S|||shilla|importtx|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Precio 2"
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
         Index           =   26
         Left            =   2520
         TabIndex        =   82
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Precio 1"
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
         Index           =   25
         Left            =   2520
         TabIndex        =   81
         Top             =   750
         Width           =   1065
      End
      Begin VB.Line Line2 
         X1              =   210
         X2              =   4740
         Y1              =   2130
         Y2              =   2130
      End
      Begin VB.Label Label1 
         Caption         =   "Tpo Espera"
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
         Index           =   19
         Left            =   180
         TabIndex        =   78
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Suplidos"
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
         Index           =   31
         Left            =   180
         TabIndex        =   68
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Suplemento"
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
         Index           =   30
         Left            =   180
         TabIndex        =   67
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Distancia"
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
         Index           =   29
         Left            =   180
         TabIndex        =   66
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Importe a Cobrar:"
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
         Index           =   24
         Left            =   180
         TabIndex        =   64
         Top             =   2610
         Width           =   2955
      End
      Begin VB.Label Label1 
         Caption         =   "Importe a Pagar:"
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
         Index           =   23
         Left            =   180
         TabIndex        =   63
         Top             =   2250
         Width           =   2685
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. TX:"
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
         Index           =   22
         Left            =   180
         TabIndex        =   62
         Top             =   330
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3105
      Left            =   120
      TabIndex        =   58
      Top             =   4350
      Width           =   5565
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   2670
         TabIndex        =   77
         Text            =   "Text2"
         Top             =   2280
         Width           =   1035
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
         Height          =   360
         Index           =   11
         Left            =   1260
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "DNI|T|S|||shilla|dnivip|||"
         Text            =   "123456789012345"
         Top             =   1080
         Width           =   1965
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
         Height          =   360
         Index           =   8
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   18
         Tag             =   "Usuario|T|S|||shilla|npvip|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   690
         Width           =   4095
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
         Height          =   360
         Index           =   14
         Left            =   1260
         MaxLength       =   30
         TabIndex        =   17
         Tag             =   "Usuario|T|S|||shilla|codusuar|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   300
         Width           =   4095
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
         Height          =   360
         Index           =   26
         Left            =   1050
         MaxLength       =   40
         TabIndex        =   24
         Tag             =   "Centro VIP|T|S|||shilla|centrovip|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2670
         Width           =   4335
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
         Height          =   360
         Index           =   25
         Left            =   2670
         MaxLength       =   8
         TabIndex        =   23
         Tag             =   "Hora Final|H|S|||shilla|horfinal|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   1830
         Width           =   1035
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
         Height          =   360
         Index           =   24
         Left            =   2670
         MaxLength       =   8
         TabIndex        =   21
         Tag             =   "Hora Ocupado|H|S|||shilla|horocupa|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   1470
         Width           =   1035
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
         Height          =   360
         Index           =   21
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   22
         Tag             =   "Fecha Final|F|S|||shilla|fecfinal|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   1830
         Width           =   1245
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
         Height          =   360
         Index           =   20
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "Fecha ocupado|F|S|||shilla|fecocupa|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   1470
         Width           =   1245
      End
      Begin VB.Line Line1 
         X1              =   270
         X2              =   3720
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Label Label1 
         Caption         =   "Total"
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
         Index           =   18
         Left            =   240
         TabIndex        =   76
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "DNI"
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
         Left            =   210
         TabIndex        =   75
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "NP"
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
         Left            =   210
         TabIndex        =   74
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
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
         Index           =   15
         Left            =   210
         TabIndex        =   73
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Centro"
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
         Index           =   27
         Left            =   240
         TabIndex        =   65
         Top             =   2670
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1080
         ToolTipText     =   "Buscar fecha"
         Top             =   1830
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   1080
         ToolTipText     =   "Buscar fecha"
         Top             =   1470
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Finalizado:"
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
         Index           =   21
         Left            =   240
         TabIndex        =   61
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio"
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
         Index           =   20
         Left            =   240
         TabIndex        =   60
         Top             =   1470
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5400
      Top             =   9000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.CommandButton cmdAceptar 
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
      Left            =   8370
      TabIndex        =   37
      Top             =   9240
      Width           =   1135
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Left            =   9660
      TabIndex        =   38
      Top             =   9240
      Width           =   1135
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   9660
      TabIndex        =   39
      Top             =   9240
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   150
      TabIndex        =   41
      Top             =   9000
      Width           =   3975
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Menu mnopciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnvertodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnbarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnnuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnInforme 
         Caption         =   "Informe de Servicios"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnsalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmGesHisLlamVIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public FechaServ As String
Public HoraServ As String
Public NumerUve As String

Public WithEvents frmLlamPre As frmGesHisLlamPrev
Attribute frmLlamPre.VB_VarHelpID = -1
Public WithEvents frmCP As frmCPostal
Attribute frmCP.VB_VarHelpID = -1
Public WithEvents frmV As frmGesVehic
Attribute frmV.VB_VarHelpID = -1
Public WithEvents frmCal As frmCal
Attribute frmCal.VB_VarHelpID = -1
Public WithEvents frmS As frmGesSocios
Attribute frmS.VB_VarHelpID = -1
Public WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim btnAnyadir As Byte
Dim btnPrimero As Byte
Dim NombreTabla As String
Dim Ordenacion As String
Dim CadenaConsulta As String
Dim HaDevueltoDatos As Boolean
Private Modo As Byte
Dim kCampo As Byte
Dim ModificaLineas As Byte
Dim Fecha As Date

Private BuscaChekc As String


Dim cadB1 As String
'-------------------------------------------------------
'Se distinguen varios MODOS
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------
Private VieneDeBuscar As Boolean

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then CheckCadenaBusqueda Check1(Index), BuscaChekc
End Sub

Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub


Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim I As Long
Dim CadB As String
Dim Cad As String
Dim Indicador As String


    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
              If InsertarDesdeForm(Me) Then PosicionarData
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    
                    '[Monica] 02/06/2011: tras modificar volvemos al registro correspondiente
                    NumRegElim = adodc1.Recordset.AbsolutePosition
                    Me.adodc1.Refresh
                    If SituarDataPosicion(adodc1, NumRegElim, "") Then
                        PonerCampos
                    End If
                    PonerModo 2
                    'fin
'[Monica] 02/06/2011: comentado
'                    PosicionarData
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
Error1:
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False, BuscaChekc)
    cadB1 = ObtenerBusqueda(Me, True, BuscaChekc)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        PonerFoco Text1(1)
    End If
End Sub
Private Function DatosOk() As Boolean
Dim b As Boolean

DatosOk = False
'If Modo = 4 Then
'    DatosOk = True
'    Exit Function
'End If
''CODIGO DE socio
'If Text1(0).Text = "" Then
'    MsgBox "Debe introducir el código de socio.", vbExclamation
'    PonerFoco Text1(0)
'    Exit Function
'ElseIf Not IsNumeric(Text1(0).Text) Then
'        MsgBox "El código de socio debe ser numérico.", vbExclamation
'        PonerFoco Text1(0)
'        Exit Function
'End If
'
''Fecha
'If Text1(1).Text = "" Then
'    MsgBox "Debe introducir la fecha de la llamada.", vbExclamation
'    PonerFoco Text1(1)
'    Exit Function
'ElseIf Not IsDate(Text1(1).Text) Then
'    MsgBox "La fecha no tiene formato correcto.", vbExclamation
'    PonerFoco Text1(1)
'    Exit Function
'End If
'
''Hora
'If Text1(7).Text = "" Then
'    MsgBox "Debe introducir la hora de la llamada.", vbExclamation
'    PonerFoco Text1(7)
'    Exit Function
'ElseIf Not IsDate(Text1(7).Text) Then
'    MsgBox "La hora no tiene formato correcto.", vbExclamation
'    PonerFoco Text1(7)
'    Exit Function
'End If
'
''numero de vehiculo
'If Text1(8).Text = "" Then
'    MsgBox "Debe introducir el número de vehiculo.", vbExclamation
'    PonerFoco Text1(8)
'    Exit Function
'ElseIf Not IsNumeric(Text1(8).Text) Then
'        MsgBox "El número de vehiculo debe ser numérico.", vbExclamation
'        PonerFoco Text1(8)
'        Exit Function
'End If

    b = CompForm(Me, 1)

    



    DatosOk = b
    
End Function

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
        If adodc1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = adodc1.Recordset.Fields(0) & "|"
        Cad = Cad & adodc1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus()

If Combo1.Text = "0" Then
    Text1(13).Text = ""
    Text1(13).BackColor = &H80000018
    Text1(13).Enabled = False
    imgBuscar(1).Enabled = False
Else
    Text1(13).Enabled = True
    Text1(13).BackColor = &H80000005
    imgBuscar(1).Enabled = True
End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If PrimeraVez Then
        PrimeraVez = False
        If FechaServ <> "" Then
            If Me.adodc1.Recordset.EOF Then
                PonerCadenaBusqueda
            Else
                PonerCampos
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim I As Integer

    'Icono del form
    Me.Icon = frmPpal.Icon

    'ICONITOS DE LA BARRA
'    btnAnyadir = 5
'    btnPrimero = 13
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'        '.Buttons(9).Image = 10 'Lineas
'        .Buttons(9).Image = 16 'Imprmir
'        .Buttons(10).Image = 40 'Informe de servicios
'        .Buttons(11).Image = 15 'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'    End With

    With Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun1
        'ASignamos botones
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2 'Ver Todos
        .Buttons(1).Image = 3 'Añadir
        .Buttons(2).Image = 4 'Modificar
        .Buttons(3).Image = 5 'Eliminar
        .Buttons(8).Image = 16 'Imprimir
    End With
    
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun1
        .Buttons(1).Image = 40 'Informe de servicios
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun1
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With

    For I = 0 To Me.imgBuscar.Count - 1
        imgBuscar(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next
    
    imgFecha(0).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    imgFecha(4).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    imgFecha(5).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    
    
      
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    PrimeraVez = True
    CargarCombo
    
    NombreTabla = "shilla"
    CadenaConsulta = "Select * from " & NombreTabla
    
    Label1(25).Caption = "a " & Format(vParamAplic.PrecioxDistancia, "##0.0000")
    Label1(26).Caption = "a " & Format(vParamAplic.PrecioxTpoEspera, "##0.0000")
    
    If FechaServ <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        'como tenemos aqui el nº de albaran, buscar a que factura corresponde
        'en la scafaccli1
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura
        If cadB1 <> "" Then CadenaConsulta = CadenaConsulta & " and " & cadB1
'        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        'CadenaConsulta = CadenaConsulta & " where numfactu=-1"
        'Cambio sugerido por Msoler
        'mariela 01/07/2010
        'If Not publicidad Then
        CadenaConsulta = CadenaConsulta & " WHERE fecha is null  "
    End If
    
    '## A mano
    Ordenacion = " ORDER BY fecha,hora"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    adodc1.ConnectionString = conn
    adodc1.RecordSource = CadenaConsulta ' "Select * from " & NombreTabla & " where numeruve=-1"
    adodc1.Refresh
    
    If FechaServ = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
        End If
    Else
        If adodc1.Recordset.EOF Then
            PonerModo 0
        Else
            PonerModo 2
        End If
    End If
End Sub


Private Sub CargarCombo()
    Combo1.AddItem "Normal"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Crédito"
    Combo1.ItemData(Combo1.NewIndex) = 1
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo


    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    BuscaChekc = ""
    
    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I
    
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not adodc1.Recordset.EOF Then
        If adodc1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And adodc1.Recordset.RecordCount > 1
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    BloquearCmb Combo1, (Modo <> 1)
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
'    Combo1.Enabled = b
'    For i = 0 To 2
'        Check1(i).Enabled = b
'    Next i
    For I = 3 To 4
        Check1(I).Enabled = (Modo = 1)
    Next I
    
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
    
    ' No hay icono para las observaciones de 60 de longitud maxima
    Me.imgBuscar(2).Enabled = False
    Me.imgBuscar(2).visible = False
    
    Me.imgBuscar(4).Enabled = (Modo > 0)
    Me.imgBuscar(4).visible = (Modo > 0)
    Me.imgBuscar(5).Enabled = (Modo > 0)
    Me.imgBuscar(5).visible = (Modo > 0)
    
    BloquearTxt Text1(6), (Modo <> 1)
    BloquearTxt Text1(10), (Modo <> 1)
    BloquearTxt Text1(16), (Modo <> 1)
    
    
    
    For I = 4 To 5
        Me.imgFecha(I).Enabled = b
    Next I
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    '-----------------------------
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub



Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub
Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean
Dim I As Byte

    b = (Modo = 2 Or Modo = 5 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(1).Enabled = b
    
    b = (Modo = 2 Or Modo = 5)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    'imprimir
    Toolbar1.Buttons(8).Enabled = b
    
    '------------------------------------------
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b
End Sub

Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    Me.Combo1.ListIndex = -1
    Check1(3).Value = 0
    Check1(4).Value = 0
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub frmCal_Selec(vFecha As Date)
    Fecha = vFecha
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Text1(13).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(16).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim indice As Byte
Dim devuelve As String
    
    indice = 3
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'Poblacion
    Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve)
    'provincia
    Text1(indice + 2).Text = devuelve

End Sub

Private Sub frmLlamPre_DatoSeleccionado(CadenaSeleccion As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaSeleccion, 1)
        CadB = Aux
        Aux = Format(ValorDevueltoFormGrid(Text1(7), CadenaSeleccion, 2), FormatoHora)
        CadB = CadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 3)
        CadB = CadB & " AND " & Aux
'        Aux = ValorDevueltoFormGrid(Text1(8), CadenaSeleccion, 4)
'        CadB = CadB & " AND " & Aux
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte
Dim Observaciones As String

Select Case Index
    Case 4 ' observaciones
        If Modo = 3 Or Modo = 4 Then
            CadenaDesdeOtroForm = Text1(22).Text
        Else
            CadenaDesdeOtroForm = ""
            If Not adodc1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(adodc1.Recordset!observa1, "T")
        End If
        frmFacClienteObser.Modificar = Modo >= 3
        frmFacClienteObser.Text1 = CadenaDesdeOtroForm
        frmFacClienteObser.Show vbModal
        'Llevara DOS VALORES.
        'Si modifica y el texto
        If Modo = 3 Or Modo = 4 Then
            If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(22).Text = Mid(CadenaDesdeOtroForm, 3)
        End If
        CadenaDesdeOtroForm = ""

    
    Case 5 ' observaciones
        If Modo = 3 Or Modo = 4 Then
            CadenaDesdeOtroForm = Text1(19).Text
        Else
            CadenaDesdeOtroForm = ""
            If Not adodc1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(adodc1.Recordset!observa2, "T")
        End If
        frmFacClienteObser.Modificar = Modo >= 3
        frmFacClienteObser.Text1 = CadenaDesdeOtroForm
        frmFacClienteObser.Show vbModal
        'Llevara DOS VALORES.
        'Si modifica y el texto
        If Modo = 3 Or Modo = 4 Then
            If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(19).Text = Mid(CadenaDesdeOtroForm, 3)
        End If
        CadenaDesdeOtroForm = ""

    Case 0 'población
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            If Index = 9 Then
                indice = 4
            Else
                PonerFoco Text1(3)
            End If
    Case 1 'clientes
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0|1|"
            frmCli.Show vbModal
            Set frmCli = Nothing
    Case 3 'socios
            Set frmS = New frmGesSocios
            frmS.DatosADevolverBusqueda = "0|1|"
            frmS.Show vbModal
            Set frmS = Nothing
    End Select
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte
    Select Case Index
        Case 0
            indice = 1
        Case 1
            indice = 11
        Case 2
            indice = 18
        Case 3
            indice = 19
        Case 4
            indice = 20
        Case 5
            indice = 21
    End Select
    Set frmCal = New frmCal
    If Text1(indice).Text <> "" Then PonerFormatoFecha Text1(indice)
    frmCal.Fecha = Now
    If Text1(indice).Text <> "" Then frmCal.Fecha = CDate(Text1(indice).Text)
    Screen.MousePointer = vbDefault
    frmCal.Show vbModal
    If IsDate(Fecha) Then
        If Fecha <> "0:00:00 " Then Text1(indice) = Fecha
    End If
    Set frmCal = Nothing
    PonerFoco Text1(indice)
End Sub



Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnInforme_Click()
    AbrirListado 120
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text1_GotFocus(Index As Integer)
kCampo = Index
ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim encontrado As String

If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
If Text1(Index).Text <> "" Then Text1(Index) = UCase(Text1(Index).Text)

Select Case Index
    Case 0 'socio
        If Modo = 1 Then Exit Sub
        If Text1(Index).Text <> "" Then
            Text1(18).Text = Text1(0).Text
            Text1(Index).Text = Format(Text1(Index).Text, "00000")
            encontrado = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(Index).Text, "T")
            If encontrado = "" Then
                MsgBox "El código de socio introducido no existe.", vbExclamation
                PonerFoco Text1(Index)
            Else
                TraerDatosSocio Text1(Index).Text
            End If
        End If
'    Case 8
'        If Text1(Index).Text <> "" Then Text1(Index).Text = Format(Text1(Index).Text, "00000")
    Case 13 'cliente
        If Modo = 1 Then Exit Sub
        If Text1(Index).Text <> "" Then
            Text1(Index).Text = Format(Text1(Index).Text, "000000")
            encontrado = DevuelveDesdeBD(conAri, "nomclien", "scliente", "codclien", Text1(Index).Text, "T")
            If encontrado = "" Then
                MsgBox "El código de cliente introducido no existe.", vbExclamation
                PonerFoco Text1(Index)
            Else
            
                TraerDatosCliente Text1(Index).Text
            
            End If
        End If
    Case 3 'CP
        If Text1(Index) <> "" Then
            'Poblacion
            Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, encontrado)
            'provincia
            Text1(Index + 2).Text = encontrado
        End If
    
    Case 27 To 31
        PonerFormatoDecimal Text1(Index), 3
        
    Case 7, 24 To 25
        If Text1(Index).Text <> "" Then PonerFormatoHora Text1(Index)
    
        If Index = 7 And Modo = 3 Then
            Text1(24).Text = Text1(7).Text
            Text1(25).Text = Text1(7).Text
        End If
    
    Case 1, 20 To 21
        If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
        If Index = 1 And Modo = 3 Then
            Text1(20).Text = Text1(1).Text
            Text1(21).Text = Text1(1).Text
        End If
        
    Case 32 ' distancia
        If Index = 32 Then ' distancia
            PonerFormatoDecimal Text1(Index), 6
            Text2(2).Text = ""
            If Text1(Index).Text <> "" Then
                Text2(2).Text = Round2(CCur(Text1(Index).Text) * vParamAplic.PrecioxDistancia, 2)
                Text1(34).Text = Text2(2).Text
            End If
        End If
       
    Case 33 ' tpoespera
        PonerFormatoEntero Text1(Index)
        Text2(3).Text = ""
        If Text1(Index).Text <> "" Then
            Text2(3).Text = Round2(CCur(ImporteSinFormato(Text1(Index).Text)) * vParamAplic.PrecioxTpoEspera, 2)
            Text1(35).Text = Text2(3).Text
        End If
        
    Case 34, 35 ' importe de distancia y de tpo de espera
        PonerFormatoDecimal Text1(Index), 3
End Select

    If Modo = 3 Or Modo = 4 Then
        If Index = 20 Or Index = 21 Or Index = 24 Or Index = 25 Then
            CalcularDuracion
        End If
            
        If Index = 27 Or Index = 28 Or Index = 31 Or Index = 32 Or Index = 33 Or Index = 34 Or Index = 35 Then
            CalcularImportes
        End If
    End If

End Sub

Private Sub CalcularImportes()
Dim ImpTX As Single
Dim ImpDis As Single
Dim ImpSup As Single
Dim ImpPea As Single
Dim ImpEsp As Single
Dim ImpPag As Single
Dim ImpCob As Single
    
    On Error GoTo eCalcularImportes
    
    ImpTX = 0
    ImpDis = 0
    ImpSup = 0
    ImpPea = 0
    ImpEsp = 0
    ImpPag = 0
    ImpCob = 0
    
    If Text1(28).Text = "" Then
        ImpTX = 0
    Else
        ImpTX = ImporteSinFormato(Text1(28).Text)
    End If
    
    If Text1(32).Text = "" Or Text1(34).Text = "" Then
        ImpDis = 0
    Else
        ImpDis = ImporteSinFormato(Text1(34).Text)
    End If
    If Text1(31).Text = "" Then
        ImpSup = 0
    Else
        ImpSup = ImporteSinFormato(Text1(31).Text)
    End If
    If Text1(27).Text = "" Then
        ImpPea = 0
    Else
        ImpPea = ImporteSinFormato(Text1(27).Text)
    End If
    If Text1(33).Text = "" Or Text1(35).Text = "" Then
        ImpEsp = 0
    Else
        ImpEsp = ImporteSinFormato(Text1(35).Text)
    End If
    
    ImpCob = ImpTX + ImpDis + ImpSup + ImpPea + ImpEsp
    
    '[Monica]27/10/2011: no descontamos el 3% pq se descuenta en retencion
    ImpPag = ImpCob ' - Round2(ImpCob * 0.03, 2)
    
    Text1(29).Text = Format(ImpPag, "##,###,##0.00")
    Text1(30).Text = Format(ImpCob, "##,###,##0.00")
    Exit Sub
    
eCalcularImportes:
    MuestraError Err.Number, "Calcular importes", Err.Description
End Sub

Private Sub CalcularDuracion()
Dim Inicio As Date
Dim Fin As Date
Dim Duracion As Integer
Dim Horas As Integer
Dim Minutos As Integer
Dim Diferencia As Single

    If Text1(20).Text = "" Or Text1(21).Text = "" Or Text1(24).Text = "" Or Text1(25).Text = "" Then Exit Sub
    
    Inicio = CDate(Format(Text1(20).Text, "dd/mm/yyyy") & " " & Format(Text1(24).Text, "hh:mm:ss"))
    Fin = CDate(Format(Text1(21).Text, "dd/mm/yyyy") & " " & Format(Text1(25).Text, "hh:mm:ss"))
    If Inicio <= Fin Then
        Horas = DateDiff("h", Inicio, Fin)
        Minutos = DateDiff("n", Inicio, Fin)
        
        '[Monica]07/09/2011: la diferencia en minutos pasa al tpo de espera
        Text1(33).Text = Format(Minutos, "###,###,##0")
        
        Horas = Minutos \ 60
        Minutos = Minutos Mod 60
        
        Diferencia = CSng(CInt(Horas) & "," & CInt(Minutos))

        Text2(0).Text = Format(Diferencia, "##0.00")
        
        If Text1(33).Text <> "" Then
            Text1(35).Text = Round2(CCur(ImporteSinFormato(Text1(33).Text)) * vParamAplic.PrecioxTpoEspera, 2)
            CalcularImportes
        End If

    Else
        MsgBox "Error entre rangos de fecha. Revise", vbExclamation
        Text2(0).Text = ""
'        PonerFoco Text1(25)
    End If

End Sub


Private Sub VisualizarDuracion2()
Dim Inicio As Date
Dim Fin As Date
Dim Duracion As Integer
Dim Horas As Integer
Dim Minutos As Integer
Dim Diferencia As Single

    
    If Text1(20).Text = "" Or Text1(21).Text = "" Or Text1(24).Text = "" Or Text1(25).Text = "" Then Exit Sub
    
    Inicio = CDate(Format(Text1(20).Text, "dd/mm/yyyy") & " " & Format(Text1(24).Text, "hh:mm:ss"))
    Fin = CDate(Format(Text1(21).Text, "dd/mm/yyyy") & " " & Format(Text1(25).Text, "hh:mm:ss"))
    If Inicio <= Fin Then
        Horas = DateDiff("h", Inicio, Fin)
        Minutos = DateDiff("n", Inicio, Fin)
        
        Horas = Minutos \ 60
        Minutos = Minutos Mod 60
        
        Diferencia = CSng(CInt(Horas) & "," & CInt(Minutos))

        
        
        Text2(0).Text = Format(Diferencia, "##0.00")
    Else
    
    End If

End Sub



Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
            
    End Select
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub
Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5  'Buscar
           mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
            
        Case 1  'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
'        Case 9
'            mnLineas_Click
        Case 8  'imprimir
            mnImprimir_Click
        
'        Case 10 ' Informe de servicios por socio o cliente
'            mnInforme_Click
'        Case 11  'Salir
'            mnSalir_Click
'        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
'            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub
Private Sub mnNuevo_Click()
         BotonAnyadir
End Sub
Private Sub BotonAnyadir()
'Añadir registro en tabla de trabajadores: straba (Cabecera)

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(1).Text = Date
    Text1(7).Text = Format(Now, "hh:mm:ss")
    
    Text1(20).Text = Date
    Text1(24).Text = Format(Now, "hh:mm:ss")
    
    Text1(21).Text = Date
    Text1(25).Text = Format(Now, "hh:mm:ss")
    
    Combo1.ListIndex = 1
    Combo1.Enabled = False
    Check1(3).Value = 0
    Check1(4).Value = 0
    
    PonerFoco Text1(1)
End Sub
Private Sub mnModificar_Click()
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub
Private Sub BotonModificar()
'Prepara el Form para Modificar
Dim DeVarios As Boolean
Dim Sql As String
On Error GoTo EModificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    Combo1.Enabled = False
    PonerModo 4
    If Combo1.Text = "0" Then Text1(13).Enabled = False
    imgFecha(0).Enabled = False
'    imgBuscar(3).Enabled = False
    If Combo1.Text = "0" Then
        imgBuscar(1).Enabled = False
    Else
        imgBuscar(1).Enabled = True
    End If
    PonerFoco Text1(0)
   
EModificar:
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub mnEliminar_Click()
    BotonEliminar
End Sub
Private Sub mnSalir_Click()
    Unload Me
End Sub
Private Sub BotonEliminar()
Dim msg As String
Dim Sql As String

On Error GoTo EEliminar

msg = "Esta seguro que desea eliminar la llamada del día:" & Text1(1).Text & "?"
If MsgBox(msg, vbYesNo) = vbYes Then
    NumRegElim = adodc1.Recordset.AbsolutePosition
    Sql = "Delete from shilla where fecha='" & Format(Text1(1).Text, FormatoFecha) & "' and hora='" & Format(Text1(7).Text, FormatoHora)
    Sql = Sql & "' and numeruve=" & Text1(0).Text
    conn.Execute Sql
End If

If SituarDataTrasEliminar(adodc1, NumRegElim) Then
    PonerCampos
End If

EEliminar:
If Err.Number <> 0 Then
    MsgBox "Error al eliminar conductor." & Err.Description
End If
End Sub

Private Sub BotonVerTodos()
'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
    cadB1 = ""
End Sub
Private Sub mnBuscar_Click()
    BotonBuscar
End Sub
Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(1)
        Text1(1).BackColor = vbLightBlue 'vbYellow
    Else
        HacerBusqueda
        If adodc1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue 'vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If adodc1.Recordset.EOF Then Exit Sub
    DesplazamientoData adodc1, Index, True
    PonerCampos
End Sub

Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    adodc1.RecordSource = CadenaConsulta
    adodc1.Refresh
    If adodc1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(0)
            Text1(0).BackColor = vbLightBlue 'vbYellow
        End If
        Exit Sub
    Else
        adodc1.Recordset.MoveFirst
        PonerCampos
        PonerModo 2
    End If

Screen.MousePointer = vbDefault
Exit Sub
EEPonerBusq:
    
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub
Private Sub PonerCampos()
Dim encontrado As String

On Error Resume Next

    
    If adodc1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, adodc1
    
    If Combo1.Text = "1" Then
        If Text1(13).Text <> "" Then
            encontrado = DevuelveDesdeBD(conAri, "nomclien", "scliente", "codclien", Text1(13).Text, "T")
            If encontrado = "" Then
                Text1(16).Text = encontrado
            End If
        End If
    End If
    If Text1(0).Text <> "" Then
        encontrado = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(0).Text, "T")
        If encontrado <> "" Then
            Text2(1).Text = encontrado
        End If
    End If
    
    Text1(18).Text = Text1(0).Text
    
'    ' distancia
'    Text2(2).Text = ""
'    If Text1(32).Text <> "" Then
'        Text2(2).Text = Round2(CCur(Text1(32).Text) * vParamAplic.PrecioxDistancia, 2)
'        Text1(34).Text = Text2(2).Text
'    End If
'
'    ' tpoespera
'    Text2(3).Text = ""
'    If Text1(33).Text <> "" Then
'        Text2(3).Text = Round2(CCur(ImporteSinFormato(Text1(33).Text)) * vParamAplic.PrecioxTpoEspera, 2)
'        Text1(35).Text = Text2(3).Text
'    End If
    
    VisualizarDuracion2
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub MandaBusquedaPrevia(CadB As String)
''Carga el formulario frmBuscaGrid con los valores correspondientes
'Dim Cad As String
'Dim Tabla As String
'Dim Titulo As String
'
'    'Llamamos a al form
'    '##A mano
'    Cad = ""
'    Cad = Cad & ParaGrid(Text1(1), 14, "Fecha")
'    Cad = Cad & ParaGrid(Text1(7), 14, "Hora")
'    Cad = Cad & ParaGrid(Text1(0), 14, "Codigo")
'    Cad = Cad & "Socio|sclien|nomclien|N||45·"
'    Cad = Cad & ParaGrid(Text1(13), 14, "Cliente")
'
'    Tabla = "shilla INNER JOIN sclien ON shilla.numeruve = sclien.numeruve"
'    Titulo = "Histórico Servicios"
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vTabla = Tabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|2|3|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 1
'        frmB.vConexionGrid = conAri
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If
'    Screen.MousePointer = vbDefault

    Set frmLlamPre = New frmGesHisLlamPrev

    frmLlamPre.DatosADevolverBusqueda = "0|1|2|4|"
    frmLlamPre.cWhere = CadB
    frmLlamPre.Show vbModal

    Set frmLlamPre = Nothing


End Sub

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

'    cad = "(fecha='" & Format(Text1(1).Text, FormatoFecha) & "' and hora='" & Format(Text1(7).Text, FormatoHora)
'    cad = cad & "' and numeruve=" & Text1(8).Text & ")"
'    If SituarDataMULTI(Adodc1, cad, Indicador) Then
'       PonerModo 2
'       lblIndicador(0).Caption = Indicador
'    Else
'       LimpiarCampos
'       PonerModo 0
'    End If

Dim vWhere As String

    If Not adodc1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Me.adodc1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If

End Sub

Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = "(fecha='" & Format(Text1(1).Text, FormatoFecha) & "' and hora='" & Format(Text1(7).Text, FormatoHora)
    Sql = Sql & "' and numeruve=" & Text1(0).Text & ")"
    
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function





Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "shilla"
        .Informe2 = "rGesHisLlam.rpt"
        If cadB1 <> "" Then
            .cadRegSelec = cadB1 'SQL2SF(cadB1)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(adodc1, Me)
        .cadTodosReg = ""
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|pOrden={tarjbanc.nomtarje}|"
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|"
        .NumeroParametros2 = 1
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub


Private Function ObtenerSelFactura() As String
Dim Cad As String
Dim RS As ADODB.Recordset

    On Error Resume Next

    Cad = ""
'    If Me.DesdeFichaCliente Then
        '
    Cad = " WHERE fecha=" & DBSet(FechaServ, "F") & " AND hora= " & DBSet(HoraServ, "H") & " AND numeruve=" & DBSet(NumerUve, "N")
        
'    Else
'        'Tengo YA el codigo de la factura
'                '******************************************************
'                'laura: esto se puede comentar, ya no hay movimiento FTI en la smoval
'                If hcoCodTipoM = "FTI" Then
'                    'no hay albaran directamente va a factura de ticket
'
'                    'ver si lo encontramos como factura: codtipom, numfactu,fecfactu
'                    cad = "SELECT COUNT(*) FROM scafaccli "
'                    cad = cad & " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
'                    If RegistrosAListar(cad) > 0 Then
'                        cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
'                    Else
'                        cad = ""
'                    End If
'                Else
'                    If hcoCodTipoM = "FAM" Then
'                        cad = "  WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
'                    End If
'                End If
'                '******************************************************
'
'                If cad = "" Then
'                    'En la smoval estaba e mov. de ALbaran
'                    cad = "SELECT codtipom,numfactu,fecfactu FROM scafaccli1 "
'                    cad = cad & " WHERE codtipoa=" & DBSet(hcoCodTipoM, "T") & " AND numalbar=" & hcoCodMovim & " AND fechaalb=" & DBSet(hcoFechaMov, "F")
'
'                    Set RS = New ADODB.Recordset
'                    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                    If Not RS.EOF Then 'where para la factura
'                        cad = " WHERE codtipom='" & RS!codtipom & "' AND numfactu= " & RS!NumFactu & " AND fecfactu=" & DBSet(RS!FecFactu, "F")
'                    Else
'                        cad = " WHERE numfactu=-1"
'                    End If
'                    RS.Close
'                    Set RS = Nothing
'                End If
'
'    End If
    ObtenerSelFactura = Cad
End Function

Private Sub TraerDatosCliente(CodClien As String)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim vCliente As CCliente

    If CodClien = "" Then Exit Sub
    Set vCliente = New CCliente
    
    If vCliente.LeerDatos(CodClien, False) Then
        Text1(12).Text = vCliente.TfnoClien
        Text1(16).Text = vCliente.Nombre
        Text1(2).Text = vCliente.Domicilio
        Text1(3).Text = vCliente.CPostal
        Text1(4).Text = vCliente.Poblacion
        Text1(5).Text = vCliente.Provincia
    End If
    
    Set vCliente = Nothing
    
End Sub

Private Sub TraerDatosSocio(codSocio As String)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim vSocio As CSocio

    If codSocio = "" Then Exit Sub
    Set vSocio = New CSocio
    
    If vSocio.LeerDatos(codSocio) Then
        Text2(1).Text = vSocio.Nombre
        Text1(10).Text = DevuelveDesdeBDNew(conAri, "sclien", "matricul", "codclien", codSocio, "N")
        Text1(6).Text = vSocio.Licencia
    End If
    
    Set vSocio = Nothing
    
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' Informe de servicios por socio o cliente
            mnInforme_Click
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
