VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "L"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrEstadisticasReparacionTecnico 
      Height          =   3495
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtDescTra 
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   1320
         Width           =   3585
      End
      Begin VB.TextBox txtTrab 
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
         Left            =   1080
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdEstadisticaReparacionTecnico 
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
         Left            =   3360
         TabIndex        =   19
         Top             =   2760
         Width           =   1095
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
         Left            =   4560
         TabIndex        =   20
         Top             =   2760
         Width           =   975
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
         Index           =   3
         Left            =   4200
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2040
         Width           =   1350
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
         Index           =   2
         Left            =   1680
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2040
         Width           =   1350
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   63
         Left            =   240
         TabIndex        =   108
         Top             =   2880
         Width           =   2865
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   0
         Left            =   720
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Técnico"
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
         Index           =   4
         Left            =   240
         TabIndex        =   36
         Top             =   1080
         Width           =   795
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   3900
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   3
         Left            =   3210
         TabIndex        =   34
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1380
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   2
         Left            =   690
         TabIndex        =   33
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albarán"
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
         Index           =   3
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Estadísticas reparación técnico"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   210
         TabIndex        =   31
         Top             =   330
         Width           =   5415
      End
   End
   Begin VB.Frame FrameFacturarCliente 
      Height          =   3015
      Left            =   0
      TabIndex        =   149
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkFacturarCliente 
         Caption         =   "Imprimir facturas generadas"
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
         Left            =   1740
         TabIndex        =   158
         Top             =   2040
         Value           =   1  'Checked
         Width           =   3465
      End
      Begin VB.CommandButton cmdFacturarCli 
         Caption         =   "Facturar"
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
         Left            =   4080
         TabIndex        =   153
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtBancoPr 
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
         Left            =   2040
         TabIndex        =   152
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtDescBancoPr 
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   156
         Text            =   "Text5"
         Top             =   1560
         Width           =   3615
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
         Index           =   31
         Left            =   2040
         TabIndex        =   151
         Text            =   "Text1"
         Top             =   1080
         Width           =   1350
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
         Index           =   25
         Left            =   5280
         TabIndex        =   154
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
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
         Index           =   34
         Left            =   240
         TabIndex        =   157
         Top             =   1560
         Width           =   660
      End
      Begin VB.Image imgBancoPr 
         Height          =   240
         Index           =   2
         Left            =   1740
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Index           =   33
         Left            =   240
         TabIndex        =   155
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   31
         Left            =   1740
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Facturación cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   23
         Left            =   240
         TabIndex        =   150
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame FrameListTrabajadores 
      Height          =   2535
      Left            =   3240
      TabIndex        =   98
      Top             =   480
      Width           =   5895
      Begin VB.TextBox txtTrab 
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
         Left            =   1080
         TabIndex        =   105
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtDescTra 
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
         Index           =   4
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   104
         Text            =   "Text1"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtTrab 
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
         Left            =   1080
         TabIndex        =   102
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtDescTra 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   101
         Text            =   "Text1"
         Top             =   840
         Width           =   3495
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
         Index           =   17
         Left            =   4560
         TabIndex        =   100
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdListTrabja 
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
         Left            =   3480
         TabIndex        =   99
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
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
         Index           =   59
         Left            =   120
         TabIndex        =   107
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
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
         Index           =   58
         Left            =   120
         TabIndex        =   106
         Top             =   840
         Width           =   615
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   4
         Left            =   780
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado trabajadores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   103
         Top             =   240
         Width           =   5415
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   3
         Left            =   780
         Top             =   840
         Width           =   240
      End
   End
   Begin VB.Frame FrameSituAlbaranes 
      Height          =   5055
      Left            =   0
      TabIndex        =   131
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdSituAlbaran 
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
         Left            =   4080
         TabIndex        =   137
         Top             =   4200
         Width           =   1095
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
         Index           =   23
         Left            =   5400
         TabIndex        =   138
         Top             =   4200
         Width           =   1095
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   360
         Style           =   1  'Checkbox
         TabIndex        =   136
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox txtDescClie 
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
         Index           =   6
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   146
         Text            =   "Text1"
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   1350
         TabIndex        =   135
         Text            =   "Text1"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtDescClie 
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
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   143
         Text            =   "Text1"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   1350
         TabIndex        =   134
         Text            =   "Text1"
         Top             =   1920
         Width           =   1335
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
         Index           =   30
         Left            =   4680
         TabIndex        =   133
         Text            =   "Text1"
         Top             =   1080
         Width           =   1350
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
         Index           =   29
         Left            =   1560
         TabIndex        =   132
         Text            =   "Text1"
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   1500
         Picture         =   "frmListado2.frx":0000
         ToolTipText     =   "Quitar al haber"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   1860
         Picture         =   "frmListado2.frx":014A
         ToolTipText     =   "Puntear al haber"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Albaranes"
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
         Index           =   32
         Left            =   360
         TabIndex        =   148
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   6
         Left            =   1020
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   75
         Left            =   360
         TabIndex        =   147
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   31
         Left            =   240
         TabIndex        =   145
         Top             =   1560
         Width           =   765
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   5
         Left            =   1020
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   74
         Left            =   360
         TabIndex        =   144
         Top             =   1920
         Width           =   705
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   30
         Left            =   4380
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   73
         Left            =   3720
         TabIndex        =   142
         Top             =   1125
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   29
         Left            =   1260
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   72
         Left            =   600
         TabIndex        =   141
         Top             =   1125
         Width           =   615
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         Index           =   30
         Left            =   240
         TabIndex        =   140
         Top             =   840
         Width           =   690
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Informe situación albaranes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   21
         Left            =   720
         TabIndex        =   139
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame FrProveedorxVenta 
      Height          =   5895
      Left            =   0
      TabIndex        =   43
      Top             =   30
      Width           =   6375
      Begin VB.TextBox txtDescProve 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text5"
         Top             =   4680
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
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
         Left            =   1140
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "Text5"
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
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
         Left            =   1140
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtArticulo 
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
         Left            =   1140
         MaxLength       =   16
         TabIndex        =   49
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   3480
         Width           =   3495
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
         Index           =   10
         Left            =   4560
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdVentaxProv 
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
         Left            =   3960
         TabIndex        =   52
         Top             =   5280
         Width           =   975
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
         Index           =   6
         Left            =   5160
         TabIndex        =   53
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtArticulo 
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
         Left            =   1140
         MaxLength       =   16
         TabIndex        =   48
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "Text5"
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   1260
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtDescClie 
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
         Index           =   4
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   1260
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtDescClie 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   1800
         Width           =   3495
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
         Index           =   9
         Left            =   1260
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Index           =   18
         Left            =   180
         TabIndex        =   72
         Top             =   4680
         Width           =   585
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   1
         Left            =   840
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Index           =   4
         Left            =   120
         TabIndex        =   70
         Top             =   3960
         Width           =   1125
      End
      Begin VB.Label Label3 
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
         Index           =   17
         Left            =   180
         TabIndex        =   69
         Top             =   4320
         Width           =   585
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   0
         Left            =   840
         Top             =   4320
         Width           =   240
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   2
         Left            =   840
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   16
         Left            =   180
         TabIndex        =   66
         Top             =   3525
         Width           =   585
      End
      Begin VB.Label Label3 
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
         Left            =   180
         TabIndex        =   65
         Top             =   3120
         Width           =   585
      End
      Begin VB.Label Label3 
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
         Left            =   180
         TabIndex        =   64
         Top             =   2205
         Width           =   585
      End
      Begin VB.Label Label3 
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
         Index           =   13
         Left            =   180
         TabIndex        =   63
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label Label3 
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
         Index           =   12
         Left            =   3660
         TabIndex        =   62
         Top             =   945
         Width           =   585
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   4290
         Top             =   975
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
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
         Index           =   3
         Left            =   120
         TabIndex        =   61
         Top             =   2760
         Width           =   810
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   1
         Left            =   840
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   4
         Left            =   960
         Top             =   2175
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   3
         Left            =   960
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   11
         Left            =   120
         TabIndex        =   58
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado venta x proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   5
         Left            =   330
         TabIndex        =   56
         Top             =   210
         Width           =   5415
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Index           =   10
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
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
         Index           =   11
         Left            =   180
         TabIndex        =   54
         Top             =   945
         Width           =   585
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   9
         Left            =   960
         Top             =   960
         Width           =   240
      End
   End
   Begin VB.Frame FrameLlamadas 
      Height          =   3975
      Left            =   360
      TabIndex        =   115
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtDescTra 
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
         Index           =   6
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "Text1"
         Top             =   2640
         Width           =   3435
      End
      Begin VB.TextBox txtTrab 
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
         Left            =   1200
         TabIndex        =   120
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtDescTra 
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
         Index           =   5
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   126
         Text            =   "Text1"
         Top             =   2280
         Width           =   3435
      End
      Begin VB.TextBox txtTrab 
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
         Left            =   1200
         TabIndex        =   119
         Top             =   2280
         Width           =   855
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
         Index           =   28
         Left            =   4200
         TabIndex        =   118
         Text            =   "Text1"
         Top             =   1320
         Width           =   1350
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
         Index           =   27
         Left            =   1320
         TabIndex        =   117
         Text            =   "Text1"
         Top             =   1320
         Width           =   1350
      End
      Begin VB.CommandButton cmdLlamadas 
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
         Left            =   3240
         TabIndex        =   121
         Top             =   3120
         Width           =   1095
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
         Index           =   22
         Left            =   4440
         TabIndex        =   122
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Trabajadores"
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
         Index           =   29
         Left            =   120
         TabIndex        =   130
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   6
         Left            =   900
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   71
         Left            =   150
         TabIndex        =   129
         Top             =   2640
         Width           =   675
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   5
         Left            =   900
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   70
         Left            =   150
         TabIndex        =   127
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         Index           =   28
         Left            =   120
         TabIndex        =   125
         Top             =   960
         Width           =   690
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   28
         Left            =   3840
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   69
         Left            =   3150
         TabIndex        =   124
         Top             =   1365
         Width           =   675
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   27
         Left            =   960
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   68
         Left            =   270
         TabIndex        =   123
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Listado Llamadas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   20
         Left            =   1320
         TabIndex        =   116
         Top             =   360
         Width           =   2925
      End
   End
   Begin VB.Frame FrameOtrasOfertas 
      Height          =   4455
      Left            =   120
      TabIndex        =   109
      Top             =   120
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton cmdAceptarOfertas 
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
         Left            =   5520
         TabIndex        =   113
         Top             =   3960
         Width           =   1215
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
         Index           =   21
         Left            =   6960
         TabIndex        =   112
         Top             =   3960
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   2775
         Left            =   240
         TabIndex        =   111
         Top             =   960
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Num. ofer"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "F. Entrega"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Di"
            Object.Width           =   6703
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmListado2.frx":0294
         ToolTipText     =   "Puntear al haber"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmListado2.frx":03DE
         ToolTipText     =   "Quitar al haber"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
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
         Index           =   27
         Left            =   240
         TabIndex        =   114
         Top             =   720
         Width           =   7725
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Otras ofertas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   19
         Left            =   2640
         TabIndex        =   110
         Top             =   240
         Width           =   2925
      End
   End
   Begin VB.Frame FrListadoReparaciones 
      Height          =   4335
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6495
      Begin VB.OptionButton optReparaciones 
         Caption         =   "Fecha entrada"
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
         Left            =   240
         TabIndex        =   6
         Top             =   3720
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optReparaciones 
         Caption         =   "Fecha albarán"
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
         Left            =   2070
         TabIndex        =   7
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CommandButton cmdReparaEfect 
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
         Left            =   4080
         TabIndex        =   8
         Top             =   3720
         Width           =   975
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
         Left            =   4380
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3000
         Width           =   1350
      End
      Begin VB.TextBox txtDpto 
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
         Left            =   1740
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtDescDpto 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2280
         Width           =   3615
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
         Index           =   1
         Left            =   5280
         TabIndex        =   15
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtDescDpto 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtDpto 
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
         Left            =   1740
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1920
         Width           =   855
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
         Left            =   1740
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3000
         Width           =   1350
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   1740
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtDescClie 
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
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   1740
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtDescClie 
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
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   840
         Width           =   3375
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4080
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   2760
         Width           =   630
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   270
         TabIndex        =   28
         Top             =   570
         Width           =   765
      End
      Begin VB.Label Label4 
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
         Index           =   1
         Left            =   3420
         TabIndex        =   27
         Top             =   3000
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Index           =   1
         Left            =   660
         TabIndex        =   26
         Top             =   3000
         Width           =   585
      End
      Begin VB.Label Label4 
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
         Index           =   0
         Left            =   660
         TabIndex        =   25
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Left            =   660
         TabIndex        =   24
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label4 
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
         Index           =   23
         Left            =   660
         TabIndex        =   23
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Index           =   42
         Left            =   660
         TabIndex        =   22
         Top             =   840
         Width           =   585
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   1
         Left            =   1440
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   765
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   0
         Left            =   1440
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1440
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Reparaciones efectuadas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   150
         Width           =   5895
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   1
         Left            =   1440
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   0
         Left            =   1440
         Top             =   840
         Width           =   240
      End
   End
   Begin VB.Frame FrameVEntasAgente 
      Height          =   3855
      Left            =   0
      TabIndex        =   80
      Top             =   120
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox chkAgentes 
         Caption         =   "Presupuestos"
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
         Left            =   3240
         TabIndex        =   86
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkAgentes 
         Caption         =   "Facturas"
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
         Left            =   1560
         TabIndex        =   85
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgentes 
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
         Left            =   3570
         TabIndex        =   87
         Top             =   3360
         Width           =   975
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
         Index           =   16
         Left            =   4650
         TabIndex        =   88
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtDescAgente 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   96
         Text            =   "Text1"
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtAgente 
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
         Left            =   1560
         TabIndex        =   84
         Text            =   "Text1"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtDescAgente 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   93
         Text            =   "Text1"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtAgente 
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
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   1800
         Width           =   615
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
         Index           =   25
         Left            =   4320
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   1020
         Width           =   1350
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
         Index           =   24
         Left            =   1560
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   1020
         Width           =   1350
      End
      Begin VB.Label Label3 
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
         Index           =   57
         Left            =   480
         TabIndex        =   97
         Top             =   2160
         Width           =   675
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   0
         Left            =   1200
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   56
         Left            =   480
         TabIndex        =   95
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
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
         Index           =   24
         Left            =   120
         TabIndex        =   94
         Top             =   1440
         Width           =   765
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   1
         Left            =   1200
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   25
         Left            =   4050
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   55
         Left            =   3390
         TabIndex        =   92
         Top             =   1065
         Width           =   675
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Index           =   23
         Left            =   120
         TabIndex        =   91
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
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
         Index           =   54
         Left            =   510
         TabIndex        =   90
         Top             =   1065
         Width           =   675
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   1200
         Top             =   1065
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Ventas por agente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   14
         Left            =   540
         TabIndex        =   89
         Top             =   240
         Width           =   4995
      End
   End
   Begin VB.Frame FrameAlbaProv 
      Height          =   4095
      Left            =   0
      TabIndex        =   159
      Top             =   0
      Width           =   6615
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
         Index           =   10
         Left            =   5400
         TabIndex        =   169
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmdAlbaranProv 
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
         Left            =   4320
         TabIndex        =   168
         Top             =   3480
         Width           =   975
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
         Index           =   18
         Left            =   1920
         TabIndex        =   167
         Text            =   "Text1"
         Top             =   840
         Width           =   1350
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
         Index           =   19
         Left            =   4920
         TabIndex        =   166
         Text            =   "Text1"
         Top             =   840
         Width           =   1350
      End
      Begin VB.TextBox txtNumAlbar 
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
         Left            =   1920
         TabIndex        =   165
         Text            =   "Text1"
         Top             =   1635
         Width           =   1350
      End
      Begin VB.TextBox txtNumAlbar 
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
         Left            =   4920
         TabIndex        =   164
         Text            =   "Text1"
         Top             =   1680
         Width           =   1350
      End
      Begin VB.TextBox txtDescProve 
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
         Index           =   8
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   163
         Text            =   "Text5"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
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
         TabIndex        =   162
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
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
         Index           =   9
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   161
         Text            =   "Text5"
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
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
         Left            =   1260
         TabIndex        =   160
         Text            =   "Text1"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albarán"
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
         Index           =   18
         Left            =   120
         TabIndex        =   179
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label3 
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
         Index           =   39
         Left            =   960
         TabIndex        =   178
         Top             =   885
         Width           =   645
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   1650
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   4650
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   40
         Left            =   3960
         TabIndex        =   177
         Top             =   885
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Index           =   41
         Left            =   960
         TabIndex        =   176
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Albarán"
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
         Index           =   19
         Left            =   120
         TabIndex        =   175
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label Label3 
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
         Index           =   43
         Left            =   3960
         TabIndex        =   174
         Top             =   1725
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Index           =   11
         Left            =   120
         TabIndex        =   173
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label Label3 
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
         Index           =   44
         Left            =   240
         TabIndex        =   172
         Top             =   2400
         Width           =   705
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   8
         Left            =   960
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Imprimir albarán proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   9
         Left            =   720
         TabIndex        =   171
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label Label3 
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
         Index           =   45
         Left            =   240
         TabIndex        =   170
         Top             =   2760
         Width           =   705
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   9
         Left            =   960
         Top             =   2790
         Width           =   240
      End
   End
   Begin VB.Frame FrameMultibase 
      Height          =   5295
      Left            =   120
      TabIndex        =   37
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmdMultibase2 
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
         Left            =   3240
         TabIndex        =   79
         Top             =   4800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame FrameTablas 
         Height          =   3375
         Left            =   120
         TabIndex        =   74
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
         Begin VB.ComboBox cboCampos 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   1560
            Width           =   2895
         End
         Begin VB.ComboBox cboTablas 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label7 
            Caption         =   "TABLAS"
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
            TabIndex        =   77
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label6 
            Caption         =   "TABLAS"
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
            TabIndex        =   75
            Top             =   480
            Width           =   3255
         End
      End
      Begin VB.CheckBox chkRoot 
         Alignment       =   1  'Right Justify
         Caption         =   "Tablas"
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
         Left            =   4440
         TabIndex        =   73
         Top             =   4440
         Width           =   975
      End
      Begin VB.ListBox lstMultibase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3300
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   41
         Top             =   960
         Width           =   5295
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
         Index           =   3
         Left            =   4440
         TabIndex        =   39
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdMultibase 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   38
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label lblMultibase 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   4560
         Width           =   2895
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Revisar caracteres especiales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmListado2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Integer
    '1      .- Listado reparaciones efectuadas
    '2      .- Reparaciones tecnico
    
    '3      .- Revision carcteres multibase
    '4      .- Listado recargas telefonia movil
    '5      .- Facturacion de recargas
    
    '6      .- Listado de TRAZA por codprove en ventas.   ENERO 2008
    
    '       LIQUIDACION PROVEEDORES. Socios tipo TERRASANA
    '7      .- Cambio precio articulos
    '8      .- Generar facturas
    '9      .- Imprimir facturas proveedores (socios)
    '10     .-   "      ALBARANES   "           "
    
    
    '13     .- Generacion y facturacion de tickets agrupados
    '14     .- Listado del punto anterior
    
    '15     .- Listado trazabilidad albaranes
        
    '16     .- Ventas x agentes
    
    '17     .- Listado trabajadores . NO HACE DESDE HASTA
    
    '18     .- Cambio de proveedor en albaranes. Solicita el codprove
    
    '19     .- Cerrar aviso. Datos para crear albaran
    
    '20     .- Listado plantillas ofertas
    
    
    '21     .- Seleccionar otras ofertas del cliente
    '22     .- Listado llamadas
    
    '23     .- Listado situacion albaranes
    '24     .- Modificar expediente y legal en frecuencias
    
    '25     .- Datos para facturacion de cliente
    
Private IndiceImg As Integer
Private WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmMtoArticulos As frmAlmArticulos
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmPr As frmComProveedores
Attribute frmPr.VB_VarHelpID = -1
Private WithEvents frmBaPr As frmFacBancosPropios
Attribute frmBaPr.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmAg As frmFacAgentesCom
Attribute frmAg.VB_VarHelpID = -1


Private PrimeraVez As Boolean




'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
'nuevo Febrero 2010
Private cadPDFrpt As String 'Nombre del informe a enviar por email
Private vImprimedirecto As Boolean '
'-----------------------------------





'Variables comunes a todos os botones aceptar
Dim cadFrom As String
Dim campo As String, devuelve As String
Dim Codigo  As String
Dim ImpTot As Currency
Dim ImpTeo As Currency
Dim miSQL As String

Private Cadena_frmB As String
Private cadImpresion As String  'Facturacion

Dim kCampo As Integer

Private Sub cboTablas_Click()
    cboCampos.Clear
    If cboTablas.ListIndex < 0 Then Exit Sub
    CargarCamposTabla
End Sub



Private Sub chkAgentes_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkFacturPorv_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkRoot_Click()
    Me.FrameTablas.visible = Me.chkRoot.Value = 1
    cmdMultibase2.visible = Me.chkRoot.Value = 1
    cmdMultibase.visible = Me.chkRoot.Value <> 1
    If Me.chkRoot.Value = 1 Then
        If Me.cboTablas.ListCount = 0 Then
            Screen.MousePointer = vbHourglass
            Me.lblMultibase.Caption = "Cargando datos"
            Me.lblMultibase.Refresh
            
            CargaTablasCambio
            
            Screen.MousePointer = vbDefault
            Me.lblMultibase.Caption = ""
        End If
    End If
End Sub

Private Sub cmbRecargaMov_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cmdAceptarOfertas_Click()
    miSQL = ""
    For NumRegElim = 1 To lw1.ListItems.Count
        If lw1.ListItems(NumRegElim).Checked Then miSQL = miSQL & ", " & lw1.ListItems(NumRegElim).Text
    Next NumRegElim
    If miSQL = "" Then
        MsgBox "Selecciona alguna oferta", vbExclamation
        Exit Sub
    End If
    CadenaDesdeOtroForm = miSQL
    Unload Me
End Sub

Private Sub cmdAgentes_Click()
    
    If Me.chkAgentes(0).Value = 0 And Me.chkAgentes(1).Value = 0 Then
        MsgBox "Seleccione facturas", vbExclamation
        Exit Sub
    End If
    
    InicializarVbles
    'El nombre de la empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'Si lleva articulo de portes, ese NO va a las lineas
    If vParamAplic.ArtPortes <> "" Then
        cadSelect = "{slifac.codartic} <> '" & vParamAplic.ArtPortes & "'"
    Else
        cadSelect = " 1 = 1"  'Para que no de error contar registros
    End If
    cadFormula = cadSelect
   
    If txtFecha(24).Text <> "" Or txtFecha(25).Text <> "" Then
        devuelve = "vFechas=""Fecha: "
        campo = "{scafac.fecfactu}"
        If Not PonerDesdeHasta(campo, "F", 24, 25, devuelve) Then Exit Sub
        
    End If
    
    If txtAgente(0).Text <> "" Or txtAgente(1).Text <> "" Then
        devuelve = "vAgentes=""Agente: "
        campo = "{scafac.codagent}"
        If Not PonerDesdeHasta(campo, "AGT", 0, 1, devuelve) Then Exit Sub
    End If
     
    'JULIO 2009
    'Las FRT no entran en el listado
    miSQL = " ({scafac.codtipom} <> 'FRT' and {scafac.codtipom} <> 'FRC') " 'NO ponemos las rectificativas
    If Me.chkAgentes(0).Value = 1 And Me.chkAgentes(1).Value = 1 Then
        'NO poenmos nada al select ya que pide las dos
            
    Else
        If Me.chkAgentes(0).Value = 1 Then
            miSQL = miSQL & " AND {scafac.codtipom} <> 'FAZ'"  'NO ponemos las "B"
        Else
            miSQL = miSQL & " AND {scafac.codtipom} = 'FAZ'"    'SOLO las B
        End If
    End If
    AnyadirAFormula cadFormula, miSQL
    cadSelect = cadSelect & " AND " & miSQL
    miSQL = ""
    'Montamos el select para los registros
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    campo = "scafac.codtipom=slifac.codtipom and scafac.numfactu=slifac.numfactu and scafac.fecfactu=slifac.fecfactu AND "
    campo = "scafac,slifac WHERE " & campo & cadSelect
    
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay facturas con estos valores", vbExclamation
        Exit Sub
    End If
    
    LlamarImprimir False


    
End Sub

Private Sub cmdAlbaranProv_Click()

    InicializarVbles
    
    'Albaran socio
    If Not PonerParamRPT(27, cadParam, numParam, cadNomRPT, vImprimedirecto, cadPDFrpt) Then Exit Sub
    
    
    
    cadSelect = "{sprove.tipprove}>=0"   'Antes ponia un tres: Estos proveedores son los REA o estimacion directa que luego
    cadFormula = "(" & cadSelect & ")"
    If txtFecha(18).Text <> "" Or txtFecha(19).Text <> "" Then
        campo = "{scaalp.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 18, 19, devuelve) Then Exit Sub
    End If
    
    If txtCodProve(8).Text <> "" Or txtCodProve(9).Text <> "" Then
        campo = "{scaalp.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 8, 9, devuelve) Then Exit Sub
    End If
     
    If txtNumAlbar(4).Text <> "" Or txtNumAlbar(5).Text <> "" Then
        campo = "{scaalp.numalbar}"
        If Not PonerDesdeHasta(campo, "ALP", 4, 5, "") Then Exit Sub
    End If
    
    
    'Montamos el select para los registros
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    campo = "scaalp,sprove WHERE scaalp.codprove=sprove.codprove AND " & cadSelect
    
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay facturas con estos valores", vbExclamation
        Exit Sub
    End If
    
    'FALTA###
    LlamarImprimir False








    frmImprimir.Opcion = 2010
    frmImprimir.Show vbModal
End Sub


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub cmdEstadisticaReparacionTecnico_Click()

    If Me.txtTrab(0).Text = "" Then
        MsgBox "Seleccione un técnico", vbExclamation
        Exit Sub
    End If
    cadSelect = "schrep.codtrab2 = " & txtTrab(0).Text

    'Ya tenemos el tecnico. Miramos las fechas
    If txtFecha(2).Text <> "" Or txtFecha(3).Text <> "" Then
        'Parametro Desde/Hasta Cliente
        
        'Marzo 2010
        'ANtes
        'campo = "schrep.fecrepar"
        campo = "schrep.fechaalb"
        
        
        If Not PonerDesdeHasta(campo, "F", 2, 3, devuelve) Then Exit Sub
        'Aqui lo añadiremos a  cadparam
        
    End If
    
    
    
    
    Screen.MousePointer = vbHourglass
   
    NumRegElim = 0
    Set miRsAux = New ADODB.Recordset
    'Aqui iremos grabanod los datos.
    'EstadisticaReparacionTecnico
    
    
    EstadisticaReparacionTecnicoNueva
    
    Set miRsAux = Nothing
    Label3(63).Caption = ""
    Screen.MousePointer = vbDefault
    
    If NumRegElim = 0 Then
        MsgBox "Ningun dato a mostrar", vbExclamation
        Exit Sub
    End If
    
    
    'Llegados aqui imprimiremos los registros
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    cadParam = cadParam & "pDHFecha= "" Técnico: " & txtTrab(0).Text & " - " & Me.txtDescTra(0).Text & """|"
    numParam = 2
    campo = ""
    If txtFecha(2).Text <> "" Then campo = "     Desde " & txtFecha(2).Text
    If txtFecha(3).Text <> "" Then campo = campo & "      Hasta " & txtFecha(3).Text
    If campo <> "" Then
        numParam = 3
        campo = "pDHCliente= """ & Trim(campo) & """|"
        cadParam = cadParam & campo
    End If
    cadFormula = "{tmpnlotes.codusu}=" & vUsu.Codigo

    cadNomRPT = "rRepEstadisticaTec.rpt"
    conSubRPT = False
    LlamarImprimir False
    
End Sub

Private Sub cmdFacturarCli_Click()
    If txtFecha(31).Text = "" Or txtBancoPr(2).Text = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
        
    If MsgBox("¿Seguro que desa continuar con la facturación?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    CadenaDesdeOtroForm = txtFecha(31).Text & "|" & txtBancoPr(2).Text & "|" & chkFacturarCliente.Value & "|"
    Unload Me
End Sub
Private Sub cmdListTrabja_Click()

    InicializarVbles
    
    If Not PonerParamRPT(35, cadParam, numParam, cadNomRPT, vImprimedirecto, cadPDFrpt) Then Exit Sub
     
    If txtTrab(3).Text <> "" Or txtTrab(4).Text <> "" Then
        campo = "{straba.codtraba}"
        If Not PonerDesdeHasta(campo, "TRA", 3, 4, devuelve) Then Exit Sub
    End If
        
    LlamarImprimir True
End Sub

Private Sub cmdLlamadas_Click()


    InicializarVbles
    
    
    If Not PonerParamRPT(41, cadParam, numParam, cadNomRPT, vImprimedirecto, cadPDFrpt) Then Exit Sub
    
    'El nombre de la empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    cadFormula = cadSelect
   
    If txtFecha(27).Text <> "" Or txtFecha(28).Text <> "" Then
        devuelve = "pDHFecha=""Fecha: "
        campo = "{sllama.feholla}"
        If Not PonerDesdeHasta(campo, "F", 27, 28, devuelve) Then Exit Sub
        
    End If
    
    If txtTrab(5).Text <> "" Or txtTrab(6).Text <> "" Then
        devuelve = "pdhTra=""Trabajador: "
        campo = "{sllama.codtraba}"
        If Not PonerDesdeHasta(campo, "TRA", 5, 6, devuelve) Then Exit Sub
        
    End If
    
    
    '
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
    campo = "sllama "
    If cadSelect <> "" Then campo = campo & " WHERE " & cadSelect
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay registros con estos valores", vbExclamation
        Exit Sub
    End If
    
    LlamarImprimir False





End Sub

Private Sub cmdMultibase_Click()
    'Revision caracteres multibase
    numParam = 0
    For NumRegElim = 1 To Me.lstMultibase.ListCount
        If Me.lstMultibase.Selected(CInt(NumRegElim - 1)) Then numParam = numParam + 1
    Next
    
    If numParam = 0 Then
        MsgBox "Seleccion alguna tabla para cambiar", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Este proceso puede durar mucho tiempo." & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    Me.Tag = ""
    Set miRsAux = New ADODB.Recordset
    For numParam = 0 To Me.lstMultibase.ListCount - 1
        If Me.lstMultibase.Selected(CInt(numParam)) Then HacerCambiosMultibase CInt(numParam + 1)
    Next
    Me.lblMultibase.Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    If Me.Tag <> "" Then
        Codigo = "Se han realizado los siguientes cambios:" & vbCrLf & vbCrLf & Me.Tag
        Me.Tag = ""
    Else
        Codigo = "Proceso finalizado. No se efectuaron cambios"
    End If
    MsgBox Codigo, vbInformation
End Sub

Private Sub cmdMultibase2_Click()
    If cboTablas.ListIndex < 0 Then Exit Sub
    
    If MsgBox("Va a buscar en los campos seleccionados. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    UpdatearTablaRoot
    
    cadFrom = ""
    Me.lblMultibase.Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdReparaEfect_Click()
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    Codigo = "schrep"
    devuelve = ""
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCliente(0).Text <> "" Or txtCliente(1).Text <> "" Then
        campo = "{" & Codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "CLI", 0, 1, devuelve) Then Exit Sub
    End If

    
    'Cadena para seleccion DEPARTAMENTO
    '--------------------------------------------
    If txtDpto(0).Text <> "" Or txtDpto(1).Text <> "" Then
        campo = "{" & Codigo & ".coddirec}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHDpto=""Dpto: "
        If Not PonerDesdeHasta(campo, "DPT", 0, 1, devuelve) Then Exit Sub
    End If
    
    
    'Este trozo lo hace siempre
    If Me.optReparaciones(0).Value Then
        devuelve = "entrada"
        campo = "entre"
    Else
        devuelve = "reparación"
        campo = "repar"
        'AHora Marzo 2010
        campo = "haalb"  'fechaalb
    End If
    campo = "{" & Codigo & ".fec" & campo & "}"
    cadParam = cadParam & "pOrden=" & campo & "|"
    numParam = numParam + 1
    
    If txtFecha(0).Text <> "" Or txtFecha(1).Text <> "" Then
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHFecha=""Fecha " & devuelve & ": "
        If Not PonerDesdeHasta(campo, "F", 0, 1, devuelve) Then Exit Sub
    End If
    
    cadFrom = "schrep"
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    Screen.MousePointer = vbHourglass
    'Prepararo los datos
    Codigo = "DELETE from tmpnlotes where codusu = " & vUsu.Codigo
    conn.Execute Codigo
    CargaImporteRealReparaciones
    
    
    'MOSTRAMOS EL INFORME
    'Añadir el nombre de la Empresa como parametro
    If cadFormula <> "" Then cadFormula = cadFormula & " AND "
    cadFormula = cadFormula & "(isnull({tmpnlotes.codusu}) or {tmpnlotes.codusu}=1000)"
    
    conSubRPT = False
    LlamarImprimir False
    Screen.MousePointer = vbDefault
End Sub
Private Sub LlamarImprimir(PongoNombrePDF As Boolean)
    
    With frmImprimir
    
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2000 + Opcion   '2000 mas la opcion de entrada
        .NombrePDF = ""
        If PongoNombrePDF Then .NombrePDF = cadPDFrpt
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub





Private Sub cmdSituAlbaran_Click()
Dim i As Integer
    InicializarVbles
    cadTitulo = ""
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    'Comporobar si hay registros
    If txtFecha(29).Text <> "" Or txtFecha(30).Text <> "" Then
        campo = "{scaalb.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 29, 30, "") Then Exit Sub
        If txtFecha(29).Text <> "" Then cadTitulo = cadTitulo & "desde " & txtFecha(29).Text
        If txtFecha(30).Text <> "" Then cadTitulo = cadTitulo & " hasta " & txtFecha(30).Text
        If cadTitulo <> "" Then cadTitulo = "Fechas: " & cadTitulo
    End If
    
    If txtCliente(5).Text <> "" Or txtCliente(6).Text <> "" Then
        campo = "{scaalb.codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHClien=""Cliente: "
        If Not PonerDesdeHasta(campo, "CLI", 5, 6, devuelve) Then Exit Sub
    End If
    
    devuelve = ""
    miSQL = ""
    IndiceImg = 0
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            IndiceImg = IndiceImg + 1
            NumRegElim = InStrRev(List1.List(i), "(")
            If NumRegElim = 0 Then
                MsgBox "No se ha encontrado (", vbExclamation
                Exit Sub
            End If
            campo = Mid(List1.List(i), NumRegElim + 1, 3)
            miSQL = miSQL & " - " & campo
            devuelve = devuelve & ", '" & campo & "'"
            
        End If
    Next i
    If devuelve = "" Then
        MsgBox "Seleccione algun tipo de albarán", vbExclamation
        Exit Sub
    End If
    
    If IndiceImg <> List1.ListCount Then
        If cadTitulo <> "" Then cadTitulo = cadTitulo & "        "
        miSQL = Mid(miSQL, 3)
        cadTitulo = cadTitulo & "Tipo albaran: " & miSQL
        
    End If
    miSQL = cadTitulo
    cadParam = cadParam & "pDHFecha=""" & miSQL & """|"
    
    devuelve = Mid(devuelve, 2)
    If cadSelect <> "" Then cadSelect = cadSelect & " AND "
    cadSelect = cadSelect & " (codtipom IN (" & devuelve & "))"
    If cadFormula <> "" Then cadFormula = cadFormula & " AND "
    cadFormula = cadFormula & "( {scaalb.codtipom} IN [" & devuelve & "])"
    
    'Pongo en campo el select
    
    If Not HayRegParaInforme("scaalb", cadSelect) Then Exit Sub
    
    
    
    
    cadNomRPT = "rFacSituacionAlb.rpt"
    LlamarImprimir False

End Sub

Private Sub cmdTraza_Click()
    Screen.MousePointer = vbHourglass
    HacerInformeTrazabilidad
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdVentaxProv_Click()
Dim Cad As String
    
    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
     If txtCliente(3).Text <> "" Or txtCliente(4).Text <> "" Then
        campo = "{scafac.codclien}"
        'Parametro Desde/Hasta Cliente
        Cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "CLI", 3, 4, Cad) Then Exit Sub
    End If
   
    
    
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtFecha(9).Text <> "" Or txtFecha(10).Text <> "" Then
        campo = "{scafac.fecfactu}"
        Cad = "pDHFecha=""Fecha Fact.: "
        If Not PonerDesdeHasta(campo, "F", 9, 10, Cad) Then Exit Sub
    End If
    
    
    
    'Cadena para seleccion Desde y Hasta ARTICULO
    '--------------------------------------------
    If txtArticulo(1).Text <> "" Or txtArticulo(2).Text <> "" Then
        campo = "{slifac.codartic}"
        Cad = "pDHDpto=""Artículo: "
        If Not PonerDesdeHasta(campo, "ART", 1, 2, Cad) Then Exit Sub
    End If
    
    
    'Cadena para seleccion Desde y Hasta PROVEEDOR
    '---------------------------------------------
    If txtCodProve(0).Text <> "" Or txtCodProve(1).Text <> "" Then
        campo = "{slifac.codprovex}"
        Cad = "pDHPro=""Proveedor: "
        If Not PonerDesdeHasta(campo, "PRO", 0, 1, Cad) Then Exit Sub
    End If

     
    
    'Pongo en campo el select
    Codigo = " scafac.codtipom=slifac.codtipom "
    Codigo = " scafac.fecfactu = slifac.fecfactu AND scafac.numfactu=slifac.numfactu AND " & Codigo
    Cad = "scafac,slifac"
    If cadSelect <> "" Then Codigo = Codigo & " AND " & cadSelect
    campo = Codigo
    If Not HayRegParaInforme(Cad, Codigo) Then Exit Sub
    
    
    cadNomRPT = "rvtaxcodprove.rpt"
    LlamarImprimir False
    

End Sub

Private Sub Command1_Click()
    If txtCodProve(12).Text = "" Or Me.txtDescProve(12).Text = "" Then
        MsgBox "Seleccione el proveedor", vbExclamation
        Exit Sub
    End If
    
    
    
    
     'Compruebo si esta bloqueado el proveedor
    miSQL = DevuelveDesdeBDNew(conAri, "sprove", "codsitua", "codprove", txtCodProve(12).Text, "N")
    
    If Val(miSQL) > 0 Then
            devuelve = "tipositu"
            miSQL = DevuelveDesdeBDNew(conAri, "ssitua", "nomsitua", "codsitua", miSQL, "N", devuelve)
            
            
            If devuelve = "1" Then 'Cliente Bloqueado por Situación Especial.
                MsgBox UCase("Proveedor bloqueado por: ") & miSQL & "-" & devuelve, vbInformation, "Situación Especial del proveedor."
            Else
                MsgBox miSQL, vbInformation, "Situación Especial del proveedor."
            End If
            Exit Sub
    End If
    
    
    
    CadenaDesdeOtroForm = txtCodProve(12).Text
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 0
            PonerFoco txtCliente(0)
        Case 1
            PonerFoco txtTrab(0)
        Case 6 To 10
            'En ambos listados lo primero es una fecha
            If Opcion = 6 Then
                numParam = 9
            ElseIf Opcion = 7 Then
                numParam = 11
            ElseIf Opcion = 8 Then
                numParam = 13  'liquidacion factura sprov
                txtFecha(17).Text = Format(Now, "dd/mm/yyyy")
            Else
                numParam = 8 + Opcion 'impresion facturas  index:17 y 18
            End If
            PonerFoco txtFecha(CInt(numParam))
        
        Case 13
            cadParam = ""
            'Poner el nombre del trabajador que esta conectado
            Me.txtTrab(2).Text = PonerTrabajadorConectado(cadParam)
            Me.txtDescTra(2).Text = cadParam
        
        Case 21
            If vParamAplic.TipoDtos Then
                lw1.ColumnHeaders(4).Text = "Departamento"
            Else
                lw1.ColumnHeaders(4).Text = "Direccion"
            End If
            CargarOtrasOfertas
'        Case 24
'            texto(5).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
'            miSQL = RecuperaValor(CadenaDesdeOtroForm, 2)
'            Me.chkFrecu.Value = Abs(miSQL)
'            CadenaDesdeOtroForm = "" 'Para que no devulev nada
        Case 25
            PonerFoco txtFecha(31)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim IndiceCancel As Integer

    Me.Icon = frmPpal.Icon
    
    PrimeraVez = True
    limpiar Me
    
    For kCampo = 0 To 1
        Me.imgCliente(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 3 To 6
        Me.imgCliente(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 0 To Me.imgDpto.Count - 1
        Me.imgDpto(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 0 To 1
        Me.imgProveedor(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 8 To 9
        Me.imgProveedor(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    Me.imgTecnico(0).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    For kCampo = 3 To 6
        Me.imgTecnico(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 0 To 1
        Me.imgAgente(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 1 To 2
        Me.imgArticulo(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    
    Me.imgBancoPr(2).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    
    
    
    For kCampo = 0 To 3
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    For kCampo = 9 To 10
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    
    For kCampo = 18 To 19
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    For kCampo = 24 To 25
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    For kCampo = 27 To 31
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    
    
    
    
    
    
    FrListadoReparaciones.visible = False
    FrEstadisticasReparacionTecnico.visible = False
    FrameMultibase.visible = False
    FrProveedorxVenta.visible = False
    Me.FrameVEntasAgente.visible = False
    FrameListTrabajadores.visible = False
    FrameOtrasOfertas.visible = False
    FrameLlamadas.visible = False
    Me.FrameSituAlbaranes.visible = False
    FrameFacturarCliente.visible = False
    Me.FrameAlbaProv.visible = False
    FrameVEntasAgente.visible = False
    
    Caption = "Listado"
    IndiceCancel = Opcion
    Select Case Opcion
    Case 1
        'Listado reparaciones efectuadas
        PonerFrameVisible FrListadoReparaciones, H, W
        PonerLabelDptoDireccion Me.lblDpto(0)
        
        
        
    Case 2
        PonerFrameVisible Me.FrEstadisticasReparacionTecnico, H, W
        Label3(63).Caption = ""
        
    Case 3
        Caption = "MULTIBASE"
        PonerFrameVisible Me.FrameMultibase, H, W
        CargaListMultibase
        chkRoot.visible = vUsu.Nivel = 0
    Case 4
        'Informe recarga movil
       
    Case 5
        'Facturacion recargas moviles
        
    Case 6
        'Ventas por codprove
        'TRAZA enero 2008
        PonerFrameVisible FrProveedorxVenta, H, W
        
    Case 7
        'Queda Libre
    
    Case 8
        'Queda Libre
    
    Case 9
        'Queda Libre
    
    Case 10
        'Queda Libre
         PonerFrameVisible FrameAlbaProv, H, W
        
        'CadenaDesdeOtroForm
         
        Me.txtNumAlbar(4).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        Me.txtNumAlbar(5).Text = Me.txtNumAlbar(4).Text
         
        Me.txtFecha(18).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
        Me.txtFecha(19).Text = Me.txtFecha(18).Text
        
        Me.txtCodProve(8).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
        Me.txtCodProve(9).Text = Me.txtCodProve(8).Text
        
        Me.txtDescProve(8).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
        Me.txtDescProve(9).Text = Me.txtDescProve(8).Text
        
        CadenaDesdeOtroForm = ""
    Case 13
        'Queda Libre
    
    Case 14
        'Queda Libre
        
        
    Case 15
        'Queda Libre
         
       
    Case 16
        PonerFrameVisible FrameVEntasAgente, H, W
        
    Case 17
        PonerFrameVisible FrameListTrabajadores, H, W
        
        
    Case 18
        ' Libre
    
    Case 19
    Case 20
        
'        PonerFrameVisible Me.FrameListadoPlantillas, H, W
        
    Case 21
        Caption = "Seleccionar"
        'optras ofertas del cliente
        PonerFrameVisible Me.FrameOtrasOfertas, H, W
    Case 22
        
        PonerFrameVisible Me.FrameLlamadas, H, W
    Case 23
        PonerFrameVisible FrameSituAlbaranes, H, W
        CargaListMov
        
        
    Case 24
'        PonerFrameVisible FrameFrecuencia, H, W
        
    Case 25
        PonerFrameVisible FrameFacturarCliente, H, W
        txtFecha(31).Text = Format(Now, "dd/mm/yyyy")
    End Select
    Me.Height = H + 150
    Me.Width = W
    Me.cmdCancel(1).Cancel = True
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Cadena_frmB = CadenaDevuelta
    
End Sub

Private Sub frmBaPr_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtBancoPr(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescBancoPr(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtFecha(IndiceImg).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtCliente(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescClie(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtArticulo(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescArticulo(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmPr_DatoSeleccionado(CadenaSeleccion As String)
    txtCodProve(IndiceImg) = RecuperaValor(CadenaSeleccion, 1)
    txtDescProve(IndiceImg) = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    txtTrab(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescTra(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub imgAgente_Click(Index As Integer)
    IndiceImg = Index
    Set frmAg = New frmFacAgentesCom
    frmAg.DatosADevolverBusqueda = "0|1|"
    frmAg.Show vbModal
    Set frmAg = Nothing
End Sub

Private Sub imgArticulo_Click(Index As Integer)
    IndiceImg = Index
    Set frmMtoArticulos = New frmAlmArticulos
    frmMtoArticulos.DatosADevolverBusqueda2 = "@1@" 'Abrimos en Modo Busqueda
    frmMtoArticulos.Show vbModal
    Set frmMtoArticulos = Nothing
End Sub

Private Sub imgBancoPr_Click(Index As Integer)
    IndiceImg = Index
    Set frmBaPr = New frmFacBancosPropios
    frmBaPr.DatosADevolverBusqueda = "1" 'Abrimos en Modo Busqueda
    frmBaPr.Show vbModal
    Set frmBaPr = Nothing
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim i As Integer

    If Index < 2 Then
        'Seleecionar otras ofertas
        For i = 1 To Me.lw1.ListItems.Count
            lw1.ListItems(i).Checked = Index = 1
        Next i
    ElseIf Index < 4 Then
        'Seleccionar Tipos albaran para listado situacion labaranes
        For i = 0 To List1.ListCount - 1
            List1.Selected(i) = Index = 2
        Next i
    End If
    
End Sub

Private Sub imgCliente_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    IndiceImg = Index
    Set frmCli = New frmFacClientes
    frmCli.DatosADevolverBusqueda = "0|1|"
    frmCli.Show vbModal
    Set frmCli = Nothing
End Sub

Private Sub imgDpto_Click(Index As Integer)
    If Index < 2 Then
        'DPTO
        IndiceImg = Index
        Cadena_frmB = ""
        If txtCliente(0).Text <> "" And txtCliente(0).Text = txtCliente(1).Text And txtDescClie(0).Text <> "" Then
            'OK
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vTitulo = Me.lblDpto(0).Caption & " " & txtCliente(0).Text & " - " & txtDescClie(0).Text
            campo = "Cod.|sdirec|coddirec|N||20·"
            campo = campo & "Desc.|sdirec|nomdirec|T||40·"
            frmB.vCampos = campo
            frmB.vCargaFrame = False
            frmB.vDevuelve = "0|1|"
            frmB.vselElem = 1
            frmB.vConexionGrid = 1  'ODBC Aritaxi
            frmB.vTabla = "sdirec"
            frmB.vSQL = "codclien = " & txtCliente(0).Text
            frmB.Show vbModal
            Set frmB = Nothing
            Screen.MousePointer = vbDefault
            
            If Cadena_frmB <> "" Then
                txtDpto(IndiceImg).Text = RecuperaValor(Cadena_frmB, 1)
                txtDescDpto(IndiceImg) = RecuperaValor(Cadena_frmB, 2)
            End If
            
        Else
            MsgBox "Para poner el departamento cliente debe y el hasta  debe ser el mismo", vbExclamation
        End If
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
   IndiceImg = Index
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(IndiceImg).Text <> "" Then
        If IsDate(txtFecha(IndiceImg).Text) Then frmC.Fecha = CDate(txtFecha(IndiceImg).Text)
    End If
   frmC.Show vbModal
   Set frmC = Nothing
    
End Sub


Private Sub imgForPa_Click(Index As Integer)
    IndiceImg = Index
    Set frmFP = New frmFacFormasPago
    frmFP.DatosADevolverBusqueda = "0|1|"
    frmFP.Show vbModal
    Set frmFP = Nothing
End Sub

Private Sub imgProveedor_Click(Index As Integer)
    IndiceImg = Index
    Set frmPr = New frmComProveedores
    frmPr.DatosADevolverBusqueda = "0|1|"
    frmPr.Show vbModal
    Set frmPr = Nothing
End Sub

Private Sub imgTecnico_Click(Index As Integer)
    IndiceImg = Index
    If Index < 3 Then
        Set frmT = New frmAdmTrabajadores
        frmT.DatosADevolverBusqueda = "0|1|"
        frmT.Show vbModal
        Set frmT = Nothing

    Else
        'Listado trabajadores
            Cadena_frmB = ""
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vTitulo = "Trabajadores"
            campo = "Codigo|straba|codtraba|N||20·"
            campo = campo & "Nombre|straba|nomtraba|T||40·"
            campo = campo & "NIF|straba|niftraba|T||20·"
            frmB.vCampos = campo
            frmB.vCargaFrame = False
            frmB.vDevuelve = "0|1|"
            frmB.vselElem = 1
            frmB.vConexionGrid = 1  'ODBC Aritaxi
            frmB.vTabla = "straba"
            frmB.vSQL = ""
            frmB.Show vbModal
            Set frmB = Nothing
            Screen.MousePointer = vbDefault
            If Cadena_frmB <> "" Then
                Me.txtTrab(Index).Text = RecuperaValor(Cadena_frmB, 1)
                Me.txtDescTra(Index).Text = RecuperaValor(Cadena_frmB, 2)
            End If
    End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub optReparaciones_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub







Private Sub texto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtAgente_GotFocus(Index As Integer)
    ConseguirFoco txtAgente(Index), 3
End Sub

Private Sub txtAgente_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtAgente_LostFocus(Index As Integer)
    miSQL = ""
    txtAgente(Index).Text = Trim(txtAgente(Index).Text)
    If txtAgente(Index).Text <> "" Then
        If PonerFormatoEntero(txtAgente(Index)) Then
            
            miSQL = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", txtAgente(Index).Text)
            If miSQL = "" Then MsgBox "No existe el agente: " & txtArticulo(Index).Text, vbExclamation
        End If
    End If
    Me.txtDescAgente(Index).Text = miSQL
    miSQL = ""
End Sub

Private Sub txtArticulo_GotFocus(Index As Integer)
    ConseguirFoco txtArticulo(Index), 3
End Sub

Private Sub txtArticulo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
    
End Sub


Private Sub txtArticulo_LostFocus(Index As Integer)
Dim T As String
    
    txtArticulo(Index).Text = Trim(txtArticulo(Index).Text)
    If txtArticulo(Index).Text = "" Then
        'EN blanco
        txtDescArticulo(Index).Text = ""
        Exit Sub
    End If
    
    
    T = "codartic"
    Codigo = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtArticulo(Index).Text, "T", T)
    If Codigo = "" Then
        MsgBox "No existe el artículo : " & txtArticulo(Index).Text, vbExclamation
    Else
        txtArticulo(Index).Text = T
    End If
    Me.txtDescArticulo(Index).Text = Codigo
    Codigo = ""
    
End Sub



Private Sub txtBancoPr_GotFocus(Index As Integer)
    ConseguirFoco txtBancoPr(Index), 3
End Sub

Private Sub txtBancoPr_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtBancoPr_LostFocus(Index As Integer)
    txtBancoPr(Index).Text = Trim(txtBancoPr(Index).Text)
    Codigo = ""
    miSQL = ""
    If txtBancoPr(Index).Text <> "" Then
        If IsNumeric(txtBancoPr(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nombanpr", "sbanpr", "codbanpr", txtBancoPr(Index).Text, "N")
            If Codigo = "" Then miSQL = "El codigo no pertence a ningun banco propio"
        Else
            miSQL = "Campo numerico"
        End If
    End If
     
    Me.txtDescBancoPr(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtBancoPr(Index).Text = ""
        PonerFoco txtBancoPr(Index)
    End If
End Sub

Private Sub txtCliente_GotFocus(Index As Integer)
    ConseguirFoco txtCliente(Index), 3
End Sub

Private Sub txtCliente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCliente_LostFocus(Index As Integer)
Dim Descri As String
    
    Descri = ""
    txtCliente(Index).Text = Trim(txtCliente(Index).Text)
    If txtCliente(Index).Text <> "" Then
        If Not IsNumeric(txtCliente(Index).Text) Then
            MsgBox "Campo codigo cliente debe ser numérico", vbExclamation
            PonerFoco txtCliente(Index)
        Else
            Descri = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCliente(Index).Text, "N")
            If Descri = "" Then
                MsgBox "No existe el cliente : " & txtCliente(Index).Text, vbExclamation
            End If
        End If
    End If
    Me.txtDescClie(Index).Text = Descri
    
    
    
End Sub


    

Private Sub txtCodProve_GotFocus(Index As Integer)
    ConseguirFoco txtCodProve(Index), 3
End Sub

Private Sub txtCodProve_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCodProve_LostFocus(Index As Integer)
    txtCodProve(Index).Text = Trim(txtCodProve(Index).Text)
    Codigo = ""
    miSQL = ""
    If txtCodProve(Index).Text <> "" Then
        If IsNumeric(txtCodProve(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtCodProve(Index).Text, "N")
            If Codigo = "" Then MsgBox "El codigo no pertence a ningun proveedor", vbExclamation
        Else
            miSQL = "Campo numerico"
        End If
    End If
    Me.txtDescProve(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtCodProve(Index).Text = ""
        PonerFoco txtCodProve(Index)
    End If
End Sub

Private Sub txtForpa_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

' Private Sub txtImporte_GotFocus(Index As Integer)
'    ConseguirFoco txtImporte(Index), 3
' End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumAlbar_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtRecargaMov_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTrab_GotFocus(Index As Integer)
    ConseguirFoco txtTrab(Index), 3
End Sub

Private Sub txtTrab_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTrab_LostFocus(Index As Integer)


    txtTrab(Index).Text = Trim(txtTrab(Index).Text)
    Codigo = ""
    miSQL = ""

    If txtTrab(Index).Text <> "" Then
        If IsNumeric(txtTrab(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", txtTrab(Index).Text, "N")
            If Codigo = "" Then miSQL = "El código no pertence a ningun trabajador"
        Else
            miSQL = "Campo numerico"
        End If
    End If
    
    Me.txtDescTra(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        If Index < 2 Then
            txtTrab(Index).Text = ""
            PonerFoco txtTrab(Index)
        End If
    End If
End Sub



Private Sub txtDpto_GotFocus(Index As Integer)
    ConseguirFoco txtDpto(Index), 3
End Sub

Private Sub txtDpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtDpto_LostFocus(Index As Integer)
Dim vC As CCliente
    'Si el cliente ES EL MISMO
    campo = ""
    txtDpto(Index).Text = Trim(txtDpto(Index).Text)
    If Index < 2 Then
        If txtDpto(Index).Text <> "" Then
             'Index=0 or 1.  Departamento sera puesto si, y solo si, el cliente es el mismo
             If txtCliente(0).Text <> "" And txtCliente(0).Text = txtCliente(1).Text And txtDescClie(0).Text <> "" Then
                 'PERFECTO, el cliente existe y es el mismo
                 Set vC = New CCliente
                 vC.Codigo = txtCliente(0).Text
                 vC.DptoCliente txtDpto(Index).Text, campo
                 Set vC = Nothing
             Else
                 'Todavia no ha puesto el cliente
                 MsgBox "Para poner el departamento cliente debe y el hasta  debe ser el mismo", vbExclamation
                 txtDpto(Index).Text = ""
        
             End If
        End If
        Me.txtDescDpto(Index).Text = campo
    Else
    
    
    End If
End Sub




Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
Dim T As String
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        T = txtFecha(Index).Text
        If EsFechaOK(T) Then
            txtFecha(Index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub

'Dado un FRAME lo pone a true y lo situa en x:120 y:0 y devuelve lo que debe medir el form
Private Sub PonerFrameVisible(ByRef F As Frame, ByRef CH As Integer, CW As Integer)
    F.Top = 0
    F.Left = 120
    F.visible = True
    CH = F.Height + 420
    CW = F.Width + 240
End Sub


Private Sub PonerLabelDptoDireccion(L As Label)
    If vParamAplic.Departamento Then
        L.Caption = "Dpto."
    Else
        L.Caption = "Direc."
    End If
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String
Dim Subtipo As String 'F: fecha   N: numero   T: texto  H: HORA
Dim TDes As TextBox
Dim THas As TextBox
Dim DesD As TextBox 'Descripcion DESDE
Dim DesH As TextBox '    "       HASTA

    PonerDesdeHasta = False
    
    Select Case Tipo
    Case "F"
        'Campos fecha
        Set TDes = txtFecha(indD)
        Set THas = txtFecha(indH)
        Subtipo = "F"
        If indD = 27 Or indH = 28 Then Subtipo = "FH"
    Case "CLI"
        'Cliente
        Set TDes = txtCliente(indD)
        Set THas = txtCliente(indH)
        Set DesD = txtDescClie(indD)
        Set DesH = txtDescClie(indH)
        Subtipo = "N"
    Case "DPT"
        'DEpartamento
        Set TDes = txtDpto(indD)
        Set THas = txtDpto(indH)
        Set DesD = txtDescDpto(indD)
        Set DesH = txtDescDpto(indH)
        Subtipo = "N"
        
    Case "PRO"
        Set TDes = txtCodProve(indD)
        Set THas = txtCodProve(indH)
        Set DesD = txtDescProve(indD)
        Set DesH = txtDescProve(indH)
        Subtipo = "N"
 
    Case "ART", "ARC"

        Set TDes = txtArticulo(indD)
        Set THas = txtArticulo(indH)
        Set DesD = txtDescArticulo(indD)
        Set DesH = txtDescArticulo(indH)
        Subtipo = "T"
    Case "AGT"
        Set TDes = txtAgente(indD)
        Set THas = txtAgente(indH)
        Set DesD = txtDescAgente(indD)
        Set DesH = txtDescAgente(indH)
        Subtipo = "N"
    
    Case "ALP"
        Set TDes = txtNumAlbar(indD)
        Set THas = txtNumAlbar(indH)
        Subtipo = "T"
      
    Case "TRA"
        'TRABAJADOR
         
        Set TDes = Me.txtTrab(indD)
        Set THas = txtTrab(indH)
        Subtipo = "N"
        If indD = 5 Then
            'llamadas
            Set DesD = txtDescTra(indD)
            Set DesH = txtDescTra(indH)
        End If
    End Select
    
    devuelve = CadenaDesdeHasta(TDes.Text, THas.Text, campo, Subtipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If Subtipo <> "F" And Subtipo <> "FH" Then
        'Fecha para Crystal Report
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(TDes.Text, THas.Text, campo, Subtipo)
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, TDes, THas, DesD, DesH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Function AnyadirParametroDH(Cad As String, ByRef TextoDESDE As TextBox, TextoHasta As TextBox, ByRef TD As TextBox, ByRef TH As TextBox) As String
On Error Resume Next
    
     If TextoDESDE.Text <> "" Then
        Cad = Cad & "desde " & TextoDESDE.Text
        If TD.Text <> "" Then Cad = Cad & " - " & TD.Text
    End If
    If TextoHasta.Text <> "" Then
        Cad = Cad & "  hasta " & TextoHasta.Text
        If TH <> "" Then Cad = Cad & " - " & TH.Text
    End If
    
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function

'Para las reparaciones. Carga el importe real y teorico.
Private Sub CargaImporteRealReparaciones()
Dim ImpTot As Currency
Dim ImpTeo As Currency
Dim miSQL As String

    'A partir de la reparacion , mirare en los albaranes, y de los albaranes ver el coste real de la reparacion y el teorico
    Set miRsAux = New ADODB.Recordset
    
    'Meto el select para las
    If cadSelect <> "" Then
        Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If
    

 
    
    'Montamos el select al reves
    Codigo = "s.codtipom=l.codtipom and s.codtipoa=l.codtipoa " & Codigo
    Codigo = "s.fecfactu = l.fecfactu and s.numfactu=l.numfactu AND " & Codigo
    Codigo = "s.codtipom=l.codtipom AND " & Codigo
    Codigo = "sartic.codartic = l.codartic AND " & Codigo
    Codigo = "select l.*,s.fechaalb,preciove,h.numrepar,h.fecrepar from  schrep h,slifac l,scafac1 s,sartic where  h.numalbar = s.numalbar and h.fechaalb=s.fechaalb AND " & Codigo
    'EL ORDEN
    Codigo = Codigo & " ORDER BY s.numalbar ,s.fechaalb"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    NumRegElim = 1
    miSQL = ""
    While Not miRsAux.EOF
        If Codigo <> CStr(miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)) Then
            If Codigo <> "" Then
                'INSERTAMOS
                ImpTeo = Round(ImpTeo, 2) * 100
                miSQL = miSQL & NumRegElim & "," & CLng(ImpTeo) & "," & TransformaComasPuntos(CStr(ImpTot)) & ")"
                'EXEcuete
                'en codprove llevare el numero de albaran
                'en codartic llevare el importe total teorico
                'en cantidad                    TOTAL
                miSQL = "insert into tmpnlotes (codusu,codprove,fechaalb,numalbar,nomartic,numlinea,codartic,cantidad) " & miSQL
                conn.Execute miSQL
                NumRegElim = NumRegElim + 1
            End If
            Codigo = miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)
            miSQL = " VALUES (" & vUsu.Codigo & "," & miRsAux!numrepar & ",'" & Format(miRsAux!fecrepar, FormatoFecha) & "',0,0,"
            ImpTot = 0
            ImpTeo = 0
        End If
        ImpTeo = ImpTeo + (miRsAux!preciove * miRsAux!Cantidad)
        ImpTot = miRsAux!ImporteL + ImpTot
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    If miSQL <> "" Then
        'El ultimo
        ImpTeo = Round(ImpTeo, 2) * 100
        miSQL = miSQL & NumRegElim & "," & CLng(ImpTeo) & "," & TransformaComasPuntos(CStr(ImpTot)) & ")"
        'EXEcuete
        'en codprove llevare el importe total teorico
        'en cantidad                    TOTAL
        miSQL = "insert into tmpnlotes (codusu,codprove,fechaalb,numalbar,nomartic,numlinea,codartic,cantidad) " & miSQL
        conn.Execute miSQL
    End If


    'La fecha hasta la tengo en la txtfecha(1)
    'Ahora pondere, en una (Y SOLO una) de las lineas el importe del mantenimiento hasta la fecha
    ' Las demas a CERO. Con lo cual, en el report, la suma del campo dara ESE importe solo
    '

    
   
    If cadSelect <> "" Then
        Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If
    

    miSQL = "select numrepar,fecrepar,tieneman,"
    miSQL = miSQL & " mes01act,mes02act,mes03act,mes04act,mes05act,mes06act,mes07act,mes08act,mes09act,mes10act,mes11act,mes12act"
    miSQL = miSQL & " from schrep h,sserie s left join scaman m  on s.nummante=m.nummante and s.codclien=m.codclien"
    miSQL = miSQL & " where h.numserie=s.numserie and s.codartic=h.codartic "
    If Codigo <> "" Then miSQL = miSQL & Codigo
    
    'EL ORDEN
    IndiceImg = 12
    If txtFecha(1).Text <> "" Then IndiceImg = Month(CDate(txtFecha(1).Text))
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        ImpTot = 0
        If miRsAux!TieneMan = 1 Then
            '--------------------------------------------------------------------
            'OK, TIENE MANTENIMIENTO
            'Ire recorriendo los importes desde mes01act hasta el mes hasta
            'Si la fecha es fin es nada, entonces hare tooodos
            For NumRegElim = 1 To IndiceImg
                If Not IsNull(miRsAux.Fields(NumRegElim + 2)) Then ImpTot = ImpTot + miRsAux.Fields(NumRegElim + 2)
            Next
        End If
        If ImpTot <> 0 Then
            'UPDATEAMOS LA tmp
            miSQL = "UPDATE tmpnlotes set nomartic=" & CLng(ImpTot * 100) & " WHERE codusu = " & vUsu.Codigo
            miSQL = miSQL & " AND codprove = " & miRsAux!numrepar & " AND fechaalb = '" & Format(miRsAux!fecrepar, FormatoFecha) & "' AND numalbar =0"
            conn.Execute miSQL
        End If
        '--------------
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
End Sub
    


Private Sub EstadisticaReparacionTecnicoNueva()
Dim RT As ADODB.Recordset
Dim OptimizarSelect As String
Dim RAlb As ADODB.Recordset
Dim C As Long
Dim EnAlbaranes As Boolean

    Label3(63).Caption = "Obteniendo reg. albaranes"
    Label3(63).Refresh
    

    'Preparamos las temporales
    Codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    Codigo = "DELETE FROM tmpnlotes WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    
    
    
    'LOS INSERTS PARA LAS TABLAS temporales                                         numserie
    cadFormula = "insert into tmpinformes (codusu,codigo1,importe1,importe2,nombre1,nombre2) VALUES (" & vUsu.Codigo & ","
    cadFrom = "insert into tmpnlotes (codusu,codprove,numalbar,fechaalb,numlinea,nomartic) values (" & vUsu.Codigo & ","
    
    
    'Optimizacion del select
    If cadSelect <> "" Then
        'Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " WHERE " & cadSelect
    Else
        Codigo = ""
    End If
    

    Codigo = "Select distinct(codtipom) from schrep  " & Codigo
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    OptimizarSelect = ""
    While Not miRsAux.EOF
        Codigo = DBLet(miRsAux!codtipom, "T")
        If Codigo <> "" Then OptimizarSelect = OptimizarSelect & " OR scafac1.codtipoa = '" & Codigo & "'"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If OptimizarSelect <> "" Then
        OptimizarSelect = Mid(OptimizarSelect, 4) 'quito el preimer or
        OptimizarSelect = "(" & OptimizarSelect & ")"
    End If
    
    'Cargo el RS con todos los datos de los albarnes
    miSQL = "select scafac1.numalbar,scafac1.fechaalb,scafac1.codtipoa,sum(importel),sum(cantidad*preciove)"
    miSQL = miSQL & " from scafac1,slifac,sartic where  scafac1.codtipom =slifac.codtipom  and scafac1.numfactu  =slifac.numfactu"
    miSQL = miSQL & " and scafac1.fecfactu  =slifac.fecfactu"
    miSQL = miSQL & " and scafac1.codtipoa  =slifac.codtipoa  and scafac1.numalbar  =slifac.numalbar and sartic.codartic=slifac.codartic"
    
    
    cadNomRPT = ""
    If txtFecha(2).Text <> "" Then cadNomRPT = cadNomRPT & " AND fechaalb >='" & Format(txtFecha(2).Text, FormatoFecha) & "'"
    If txtFecha(3).Text <> "" Then cadNomRPT = cadNomRPT & " AND fechaalb <='" & Format(txtFecha(3).Text, FormatoFecha) & "'"
    If OptimizarSelect <> "" Then cadNomRPT = cadNomRPT & " AND " & OptimizarSelect
    

    miSQL = miSQL & cadNomRPT
    miSQL = miSQL & " group by scafac1.numalbar,scafac1.fechaalb,scafac1.codtipoa order by codtipoa,numalbar,fechaalb"

    'Cargamos las sumas en facturas
    miRsAux.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    'Cargamos las sumas en albaranaes  ############
    miSQL = "Select scaalb.numalbar,scaalb.fechaalb,scaalb.codtipom,sum(importel),sum(cantidad*preciove)"
    miSQL = miSQL & " from scaalb,slialb,sartic where  scaalb.codtipom =slialb.codtipom  and scaalb.numalbar  =slialb.numalbar"
    miSQL = miSQL & " and sartic.codartic=slialb.codartic"
    cadNomRPT = Replace(cadNomRPT, " scafac1.codtipoa", "scaalb.codtipom")
    miSQL = miSQL & cadNomRPT
    miSQL = miSQL & " group by numalbar,fechaalb,codtipom order by codtipom,numalbar,fechaalb"

    Set RAlb = New ADODB.Recordset
    RAlb.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    
    
    
    
    
    
    
    
    
    'Cargamos el rS de la reparaciones
    If cadSelect <> "" Then
        'Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " WHERE " & cadSelect
    Else
        Codigo = ""
    End If
    

    Codigo = " from schrep  " & Codigo
    
    Set RT = New ADODB.Recordset
    RT.Open "select count(*)" & Codigo, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    C = DBLet(RT.Fields(0), "N")
    RT.Close
    
    RT.Open "select * " & Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    While Not RT.EOF
        NumRegElim = NumRegElim + 1
        
        Label3(63).Caption = "Rep: " & RT!numrepar & "  (" & NumRegElim & "/" & C & ")"
        Label3(63).Refresh
        
        If IsNull(RT!codtipom) Or IsNull(RT!NumAlbar) Or IsNull(RT!FechaAlb) Then
            ImpTeo = 0
            ImpTot = 0
        Else
            
            PonerIMportesAlbaranes RAlb, RT!codtipom, RT!NumAlbar, RT!FechaAlb, ImpTot, ImpTeo, EnAlbaranes
        End If
        
 

            
                'INSERTAMOS
                'en tmpinformes
                Codigo = "'" & DevNombreSQL(RT!NomArtic) & "','" & DBLet(RT!numSerie, "T") & "')"
                Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
                Codigo = cadFormula & Codigo
                conn.Execute Codigo
                
                
                'en tmpnlotes
                '       numprove
                Codigo = DBLet(RT!nummante, "T")
                If Codigo <> "" Then
                    Codigo = UCase(Codigo)
                    Debug.Print Codigo
                    If Codigo = "S/MTO" Or Codigo = "SIN ESPC." Then
                        Codigo = "0"
                    Else
                        Codigo = "1"
                    End If
                Else
                    Codigo = "0"
                End If
                Codigo = Abs(EnAlbaranes) & ",'" & Format(RT!fecrepar, FormatoFecha) & "'," & Codigo & ",'" & Trim(DevNombreSQL(RT!nomclien)) & "')"
                Codigo = NumRegElim & "," & Codigo
                Codigo = cadFrom & Codigo
                conn.Execute Codigo
                
                 
      
        
        RT.MoveNext
    Wend
    RT.Close
    RAlb.Close
    miRsAux.Close
  Set RT = Nothing
  Set RAlb = Nothing
 

    
    


End Sub



Private Sub PonerIMportesAlbaranes(ByRef RAlba As ADODB.Recordset, codtipom As String, alb As Long, fech As Date, ByRef Impt As Currency, ByRef ImpTeor As Currency, ByRef EnAlbaranes As Boolean)
Dim Fin As Boolean
Dim Esta As Boolean

    Impt = 0
    ImpTeor = 0
    
    'Comprobamos en albaranes primero
    EnAlbaranes = True
    Esta = False
    If Not RAlba.EOF Then
        
        Fin = False
        While Not Fin
            If RAlba!codtipom = codtipom Then
                If RAlba!NumAlbar = alb Then
                    If RAlba!FechaAlb = fech Then
                        'AQUI ESTA
                        Fin = True
                        Esta = True
                        Impt = RAlba.Fields(3)
                        ImpTeor = RAlba.Fields(4)
                    End If
                    
                Else
                    'SI ha sobrepasado YA no esta
                    If RAlba!NumAlbar > alb Then Fin = True
                End If
            End If
            RAlba.MoveNext
            If RAlba.EOF Then Fin = True
        Wend
    
        RAlba.MoveFirst
        If Esta Then Exit Sub  'Ya lo hemos encontrado
    End If
    
    
    EnAlbaranes = False
    If miRsAux.EOF Then Exit Sub
    Fin = False
    While Not Fin
        If miRsAux!codtipoa = codtipom Then
            If miRsAux!NumAlbar = alb Then
                If miRsAux!FechaAlb = fech Then
                    'AQUI ESTA
                    Fin = True
                    Impt = miRsAux.Fields(3)
                    ImpTeor = miRsAux.Fields(4)
                End If
            Else
                miRsAux.MoveLast
            End If
        End If
        miRsAux.MoveNext
        If miRsAux.EOF Then Fin = True
    Wend
    miRsAux.MoveFirst
End Sub

Private Sub EstadisticaReparacionTecnico()
    'Preparamos las temporales
    Codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    Codigo = "DELETE FROM tmpnlotes WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    

    
    'LOS INSERTS PARA LAS TABLAS temporales                                         numserie
    cadFormula = "insert into tmpinformes (codusu,codigo1,importe1,importe2,nombre1,nombre2) VALUES (" & vUsu.Codigo & ","
    cadFrom = "insert into tmpnlotes (codusu,codprove,numalbar,fechaalb,numlinea,nomartic) values (" & vUsu.Codigo & ","
    
    'Montamos el select al reves
    'PARA LAS FACTURAS
    If cadSelect <> "" Then
        Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If
    Codigo = " s.codtipom=l.codtipom and s.codtipoa=l.codtipoa " & Codigo
    Codigo = " s.fecfactu = l.fecfactu and s.numfactu=l.numfactu AND " & Codigo
    Codigo = " s.codtipom=l.codtipom AND " & Codigo
    Codigo = " sartic.codartic = l.codartic AND " & Codigo
    Codigo = " h.numserie=sserie.numserie AND h.codclien=sserie.codclien AND " & Codigo
    Codigo = " sclien.codclien = h.codclien AND " & Codigo
    Codigo = " where  h.numalbar = s.numalbar and h.fechaalb=s.fechaalb AND " & Codigo
    'Las tablas
    Codigo = " from schrep h,slifac l,scafac1 s,sclien , sserie,sartic" & Codigo
    Codigo = "select l.*,s.fechaalb,preciove,h.fecrepar,h.nomclien,tieneman,h.nomartic,h.numserie " & Codigo
    'EL ORDEN
    Codigo = Codigo & " ORDER BY s.numalbar ,s.fechaalb"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    miSQL = ""
    While Not miRsAux.EOF
    

    
        If Codigo <> CStr(miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)) Then
            If Codigo <> "" Then
                NumRegElim = NumRegElim + 1
                'INSERTAMOS
                'en tmpinformes
                Codigo = RecuperaValor(miSQL, 1)
                Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
                Codigo = cadFormula & Codigo
                conn.Execute Codigo
                
                
                'en tmpnlotes
                '       numprove
                Codigo = RecuperaValor(miSQL, 2)
                Codigo = NumRegElim & "," & Codigo
                Codigo = cadFrom & Codigo
                conn.Execute Codigo
                
                
            End If
            'Meto dos datos enpipados
            miSQL = "'" & DevNombreSQL(miRsAux!NomArtic) & "','" & miRsAux!numSerie & "')"
            miSQL = miSQL & "|"
            miSQL = miSQL & "0,'" & Format(miRsAux!fecrepar, FormatoFecha) & "'," & miRsAux!TieneMan & ",'" & DevNombreSQL(miRsAux!nomclien) & "')|"
            Codigo = miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)
            ImpTot = 0
            ImpTeo = 0
        End If
        ImpTeo = ImpTeo + (miRsAux!preciove * miRsAux!Cantidad)
        ImpTot = miRsAux!ImporteL + ImpTot
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If miSQL <> "" Then
        NumRegElim = NumRegElim + 1
        
        'en tmpinformes
        Codigo = RecuperaValor(miSQL, 1)
        Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
        Codigo = cadFormula & Codigo
        conn.Execute Codigo
        
        
        'en tmpnlotes
        '       numprove
        Codigo = RecuperaValor(miSQL, 2)
        Codigo = NumRegElim & "," & Codigo
        Codigo = cadFrom & Codigo
        conn.Execute Codigo
        
        
    End If
    

    
    'AHORA HAGO EL INSERT PARA LOS ALBARANES QUE NO HAN SIDO FACTURADOS
    'PARA LOS ALBARANES
    If cadSelect <> "" Then
        Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If

    miSQL = "select l.*,preciove,tieneman,h.fechaalb,h.numserie,h.nomclien,fecrepar"
    miSQL = miSQL & " from schrep h,scaalb c,slialb l,sartic a,sserie s "
    miSQL = miSQL & " WHERE h.codtipom=c.codtipom and h.numalbar=c.numalbar and h.fechaalb=c.fechaalb and"
    miSQL = miSQL & " l.numalbar=c.numalbar and l.codtipom=c.codtipom and l.codartic=a.codartic"
    miSQL = miSQL & " and h.numserie=s.numserie and h.codclien =s.codclien" & Codigo
    
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    miSQL = ""
    While Not miRsAux.EOF
        If Codigo <> CStr(miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)) Then
            If Codigo <> "" Then
                NumRegElim = NumRegElim + 1
                'INSERTAMOS
                'en tmpinformes
                Codigo = RecuperaValor(miSQL, 1)
                Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
                Codigo = cadFormula & Codigo
                conn.Execute Codigo
                
                
                'en tmpnlotes
                '       numprove
                Codigo = RecuperaValor(miSQL, 2)
                Codigo = NumRegElim & "," & Codigo
                Codigo = cadFrom & Codigo
                conn.Execute Codigo
                
            End If
            'Meto dos datos enpipados
            miSQL = "'" & DevNombreSQL(miRsAux!NomArtic) & "','" & miRsAux!numSerie & "')"
            miSQL = miSQL & "|"
            miSQL = miSQL & "1,'" & Format(miRsAux!fecrepar, FormatoFecha) & "'," & miRsAux!TieneMan & ",'" & DevNombreSQL(miRsAux!nomclien) & "')|"
            Codigo = miRsAux!NumAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)
            ImpTot = 0
            ImpTeo = 0
        End If
        ImpTeo = ImpTeo + (miRsAux!preciove * miRsAux!Cantidad)
        ImpTot = miRsAux!ImporteL + ImpTot
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If miSQL <> "" Then
        NumRegElim = NumRegElim + 1
        
        'en tmpinformes
        Codigo = RecuperaValor(miSQL, 1)
        Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
        Codigo = cadFormula & Codigo
        conn.Execute Codigo
        
        
        'en tmpnlotes
        '       numprove
        Codigo = RecuperaValor(miSQL, 2)
        Codigo = NumRegElim & "," & Codigo
        Codigo = cadFrom & Codigo
        conn.Execute Codigo
        
        
    End If
    


End Sub










'------------------------------------------------------------------
'------------------------------------------------------------------
'               M U L T I B A S E
'------------------------------------------------------------------
Private Sub CargaListMultibase()
    Me.lstMultibase.Clear
    miSQL = "Clientes|Proveedores|Trabajadores|Direcciones|"
    For numParam = 1 To 4
        Me.lstMultibase.AddItem RecuperaValor(miSQL, CInt(numParam))
    Next numParam
    'Como organiza informacion
    '         tabla  clave    campos a cambiar(empieza con coma) tipodatos clave.
    'Clientes
    miSQL = "sclien:codclien:,nomclien,nomcomer ,domclien ,codpobla ,pobclien,perclie1,perclie2:N|"
    miSQL = miSQL & "sprove:codprove:,nomprove,nomcomer ,domprove ,codpobla ,pobprove ,perprov1 ,perprov2:N|"
    miSQL = miSQL & "straba:codtraba:,nomtraba,domtraba,codpobla,pobtraba:N|"
    miSQL = miSQL & "sdirec:codclien,coddirec:,nomdirec ,domdirec ,pobdirec ,prodirec ,perdirec:N,N|"
        
End Sub


Private Sub HacerCambiosMultibase(numlinea As Integer)
Dim TotalReg As Long
Dim i As Integer
Dim J As Integer
Dim Claves As Integer
Dim Campos As Integer
Dim Cambios As Long
Dim T1 As Single
'Reutilizacion de variables
'cadTitulo cadNomRPT  conSubRPT

    On Error GoTo EHacerCambiosMultibase
    campo = lstMultibase.List(numlinea - 1)
    lblMultibase.Caption = "Preparando datos: " & campo
    
    lblMultibase.Refresh

    cadFormula = RecuperaValor(miSQL, numlinea)
    cadFormula = Replace(cadFormula, ":", "|")
    cadFormula = cadFormula & "|"  'Le añado el pipe final
    'Primero el conteo
    cadParam = "Select count(*) from " & RecuperaValor(cadFormula, 1)
    miRsAux.Open cadParam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalReg = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then TotalReg = miRsAux.Fields(0)
    End If
    miRsAux.Close
    
    DoEvents
    If TotalReg = 0 Then
        lblMultibase.Caption = "Tabla vacia " & campo
        lblMultibase.Refresh
        Espera 1
    End If
    
    'Veamos cuantos campos hay que ver la conversion de campos, y las claves
    cadParam = RecuperaValor(cadFormula, 2)
    Claves = 1
    Cambios = 0
    While cadParam <> ""
        NumRegElim = InStr(1, cadParam, ",")
        If NumRegElim = 0 Then
            cadParam = ""
        Else
            Claves = Claves + 1
            cadParam = Mid(cadParam, NumRegElim + 1)
        End If
    Wend
    cadParam = RecuperaValor(cadFormula, 3)
    Campos = 0 'aqui cero pq empieza con la coma
    While cadParam <> ""
        NumRegElim = InStr(1, cadParam, ",")
        If NumRegElim = 0 Then
            cadParam = ""
        Else
            Campos = Campos + 1
            cadParam = Mid(cadParam, NumRegElim + 1)
        End If
    Wend
        

                            'claves                                 'campos cambiar
    cadParam = "SELECT " & RecuperaValor(cadFormula, 2) & RecuperaValor(cadFormula, 3)
    cadParam = cadParam & " FROM " & RecuperaValor(cadFormula, 1)
    miRsAux.Open cadParam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    Cambios = 0
    T1 = Timer   'Para hacer doevents cada 3 segundos
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        'Los labels
        lblMultibase.Caption = campo & " ( " & NumRegElim & " / " & TotalReg & " )"
        lblMultibase.Refresh
        If Timer - T1 > 3 Then
            DoEvents
            Me.Refresh
            Espera 0.2
            T1 = Timer
        End If
        
        cadSelect = "" 'LOS UPDATES
        For i = Claves To Campos
            If Not IsNull(miRsAux.Fields(i)) Then
                cadParam = miRsAux.Fields(i)  'Cojo el valor del field
                cadNomRPT = RevisaCaracterMultibase(cadParam)  'Obtengo la modificaicon por campos multibase
                If cadParam <> cadNomRPT Then
                    'HAY que modificar ya que son disitintos el de laBD y el calculado por el modulo de multibase
                    cadSelect = cadSelect & ", " & miRsAux.Fields(i).Name & " = '" & DevNombreSQL(cadNomRPT) & "'"
                End If
            End If
        Next
        'SI cadselect <>"" entonces hay que ejecutar SQL
        If cadSelect <> "" Then
            'Los campos claves van del 0 a claves -1
            cadParam = ""
            cadTitulo = RecuperaValor(cadFormula, 4) 'los tipos de datos
            cadTitulo = Replace(cadTitulo, ",", "|") & "|"
            For J = 0 To Claves - 1
                cadParam = cadParam & " AND " & miRsAux.Fields(J).Name & " = "
                Codigo = RecuperaValor(cadTitulo, J + 1)

                Select Case Codigo
                Case "F"
                    cadParam = cadParam & "'" & Format(miRsAux.Fields(i).Value, FormatoFecha) & "'"
                Case "T"
                    cadParam = cadParam & "'" & miRsAux.Fields(i).Value & "'"
                Case Else  'NUMERICO
                    cadParam = cadParam & miRsAux.Fields(J).Value
                End Select
            Next J
            
            
            'Acabas de montar el UPDATE
            cadTitulo = "UPDATE " & RecuperaValor(cadFormula, 1)
            cadSelect = Mid(cadSelect, 2)   'QUITO la coma
            cadParam = Mid(cadParam, 5)     'QUITO el primer AND
            cadTitulo = cadTitulo & " SET " & cadSelect & " WHERE " & cadParam
            conn.Execute cadTitulo
            Cambios = Cambios + 1
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    lblMultibase.Caption = "FIN " & campo
    lblMultibase.Refresh
    If Cambios > 0 Then Me.Tag = Me.Tag & vbCrLf & "   .- " & campo & " : " & Cambios
    Exit Sub
EHacerCambiosMultibase:
    MuestraError Err.Number
End Sub
'       fin mULTIBASE
'------------------------------------------------------------------'------------------------------------------------------------------


'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Facturacion de recargas de telefonia
'
'------------------------------------------------------------------
'------------------------------------------------------------------



Private Sub HacerFacturacionTelefonia(vAlbaranes As Collection, MenError As String)
Dim RT As ADODB.Recordset
Dim b As Boolean
Dim NumAlb As String
Dim Almacen As Integer



    'El proceso sera el siguiente:
    'Voy a agrupar por dia (podria ser por mes),trabajador
    'Y para cada uno de los resultados del recodset voy a generar un albaran.
    'Me guardare los albaranes generados y despues los facturare.
    'Para ello
    campo = "Select codtraba,count(*) as cantidad,sum(importe)as total from stelefonia WHERE " & cadSelect & " GROUP by codtraba"
    
    Set RT = New ADODB.Recordset
    RT.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RT.EOF
        
        Almacen = DevuelveDesdeBD(conAri, "codalmac", "straba", "codtraba", CStr(RT!CodTraba), "N")
        
        conn.BeginTrans
        
        'Obtener el contador de Albaran (ALV).
        b = ObtenerContadorAlbaran(NumAlb)
        
        If b Then
            'Actualizar los stocks de todos los articulos comprados
            'Insertar movimiento en smoval
            'B = InsertarMovAlmacen(NumAlb)  ¿ FALTA### ?
    
            'Insertar en las tablas de Albaranes: scaalb, slialb
            'en el campo scafac1.numalbar guardamos el nº de ticket
            If b Then b = InsertarAlbaran(NumAlb, CStr(RT!CodTraba), 1, RT!Cantidad, RT!total, MenError)
        
        End If



       
        If Not b Then
            conn.RollbackTrans
            RT.Close
            Set RT = Nothing
            Exit Sub
        Else
            vAlbaranes.Add CStr(NumAlb)
            conn.CommitTrans
            
            'Le pongo a facturado en la telefonia
            miSQL = "UPDATE stelefonia SET facturado = 1 WHERE " & cadSelect & " AND codtraba = " & RT!CodTraba
            conn.Execute miSQL
        End If
    
    
        RT.MoveNext
    Wend
    RT.Close
    


End Sub


Private Function ObtenerContadorAlbaran(NumAlb As String) As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean

    On Error GoTo ErrConAlb

    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer("ALV") Then
        Do
            NumAlb = vTipoMov.ConseguirContador("ALV")
            vTipoMov.IncrementarContador ("ALV")
            miSQL = "select count(*) from scaalb where codtipom='ALV' and numalbar=" & NumAlb
            Existe = (RegistrosAListar(miSQL) > 0)
        Loop Until Existe = False
        ObtenerContadorAlbaran = True
    Else
        ObtenerContadorAlbaran = False
    End If
    Set vTipoMov = Nothing
    Exit Function
    
ErrConAlb:
    ObtenerContadorAlbaran = False
    MuestraError Err.Number, "Obtener contador albaran", Err.Description
End Function

Private Function InsertarAlbaran(NumAlb As String, CodTraba As String, CodAlmc As Integer, Cantidad As Currency, Importe As Currency, menErr As String) As Boolean
Dim b As Boolean
Dim vClien As CCliente
Dim Sql As String

    On Error GoTo EInsAlb



    'Cabecera de albaran
    '----------------------------------
    Sql = "INSERT INTO scaalb (codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    Sql = Sql & "coddirec,nomdirec,referenc,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,"
    Sql = Sql & "observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,esticket,numtermi,numventa) "
                                                                    'Facturar   cliente
    Sql = Sql & " VALUES ('ALV'," & NumAlb & "," & DBSet(Now, "F") & ",1," & txtCliente(2).Text & ","
    
    'Obtenemos los datos del cliente
    Set vClien = New CCliente
    If vClien.Existe(txtCliente(2).Text) Then
        If vClien.LeerDatos(txtCliente(2).Text) Then
            Sql = Sql & DBSet(vClien.Nombre, "T", "N") & ", " & DBSet(vClien.Domicilio, "T", "N") & ","
            Sql = Sql & DBSet(vClien.CPostal, "T", "N") & ", " & DBSet(vClien.Poblacion, "T", "N") & "," & DBSet(vClien.Provincia, "T", "N") & ","
            Sql = Sql & DBSet(vClien.NIF, "T", "N") & "," & DBSet(vClien.TfnoClien, "T") & ","
            'coddirec,nomdirec,referenc a nulo
            Sql = Sql & "NULL,NULL,NULL,"
            
            Sql = Sql & CodTraba & "," & CodTraba & "," & CodTraba & "," 'trabajador
            '                              cod forpa
            Sql = Sql & vClien.Agente & ",1," & vClien.FEnvio & ",0,0," & vClien.TipoFactu & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'observaciones
            Sql = Sql & ValorNulo & "," & ValorNulo & "," 'datos oferta: aqui guardamos nº venta
            'En los campos de datos del pedido guardamos los datos del ticket
            'SQL = SQL & NumTicket & "," & DBSet(RSVenta!fecventa, "F") & "," & ValorNulo & "," & ValorNulo & ",1," & DBSet(RSVenta!NumTermi, "N") & "," & DBSet(RSVenta!NumVenta, "N", "S") & ")" 'esticket=1, terminal
            Sql = Sql & "NULL,NULL," & ValorNulo & "," & ValorNulo & ",0,NULL,NULL)"
            b = vClien.ActualizaUltFecMovim(Now)
        Else
            b = False
        End If
    End If
    Set vClien = Nothing
    
    
    If b Then
        'Insertar Cabecera
'    MenError = "Error al insertar en la tabla Cabecera de Albaranes (scaalb )."
        conn.Execute Sql, , adCmdText
        
        'Lineas del albaran
        'Inserta en tabla "slialb" todas las lineas de venta
        Sql = "INSERT INTO slialb "
        Sql = Sql & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, "
        Sql = Sql & "dtoline1, dtoline2, importel, origpre) VALUES ("
        Sql = Sql & "'ALV'," & DBSet(NumAlb, "N") & ",1," & CodAlmc & ",'" & DevNombreSQL(txtArticulo(0).Text) & "','" & DevNombreSQL(txtDescArticulo(0).Text)
        Sql = Sql & "',NULL," & Cantidad & "," & TransformaComasPuntos(CStr(Round(Importe / Cantidad, 4))) & ",0,0," & TransformaComasPuntos(CStr(Importe)) & ",'')"
        'SQL = SQL & " FROM sliven WHERE " & Replace(cadSel, "scaven", "sliven")
        conn.Execute Sql, , adCmdText
    End If


    
    'Guardamos los valores identificativos de la factura generada
    'para imprimirla posteriormente
    If b Then cadImpresion = "{scaalb.codtipom}='ALV' and {scaalb.numalbar}=" & DBSet(NumAlb, "N")

EInsAlb:
    If Err.Number <> 0 Then
        menErr = "Insertando el Albaran: " & vbCrLf & Err.Description
        b = False
    End If
    InsertarAlbaran = b
End Function


Private Function ObtenerDatosTickets(Diario As Boolean, CodCliVarios As Long) As Boolean
Dim TiposIva As Byte
Dim vTipom As CTiposMov
Dim vCli As CCliente

        On Error GoTo EObteniendoDatosTickets


        ObtenerDatosTickets = False



        'En la tabla tmpspla pondre todos los importes por tp iva
        conn.Execute "DELETE from tmpinformes where codusu = " & vUsu.Codigo
        
        
        'Veo todos los importes y bases imponibles etc
        'Para no tener que hacer selects y demas me guardare que tipos de iva estoy tratatando
        '
        cadNomRPT = "|"
        TiposIva = 0
        For numParam = 1 To 3
            miSQL = "SELECT codigiv" & numParam & " tipodeiva,sum(baseimp" & numParam & ") labase,sum(imporiv" & numParam & ") importeiva FROM SCafac where "
            miSQL = miSQL & " intconta=0 and codtipom='FTI'"
            If Diario Then
                miSQL = miSQL & " AND fecfactu='" & devuelve & "'"
            Else
                'MOdificacion 13 - Agosto - 2008
                'Si no pongo esto suma tooooodas las facturas FTI que no esten contabilizadas
                'Desde
                If txtFecha(20).Text <> "" Then miSQL = miSQL & " AND fecfactu>='" & Format(txtFecha(20).Text, FormatoFecha) & "'"
                'El campo HASTA es obligado
                miSQL = miSQL & " AND fecfactu<='" & Format(txtFecha(21).Text, FormatoFecha) & "'"
            End If
            
            miSQL = miSQL & " group by 1 "
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
            
                If Not IsNull(miRsAux!tipodeiva) Then
                    ImpTot = DBLet(miRsAux!labase, "N")
                    ImpTeo = DBLet(miRsAux!ImporteIva, "N")
                    miSQL = "|" & miRsAux!tipodeiva & "|"
                    
                    If InStr(1, cadNomRPT, miSQL) > 0 Then
                        'YA LO HE INSERTADO. UPDATEO
                        miSQL = "UPDATE tmpinformes SET importe1=importe1 + " & TransformaComasPuntos(CStr(ImpTot))
                        miSQL = miSQL & " ,importe2=importe2 + " & TransformaComasPuntos(CStr(ImpTeo))
                        miSQL = miSQL & " WHERE codusu = " & vUsu.Codigo & " AND codigo1 = " & miRsAux!tipodeiva
                    Else
                        miSQL = "INSERT INTO `tmpinformes` (`codusu`,`codigo1`,`importe1`,importe2) values (" & vUsu.Codigo & "," & miRsAux!tipodeiva
                        miSQL = miSQL & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & ")"
                        TiposIva = TiposIva + 1
                        cadNomRPT = cadNomRPT & miRsAux!tipodeiva & "|"
                    End If
                    conn.Execute miSQL
                
                End If
                
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        Next numParam
        
        If TiposIva > 3 Or cadNomRPT = "" Then
            'ERROR  ERROR ERROR
            'ERROR en los tipos de iva. Hay mas de 3 o no hay ninguno
            If cadNomRPT = "" Then TiposIva = 0
            cadNomRPT = "Error en los tipos de IVA." & vbCrLf & "Total IVAS: " & TiposIva & vbCrLf & " Fec: " & devuelve
            MsgBox cadNomRPT, vbExclamation
            Exit Function
        End If
        
        'Ya tengo las bases ivas para las facturas
        'Ahora creo la FTG para poder utilizar las funciones ya realizadas
        Set vCli = New CCliente
        
        vCli.LeerDatos CStr(CodCliVarios)
        
             Set vTipom = New CTiposMov
             vTipom.Leer "FTG"
             vTipom.ConseguirContador vTipom.TipoMovimiento
             
             miSQL = "INSERT INTO `scafac` (`codtipom`,`numfactu`,`fecfactu`,`codclien`,`nomclien`,`domclien`,`codpobla`,"
             miSQL = miSQL & "`pobclien`,`proclien`,`nifclien`,`telclien`,`coddirec`,`nomdirec`,"
             miSQL = miSQL & "`codagent`,`codforpa`,`dtoppago`,`dtognral`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,"
             miSQL = miSQL & "`brutofac`,`impdtopp`,`impdtogr`,`intconta`,`totalfac`,"
             'LOS IVAS
             miSQL = miSQL & "`baseimp1`,`codigiv1`,`porciva1`,`imporiv1`,"
             miSQL = miSQL & "`baseimp2`,`codigiv2`,`porciva2`,`imporiv2`,"
             miSQL = miSQL & "`baseimp3`,`codigiv3`,`porciva3`,`imporiv3`)"
             
             'Cargo los ivas
             cadNomRPT = "Select codigo1,importe1,importe2 from tmpinformes where codusu = " & vUsu.Codigo
             miRsAux.Open cadNomRPT, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
             cadNomRPT = ""
             TiposIva = 0
             ImpTot = 0
             ImpTeo = 0
             While Not miRsAux.EOF
                 TiposIva = TiposIva + 1
                 Codigo = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", miRsAux!Codigo1)
                 cadFrom = "," & TransformaComasPuntos(CStr(miRsAux!Importe1)) & "," & miRsAux!Codigo1 & "," & TransformaComasPuntos(Codigo) & ","
                 cadFrom = cadFrom & TransformaComasPuntos(CStr(miRsAux!importe2))
                 
                 'Meto en el select
                 cadNomRPT = cadNomRPT & cadFrom
                 
                 'ImpTot
                 ImpTot = ImpTot + miRsAux!Importe1
                 ImpTeo = ImpTeo + miRsAux!importe2
                 miRsAux.MoveNext
             Wend
             miRsAux.Close
                 
                 
             'Si no tiene 3 tipos de ivas meter los null
             For numParam = TiposIva + 1 To 3
                 cadNomRPT = cadNomRPT & ",NULL,NULL,NULL,NULL"
             Next
             
             
             'Ahora relleno los datos que faltan
             'INSERT INTO `scafac` (`codtipom`,`numfactu`,`fecfactu`,`codclien`,`nomclien`,`domclien`,`codpobla`,"
             '`pobclien`,`proclien`,`nifclien`,`telclien`,`coddirec`,`nomdirec`,"
             '`codagent`,`codforpa`,`dtoppago`,`dtognral`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,"
             '`brutofac`,`impdtopp`,`impdtogr`,`intconta`,`totalfac`,"
                         
             cadFrom = " VALUES ('" & vTipom.TipoMovimiento & "'," & vTipom.Contador & ",'" & devuelve & "'," & vCli.Codigo
             cadFrom = cadFrom & ",'" & vCli.Nombre & "','','0','','','0',NULL,NULL,NULL" '0: codpos y nif
             'Agente:
             cadFrom = cadFrom & "," & vCli.Agente & "," & vCli.ForPago & ",0,0,NULL,NULL,NULL,NULL,"
             'Bruto factra
             cadFrom = cadFrom & "" & TransformaComasPuntos(CStr(ImpTot)) & ",0,0,0," & TransformaComasPuntos(CStr(ImpTot + ImpTeo))
              
             miSQL = miSQL & cadFrom & cadNomRPT & ")"
             conn.Execute miSQL
             
            'Si lleva la analitica metere una linea en slifac1 que es desde donde,
            ' el proceso de contabilizacion cojera EL CODTRABA para obtener el CC
                
                miSQL = "insert into `scafac1` (`codtipom`,`numfactu`,`fecfactu`,codtipoa,numalbar,`codenvio`,`codtraba`,`codtrab1`,`codtrab2`)"
                miSQL = miSQL & " VALUES ('FTG'," & vTipom.Contador & ",'" & devuelve & "','DAV','8',"  'Pongo tipoa y numalbar a piñon
                miSQL = miSQL & vParamAplic.PorDefecto_Envio & "," & txtTrab(2).Text & "," & txtTrab(2).Text & "," & txtTrab(2).Text & ")"
                conn.Execute miSQL
            
            
            
            
            'Ahora, despues de crear la factura temporal FTG, insertare en la tabla
            'que lleva la relacion, numfactura, codticket
            miSQL = "INSERT INTO sfactik(`numfacFTG`,`fecfacFTG`,`numfactu`,`fecfactu`,`codtraba`)"
            miSQL = miSQL & " SELECT " & vTipom.Contador & ",'" & devuelve & "',numfactu,fecfactu," & txtTrab(2).Text & " FROM scafac where "
            miSQL = miSQL & cadSelect
            If Diario Then miSQL = miSQL & " AND fecfactu='" & devuelve & "'"
            conn.Execute miSQL
             
            'Pongo la marca de contabilizado
            miSQL = "UPDATE scafac SET intconta = 1 WHERE " & cadSelect
            If Diario Then miSQL = miSQL & " AND fecfactu='" & devuelve & "'"
            conn.Execute miSQL
             
            vTipom.IncrementarContador vTipom.TipoMovimiento
            ObtenerDatosTickets = True
            

    

EObteniendoDatosTickets:

    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description & vbCrLf & miSQL
    End If
    Set vCli = Nothing
    Set vTipom = Nothing
End Function




'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Informe de trazabilidad
'       ========================================
'
'
'
'
'       A partir del desde /hasta mostraremos el informe que tiene la asociacion
'       entre albaranes de compra / venta
'
'
'       Hay datos tanto en albaranes como en facturas, con lo cual insertare sobre tmp
'
'
'----------------------------------------------------------------------
'----------------------------------------------------------------------
Private Sub HacerInformeTrazabilidad()

    
    
    InicializarVbles
    
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    
    If txtFecha(22).Text <> "" Or txtFecha(23).Text <> "" Then
        campo = "{slcomven.fechaalbc}"
        devuelve = "pDHFamilia=""Fecha: "
        If Not PonerDesdeHasta(campo, "F", 22, 23, devuelve) Then Exit Sub
    End If
    
    If txtCodProve(10).Text <> "" Or txtCodProve(11).Text <> "" Then
        campo = "{slcomven.codprovec}"
        devuelve = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(campo, "PRO", 10, 11, devuelve) Then Exit Sub
    End If
     
    If txtArticulo(4).Text <> "" Or txtArticulo(5).Text <> "" Then
        campo = "{slcomven.codartic}"
        devuelve = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(campo, "ART", 4, 5, devuelve) Then Exit Sub
    End If
    
    
    'Montamos el select para los registros
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    If cadSelect = "" Then cadSelect = " 1 = 1 "
    campo = "slcomven WHERE  " & cadSelect
    
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay facturas con estos valores", vbExclamation
        Exit Sub
    End If
    
    
    
    cadNomRPT = "rTraza.rpt"
    LlamarImprimir False
    
End Sub




Private Sub CargaTablasCambio()


    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open "show tables", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Me.cboTablas.AddItem miRsAux.Fields(0)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing


End Sub



Private Sub CargarCamposTabla()
'Dim Cad As String
'Dim Aux As String
Dim RS As ADODB.Recordset
Dim i As Integer
Dim TieneClaves As Boolean

    
    miSQL = "Select * from " & Me.cboTablas.List(cboTablas.ListIndex) & " LIMIT 1,1"
    Set RS = New ADODB.Recordset
    RS.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
 
        TieneClaves = False
        For i = 0 To RS.Fields.Count - 1
           
            
            
            'SOLO TEXTOS
            If RS.Fields(i).Type = 129 Or RS.Fields(i).Type = 200 Or RS.Fields(i).Type = adVarChar Then
    
       
  
                If RS.Fields(i).Properties(18).Value Then
                    'NO HACEMOS NADA. Es campo clave
                
                Else
                    cboCampos.AddItem RS.Fields(i).Name
                End If
                
            End If
            
            'Para saber si tiene claves
            If RS.Fields(i).Properties(18).Value Then TieneClaves = True
            
        Next i
        
        
        
    RS.Close
    Set RS = Nothing

    If cboCampos.ListCount > 0 And Not TieneClaves Then
        MsgBox "No tiene campos clave", vbInformation
        Me.cboCampos.Clear
    End If
End Sub




Private Sub UpdatearTablaRoot()
Dim i As Integer
Dim TienDatos As Boolean

    On Error GoTo EUpdatearTablaRoot
    
    devuelve = Me.cboTablas.List(cboTablas.ListIndex)
    miSQL = "Select " & Me.cboCampos.List(cboCampos.ListIndex) & "," & devuelve & ".* from " & devuelve

    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadFrom = ""
    miSQL = ""
    While Not miRsAux.EOF
        If IsNull(miRsAux.Fields(0)) Then
            Me.lblMultibase.Caption = ""
            Me.lblMultibase.Refresh
        Else
            miSQL = miRsAux.Fields(0)
            Me.lblMultibase.Caption = miSQL
            Me.lblMultibase.Refresh
            devuelve = RevisaCaracterMultibase(miSQL)
            
            If miSQL <> devuelve Then
                    'La clave
                    cadFrom = ""
                    For i = 0 To miRsAux.Fields.Count - 1
                        If miRsAux.Fields(i).Properties(18).Value Then
                            Select Case miRsAux.Fields(i).Type
                            Case 133
                                campo = CStr(miRsAux.Fields(i))
                                campo = "'" & Format(campo, "yyyy-mm-dd") & "'"
            
                            Case 135 'Fecha/Hora
                                campo = DBSet(miRsAux.Fields(i), "FH", "S")
                            'Numero normal, sin decimales
                            Case 2, 3, 16 To 19
                                campo = miRsAux.Fields(i)
                            Case 129, 200
                                campo = DBSet(miRsAux.Fields(i), "T")
                            Case Else
                                MsgBox "No tratado: " & miRsAux.Fields(i).Type, vbExclamation
                                Exit Sub
                                
                            End Select
                            cadFrom = cadFrom & " AND " & miRsAux.Fields(i).Name & " = " & campo
                        End If
                    Next i
                    cadFrom = Mid(cadFrom, 6)
                    devuelve = DevNombreSQL(devuelve)
                    miSQL = "UPDATE " & Me.cboTablas.List(cboTablas.ListIndex) & " SET " & Me.cboCampos.List(cboCampos.ListIndex)
                    miSQL = miSQL & " = '" & devuelve & "' WHERE " & cadFrom
                    conn.Execute miSQL
            End If 'DEl campo <>
        End If 'de ISNULL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'If miSQL <> "" Then
        MsgBox "Proceso finalizado", vbInformation
    'Else
    '    MsgBox "No hay registros", vbInformation
    'End If
    Exit Sub
EUpdatearTablaRoot:
    MuestraError Err.Number, Err.Description
End Sub


Private Sub CargarOtrasOfertas()
Dim It 'As ListItem
    Me.lw1.ListItems.Clear
    lblDpto(27).Caption = miRsAux!nomclien
    While Not miRsAux.EOF
        Set It = lw1.ListItems.Add()
        It.Text = Format(miRsAux!NumOfert, "000000")
        It.SubItems(1) = Format(miRsAux!fecofert, "dd/mm/yyyy")
        It.SubItems(2) = Format(miRsAux!FecEntre, "dd/mm/yyyy")
        It.SubItems(3) = DBLet(miRsAux!nomdirec, "T") & " "
        If Val(miRsAux!aceptado) = 0 Then
            It.Checked = True
        Else
            It.Checked = False
        End If
        miRsAux.MoveNext
    Wend
    
End Sub


Private Sub CargaListMov()
Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    If Me.Opcion = 23 Then
        'Estoy cargando el list de las fras
        Me.List1.Clear
        miSQL = "select * from stipom where codtipom like 'AL%' order by codtipom"
        R.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not R.EOF
            Me.List1.AddItem R!nomtipom & " (" & R!codtipom & ")"
            List1.Selected(List1.NewIndex) = True
            R.MoveNext
        Wend
        R.Close
        
    End If
    Set R = Nothing
End Sub
