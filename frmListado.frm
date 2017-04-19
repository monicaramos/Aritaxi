VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   11175
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   10680
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameDtosFM 
      Height          =   5415
      Left            =   510
      TabIndex        =   276
      Top             =   600
      Width           =   6915
      Begin VB.OptionButton optFrDto 
         Caption         =   "Marca"
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
         Left            =   3570
         TabIndex        =   523
         Top             =   4920
         Width           =   1095
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Familia"
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
         Left            =   2520
         TabIndex        =   522
         Top             =   4920
         Width           =   975
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Actividad"
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
         Left            =   1200
         TabIndex        =   521
         Top             =   4920
         Width           =   1455
      End
      Begin VB.OptionButton optFrDto 
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
         Index           =   0
         Left            =   120
         TabIndex        =   520
         Top             =   4920
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   313
         Top             =   810
         Width           =   6135
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
            Index           =   74
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   315
            Text            =   "Text5"
            Top             =   720
            Width           =   3645
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
            Index           =   74
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   266
            Top             =   720
            Width           =   975
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
            Index           =   73
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   314
            Text            =   "Text5"
            Top             =   360
            Width           =   3645
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
            Index           =   73
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   265
            Top             =   360
            Width           =   975
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   1
            Left            =   1275
            ToolTipText     =   "Buscar cliente"
            Top             =   720
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
            Index           =   61
            Left            =   540
            TabIndex        =   318
            Top             =   360
            Width           =   585
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   0
            Left            =   1275
            ToolTipText     =   "Buscar cliente"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
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
            Index           =   44
            Left            =   240
            TabIndex        =   317
            Top             =   120
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
            Index           =   45
            Left            =   540
            TabIndex        =   316
            Top             =   720
            Width           =   690
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   1005
         Left            =   360
         TabIndex        =   307
         Top             =   2820
         Width           =   6135
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
            Index           =   77
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   269
            Top             =   270
            Width           =   735
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
            Index           =   78
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   270
            Top             =   630
            Width           =   735
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
            Index           =   77
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   309
            Text            =   "Text5"
            Top             =   270
            Width           =   3825
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
            Index           =   78
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   308
            Text            =   "Text5"
            Top             =   630
            Width           =   3825
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
            Index           =   66
            Left            =   540
            TabIndex        =   312
            Top             =   270
            Width           =   675
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
            Index           =   67
            Left            =   540
            TabIndex        =   311
            Top             =   630
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
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
            Index           =   42
            Left            =   240
            TabIndex        =   310
            Top             =   30
            Width           =   645
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   4
            Left            =   1275
            ToolTipText     =   "Buscar marca"
            Top             =   270
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   5
            Left            =   1275
            ToolTipText     =   "Buscar marca"
            Top             =   630
            Width           =   240
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   301
         Top             =   3720
         Width           =   6255
         Begin VB.TextBox txtCodigo 
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
            Index           =   79
            Left            =   1560
            TabIndex        =   263
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   80
            Left            =   1560
            TabIndex        =   264
            Top             =   720
            Width           =   975
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
            Index           =   79
            Left            =   2565
            Locked          =   -1  'True
            TabIndex        =   303
            Text            =   "Text5"
            Top             =   360
            Width           =   3585
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
            Index           =   80
            Left            =   2565
            Locked          =   -1  'True
            TabIndex        =   302
            Text            =   "Text5"
            Top             =   720
            Width           =   3585
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
            Index           =   46
            Left            =   240
            TabIndex        =   304
            Top             =   120
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
            Index           =   65
            Left            =   540
            TabIndex        =   306
            Top             =   360
            Width           =   705
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
            Index           =   64
            Left            =   540
            TabIndex        =   305
            Top             =   720
            Width           =   660
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   63
            Left            =   1275
            ToolTipText     =   "Buscar proveedor"
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   64
            Left            =   1275
            ToolTipText     =   "Buscar proveedor"
            Top             =   720
            Width           =   240
         End
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
         Index           =   12
         Left            =   5760
         TabIndex        =   273
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarDtosFM 
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
         Left            =   4680
         TabIndex        =   272
         Top             =   4920
         Width           =   975
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
         Index           =   75
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   267
         Top             =   2160
         Width           =   735
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
         Index           =   76
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   268
         Top             =   2520
         Width           =   735
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
         Index           =   75
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   278
         Text            =   "Text5"
         Top             =   2160
         Width           =   3825
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
         Index           =   76
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   277
         Text            =   "Text5"
         Top             =   2520
         Width           =   3825
      End
      Begin VB.Label Label10 
         Caption         =   "Listado Descuentos Familia/Marca"
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
         Left            =   480
         TabIndex        =   282
         Top             =   360
         Width           =   5655
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
         Index           =   63
         Left            =   900
         TabIndex        =   281
         Top             =   2160
         Width           =   675
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
         Index           =   62
         Left            =   900
         TabIndex        =   280
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Index           =   40
         Left            =   600
         TabIndex        =   279
         Top             =   1920
         Width           =   780
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   2
         Left            =   1635
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   3
         Left            =   1635
         ToolTipText     =   "Buscar familia"
         Top             =   2520
         Width           =   240
      End
   End
   Begin VB.Frame FrameEtiqEstanteria 
      ForeColor       =   &H00000000&
      Height          =   5175
      Left            =   0
      TabIndex        =   366
      Top             =   0
      Width           =   7815
      Begin VB.CheckBox chkDtoFM 
         Caption         =   "Mostrar descuento fam/marca"
         Height          =   255
         Left            =   960
         TabIndex        =   375
         Top             =   4680
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   124
         Left            =   4140
         TabIndex        =   372
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   123
         Left            =   1800
         TabIndex        =   371
         Top             =   3480
         Width           =   1095
      End
      Begin VB.ComboBox cboDecimal 
         Height          =   315
         ItemData        =   "frmListado.frx":000C
         Left            =   1800
         List            =   "frmListado.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   373
         Top             =   4080
         Width           =   735
      End
      Begin VB.CheckBox chkImprimeCodigoBarras 
         Caption         =   "Impime codigo barras"
         Height          =   255
         Left            =   2760
         TabIndex        =   374
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   95
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   381
         Text            =   "Text5"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   94
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   380
         Text            =   "Text5"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   95
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   368
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   94
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   367
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   93
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   379
         Text            =   "Text5"
         Top             =   2760
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   92
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   377
         Text            =   "Text5"
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   93
         Left            =   1800
         TabIndex        =   370
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   92
         Left            =   1800
         TabIndex        =   369
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdEtiqEstanteria 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5640
         TabIndex        =   376
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   94
         Left            =   6720
         TabIndex        =   378
         Top             =   4560
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   3840
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1515
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   96
         Left            =   3315
         TabIndex        =   493
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   492
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha ult. cambio precio P.V.P."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   89
         Left            =   480
         TabIndex        =   491
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Decimales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   57
         Left            =   480
         TabIndex        =   389
         Top             =   4200
         Width           =   870
      End
      Begin VB.Label lblTituloEst 
         Caption         =   "Etiquetas estanterias"
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
         Left            =   480
         TabIndex        =   388
         Top             =   360
         Width           =   5895
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   74
         Left            =   1515
         ToolTipText     =   "Buscar familia"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   73
         Left            =   1515
         ToolTipText     =   "Buscar familia"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   56
         Left            =   480
         TabIndex        =   387
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   78
         Left            =   960
         TabIndex        =   386
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   77
         Left            =   960
         TabIndex        =   385
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   72
         Left            =   1515
         ToolTipText     =   "Buscar artículo"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   71
         Left            =   1515
         ToolTipText     =   "Buscar artículo"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   55
         Left            =   480
         TabIndex        =   384
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   76
         Left            =   960
         TabIndex        =   383
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   75
         Left            =   960
         TabIndex        =   382
         Top             =   2400
         Width           =   465
      End
   End
   Begin VB.Frame FrameFrecuencia 
      Height          =   3855
      Left            =   120
      TabIndex        =   409
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame13 
         Height          =   615
         Left            =   360
         TabIndex        =   517
         Top             =   2880
         Width           =   2655
         Begin VB.OptionButton OptFrecFicha 
            Caption         =   "Ficha"
            Height          =   255
            Left            =   120
            TabIndex        =   519
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptFrecResumen 
            Caption         =   "Resumen"
            Height          =   255
            Left            =   1320
            TabIndex        =   518
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   99
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   420
         Text            =   "Text5"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   99
         Left            =   1320
         TabIndex        =   412
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   101
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   419
         Text            =   "Text5"
         Top             =   2400
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   100
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   418
         Text            =   "Text5"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   101
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   414
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   100
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   413
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   98
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   417
         Text            =   "Text5"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   98
         Left            =   1320
         TabIndex        =   411
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdFrecuencias 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   415
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   96
         Left            =   4800
         TabIndex        =   416
         Top             =   3120
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   77
         Left            =   1035
         ToolTipText     =   "Buscar cliente"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   81
         Left            =   480
         TabIndex        =   426
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   79
         Left            =   1035
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2400
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   78
         Left            =   1035
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2040
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direc/Dpto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   67
         Left            =   120
         TabIndex        =   425
         Top             =   1800
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   80
         Left            =   480
         TabIndex        =   424
         Top             =   2400
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   79
         Left            =   480
         TabIndex        =   423
         Top             =   2040
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   76
         Left            =   1035
         ToolTipText     =   "Buscar cliente"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   66
         Left            =   120
         TabIndex        =   422
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   65
         Left            =   480
         TabIndex        =   421
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Datos de frecuencias  clientes"
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
         Index           =   5
         Left            =   480
         TabIndex        =   410
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame FrameEstMargenes 
      Height          =   5295
      Left            =   2400
      TabIndex        =   342
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   131
         Left            =   5040
         TabIndex        =   347
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   130
         Left            =   1800
         TabIndex        =   346
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Frame FrameValorar2 
         Caption         =   "Valorar Con:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   360
         TabIndex        =   362
         Top             =   3960
         Width           =   2535
         Begin VB.OptionButton optPrecioMP2 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Left            =   240
            TabIndex        =   365
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioUC2 
            Caption         =   "Precio Ultima Compra"
            Height          =   255
            Left            =   240
            TabIndex        =   364
            Top             =   525
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioStd2 
            Caption         =   "Precio Standard"
            Height          =   255
            Left            =   240
            TabIndex        =   363
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   6360
         TabIndex        =   353
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarEst 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   352
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   90
         Left            =   1800
         TabIndex        =   350
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   91
         Left            =   1800
         TabIndex        =   351
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   90
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   358
         Text            =   "Text5"
         Top             =   2880
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   91
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   357
         Text            =   "Text5"
         Top             =   3240
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   88
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   348
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   89
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   349
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   88
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   345
         Text            =   "Text5"
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   89
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   344
         Text            =   "Text5"
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   100
         Left            =   4200
         TabIndex        =   526
         Top             =   1200
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   4680
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   99
         Left            =   960
         TabIndex        =   525
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   1440
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   95
         Left            =   480
         TabIndex        =   524
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   74
         Left            =   960
         TabIndex        =   361
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   73
         Left            =   960
         TabIndex        =   360
         Top             =   3240
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   54
         Left            =   480
         TabIndex        =   359
         Top             =   2640
         Width           =   660
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   69
         Left            =   1515
         ToolTipText     =   "Buscar artículo"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   70
         Left            =   1515
         ToolTipText     =   "Buscar artículo"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   72
         Left            =   960
         TabIndex        =   356
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   71
         Left            =   960
         TabIndex        =   355
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   53
         Left            =   480
         TabIndex        =   354
         Top             =   1560
         Width           =   600
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   67
         Left            =   1515
         ToolTipText     =   "Buscar familia"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   68
         Left            =   1515
         ToolTipText     =   "buscar familia"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label lblTituloEst 
         Caption         =   "Informe Margenes de Venta por Artículo"
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
         Left            =   720
         TabIndex        =   343
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame FrameTarifas 
      Height          =   6375
      Left            =   480
      TabIndex        =   97
      Top             =   120
      Width           =   7635
      Begin VB.ComboBox cboDecimales 
         Height          =   315
         ItemData        =   "frmListado.frx":0032
         Left            =   3480
         List            =   "frmListado.frx":0045
         Style           =   2  'Dropdown List
         TabIndex        =   485
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CheckBox chkMostrarErrores 
         Caption         =   "Mostrar solo tarifas con error"
         Height          =   255
         Left            =   960
         TabIndex        =   341
         Top             =   5880
         Width           =   2415
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   24
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   139
         Text            =   "Text5"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1920
         TabIndex        =   99
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkSaltaPagTarif 
         Caption         =   "Salta pág. en Familia"
         Height          =   255
         Left            =   960
         TabIndex        =   115
         Top             =   5520
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   26
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "Text5"
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   25
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Text5"
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   30
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "Text5"
         Top             =   4680
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   29
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   111
         Text            =   "Text5"
         Top             =   4320
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   101
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   100
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   30
         Left            =   1920
         TabIndex        =   105
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   29
         Left            =   1920
         TabIndex        =   104
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptarTarif 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5280
         TabIndex        =   106
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   6360
         TabIndex        =   107
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1920
         TabIndex        =   102
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1920
         TabIndex        =   103
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   27
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   110
         Text            =   "Text5"
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   28
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   109
         Text            =   "Text5"
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   23
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   108
         Text            =   "Text5"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1920
         TabIndex        =   98
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Decimales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   88
         Left            =   3480
         TabIndex        =   484
         Top             =   5160
         Width           =   870
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   56
         Left            =   1635
         ToolTipText     =   "Buscar tarifa"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   21
         Left            =   1080
         TabIndex        =   138
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   58
         Left            =   1635
         ToolTipText     =   "Buscar familia"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   57
         Left            =   1635
         ToolTipText     =   "Buscar familia"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   600
         TabIndex        =   128
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   18
         Left            =   1080
         TabIndex        =   127
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   1080
         TabIndex        =   126
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   62
         Left            =   1635
         ToolTipText     =   "Buscar artículo"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   61
         Left            =   1635
         ToolTipText     =   "Buscar artículo"
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   600
         TabIndex        =   125
         Top             =   4080
         Width           =   660
      End
      Begin VB.Label lblTituloTarif 
         Caption         =   "Informe Precios y Descuentos"
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
         Left            =   480
         TabIndex        =   124
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   16
         Left            =   1080
         TabIndex        =   123
         Top             =   4680
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   1080
         TabIndex        =   122
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   1080
         TabIndex        =   121
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   1080
         TabIndex        =   120
         Top             =   3600
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   600
         TabIndex        =   119
         Top             =   3000
         Width           =   525
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   59
         Left            =   1635
         ToolTipText     =   "Buscar marca"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   60
         Left            =   1635
         ToolTipText     =   "Buscar marca"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Opciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   600
         TabIndex        =   118
         Top             =   5160
         Width           =   765
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   55
         Left            =   1635
         ToolTipText     =   "Buscar tarifa"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   600
         TabIndex        =   117
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   116
         Top             =   1560
         Width           =   420
      End
   End
   Begin VB.Frame FrameAlbaranesMarcaFacturar 
      Height          =   3735
      Left            =   0
      TabIndex        =   467
      Top             =   0
      Width           =   6495
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         TabIndex        =   483
         Top             =   1680
         Width           =   6135
      End
      Begin VB.CommandButton cmdFactAlbaranes 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   473
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   82
         Left            =   5160
         TabIndex        =   474
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   120
         Left            =   3960
         TabIndex        =   470
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   119
         Left            =   1680
         TabIndex        =   469
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   118
         Left            =   1680
         TabIndex        =   472
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   118
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   476
         Text            =   "Text5"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   117
         Left            =   1680
         TabIndex        =   471
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   117
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   475
         Text            =   "Text5"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   3600
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   87
         Left            =   3000
         TabIndex        =   482
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   86
         Left            =   240
         TabIndex        =   481
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   93
         Left            =   720
         TabIndex        =   480
         Top             =   1440
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1320
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   85
         Left            =   840
         TabIndex        =   479
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   84
         Left            =   360
         TabIndex        =   478
         Top             =   1920
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   94
         Left            =   1395
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   92
         Left            =   840
         TabIndex        =   477
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   93
         Left            =   1395
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Marcar facturar albaranes"
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
         Left            =   360
         TabIndex        =   468
         Top             =   360
         Width           =   5775
      End
   End
   Begin VB.Frame FrameInvArtComp 
      Height          =   4335
      Left            =   360
      TabIndex        =   501
      Top             =   1320
      Width           =   7935
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   506
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5280
         TabIndex        =   516
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   126
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   514
         Text            =   "Text5"
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   126
         Left            =   1680
         MaxLength       =   16
         TabIndex        =   505
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   125
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   511
         Text            =   "Text5"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   125
         Left            =   1680
         MaxLength       =   16
         TabIndex        =   504
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Frame FrameValorar3 
         Caption         =   "Valorar Con:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   360
         TabIndex        =   503
         Top             =   2280
         Width           =   2535
         Begin VB.OptionButton optPrecioMP3 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Left            =   240
            TabIndex        =   507
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioMA3 
            Caption         =   "Precio Medio Acumulado"
            Height          =   255
            Left            =   240
            TabIndex        =   508
            Top             =   560
            Width           =   2175
         End
         Begin VB.OptionButton optPrecioUC3 
            Caption         =   "Precio Ultima Compra"
            Height          =   255
            Left            =   240
            TabIndex        =   509
            Top             =   880
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioStd3 
            Caption         =   "Precio Standard"
            Height          =   255
            Left            =   240
            TabIndex        =   510
            Top             =   1200
            Width           =   2055
         End
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   98
         Left            =   1395
         ToolTipText     =   "Buscar artículo"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   98
         Left            =   840
         TabIndex        =   515
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   95
         Left            =   1395
         ToolTipText     =   "Buscar artículo"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   92
         Left            =   360
         TabIndex        =   513
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   97
         Left            =   840
         TabIndex        =   512
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label12 
         Caption         =   "Listado Artículos con componentes"
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
         Left            =   240
         TabIndex        =   502
         Top             =   360
         Width           =   6855
      End
   End
   Begin VB.Frame FrameMovArtic 
      Height          =   5535
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   10635
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   86
         Left            =   1485
         TabIndex        =   25
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   87
         Left            =   1485
         TabIndex        =   26
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   5
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   17
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeselTodos 
         Height          =   435
         Left            =   9000
         Picture         =   "frmListado.frx":0066
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   740
         Width           =   585
      End
      Begin VB.CommandButton cmdSelTodos 
         Height          =   435
         Left            =   9720
         Picture         =   "frmListado.frx":0750
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   740
         Width           =   585
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   6960
         TabIndex        =   27
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   3960
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "Text5"
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text5"
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text5"
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Text5"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text5"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   24
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   23
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   3600
         TabIndex        =   22
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1440
         TabIndex        =   21
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   20
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   19
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   6
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   28
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   29
         Top             =   4800
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   65
         Left            =   1200
         ToolTipText     =   "Cliente"
         Top             =   4560
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   66
         Left            =   1200
         Top             =   4920
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   70
         Left            =   600
         TabIndex        =   340
         Top             =   4560
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   69
         Left            =   600
         TabIndex        =   339
         Top             =   4920
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente/Proveedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   68
         Left            =   360
         TabIndex        =   338
         Top             =   4320
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipos de Movimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   6960
         TabIndex        =   61
         Top             =   960
         Width           =   1755
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3315
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1155
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   34
         Left            =   1155
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   33
         Left            =   1155
         ToolTipText     =   "Buscar almacen"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Almacen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   60
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   59
         Top             =   3960
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   58
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   57
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   2760
         TabIndex        =   56
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   43
         Top             =   3000
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   32
         Left            =   1155
         ToolTipText     =   "Buscar familia"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   31
         Left            =   1155
         ToolTipText     =   "Buscar familia"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   42
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   41
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   40
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   30
         Left            =   1155
         ToolTipText     =   "Buscar artículo"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   29
         Left            =   1155
         ToolTipText     =   "Buscar artículo"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   39
         Top             =   840
         Width           =   660
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Informes Movimiento Artículos"
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
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   600
         TabIndex        =   37
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   36
         Top             =   1080
         Width           =   465
      End
   End
   Begin VB.Frame FrameInfArticulos 
      Height          =   6855
      Left            =   240
      TabIndex        =   218
      Top             =   0
      Width           =   7395
      Begin VB.Frame FrameSituacionArticulo 
         Caption         =   "Situación artículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   360
         TabIndex        =   496
         Top             =   5880
         Width           =   4575
         Begin VB.CheckBox chkSitaucionArticulo 
            Caption         =   "Caducado"
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   499
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSitaucionArticulo 
            Caption         =   "Bloqueado"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   498
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSitaucionArticulo 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   497
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkMinimoCorreg 
         Caption         =   "No mostrar tarifas por encima de margen"
         Height          =   195
         Left            =   600
         TabIndex        =   450
         Top             =   5280
         Width           =   6015
      End
      Begin VB.Frame FrameStockMaxMin 
         Caption         =   "Imprimir Stocks"
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
         Height          =   615
         Left            =   360
         TabIndex        =   275
         Top             =   5880
         Width           =   4455
         Begin VB.OptionButton optPuntoPedido 
            Caption         =   "Punto de pedido"
            Height          =   255
            Left            =   2520
            TabIndex        =   235
            Top             =   280
            Width           =   1575
         End
         Begin VB.OptionButton optStockMin 
            Caption         =   "Mínimos"
            Height          =   255
            Left            =   1320
            TabIndex        =   234
            Top             =   280
            Width           =   975
         End
         Begin VB.OptionButton optStockMax 
            Caption         =   "Máximos"
            Height          =   255
            Left            =   120
            TabIndex        =   233
            Top             =   280
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbDecimales 
         Height          =   315
         ItemData        =   "frmListado.frx":0E3A
         Left            =   600
         List            =   "frmListado.frx":0E47
         Style           =   2  'Dropdown List
         TabIndex        =   237
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Frame FrameTapaINCORRECTO 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1440
         TabIndex        =   439
         Top             =   840
         Width           =   4215
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   107
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   440
            Text            =   "Text5"
            Top             =   45
            Width           =   3015
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   107
            Left            =   360
            MaxLength       =   4
            TabIndex        =   221
            Top             =   45
            Width           =   615
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   87
            Left            =   80
            ToolTipText     =   "Buscar almacen"
            Top             =   60
            Width           =   240
         End
      End
      Begin VB.Frame FrameOrden 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   5760
         TabIndex        =   319
         Top             =   840
         Width           =   2655
         Begin VB.CommandButton cmdBajar 
            Height          =   510
            Left            =   1920
            Picture         =   "frmListado.frx":0E66
            Style           =   1  'Graphical
            TabIndex        =   321
            Top             =   1305
            Width           =   510
         End
         Begin VB.CommandButton cmdSubir 
            Caption         =   " "
            Height          =   510
            Left            =   1920
            Picture         =   "frmListado.frx":1170
            Style           =   1  'Graphical
            TabIndex        =   320
            Top             =   600
            Width           =   510
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1335
            Left            =   120
            TabIndex        =   322
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   2355
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Orden del Informe"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   31
            Left            =   120
            TabIndex        =   323
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   72
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   222
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   72
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   271
         Text            =   "Text5"
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   69
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   259
         Text            =   "Text5"
         Top             =   4470
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   68
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   258
         Text            =   "Text5"
         Top             =   4150
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   69
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   230
         Top             =   4470
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   68
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   229
         Top             =   4155
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   65
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   254
         Text            =   "Text5"
         Top             =   2590
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   64
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   253
         Text            =   "Text5"
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   65
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   226
         Top             =   2590
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   64
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   225
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   63
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   242
         Text            =   "Text5"
         Top             =   1750
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   62
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   241
         Text            =   "Text5"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   71
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   240
         Text            =   "Text5"
         Top             =   5400
         Width           =   3855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   70
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   239
         Text            =   "Text5"
         Top             =   5080
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   63
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   224
         Top             =   1750
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   62
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   223
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   71
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   232
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   70
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   231
         Top             =   5080
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptarArtic 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   236
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   6240
         TabIndex        =   238
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   66
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   227
         Top             =   3190
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   67
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   228
         Top             =   3510
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   66
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   220
         Text            =   "Text5"
         Top             =   3190
         Width           =   4575
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   67
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   219
         Text            =   "Text5"
         Top             =   3510
         Width           =   4575
      End
      Begin VB.ComboBox cmbProduccion 
         Height          =   315
         ItemData        =   "frmListado.frx":147A
         Left            =   2280
         List            =   "frmListado.frx":1484
         Style           =   2  'Dropdown List
         TabIndex        =   494
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Verificar sobre"
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
         Index           =   90
         Left            =   2280
         TabIndex        =   495
         Top             =   5880
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Decimales"
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
         Index           =   75
         Left            =   600
         TabIndex        =   441
         Top             =   5880
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   39
         Left            =   600
         TabIndex        =   251
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   36
         Left            =   600
         TabIndex        =   274
         Top             =   890
         Width           =   735
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   18
         Left            =   1515
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   26
         Left            =   1515
         ToolTipText     =   "Buscar tipo articulo"
         Top             =   4485
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   25
         Left            =   1515
         ToolTipText     =   "Buscar tipo articulo"
         Top             =   4155
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Articulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   30
         Left            =   600
         TabIndex        =   262
         Top             =   3900
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   60
         Left            =   960
         TabIndex        =   261
         Top             =   4470
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   59
         Left            =   960
         TabIndex        =   260
         Top             =   4155
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   22
         Left            =   1515
         ToolTipText     =   "Buscar marca"
         Top             =   2625
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   21
         Left            =   1515
         ToolTipText     =   "Buscar marca"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   35
         Left            =   600
         TabIndex        =   257
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   58
         Left            =   960
         TabIndex        =   256
         Top             =   2595
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   57
         Left            =   960
         TabIndex        =   255
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label9 
         Caption         =   "Informe de Articulos"
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
         Left            =   600
         TabIndex        =   252
         Top             =   360
         Width           =   6735
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   20
         Left            =   1515
         ToolTipText     =   "Buscar familia"
         Top             =   1770
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   19
         Left            =   1515
         ToolTipText     =   "Buscar familia"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   56
         Left            =   960
         TabIndex        =   250
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   55
         Left            =   960
         TabIndex        =   249
         Top             =   1440
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   28
         Left            =   1515
         ToolTipText     =   "Buscar artículo"
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   27
         Left            =   1515
         ToolTipText     =   "Buscar artículo"
         Top             =   5085
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   38
         Left            =   600
         TabIndex        =   248
         Top             =   4820
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   54
         Left            =   960
         TabIndex        =   247
         Top             =   5400
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   51
         Left            =   960
         TabIndex        =   246
         Top             =   5085
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   50
         Left            =   960
         TabIndex        =   245
         Top             =   3195
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   48
         Left            =   960
         TabIndex        =   244
         Top             =   3510
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   37
         Left            =   600
         TabIndex        =   243
         Top             =   2950
         Width           =   885
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   23
         Left            =   1515
         ToolTipText     =   "Buscar proveedor"
         Top             =   3195
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   24
         Left            =   1515
         ToolTipText     =   "Buscar proveedor"
         Top             =   3540
         Width           =   240
      End
   End
   Begin VB.Frame FrameRepNSerie 
      Height          =   4995
      Left            =   240
      TabIndex        =   141
      Top             =   60
      Width           =   6795
      Begin VB.CheckBox Check2 
         Caption         =   "Incluir los de baja"
         Height          =   195
         Left            =   570
         TabIndex        =   563
         Top             =   4410
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Clasificado por Artículo"
         Height          =   195
         Left            =   570
         TabIndex        =   135
         Top             =   4020
         Width           =   6015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   42
         Left            =   3435
         Locked          =   -1  'True
         TabIndex        =   533
         Text            =   "Text5"
         Top             =   3540
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   41
         Left            =   3435
         Locked          =   -1  'True
         TabIndex        =   532
         Text            =   "Text5"
         Top             =   3180
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   42
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   134
         Top             =   3540
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   41
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   133
         Top             =   3180
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   40
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   528
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   39
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   527
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   132
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   131
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1920
         TabIndex        =   129
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   37
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   143
         Text            =   "Text5"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   5430
         TabIndex        =   137
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarNSerie 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4380
         TabIndex        =   136
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1920
         TabIndex        =   130
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   38
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   142
         Text            =   "Text5"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   36
         Left            =   1635
         ToolTipText     =   "Buscar artículo"
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   35
         Left            =   1635
         ToolTipText     =   "Buscar artículo"
         Top             =   3180
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   600
         TabIndex        =   536
         Top             =   2940
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   32
         Left            =   1080
         TabIndex        =   535
         Top             =   3540
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   31
         Left            =   1080
         TabIndex        =   534
         Top             =   3180
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   52
         Left            =   1635
         ToolTipText     =   "Buscar tipo artículo"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   51
         Left            =   1635
         ToolTipText     =   "Buscar tipo artículo"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Artículo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   600
         TabIndex        =   531
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   27
         Left            =   1080
         TabIndex        =   530
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   28
         Left            =   1080
         TabIndex        =   529
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   1080
         TabIndex        =   150
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   19
         Left            =   600
         TabIndex        =   148
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   49
         Left            =   1635
         ToolTipText     =   "Buscar socio"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Equipamiento Socios"
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
         Left            =   600
         TabIndex        =   145
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   1080
         TabIndex        =   144
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   50
         Left            =   1635
         ToolTipText     =   "Buscar socio"
         Top             =   1680
         Width           =   240
      End
   End
   Begin VB.Frame FrameServicios 
      Height          =   4995
      Left            =   0
      TabIndex        =   537
      Top             =   0
      Width           =   6795
      Begin VB.Frame Frame8 
         Caption         =   "Clasificado por"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   600
         TabIndex        =   561
         Top             =   3630
         Width           =   3345
         Begin VB.OptionButton Option1 
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
            Height          =   285
            Index           =   1
            Left            =   1770
            TabIndex        =   562
            Top             =   300
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Socio "
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
            Left            =   480
            TabIndex        =   547
            Top             =   300
            Width           =   855
         End
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
         Index           =   56
         Left            =   4050
         TabIndex        =   546
         Top             =   3240
         Width           =   1350
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
         Index           =   55
         Left            =   1710
         TabIndex        =   545
         Top             =   3240
         Width           =   1350
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
         Height          =   285
         Index           =   60
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   549
         Text            =   "Text5"
         Top             =   1680
         Width           =   3735
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
         Index           =   60
         Left            =   1920
         TabIndex        =   541
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton CmdAceptarServicios 
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
         Left            =   4380
         TabIndex        =   548
         Top             =   4320
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
         Index           =   10
         Left            =   5430
         TabIndex        =   550
         Top             =   4320
         Width           =   975
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
         Index           =   59
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   542
         Text            =   "Text5"
         Top             =   1320
         Width           =   3735
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
         Index           =   59
         Left            =   1920
         TabIndex        =   540
         Top             =   1320
         Width           =   735
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
         Index           =   58
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   544
         Top             =   2640
         Width           =   735
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
         Index           =   57
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   543
         Top             =   2280
         Width           =   735
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
         Index           =   57
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   539
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
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
         Index           =   58
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   538
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
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
         Index           =   84
         Left            =   720
         TabIndex        =   560
         Top             =   3240
         Width           =   675
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
         Index           =   47
         Left            =   3180
         TabIndex        =   559
         Top             =   3240
         Width           =   570
      End
      Begin VB.Label Label3 
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
         Index           =   46
         Left            =   600
         TabIndex        =   558
         Top             =   3000
         Width           =   630
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   3780
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1410
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   54
         Left            =   1635
         ToolTipText     =   "Buscar socio"
         Top             =   1680
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
         Index           =   53
         Left            =   900
         TabIndex        =   557
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label8 
         Caption         =   "Informe de Servicios"
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
         Left            =   600
         TabIndex        =   556
         Top             =   360
         Width           =   4815
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   53
         Left            =   1635
         ToolTipText     =   "Buscar socio"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Left            =   600
         TabIndex        =   555
         Top             =   1080
         Width           =   555
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
         Index           =   32
         Left            =   900
         TabIndex        =   554
         Top             =   1680
         Width           =   780
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
         Index           =   52
         Left            =   900
         TabIndex        =   553
         Top             =   2280
         Width           =   675
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
         Index           =   49
         Left            =   900
         TabIndex        =   552
         Top             =   2640
         Width           =   630
      End
      Begin VB.Label Label4 
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
         Index           =   29
         Left            =   600
         TabIndex        =   551
         Top             =   2040
         Width           =   765
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   40
         Left            =   1650
         ToolTipText     =   "Buscar tipo artículo"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   39
         Left            =   1650
         ToolTipText     =   "Buscar tipo artículo"
         Top             =   2280
         Width           =   240
      End
   End
   Begin VB.Frame FrameInventario 
      Height          =   6495
      Left            =   360
      TabIndex        =   68
      Top             =   360
      Width           =   7815
      Begin VB.TextBox txtCodigo 
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
         Left            =   4920
         TabIndex        =   52
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1920
         TabIndex        =   53
         Top             =   4680
         Width           =   735
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
         Index           =   21
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "Text5"
         Top             =   4680
         Width           =   4605
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
         Index           =   19
         Left            =   2720
         Locked          =   -1  'True
         TabIndex        =   84
         Text            =   "Text5"
         Top             =   3960
         Width           =   4575
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
         Index           =   18
         Left            =   2720
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "Text5"
         Top             =   3600
         Width           =   4575
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1920
         TabIndex        =   50
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
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
         TabIndex        =   49
         Top             =   3600
         Width           =   735
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
         Index           =   4
         Left            =   6330
         TabIndex        =   55
         Top             =   5850
         Width           =   1005
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
         Index           =   4
         Left            =   5250
         TabIndex        =   54
         Top             =   5850
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   45
         Top             =   1680
         Width           =   1485
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   15
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   46
         Top             =   2040
         Width           =   1485
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   47
         Top             =   2640
         Width           =   645
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   17
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   48
         Top             =   3000
         Width           =   645
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   2440
         TabIndex        =   51
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   44
         Top             =   1080
         Width           =   525
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
         Index           =   14
         Left            =   3440
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "Text5"
         Top             =   1680
         Width           =   3855
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
         Index           =   15
         Left            =   3440
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text5"
         Top             =   2040
         Width           =   3855
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
         Index           =   16
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text5"
         Top             =   2640
         Width           =   4695
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
         Index           =   17
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   3000
         Width           =   4695
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
         Index           =   13
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text5"
         Top             =   1080
         Width           =   4845
      End
      Begin VB.Frame FrameOpciones 
         Height          =   1575
         Left            =   3630
         TabIndex        =   324
         Top             =   4770
         Width           =   3435
         Begin VB.CheckBox chkValorado 
            Caption         =   "Valorado"
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
            TabIndex        =   328
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1485
         End
         Begin VB.CheckBox chkImprimeStock 
            Caption         =   "Imprimir Stock"
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
            TabIndex        =   327
            Top             =   870
            Value           =   1  'Checked
            Width           =   1845
         End
         Begin VB.CheckBox chkSinStock 
            Caption         =   "Imprimir Artículos sin Stock"
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
            TabIndex        =   326
            Top             =   510
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3045
         End
         Begin VB.CheckBox chkSaltaPag 
            Caption         =   "Salta pág. en Familia"
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
            TabIndex        =   325
            Top             =   150
            Width           =   2565
         End
      End
      Begin VB.Frame FrameValorar 
         Caption         =   "Valorar Con:"
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
         Height          =   1575
         Left            =   270
         TabIndex        =   90
         Top             =   4770
         Visible         =   0   'False
         Width           =   3105
         Begin VB.OptionButton optPrecioStd 
            Caption         =   "Precio Standard"
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
            TabIndex        =   94
            Top             =   1200
            Width           =   2745
         End
         Begin VB.OptionButton optPrecioUC 
            Caption         =   "Precio Ultima Compra"
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
            TabIndex        =   93
            Top             =   880
            Width           =   2805
         End
         Begin VB.OptionButton optPrecioMA 
            Caption         =   "Precio Medio Acumulado"
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
            TabIndex        =   92
            Top             =   560
            Width           =   2745
         End
         Begin VB.OptionButton optPrecioMP 
            Caption         =   "Precio Medio Ponderado"
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
            TabIndex        =   91
            Top             =   240
            Value           =   -1  'True
            Width           =   2685
         End
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   4670
         Top             =   4440
         Width           =   240
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
         Index           =   9
         Left            =   4350
         TabIndex        =   96
         Top             =   4440
         Width           =   1050
      End
      Begin VB.Label Label4 
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
         Index           =   8
         Left            =   3720
         TabIndex        =   95
         Top             =   4440
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
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
         Left            =   210
         TabIndex        =   89
         Top             =   4680
         Width           =   1185
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   17
         Left            =   1635
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   16
         Left            =   1635
         ToolTipText     =   "Buscar provedor"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   15
         Left            =   1635
         ToolTipText     =   "Buscar proveedor"
         Top             =   3600
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
         Left            =   210
         TabIndex        =   87
         Top             =   3360
         Width           =   1125
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
         Index           =   25
         Left            =   840
         TabIndex        =   86
         Top             =   3960
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
         Index           =   24
         Left            =   840
         TabIndex        =   85
         Top             =   3600
         Width           =   705
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
         Index           =   23
         Left            =   840
         TabIndex        =   82
         Top             =   1680
         Width           =   735
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
         Index           =   22
         Left            =   840
         TabIndex        =   81
         Top             =   2040
         Width           =   660
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
         Index           =   2
         Left            =   210
         TabIndex        =   79
         Top             =   1440
         Width           =   810
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   11
         Left            =   1635
         ToolTipText     =   "Buscar artículo"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   12
         Left            =   1635
         ToolTipText     =   "Buscar artículo"
         Top             =   2040
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
         Index           =   20
         Left            =   840
         TabIndex        =   78
         Top             =   2640
         Width           =   705
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
         Index           =   19
         Left            =   840
         TabIndex        =   77
         Top             =   3000
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Left            =   210
         TabIndex        =   76
         Top             =   2400
         Width           =   780
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   13
         Left            =   1635
         ToolTipText     =   "Buscar familia"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   14
         Left            =   1635
         ToolTipText     =   "Buscar familia"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inventario"
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
         Index           =   5
         Left            =   210
         TabIndex        =   75
         Top             =   4410
         Width           =   1830
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Left            =   210
         TabIndex        =   74
         Top             =   1050
         Width           =   915
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   10
         Left            =   1635
         ToolTipText     =   "Buscar almacen"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   2140
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label lbltituloInven 
         Caption         =   "Informe Toma de Inventario Articulos"
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
         Left            =   240
         TabIndex        =   80
         Top             =   360
         Width           =   7455
      End
   End
   Begin VB.Frame FrameEnvioMail 
      Height          =   1215
      Left            =   120
      TabIndex        =   451
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ProgressBar PBMail 
         Height          =   375
         Left            =   360
         TabIndex        =   452
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Preparando datos envio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   360
         TabIndex        =   453
         Top             =   840
         Width           =   5805
      End
   End
   Begin VB.Frame FrameBultos 
      Height          =   6975
      Left            =   240
      TabIndex        =   390
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtBultos 
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
         Left            =   3150
         TabIndex        =   401
         Text            =   "Text1"
         Top             =   6480
         Width           =   615
      End
      Begin VB.TextBox txtBultos 
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
         Left            =   1320
         TabIndex        =   398
         Text            =   "Text1"
         Top             =   3600
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
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
         TabIndex        =   397
         Text            =   "Text1"
         Top             =   3120
         Width           =   4215
      End
      Begin VB.TextBox txtBultos 
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
         Left            =   1320
         TabIndex        =   396
         Text            =   "Text1"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtBultos 
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
         Left            =   1320
         TabIndex        =   395
         Text            =   "Text1"
         Top             =   2640
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
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
         Left            =   1320
         TabIndex        =   394
         Text            =   "Text1"
         Top             =   2160
         Width           =   5175
      End
      Begin VB.ComboBox cmbBulto 
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   393
         Top             =   1620
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
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
         TabIndex        =   400
         Text            =   "Text1"
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton cmdEtiqBulto 
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
         Left            =   4440
         TabIndex        =   402
         Top             =   6480
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
         Index           =   95
         Left            =   5520
         TabIndex        =   403
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtBultos 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   0
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   399
         Text            =   "frmListado.frx":14B5
         Top             =   4200
         Width           =   5175
      End
      Begin VB.TextBox txtClie 
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
         Left            =   1320
         TabIndex        =   392
         Top             =   840
         Width           =   735
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
         Index           =   10
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   404
         Text            =   "Text5"
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "En blanco"
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
         Index           =   91
         Left            =   2040
         TabIndex        =   500
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pais"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   71
         Left            =   240
         TabIndex        =   430
         Top             =   3660
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Población"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   70
         Left            =   240
         TabIndex        =   429
         Top             =   2700
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   69
         Left            =   240
         TabIndex        =   428
         Top             =   3180
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   68
         Left            =   240
         TabIndex        =   427
         Top             =   2220
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Copias"
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
         Index           =   64
         Left            =   210
         TabIndex        =   408
         Top             =   6480
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Texto"
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
         Index           =   63
         Left            =   180
         TabIndex        =   407
         Top             =   4200
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
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
         Index           =   62
         Left            =   180
         TabIndex        =   406
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label Label4 
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
         Index           =   61
         Left            =   180
         TabIndex        =   405
         Top             =   840
         Width           =   765
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   75
         Left            =   1020
         ToolTipText     =   "Buscar cliente"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Etiquetas de bultos"
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
         Left            =   480
         TabIndex        =   391
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame FrameRepxDia 
      Height          =   5415
      Left            =   480
      TabIndex        =   155
      Top             =   480
      Width           =   6075
      Begin VB.Frame FrameProgress 
         Height          =   1200
         Left            =   360
         TabIndex        =   291
         Top             =   4080
         Visible         =   0   'False
         Width           =   4935
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   400
            Left            =   120
            TabIndex        =   293
            Top             =   640
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   714
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess 
            Caption         =   "Comprobaciones:"
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
            Index           =   0
            Left            =   120
            TabIndex        =   294
            Top             =   135
            Width           =   4455
         End
         Begin VB.Label lblProgess 
            Caption         =   "Iniciando el proceso ..."
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
            Index           =   1
            Left            =   120
            TabIndex        =   292
            Top             =   375
            Width           =   4575
         End
      End
      Begin VB.Frame FrameTipMov 
         BorderStyle     =   0  'None
         Caption         =   "Nº Factura"
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
         Height          =   990
         Left            =   360
         TabIndex        =   486
         Top             =   2560
         Width           =   4815
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
            Index           =   122
            Left            =   3555
            TabIndex        =   154
            Top             =   495
            Width           =   1040
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
            Index           =   121
            Left            =   2360
            TabIndex        =   153
            Top             =   495
            Width           =   1040
         End
         Begin VB.ComboBox cboTipMov 
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
            ItemData        =   "frmListado.frx":14BB
            Left            =   110
            List            =   "frmListado.frx":14BD
            Style           =   2  'Dropdown List
            TabIndex        =   152
            Top             =   495
            Width           =   2060
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Factura: "
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
            Index           =   8
            Left            =   120
            TabIndex        =   490
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Tip. Mov."
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
            Index           =   95
            Left            =   105
            TabIndex        =   489
            Top             =   240
            Width           =   2040
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
            Index           =   94
            Left            =   3555
            TabIndex        =   488
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Label2 
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
            Index           =   5
            Left            =   2355
            TabIndex        =   487
            Top             =   240
            Width           =   1005
         End
      End
      Begin VB.CommandButton cmdAceptarRepxDia 
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
         Left            =   2640
         TabIndex        =   156
         Top             =   3600
         Width           =   975
      End
      Begin VB.Frame FrameContab 
         Caption         =   " Facturas "
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
         Height          =   620
         Left            =   480
         TabIndex        =   290
         Top             =   1080
         Width           =   4455
         Begin VB.OptionButton OptProve 
            Caption         =   "Proveedores"
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
            Left            =   2280
            TabIndex        =   147
            Top             =   250
            Width           =   1695
         End
         Begin VB.OptionButton OptClientes 
            Caption         =   "Clientes"
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
            Left            =   600
            TabIndex        =   146
            Top             =   250
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         TabIndex        =   286
         Top             =   1680
         Width           =   5415
         Begin VB.TextBox txtCodigo 
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
            Left            =   1260
            TabIndex        =   149
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   3660
            TabIndex        =   151
            Top             =   480
            Width           =   1350
         End
         Begin VB.Label Label2 
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
            Left            =   360
            TabIndex        =   289
            Top             =   480
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
            Index           =   29
            Left            =   2745
            TabIndex        =   288
            Top             =   480
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Reparación:"
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
            Left            =   360
            TabIndex        =   287
            Top             =   195
            Width           =   1980
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   4
            Left            =   975
            Top             =   480
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   5
            Left            =   3360
            Top             =   480
            Width           =   240
         End
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
         Index           =   7
         Left            =   3840
         TabIndex        =   157
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Reparaciones por Día"
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
         Left            =   480
         TabIndex        =   158
         Top             =   465
         Width           =   5055
      End
   End
   Begin VB.Frame FrameRepSustNSerie 
      Height          =   3735
      Left            =   240
      TabIndex        =   329
      Top             =   0
      Width           =   5715
      Begin VB.TextBox txtCodigo 
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
         Index           =   81
         Left            =   1260
         MaxLength       =   15
         TabIndex        =   330
         Top             =   2160
         Width           =   2475
      End
      Begin VB.CommandButton cmdAceptarSustNSerie 
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
         Index           =   0
         Left            =   3180
         TabIndex        =   331
         Top             =   3000
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
         Index           =   13
         Left            =   4260
         TabIndex        =   332
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblNumSerie 
         Caption         =   "num serie"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   337
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label lblNumSerie 
         Caption         =   "num serie"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   336
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Introduce el nuevo Nº de Serie que va a sustituir al: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   360
         TabIndex        =   335
         Top             =   1005
         Width           =   4995
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Sustitución Nº de Serie"
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
         TabIndex        =   334
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Serie"
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
         Left            =   360
         TabIndex        =   333
         Top             =   2160
         Width           =   975
      End
   End
   Begin VB.Frame FrEliminarFacturas 
      Height          =   4215
      Left            =   120
      TabIndex        =   431
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdElimiaFacturas 
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
         Left            =   3840
         TabIndex        =   435
         Top             =   3600
         Width           =   975
      End
      Begin VB.ComboBox cmbEliFac 
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
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   434
         Top             =   3000
         Width           =   2655
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
         Index           =   97
         Left            =   5040
         TabIndex        =   432
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "lore ipsum lorem ipsum lorem ipsum"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   438
         Top             =   2160
         Width           =   5775
      End
      Begin VB.Label Label11 
         Caption         =   "lore"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   0
         Left            =   360
         TabIndex        =   437
         Top             =   360
         Width           =   5775
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   83
         Left            =   120
         TabIndex        =   436
         Top             =   3600
         Width           =   3525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Eliminar facturas hasta: "
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
         Index           =   82
         Left            =   360
         TabIndex        =   433
         Top             =   3000
         Width           =   2655
      End
   End
   Begin VB.Frame frameListado 
      Height          =   4695
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   6555
      Begin VB.Frame frameOrdenar 
         Caption         =   "Ordenar por"
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
         Height          =   735
         Left            =   720
         TabIndex        =   140
         Top             =   2640
         Width           =   3375
         Begin VB.OptionButton OptNombre 
            Caption         =   "Descripción"
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
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Optcodigo 
            Caption         =   "Código"
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
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
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
         Index           =   2
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2040
         Width           =   3255
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
         Index           =   1
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1560
         Width           =   3255
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
         Index           =   2
         Left            =   1605
         TabIndex        =   1
         Top             =   2040
         Width           =   830
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
         Left            =   1605
         TabIndex        =   0
         Top             =   1560
         Width           =   830
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
         Index           =   1
         Left            =   3600
         TabIndex        =   4
         Top             =   3960
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
         Index           =   1
         Left            =   4680
         TabIndex        =   5
         Top             =   3960
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1320
         ToolTipText     =   "Buscar marca"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1320
         ToolTipText     =   "Buscar marca"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Left            =   570
         TabIndex        =   16
         Top             =   1140
         Width           =   900
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
         Height          =   195
         Index           =   3
         Left            =   570
         TabIndex        =   14
         Top             =   1980
         Width           =   570
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
         Height          =   195
         Index           =   2
         Left            =   570
         TabIndex        =   13
         Top             =   1545
         Width           =   615
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Listado Marcas"
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
         Left            =   600
         TabIndex        =   15
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.Frame FrameInfAlmacen 
      Height          =   3495
      Left            =   1560
      TabIndex        =   30
      Top             =   1080
      Width           =   5835
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
         Left            =   3720
         TabIndex        =   9
         Top             =   2880
         Width           =   975
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
         Index           =   2
         Left            =   2640
         TabIndex        =   8
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1200
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   3480
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Index           =   6
         Left            =   270
         TabIndex        =   34
         Top             =   1800
         Width           =   615
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
         Left            =   2550
         TabIndex        =   33
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Informes Almacenes"
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
         Left            =   270
         TabIndex        =   32
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Traspaso"
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
         TabIndex        =   31
         Top             =   1320
         Width           =   1470
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   920
         ToolTipText     =   "Buscar almacén"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   3200
         ToolTipText     =   "Buscar almacén"
         Top             =   1800
         Width           =   240
      End
   End
   Begin VB.Frame FrameHcoMante 
      Height          =   3495
      Left            =   0
      TabIndex        =   455
      Top             =   -120
      Width           =   6495
      Begin VB.CommandButton cmdHcoMante 
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
         Left            =   3840
         TabIndex        =   460
         Top             =   2880
         Width           =   1095
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
         Index           =   112
         Left            =   1830
         MaxLength       =   4
         TabIndex        =   459
         Top             =   2280
         Width           =   735
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
         Index           =   112
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   465
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
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
         Index           =   111
         Left            =   1830
         MaxLength       =   4
         TabIndex        =   458
         Top             =   1560
         Width           =   735
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
         Index           =   111
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   463
         Text            =   "Text5"
         Top             =   1560
         Width           =   3735
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
         Index           =   110
         Left            =   1830
         TabIndex        =   457
         Top             =   960
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
         Index           =   99
         Left            =   5160
         TabIndex        =   462
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo baja"
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
         Index           =   81
         Left            =   240
         TabIndex        =   466
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   90
         Left            =   1545
         ToolTipText     =   "Buscar motivo baja"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
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
         Index           =   80
         Left            =   240
         TabIndex        =   464
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   89
         Left            =   1545
         ToolTipText     =   "Buscar trabajador"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   1545
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha baja"
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
         Index           =   79
         Left            =   240
         TabIndex        =   461
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label7 
         Caption         =   "Paso a mantenimientos anulados"
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
         Left            =   240
         TabIndex        =   456
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame FrameRepxClien 
      Height          =   5415
      Left            =   240
      TabIndex        =   161
      Top             =   240
      Width           =   6795
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   3720
         TabIndex        =   283
         Top             =   3240
         Width           =   2715
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
            Left            =   120
            MaxLength       =   4
            TabIndex        =   168
            Text            =   "1"
            Top             =   420
            Width           =   945
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "reparaciones"
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
            Index           =   43
            Left            =   1200
            TabIndex        =   285
            Top             =   450
            Width           =   1260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mostrar equipos con más de:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   41
            Left            =   120
            TabIndex        =   284
            Top             =   120
            Width           =   2490
         End
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
         Index           =   34
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   174
         Text            =   "Text5"
         Top             =   1680
         Width           =   3645
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
         Index           =   34
         Left            =   1740
         TabIndex        =   163
         Top             =   1680
         Width           =   1035
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
         Index           =   36
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   173
         Text            =   "Text5"
         Top             =   2640
         Width           =   3645
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
         Index           =   35
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   172
         Text            =   "Text5"
         Top             =   2280
         Width           =   3645
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
         Index           =   36
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   165
         Top             =   2640
         Width           =   1035
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
         Index           =   35
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   164
         Top             =   2280
         Width           =   1035
      End
      Begin VB.CommandButton cmdAceptarRepxClien 
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
         TabIndex        =   169
         Top             =   4680
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
         Index           =   8
         Left            =   5040
         TabIndex        =   170
         Top             =   4680
         Width           =   975
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
         Index           =   43
         Left            =   1740
         TabIndex        =   166
         Top             =   3360
         Width           =   1350
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
         Index           =   44
         Left            =   1740
         TabIndex        =   167
         Top             =   3720
         Width           =   1350
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
         Index           =   33
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   171
         Text            =   "Text5"
         Top             =   1320
         Width           =   3645
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
         Index           =   33
         Left            =   1740
         TabIndex        =   162
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1455
         Top             =   3750
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1455
         Top             =   3390
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   7
         Left            =   1455
         ToolTipText     =   "Buscar cliente"
         Top             =   1680
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
         Index           =   37
         Left            =   750
         TabIndex        =   184
         Top             =   1320
         Width           =   615
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   9
         Left            =   1455
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   8
         Left            =   1455
         ToolTipText     =   "buscar dir/dpto"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direc/Dpto"
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
         Index           =   22
         Left            =   420
         TabIndex        =   183
         Top             =   2040
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
         Index           =   36
         Left            =   750
         TabIndex        =   182
         Top             =   2640
         Width           =   570
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
         Index           =   35
         Left            =   750
         TabIndex        =   181
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Reparaciones  por Cliente"
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
         Left            =   420
         TabIndex        =   180
         Top             =   360
         Width           =   4815
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
         Index           =   34
         Left            =   750
         TabIndex        =   179
         Top             =   3360
         Width           =   615
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
         Index           =   30
         Left            =   750
         TabIndex        =   178
         Top             =   3720
         Width           =   570
      End
      Begin VB.Label Label4 
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
         Index           =   21
         Left            =   420
         TabIndex        =   177
         Top             =   3120
         Width           =   1515
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   6
         Left            =   1455
         ToolTipText     =   "Buscar cliente"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   18
         Left            =   420
         TabIndex        =   176
         Top             =   1080
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
         Index           =   16
         Left            =   750
         TabIndex        =   175
         Top             =   1680
         Width           =   720
      End
   End
   Begin VB.Frame FrameMantenimientos 
      Height          =   6975
      Left            =   360
      TabIndex        =   185
      Top             =   0
      Width           =   6735
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   270
         TabIndex        =   445
         Top             =   840
         Width           =   6255
         Begin VB.CheckBox chkMante 
            Caption         =   "Copia remitente"
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
            Left            =   360
            TabIndex        =   454
            Top             =   360
            Width           =   1935
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
            Index           =   109
            Left            =   1650
            TabIndex        =   189
            Top             =   720
            Width           =   1350
         End
         Begin VB.OptionButton optMante 
            Caption         =   "Comercial"
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
            Left            =   4200
            TabIndex        =   448
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optMante 
            Caption         =   "Administracion"
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
            Left            =   2400
            TabIndex        =   447
            Top             =   0
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.CheckBox chkMante 
            Caption         =   "Enviar e-mail"
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
            Left            =   360
            TabIndex        =   446
            Top             =   0
            Width           =   1845
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha carta"
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
            Index           =   77
            Left            =   60
            TabIndex        =   449
            Top             =   720
            Width           =   1245
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   109
            Left            =   1365
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   1
         Left            =   270
         TabIndex        =   442
         Top             =   5040
         Width           =   5895
         Begin VB.ComboBox cboTipoList 
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
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   443
            Tag             =   "Tipo Facturación|N|N|||scaalb|tipofact||N|"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Listado"
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
            Index           =   4
            Left            =   60
            TabIndex        =   444
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   45
         Left            =   1920
         TabIndex        =   187
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   45
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   214
         Text            =   "Text5"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   46
         Left            =   1920
         TabIndex        =   188
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   46
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   213
         Text            =   "Text5"
         Top             =   1560
         Width           =   3735
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
         Index           =   51
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   212
         Text            =   "Text5"
         Top             =   4080
         Width           =   3735
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
         Index           =   52
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   211
         Text            =   "Text5"
         Top             =   4440
         Width           =   3735
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
         Index           =   48
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   200
         Text            =   "Text5"
         Top             =   2520
         Width           =   3735
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
         Index           =   48
         Left            =   1920
         TabIndex        =   191
         Top             =   2520
         Width           =   735
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
         Index           =   50
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   199
         Text            =   "Text5"
         Top             =   3480
         Width           =   3735
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
         Index           =   49
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   198
         Text            =   "Text5"
         Top             =   3120
         Width           =   3735
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
         Index           =   50
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   193
         Top             =   3480
         Width           =   735
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
         Index           =   49
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   192
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton cmdAceptarMante 
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
         TabIndex        =   196
         Top             =   6360
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
         Index           =   9
         Left            =   5280
         TabIndex        =   197
         Top             =   6360
         Width           =   975
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
         Index           =   51
         Left            =   1920
         TabIndex        =   194
         Top             =   4080
         Width           =   735
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
         Index           =   52
         Left            =   1920
         TabIndex        =   195
         Top             =   4440
         Width           =   735
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
         Index           =   47
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   186
         Text            =   "Text5"
         Top             =   2160
         Width           =   3735
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
         Index           =   47
         Left            =   1920
         TabIndex        =   190
         Top             =   2160
         Width           =   735
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   0
         Left            =   600
         TabIndex        =   295
         Top             =   4920
         Width           =   5415
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   53
            Left            =   1560
            TabIndex        =   298
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   54
            Left            =   3840
            TabIndex        =   297
            Top             =   360
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   9
            Left            =   3555
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   8
            Left            =   1275
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   44
            Left            =   720
            TabIndex        =   300
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   45
            Left            =   3000
            TabIndex        =   299
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Revisiones Efectuadas"
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
            Index           =   5
            Left            =   240
            TabIndex        =   296
            Top             =   120
            Width           =   4335
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   1080
         TabIndex        =   217
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
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
         Left            =   600
         TabIndex        =   216
         Top             =   960
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   41
         Left            =   1635
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   43
         Left            =   1080
         TabIndex        =   215
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   42
         Left            =   1635
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   44
         Left            =   1635
         ToolTipText     =   "Buscar cliente"
         Top             =   2520
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
         Index           =   42
         Left            =   780
         TabIndex        =   210
         Top             =   2160
         Width           =   705
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   46
         Left            =   1635
         ToolTipText     =   "Buscar agente"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   45
         Left            =   1635
         ToolTipText     =   "Buscar agente"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   26
         Left            =   300
         TabIndex        =   209
         Top             =   2880
         Width           =   1005
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
         Index           =   41
         Left            =   780
         TabIndex        =   208
         Top             =   3480
         Width           =   660
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
         Index           =   40
         Left            =   780
         TabIndex        =   207
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Informe de Mantenimientos"
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
         Left            =   360
         TabIndex        =   206
         Top             =   360
         Width           =   5775
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
         Left            =   780
         TabIndex        =   205
         Top             =   4080
         Width           =   705
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
         Index           =   38
         Left            =   780
         TabIndex        =   204
         Top             =   4440
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Contrato"
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
         Index           =   25
         Left            =   300
         TabIndex        =   203
         Top             =   3840
         Width           =   1710
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   47
         Left            =   1635
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   48
         Left            =   1635
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   43
         Left            =   1635
         ToolTipText     =   "Buscar cliente"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   24
         Left            =   300
         TabIndex        =   202
         Top             =   1920
         Width           =   1005
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
         Left            =   780
         TabIndex        =   201
         Top             =   2520
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ====  MODIFICACIONES  ==========================================
' ====  [16/09/2009] LAURA : Añadir el frame "FrameInvArtComp" para sacar listado articulos con componentes
' ====  [17/09/2009] LAURA : Cambiar conexion CONTAB si FAZ
' ================================================================


Public OpcionListado As Integer

    '==== Listados de ALMACEN ====
    '=============================
    ' 1 .- Listados Marcas.
    ' 2 .- Listado de Almacenes Propios
    ' 3 .- Listado de Tipos de Unidad
    ' 4 .- Listado de Tipos de Artículos
    ' 5 .- Listado de Familias de artículos
    
    ' 6 .- Listado de Artículos
    ' 7 .- Informe de Traspaso de Almacenes
    ' 8 .- Informe de Movimientos de Almacen
    ' 9 .- Listado Busquedas de movimientos de Artículos
    '10 .-
    
    '11 .- Listado de Articulos con componentes ' ====  [16/09/2009] LAURA
    '12 .- Listado Toma de Inventario Articulos
    '13 .- Listado de Diferencias de Inventario Articulos
    '14 .- Actualizar Diferencias de Inventario (No IMPRIME INFORME)
    '15 .- Listado de Articulos Inactivos.
    
    '16 .- Listado Valoracion de Stocks Inventariados
    '17 .- Listado Valoración Stocks
    '18 .- Informe Stocks Maximos y Minimos
    '19 .- Informe de Stocks a una fecharEtiqBulto.rpt
    
    '110 .- Listado de Ubicaciones
    
    
    
    
    '==== Listados de FACTURACION ====
    '=================================
    '20 .- Listado de Actividades de Clientes
    '21 .- Listado de Zonas de Clientes
    '22 .- Listado de Rutas de Asistencia
    '23 .- Listado de Formas de Envío
    '24 .- Listado de Tarifas Ventas
    '25 .-
    
    '26 .-
    '27 .- Listado de Situaciones Especiales
    '28 .- Informe de Tarifas de Articulos
    '29 .- Informe de Promociones de Tarifas
    '30 .- Informe de Precios Especiales
    
    '31 .- Informe de Ofertas
    '32 .- Informe de Recordatorio de Ofertas
    '33 .- Informe de Valoración de Ofertas
    '34 .- Informe de Ofertas Efectuadas
    '35 .- Informe Historico de Ofertas
    
    '36 .- Traspaso de Ofertas al Historico (NO IMPRIME INFORME)
    '37 .- Solicitar datos para pasar de Oferta a Pedido (NO IMPRIME INFORME)
    '38 .- Informe de Pedidos
    '39 .- Orden de Instalacion
    '40 .- Cartas Confirmacion de Pedidos
    
    '41 .- Informe de Pedidos por Articulo
    '42 .- Informe de Disponibilidad de Stocks
    '43 .- Generar Albaran desde Pedido (NO IMPRIME LISTADO)
    '44 .- Informe de Pedidos por Cliente
    '45 .- Informe de Albaran
    
    '46 .- Informe de Clientes Inactivos
    '47 .- Informe de Clientes
    '48 .- Informe de Altas de Nuevos Cliente
    '49 .- Informe de Albaranes por Articulo
    '50 .- Prevision de Facturacion de ALbaranes
    
    '51 .- Informe Incumplimiento Plazos de Entrega
    '52 .- Facturacion de Albaranes (NO IMPRIME LISTADO?)
    '53 .- Informe de Factura
    '54 .- Listado de Descuentos Familia/Marca
    
    '59 .- Informe de Factura ProForma
    '222 .- Informe de Factura Mostrador
    '223 .- Pedir datos para contabilizar facturas CLIENTES
    '224 .- Pedir datos para contabilizar facturas PROVEEDOR
    '225 .- Pedir datos para generar Facturas Rectificativas
    '226 .- Pedir datos para reimprimir Facturas
    '227 .- Informe estadistica Ventas por cliente
    '228 .- Informe estadistica Ventas por Trabajador
    '229 .- Informe estadistica Ventas por meses
    '230 .- Informe estadistica Ventas por familia
    '231 .- Informe detalle facturacion clientes
    
    '238 .- Confirmacion entrega Pedido
    '239 .- Hco de Pedidos de venta (Historico)
    '240 .- Informe Cierre de Caja del TPV
    
    '245 .- Informe control margenes tarifas
    '246 .- Informe Margen ventas por articulo
    '247 .- Corrección de errores y acutalizacion de tarifas
    
    
    'Abril 2008
    '248 .- Contabilizar facturas de tickets AGRUPADAS
    
    
    
    '==== Listados de COMPRAS ====
    '=============================
    '55 .- Informe de Pedido Proveedor
    '56 .- Inf. Historico Pedido Proveedor
    '57 .- Pasa Pedido a Albaran compras (NO IMPRIME LISTADO)
    '58 .- Listado de Proveedores
    
    
    '305 .- Listado Etiquetas de Proveedores
    '306 .- Listado Cartas a Proveedores
    '307 .- Listado Material pendiente de recibir
    '308 .- Listado Albaranes pendientes de facturar
    '309 .- Listado  Precios de Compra
    '310 .- Listado Compras por Proveedor
    '311 .- Listado Compras por Familia
    '312 .- Listado albaranes por proveedor
    
    
    '==== Listados de REPARACIONES ====
    '==================================
    '60 .- Informe de Numeros de Serie
    '61 .- Listado Motivos Pend. Rep.
    '62 .- Listado Resguardo Reparacion
    '63 .- Listado Reparaciones por Día
    '64 .- Listado Reparaciones por Cliente
    '65 .- Listado motivos baja equipos
    
    '406 .- Listado Frecuencia de reparaciones
    '407 .- Sustitución Nº de Serie
    '408 .- Informe Aviso de Averia
    '409 .- Listado Avisos de averia pendientes
    
    
    '==== Listados de ADMINISTRACION ====
    '====================================
    
    '501 .- Listado de Nominas y Gastos
    
    
    '==== Listados de MANTENIMIENTOS ====
    '==================================
    '70 .- Listado Mantenimiento
    '71 .- Listado Revisiones de Mantenimientos
    '72 .- Informe Fichas de Mantenimientos
    '73 .- Listado Altas de Mantenimientos
    '74 .- Prefacturación Mantenimientos
    '75 .- Facturación de Mantenimientos
    '76 .- IGUAL QUE EL 70 pero en ANULADOS
        
        
        
    '77 .- Informe teórico de mantenimientos
    '78 .- Cartas de renovacion
    '79 .- Etiquetas manteimiento
    
    
    '==== Listados OTROS ====
    '==================================
    
    '80 .- Pasar Albaranes Ventas al historico (NO IMPRIME)
    '81 .- Pasar Pedidos Ventas al historico (NO IMPRIME)
       
           
    '82 .- Marcar facturar albaranes
    '83 .- Borre avisos cerrados
       
    
       
       
    '90 .- Etiquetas de Clientes
    '91 .- Cartas a Clientes
    
    '92 .- Informe de Gastos Técnicos
    '93 .- Ticket del TPV
      
    '94 .- Etiquetas estanteria
    
    '95 .- Etiquetas de bultos
    '96 .- Frecuencias
    '97 .- Eliminar facturas
    '99 .- Traspaso a mantenimientos anulados
    
    
    '120.- Informe de Servicios
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Aritaxi  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmMtoAlPropios As frmAlmAlPropios
Attribute frmMtoAlPropios.VB_VarHelpID = -1
Private WithEvents frmMtoMarcas As frmAlmMarcas
Attribute frmMtoMarcas.VB_VarHelpID = -1
Private WithEvents frmMtoTUnidad As frmAlmTipoUnidad
Attribute frmMtoTUnidad.VB_VarHelpID = -1
Private WithEvents frmMtoTArticulo As frmAlmTipoArticulo
Attribute frmMtoTArticulo.VB_VarHelpID = -1
Private WithEvents frmMtoActiv As frmFacActividades
Attribute frmMtoActiv.VB_VarHelpID = -1
Private WithEvents frmMtoFEnvio As frmFacFormasEnvio
Attribute frmMtoFEnvio.VB_VarHelpID = -1
Private WithEvents frmMtoTarifas As frmFacTarifas
Attribute frmMtoTarifas.VB_VarHelpID = -1
Private WithEvents frmMtoSituac As frmFacSituaciones
Attribute frmMtoSituac.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmAlmFamiliaArticulo
Attribute frmMtoFamilia.VB_VarHelpID = -1
Private WithEvents frmMtoProveedor As frmComProveedores
Attribute frmMtoProveedor.VB_VarHelpID = -1
Private WithEvents frmMtoArticulos As frmAlmArticulos
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Attribute frmMtoCartasOfe.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoClientes As frmFacClientes
Attribute frmMtoClientes.VB_VarHelpID = -1
Private WithEvents frmMtoMotivos As frmRepMotivosPend
Attribute frmMtoMotivos.VB_VarHelpID = -1
Private WithEvents frmMtoAgentes As frmFacAgentesCom
Attribute frmMtoAgentes.VB_VarHelpID = -1
Private WithEvents frmMtoIncid As frmIncidencias
Attribute frmMtoIncid.VB_VarHelpID = -1
Private WithEvents frmMtoSocios As frmGesSocios
Attribute frmMtoSocios.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
'-----------------------------------





'Para ademas de insertarlas en la conta, que las contabilice (pase a hsaldos)
'es decir, en el momento que inserta en cabfact tb insertaremos en hlinapu, hacabapu, hsaldos y hsaldosanal (si procede)










Dim indCodigo As Integer 'indice para txtCodigo

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim indFrame As Single


Dim kCampo As Integer

Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cboSituaAviso_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboTipMov_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboTipoList_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDtoFM_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkImprimeCodigoBarras_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkSitaucionArticulo_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cmbBulto_Click()
    PonerCamposDireccionBultos cmbBulto.ListIndex
End Sub

Private Sub cmbBulto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbProduccion_Click()
    If PrimeraVez Then Exit Sub
    PonerLabelsArticulosFrameVisible cmbProduccion.ListIndex = 1
End Sub

Private Sub CmdAceptarServicios_Click()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim cadAux As String
Dim bol As Boolean
Dim bytPrecio As Byte

   InicializarVbles

   If Me.Option1(0).Value Then
       cadAux = "{shilla.numeruve}"
   Else
       cadAux = "{shilla.codclien}"
   End If
   cadParam = "|pGrupo=" & cadAux & "|"
   numParam = 1
    
   'Añadir el parametro de Empresa
   cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
   numParam = numParam + 1
    
    'Desde/Hasta numero de socio
    '---------------------------------------------
    If txtCodigo(59).Text <> "" Or txtCodigo(60).Text <> "" Then
        Codigo = "{shilla.numeruve}"
        If Not PonerDesdeHasta(Codigo, "N", 59, 60, "pDHSocio=""Socio:") Then Exit Sub
    End If
    
    'Desde/Hasta numero de cliente
    '---------------------------------------------
    If txtCodigo(57).Text <> "" Or txtCodigo(58).Text <> "" Then
        Codigo = "{shilla.codclien}"
        If Not PonerDesdeHasta(Codigo, "N", 57, 58, "pDHCliente=""Cliente:") Then Exit Sub
    End If
    
    'Desde/Hasta fecha de servicio
    '---------------------------------------------
    If txtCodigo(55).Text <> "" Or txtCodigo(56).Text <> "" Then
        Codigo = "{shilla.fecha}"
        If Not PonerDesdeHasta(Codigo, "F", 55, 56, "pDHFecha=""Fecha") Then Exit Sub
    End If
   
    If Me.Option1(0).Value Then
        cadParam = cadParam & "pTitulo=""Cliente""|"
    Else
        cadParam = cadParam & "pTitulo=""Socio""|"
    End If
    numParam = numParam + 1
       
    cadNomRPT = "rInfServicios.rpt"
    cadTitulo = "Informe de Servicios"
       
   LlamarImprimir 0 'Movimientos almacen si tiene rpt personalizables
   
   Screen.MousePointer = vbDefault

End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim cadAux As String
Dim bol As Boolean
Dim bytPrecio As Byte

   InicializarVbles
   
   Select Case Index
   '========= Frame Listados =================================================
    
    ' ====  [16/09/2009] LAURA : Listado Articulos con componentes
    Case 0 ' Listado Articulos con componentes
        cadNomRPT = "rAlmArtCompon.rpt"
        conSubRPT = True
        
        Screen.MousePointer = vbHourglass
        DoEvents
        
        'Añadir el parametro de Empresa
        cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = 1
        
        If Trim(txtCodigo(125).Text) <> "" Or Trim(txtCodigo(126).Text) <> "" Then
            cadFormula = CadenaDesdeHasta(txtCodigo(125).Text, txtCodigo(126).Text, Codigo, "T")
            
            If cadFormula <> "" Then
                If cadFormula = "Error" Then Exit Sub
                cadAux = ""
                If txtCodigo(125).Text <> "" Then cadAux = "Desde: " & txtCodigo(125).Text & " " & txtNombre(125).Text
                If txtCodigo(126).Text <> "" Then
                    If cadAux <> "" Then cadAux = cadAux & "  -  "
                    cadAux = cadAux & " Hasta: " & txtCodigo(126).Text & " " & txtNombre(126).Text
                End If
                cadParam = cadParam & "pDesde=""" & cadAux & """|"
                numParam = numParam + 1
            End If
        End If
        AnyadirAFormula cadFormula, " {sartic.conjunto}=1"
        
        
        'Parametro Precio de Valoracion, elegir un Precio para realizar la valoracion: canstock * precio
        If Me.optPrecioMP3.Value Then bytPrecio = 1
        If Me.optPrecioMA3.Value Then bytPrecio = 2
        If Me.optPrecioUC3.Value Then bytPrecio = 3
        If Me.optPrecioStd3.Value Then bytPrecio = 4
        cadParam = cadParam & "pPrecio=" & bytPrecio & "|"
        numParam = numParam + 1
    ' ====
   
   
    Case 1 'Frame Listados
        If Me.Optcodigo.Value = True Then
            cadAux = Orden1
        Else
            cadAux = Orden2
        End If
        cadParam = "|pOrden=" & cadAux & "|"
        numParam = 1
        
        'Añadir el parametro de Empresa
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
        
        If Trim(txtCodigo(1).Text) <> "" Or Trim(txtCodigo(2).Text) <> "" Then
            'Cadena para seleccion Desde y Hasta
            If OpcionListado = 4 Or OpcionListado = 110 Then
                '4: Listado Tipos de Articulos, 110: List. Ubicaciones
                cadFormula = CadenaDesdeHasta(txtCodigo(1).Text, txtCodigo(2).Text, Codigo, "T")
            Else
                cadFormula = CadenaDesdeHasta(txtCodigo(1).Text, txtCodigo(2).Text, Codigo, "N")
            End If
            
            If cadFormula <> "" Then
                If cadFormula = "Error" Then Exit Sub
                cadAux = ""
                If txtCodigo(1).Text <> "" Then cadAux = "Desde: " & txtCodigo(1).Text & " " & txtNombre(1).Text
                If txtCodigo(2).Text <> "" Then
                    If cadAux <> "" Then cadAux = cadAux & "  -  "
                    cadAux = cadAux & " Hasta: " & txtCodigo(2).Text & " " & txtNombre(2).Text
                End If
                cadParam = cadParam & "pDesde=""" & cadAux & """|"
                numParam = numParam + 1
            End If
        End If
        
    '========= Frame Informes Almacen ========================================
    Case 2 'Frame Informes Almacen
        If OpcionListado = 7 Then '7: Traspaso Almacen
            indRPT = 1
            cadAux = "scatra"
            cadTitulo = "Informe Traspaso Almacenes"
        ElseIf OpcionListado = 8 Then '8: Movimientos Almacen
            indRPT = 3
            cadAux = "scamov"
            cadTitulo = "Informe Movimientos Almacen"
        End If
        
        cadParam = "|"
        If Not PonerParamEmpresa(cadParam, numParam) Then Exit Sub
        If PonerParamRPT(indRPT, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt) Then
            'Cadena para seleccion Desde y Hasta DOCUMENTO
            '----------------------------------------------
            If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
                If Not PonerDesdeHasta(Codigo, "N", 3, 4, "") Then Exit Sub
            End If
        
            If Not HayRegParaInforme(cadAux, cadSelect) Then Exit Sub
        End If
                       
                   
                   
    '========= Frame Listado Movimiento de Artículos ========================
    Case 3 'Frame Listado Movimiento de Artículos
        'Nombre fichero .rpt a Imprimir
        cadNomRPT = "rAlmMovim.rpt"
        
        If Not PonerFormulaYParametrosInf9() Then Exit Sub
        'comprobar que hay datos para mostrar en el Informe
        cadAux = "smoval INNER JOIN sartic ON smoval.codartic=sartic.codartic "
        If Not HayRegParaInforme(cadAux, cadSelect) Then Exit Sub
        conSubRPT = True
    
    '========= Frame de Inventario ==========================================
    Case 4 'Frame de Inventario
        If Not ValidarCamposInventario Then Exit Sub
        If OpcionListado = 19 Then
            cadNomRPT = "rAlmStocksFecha.rpt"
        Else
            'Nombre fichero .rpt a Imprimir
            If vParamAplic.InventarioxProv Then 'Se realiza inventario por Proveedor
                                                'Ordenar por: codprove, codfamia, codartic
                Select Case OpcionListado
                    Case 12: cadNomRPT = "rAlmInvenxProv.rpt"  'Listado Toma de Inventario
                    Case 13: cadNomRPT = "rAlmInvenxProvDif.rpt"  'Listado Diferencias
                    Case 15: cadNomRPT = "rAlmArtInactivos.rpt"
                    Case 16: cadNomRPT = "rAlmInvenxProvValoracion.rpt"  'Listado Valoracion Stocks Inventariados
                    Case 17: cadNomRPT = "rAlmValoracionxProv.rpt"  'Listado Valoracion Stocks (Por Proveedor)
                End Select
            Else 'Ordenar por Cod. Familia y no por Proveedor. Ordenar por: codfamia, codartic.
                Select Case OpcionListado
                    Case 12: cadNomRPT = "rAlmInventario.rpt"  'Listado Toma de Inventario
                    Case 13: cadNomRPT = "rAlmInventarioDif.rpt"  'Listado Diferencias
                    Case 15: cadNomRPT = "rAlmArtInactivos.rpt"
                    Case 16: cadNomRPT = "rAlmInvenValoracion.rpt"  'Listado Valoracion Stocks Inventariados
                    Case 17: cadNomRPT = "rAlmValoracion.rpt"  'Listado Valoracion Stocks)
                End Select
            End If
        End If
        Screen.MousePointer = vbHourglass
        DoEvents
        bol = PonerFormulaYParametrosInf12()
        Screen.MousePointer = vbDefault
        If Not bol Then Exit Sub
        
   End Select
    
       
   If OpcionListado = 14 Then 'Actualizar Inventario (NO IMPRIME INFORME)
        If Trim(txtCodigo(21).Text) <> "" Then
            'Quitar las llaves:{tabla.codigo} de la cadena consulta
            'para el FormulaSelection del informe Crystal Report y
            'Tendremos la clausula WHERE para insertar en la tabla:sinven
            cadAux = QuitarCaracterACadena(cadFormula, "{")
            cadFormula = QuitarCaracterACadena(cadAux, "}")
            If ActualizarInventario Then
                MsgBox "La Actualización de Inventario se ha realizado correctamente.", vbInformation
            End If
        Else
            MsgBox "El campo Trabajador debe tener valor", vbInformation
            PonerFoco txtCodigo(21)
            Exit Sub
        End If
        
   Else 'Listados
'        If OpcionListado = 19 Then cadFormula = ""
        If OpcionListado = 19 Then cadFormula = "({tmpstockfec.codusu} =" & vUsu.Codigo & ")"
        
        LlamarImprimir Index = 2 'Movimientos almacen si tiene rpt personalizables

        'Realizar otras acciones segun el informe que llame
        Select Case OpcionListado
            Case 12 'Toma de Inventario
                'If frmVisReport.EstaImpreso = True Then
                    PrepararTomaInventario
                'End If
            Case 7, 8 'Movimientos
                ActualizarImprimir
            Case 19
                DescargarDatosTMPStockFecha
        End Select
        
   End If
   Screen.MousePointer = vbDefault
End Sub


Private Sub PrepararTomaInventario()
Dim cadAux As String

    On Error GoTo ETomaInv
    
    If MsgBox("¿Impresión correcta para Actualizar Inventario?", vbQuestion + vbYesNo) = vbYes Then
        'Quitar las llaves:{tabla.codigo} de la cadena consulta
        'para el FormulaSelection del informe Crystal Report y
        'Tendremos la clausula WHERE para insertar en la tabla:sinven
'                cadAux = QuitarCaracterACadena(cadFormula, "{")
'                cadFormula = QuitarCaracterACadena(cadAux, "}")
       If CrearTmpInventario(cadSelect) Then
            If InsertarInventario Then
                MsgBox "Puede pasar a realizar la Entrada de Inventario Real", vbInformation
            End If
       End If
       cadAux = "DROP TABLE IF EXISTS tmpInven "
       conn.Execute cadAux
    End If
    
ETomaInv:
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub cmdAceptarArtic_Click()
'Listado de Articulos
Dim campo As String
Dim devuelve As String
Dim Opcion As Byte, numOp As Byte
Dim cadFrom As String





    InicializarVbles
    
    'Si es informe=18 de Stocks Maximos y Minimos comprobar
    'que se ha seleccionado un almacen
    Select Case OpcionListado
    Case 18
        'If OpcionListado = 18 Then
        If txtCodigo(72).Text = "" Then
            MsgBox "Se debe seleccionar un Almacen para el informe.", vbInformation
            Exit Sub
        End If
        cadNomRPT = "rAlmStocksMaxMin.rpt"
        cadFrom = " salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
    Case 247
        '
        Opcion = 0
        If vParamAplic.Produccion And Me.cmbProduccion.ListIndex = 1 Then Opcion = 1
        If Opcion = 0 Then
            If txtCodigo(107).Text = "" Or txtNombre(107) = "" Then
                MsgBox "Debe seleccionar una tarifa para el informe.", vbInformation
                Exit Sub
            End If
        Else
            'Corrector de precios de articulos con componentes
            txtCodigo(107).Text = ""
            txtNombre(107) = ""
        End If
    Case Else
        'El 6
        cadNomRPT = "rAlmListArticulos.rpt"  'Nombre fichero .rpt a Imprimir
        cadFrom = " sartic"
        cadParam = ""
        For Opcion = 0 To 2
            If Me.chkSitaucionArticulo(Opcion).Value = 1 Then cadParam = cadParam & "O"
        Next
        If cadParam = "" Then
            MsgBox "Seleccione la situacion del articulo", vbExclamation
            Exit Sub
        End If
        Opcion = 0
    End Select
    
    '===================================================
    '============ PARAMETROS ===========================
    cadParam = "|"
    'Empresa
    cadParam = cadParam & "pEmpresa=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion  ALMACEN
    '--------------------------------------------
    If OpcionListado = 18 And txtCodigo(72).Text <> "" Then
        campo = "{salmac.codalmac}"
        cadFormula = campo & "= " & txtCodigo(72).Text
        
        
    Else
        'Es tarifa para la correccion
        If OpcionListado = 247 And txtCodigo(107).Text <> "" Then
            campo = "{slista.codlista}"
            cadFormula = campo & "= " & txtCodigo(107).Text
        End If
    End If
    
    
    'Cadena para seleccion D/H FAMILIA
    '--------------------------------------------
     If txtCodigo(62).Text <> "" Or txtCodigo(63).Text <> "" Then
        campo = "{sartic.codfamia}"
        'Parametro Desde/Hasta Familila
        devuelve = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(campo, "N", 62, 63, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H MARCA
    '--------------------------------------------
    If txtCodigo(64).Text <> "" Or txtCodigo(65).Text <> "" Then
        campo = "{sartic.codmarca}"
        'Parametro Desde/Hasta Marca
        devuelve = "pDHMarca=""Marca: "
        If Not PonerDesdeHasta(campo, "N", 64, 65, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    If txtCodigo(66).Text <> "" Or txtCodigo(67).Text <> "" Then
        campo = "{sartic.codprove}"
        'Parametro Desde/Hasta Proveedor
        devuelve = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(campo, "N", 66, 67, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H TIPO ARTICULO
    '--------------------------------------------
    If txtCodigo(68).Text <> "" Or txtCodigo(69).Text <> "" Then
        campo = "{sartic.codtipar}"
        'Parametro Desde/Hasta Tipo Articulo
        devuelve = "pDHTipoArt=""Tipo Articulo: "
        If Not PonerDesdeHasta(campo, "T", 68, 69, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H ARTICULO
    '--------------------------------------------
    If txtCodigo(70).Text <> "" Or txtCodigo(71).Text <> "" Then
        campo = "{sartic.codartic}"
        'Parametro Desde/Hasta Articulo
        devuelve = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(campo, "T", 70, 71, devuelve) Then Exit Sub
    End If
    
    
    
    'Obtener el parametro con el Orden del Informe
    '---------------------------------------------
    Select Case OpcionListado
    Case 6
    
        'Veos que articulos quiere mostrar en funcion de la situacion
        '---------------------------------
        ' si los de situacion NORMAL
        devuelve = ""
        If Me.chkSitaucionArticulo(0).Value = 1 Then
            'SI los BLOQUEADO
            If Me.chkSitaucionArticulo(1).Value = 1 Then
                If Me.chkSitaucionArticulo(2).Value = 1 Then
                    'LOS QUIERE TODOS. NO PONGO NADA
                Else
                    'NO QUEIRE LOS CADUCADOS
                    devuelve = " < 2"
                End If
            Else
                'Los bloqueados NO
                '-----------------
                
                '       si los caducados
                If Me.chkSitaucionArticulo(2).Value = 1 Then
                    devuelve = " <> 1"
                    
                Else
                '       los caducados tampoco, es decir solo los normales
                    devuelve = " = 0"
                End If
            End If
        Else
            'NO QUIERE LOS NORMALES
            If Me.chkSitaucionArticulo(1).Value = 1 Then
                If Me.chkSitaucionArticulo(2).Value = 1 Then
                        devuelve = " > 0"
                Else
                        devuelve = " = 1" 'solo bloqueados
                End If
            Else
                'Es decir, NO QUIERE ni normal ni bloqueados, SOLO caducados
                devuelve = " = 2"
            End If

        End If
        If devuelve <> "" Then
            campo = "{sartic.codstatu} " & devuelve
            AnyadirAFormula cadFormula, campo
            devuelve = ""
        End If
        
    ''''If OpcionListado = 6 Then '6: Listado de Articulos
        numOp = PonerGrupo(1, ListView2.ListItems(1).Text)
        If numOp <> 0 Then Opcion = numOp
        numOp = PonerGrupo(2, ListView2.ListItems(2).Text)
        If numOp <> 0 Then Opcion = numOp
        numOp = PonerGrupo(3, ListView2.ListItems(3).Text)
        If numOp <> 0 Then Opcion = numOp
        numOp = PonerGrupo(4, ListView2.ListItems(4).Text)
        If numOp <> 0 Then Opcion = numOp
        Opcion = Opcion - 1
    
        Select Case Opcion
            Case 1 'El group2 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(3).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(4).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
            Case 2 'El Group3 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(4).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
            Case 3, 0 'El Group4 es el Proveedor
                      '0 'El Group1 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(3).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                If Opcion = 0 Then
                    campo = "pTitulo3=""" & ListView2.ListItems(4).Text & """"
                    cadParam = cadParam & campo & "|"
                    numParam = numParam + 1
                End If
        End Select
       
        'Parametro Orden del Informe
        campo = "pOrden=" & Opcion
        cadParam = cadParam & campo & "|"
        numParam = numParam + 1
        
    Case 18
    ''ElseIf OpcionListado = 18 Then
        'filtrar ademas por solo articulos con control de stock
        campo = "{sartic.ctrstock}=1"
        AnyadirAFormula cadFormula, campo
    
    
        'David.  Enero 2009
        'Los articulos cuya situacion NO este cadaducado, es decir, NORMAL y BLOQUEADO
        campo = "{sartic.codstatu}<2"
        AnyadirAFormula cadFormula, campo
    
        'Filtrar ademas por stock<stockMin o stock>stockMax
        campo = "{salmac.canstock}"
        If Me.optStockMax Then
            cadFormula = cadFormula & " AND (" & campo & "> {salmac.stockmax})"
        Else
            'David G 30/01/2007
            If optPuntoPedido.Value Then
                cadFormula = cadFormula & " AND (" & campo & "< {salmac.puntoped})"
            Else
                cadFormula = cadFormula & " AND (" & campo & "< {salmac.stockmin})"
            End If
        End If
    
        'Añadir el Parametro de Stocks Maximos o Minimos
        If Me.optStockMax.Value = True Then
            campo = "0"
        Else
            If optPuntoPedido.Value Then
                campo = "2"
            Else
                campo = "1"
            End If
        End If
        cadParam = cadParam & "pStockMax=" & campo & "|"
        numParam = numParam + 1
    Case 247

        'Correccion de importes
        '-------------------------------------------------------
        
        If BloqueoManual("CORRIGEPRECIOS", "1") Then
            
            
        
            'Mostrare el list
            cadSelect = QuitarCaracterACadena(cadFormula, "{")
            cadSelect = QuitarCaracterACadena(cadSelect, "}")
            frmMensajes.cadWHERE = cadSelect
            
            frmMensajes.OpcionMensaje = 16
                ' CORRECCION DE PRECIOS DE ARTICULOS QUE TIENEN COMPONENTES
            If vParamAplic.Produccion And Me.cmbProduccion.ListIndex = 1 Then frmMensajes.OpcionMensaje = 20
                
            frmMensajes.vCampos = txtCodigo(107).Text
            frmMensajes.cadWHERE2 = Trim(Me.cmbDecimales.Text)
            'Por no utilizar otra variable
            NumRegElim = 0
            If Me.chkMinimoCorreg.Value = 1 Then NumRegElim = 1
            frmMensajes.Show vbModal
       
        End If
        DesBloqueoManual ("CORRIGEPRECIOS")
        Exit Sub
    End Select
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
        
        
    
        
    LlamarImprimir False
End Sub


Private Sub cmdAceptarDtosFM_Click()
'54: Listado de Descuentos Familia/Marca
'309: Listado precio compras
Dim campo As String, cad As String
Dim Tabla As String

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
        
    If OpcionListado = 54 Then
        Tabla = "sdtofm"
        conSubRPT = True
    ElseIf OpcionListado = 309 Then
        Tabla = "slispr"
        cadTitulo = "Listado Precios de compra"
        cadNomRPT = "rComPrecios.rpt"
        conSubRPT = False
    End If
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H FAMILIA
    '----------------------------------
    If txtCodigo(75).Text <> "" Or txtCodigo(76).Text <> "" Then
        campo = "{" & Tabla & ".codfamia}"
        If OpcionListado = 309 Then campo = "{sartic.codfamia}"
        cad = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(campo, "N", 75, 76, cad) Then Exit Sub
    End If

    If OpcionListado = 54 Then
        'Cadena para seleccion D/H CLIENTE
        '--------------------------------------------
        If txtCodigo(73).Text <> "" Or txtCodigo(74).Text <> "" Then
            campo = "{sdtofm.codclien}"
            cad = "pDHCliente=""Cliente: "
            If Not PonerDesdeHasta(campo, "N", 73, 74, cad) Then Exit Sub
            
            If Me.optFrDto(1).Value Then
                'Va a mostrar por actividad. NO debe poner desde hasta cliente
                MsgBox "Va a mostrar los datos por actividad. No debe poner D/H cliente", vbExclamation
                Exit Sub
            End If
        End If
    
    
        'Cadena para seleccion D/H MARCA
        '--------------------------------------------
        If txtCodigo(77).Text <> "" Or txtCodigo(78).Text <> "" Then
            campo = "{sdtofm.codmarca}"
            cad = "pDHMarca=""Marca: "
            If Not PonerDesdeHasta(campo, "N", 77, 78, cad) Then Exit Sub
        End If
    ElseIf OpcionListado = 309 Then
        'Cadena para seleccion D/H PROVEEDOR
        '--------------------------------------------
        If txtCodigo(79).Text <> "" Or txtCodigo(80).Text <> "" Then
            campo = "{" & Tabla & ".codprove}"
            cad = "pDHProveedor=""Proveedor: "
            If Not PonerDesdeHasta(campo, "N", 79, 80, cad) Then Exit Sub
        End If
    End If
    
    '==============================================================
    If OpcionListado = 54 Then
        
    
        
        If Me.optFrDto(0).Value Then
            cadNomRPT = "rFacDtosFM.rpt"
            campo = "codclien"
        ElseIf Me.optFrDto(1).Value Then
            cadNomRPT = "rFacDtosFMAct.rpt"
            campo = "codactiv"
        Else
            campo = ""
            If Me.optFrDto(2).Value Then
                cadNomRPT = "rFacDtosFMF.rpt"
            Else
                cadNomRPT = "rFacDtosFMM.rpt"
            End If
        End If
        If campo <> "" Then
            'dtofm
            If cadSelect <> "" Then cadSelect = cadSelect & "  AND "
            If cadFormula <> "" Then cadFormula = cadFormula & " AND  "
            cadSelect = cadSelect & campo & " > 0"
            cadFormula = cadFormula & " ({sdtofm." & campo & "}>0)"
        End If
    End If
    
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If OpcionListado = 309 Then Tabla = Tabla & " INNER JOIN sartic ON " & Tabla & ".codartic=sartic.codartic"
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    LlamarImprimir False
End Sub


Private Sub cmdAceptarEst_Click()
'Estadistica margen ventas por articulo
Dim campo As String
Dim param As String
Dim Tabla As String
Dim opcPrecio As String
Dim desPrecio As String

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    
    'Cadena para seleccion D/H fecha.   Lo metere junto con la seelccion del prico a copger
    '--------------------------------------------
    param = ""
    If txtCodigo(130).Text <> "" Or txtCodigo(131).Text <> "" Then
        campo = "{slifac.fecfactu}"
        param = "Fecha: "
        If Not PonerDesdeHasta(campo, "F", 130, 131, param) Then Exit Sub
        
    End If
    
    
    
    'Parametro Precio de Valoracion
    'elegir un Precio para realizar la valoracion
    '==================================================
    desPrecio = "Valoración coste: "
    If Me.optPrecioMP2.Value Then
        opcPrecio = "{slifac.preciomp}" 'precio medio ponderado
        desPrecio = desPrecio & "Precio medio ponderado"
    ElseIf Me.optPrecioUC2.Value Then
        opcPrecio = "{slifac.preciouc}" 'precio ultima compra
        desPrecio = desPrecio & "Precio última compra"
    ElseIf Me.optPrecioStd2.Value Then
        opcPrecio = "{slifac.preciost}" 'precio standard
        desPrecio = desPrecio & "Precio standard"
    End If
    cadParam = cadParam & "pCampo=" & opcPrecio & "|"
    'Le pong las fechas(si es k las han puesto)
    desPrecio = Trim(desPrecio & "          " & param)
    cadParam = cadParam & "pDesCampo=""" & desPrecio & """|"
    numParam = numParam + 2
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H familia
    '--------------------------------------------
    If txtCodigo(88).Text <> "" Or txtCodigo(89).Text <> "" Then
        campo = "{sartic.codfamia}"
        param = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(campo, "N", 88, 89, param) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H artículo
    '--------------------------------------------
    If txtCodigo(90).Text <> "" Or txtCodigo(91).Text <> "" Then
        campo = "{slifac.codartic}"
        param = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(campo, "T", 90, 91, param) Then Exit Sub
    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    Tabla = " slifac INNER JOIN sartic ON slifac.codartic=sartic.codartic "
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    cadNomRPT = "rFacEstMargen.rpt"
    
    LlamarImprimir False
     
End Sub

Private Sub cmdAceptarFichas_Click()
'Fichas de Mantenimientos
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim campo As String


    InicializarVbles
    
    
    '===================================================
    '============ PARAMETROS ===========================
    cadParam = "|"
    
    If chkMante(4).Value = 1 Then
        'Enero 2010
        'Informe completo
        indRPT = 38
    Else
        indRPT = 13
    End If
    If Not PonerParamRPT(indRPT, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt) Then Exit Sub
    'Ejercicio
    cadParam = cadParam & "pEjercicio=""" & txtCodigo(61).Text & """|"
    numParam = numParam + 1
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
    If txtCodigo(55).Text <> "" Or txtCodigo(56).Text <> "" Then
        campo = "{scaman.codclien}"
        If Not PonerDesdeHasta(campo, "N", 55, 56, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H TIPO CONTRATO
    '--------------------------------------------
    If txtCodigo(57).Text <> "" Or txtCodigo(58).Text <> "" Then
        campo = "{scaman.codtipco}"
        If Not PonerDesdeHasta(campo, "T", 57, 58, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H Nº Mantenimiento
    '--------------------------------------------
    If txtCodigo(59).Text <> "" Or txtCodigo(60).Text <> "" Then
        campo = "{scaman.nummante}"
        If Not PonerDesdeHasta(campo, "T", 59, 60, "") Then Exit Sub
    End If
    
    
    
    
    
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    campo = "(scaman INNER JOIN sclien ON scaman.codclien=sclien.codclien) INNER JOIN stipco ON scaman.codtipco=stipco.codtipco"
    ' ----
    
    If Not HayRegParaInforme(campo, cadSelect) Then Exit Sub
    
    'Si detalla articulos o no
    cadParam = cadParam & "ImprimeArticulo=" & Abs(Me.chkMante(1).Value) & "|"
    numParam = numParam + 1
    LlamarImprimir True
End Sub


Private Sub cmdAceptarMante_Click()
'Listado de Mantenimientos
Dim cadFrom As String
Dim campo As String, devuelve As String
Dim Codigo  As String

    InicializarVbles
    cadFrom = ""
    
    Select Case OpcionListado
    Case 70, 76
        'comprobar que se ha seleccionado un Tipo de Informe
        If Me.cboTipoList.ListIndex = -1 Then Exit Sub
        'En funcion del valor seleccionado en Tipo Informe se abrira un listado diferente
        Select Case Me.cboTipoList.ListIndex
            Case 0 'Listado Equipos
                cadNomRPT = "rManListManEquipo"
            Case 1 'Listado Pagos
                cadNomRPT = "rManListManPago"
            Case 2 'Listado Importes Contrato
                cadNomRPT = "rManListManImporte"
        End Select
        
        cadTitulo = "Informe Mantenimientos"
        Codigo = "scaman"
        If OpcionListado = 76 Then
            'ANULADOS    rManListManImporteAnu.rpt
            cadTitulo = cadTitulo & " Anulados"
            Codigo = Codigo & "a"
            cadNomRPT = cadNomRPT & "Anu"
        End If
        cadNomRPT = cadNomRPT & ".RPT"
    Case 71
        cadNomRPT = "rManListRevisiones.rpt"
        Codigo = "scaman"
        cadTitulo = "Informe Revisiones"
    Case 78
    
        'PEqueña comprobacion.
        'Fecha obligatoria
        If txtCodigo(109).Text = "" Then
            MsgBox "Debe indicar la fecha", vbExclamation
            Exit Sub
        End If
    
    
        If Not PonerParamRPT(21, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt) Then Exit Sub
        Codigo = "scaman"
    Case 79
        If Not PonerParamRPT(45, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt) Then Exit Sub
        Codigo = "scaman"
    End Select
    cadFrom = "(" & Codigo & " INNER JOIN sclien ON " & Codigo & ".codclien=sclien.codclien) "
      
      
      
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
      
      
      
    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(47).Text <> "" Or txtCodigo(48).Text <> "" Then
        campo = "{" & Codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 47, 48, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion AGENTE
    '--------------------------------------------
    If txtCodigo(49).Text <> "" Or txtCodigo(50).Text <> "" Then
        campo = "{sclien.codagent}"
        'Parametro Desde/Hasta Agente
        devuelve = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(campo, "N", 49, 50, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion TIPO CONTRATO
    '--------------------------------------------
    If txtCodigo(51).Text <> "" Or txtCodigo(52).Text <> "" Then
        campo = "{" & Codigo & ".codtipco}"
        'Parametro Desde/Hasta Tipo Contrato
        devuelve = "pDHTipoCon=""Tipo Contrato: "
        If Not PonerDesdeHasta(campo, "T", 51, 52, devuelve) Then Exit Sub
    End If
    
    'Motivo de baja. Solo para anulados
    If OpcionListado = 76 Then
        If txtCodigo(115).Text <> "" Or txtCodigo(116).Text <> "" Then
            campo = "{scamana.fechabaj}"
            devuelve = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(campo, "F", 115, 116, devuelve) Then Exit Sub
        End If
    
        If txtCodigo(113).Text <> "" Or txtCodigo(114).Text <> "" Then
            campo = "{" & Codigo & ".codincid}"
            'Parametro Desde/Hasta Cliente
            devuelve = "pDHMotivo=""Motivo anul.: "
            If Not PonerDesdeHasta(campo, "T", 113, 114, devuelve) Then Exit Sub
        End If
        
        
    ElseIf OpcionListado = 79 Then 'solo para Etiquetas
        'Cadena para seleccion ACTIVIDAD
        '--------------------------------------------
        If txtCodigo(127).Text <> "" Or txtCodigo(128).Text <> "" Then
            campo = "{sclien.codactiv}"
            'Parametro Desde/Hasta Actividad
            devuelve = "pDHActividad=""Actividad: "
            If Not PonerDesdeHasta(campo, "N", 127, 128, devuelve) Then Exit Sub
        End If
        
        'Cadena para seleccion COD. POSTAL
        '--------------------------------------------
         If txtCodigo(129).Text <> "" Then
            campo = "{sclien.codpobla}"
            'Parametro Desde/Hasta Actividad
            devuelve = "pCodPosta=""C. Postal: " & txtCodigo(129).Text
            AnyadirAFormula cadFormula, campo & "=" & DBSet(txtCodigo(129).Text, "T")
            AnyadirAFormula cadSelect, campo & "=" & DBSet(txtCodigo(129).Text, "T")
'            If Not PonerDesdeHasta(campo, "N", 127, 128, devuelve) Then Exit Sub
         End If
    End If
    
    'Cadena para seleccion FECHA
    '--------------------------------------------
    If OpcionListado = 71 Then
        If txtCodigo(53).Text = "" Or txtCodigo(54).Text = "" Then
            MsgBox "Los campos Fecha Desde/Hasta deben tener valor", vbInformation
            Exit Sub
        End If
        If txtCodigo(53).Text <> "" Then
            devuelve = "Date(" & Year(txtCodigo(53).Text) & "," & Month(txtCodigo(53).Text) & "," & Day(txtCodigo(53).Text) & ")"
            'Parametro D/H Fecha
            If devuelve <> "" Then
                devuelve = "pDFecha=" & devuelve & "|"
                cadParam = cadParam & devuelve & """|"
                numParam = numParam + 1
            End If
        End If
        
        If txtCodigo(54).Text <> "" Then
            devuelve = "Date(" & Year(txtCodigo(54).Text) & "," & Month(txtCodigo(54).Text) & "," & Day(txtCodigo(54).Text) & ")"
            If devuelve <> "" Then
                devuelve = "pHFecha=" & devuelve & "|"
                cadParam = cadParam & devuelve & """|"
                numParam = numParam + 1
            End If
        End If
    End If
        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    'cadSelect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    
    'Esto lo hago siempre para gene temporales
    conn.Execute "DELETE from tmpnlotes where codusu =" & vUsu.Codigo
    
    If OpcionListado = 79 Then

        ' ---- [30/10/2009] (LAURA) : agrupar por cliente y departamento
        devuelve = "Select scaman.codclien,nomclien,nifclien,scaman.coddirec,nomdirec"
        devuelve = devuelve & " FROM " & cadFrom
        devuelve = devuelve & " LEFT OUTER JOIN sdirec ON scaman.codclien=sdirec.codclien and scaman.coddirec=sdirec.coddirec"
        If cadSelect <> "" Then devuelve = devuelve & " WHERE " & cadSelect
        devuelve = devuelve & " group by scaman.codclien,scaman.coddirec"
        
        ' ---- [ANTES] Mostraremos los clientes para imprimirles etiquetas
'        If cadSelect <> "" Then
'            devuelve = " WHERE " & cadSelect
'        Else
'            devuelve = ""
'        End If
'        devuelve = "Select sclien.codclien,nomclien,nifclien FROM " & cadFrom & devuelve
'        devuelve = devuelve & " group by 1"
        ' ----
        
        
        
        NumRegElim = 0
        frmMensajes.cadWHERE = devuelve
        frmMensajes.OpcionMensaje = 17 'Etiquetas clientes mantenimientos
        frmMensajes.Show vbModal
        If NumRegElim = 0 Then Exit Sub
    
        cadFormula = "({tmpnlotes.codusu} =" & vUsu.Codigo & ")"
    End If
    devuelve = ""
    If OpcionListado = 78 Then
        'Añado la fecha
        cadParam = cadParam & "|FechaImp=""" & txtCodigo(109).Text & """|"
        numParam = numParam + 1
    
    
    
        If Me.chkMante(2).Value Then devuelve = "EMAIL"
    End If
    
    If devuelve = "" Then
        LlamarImprimir True
    Else

        '------------------------------------------------------------
        'Envio por mail del desde hasta seleccionado
        'Comprobaremos los mail, que todos tienen

        'FALTA###
        
        
       
       
        DoEvents
        If Me.optMante(0).Value Then
            devuelve = "1"
        Else
            devuelve = "2"
        End If
        
        devuelve = "Select maiclie" & devuelve & " as el_mail,nomclien,scaman.* "
        devuelve = devuelve & " FROM  scaman INNER JOIN sclien ON scaman.codclien=sclien.codclien"
        If cadSelect <> "" Then devuelve = devuelve & " AND " & cadSelect
        
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
        devuelve = ""
        NumRegElim = 0
        While Not miRsAux.EOF
            If IsNull(miRsAux!el_mail) Then
                devuelve = devuelve & "    - " & miRsAux!nomclien & vbCrLf
            Else
                'INSERTAMOS
                NumRegElim = NumRegElim + 1
                Codigo = "insert into tmpnlotes (codusu,numalbar,fechaalb,codprove,numlinea,codartic) values ("
                Codigo = Codigo & vUsu.Codigo & ",1,'" & Format(txtCodigo(109).Text, FormatoFecha) & "'," & miRsAux!CodClien & ","
                Codigo = Codigo & NumRegElim & ",'" & miRsAux!nummante & "')"
                conn.Execute Codigo
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
        If NumRegElim = 0 Then
            MsgBox "No hay datos para poder enviar por email", vbExclamation
            Exit Sub
        End If
        
        
        If devuelve <> "" Then
            If Len(devuelve) > 500 Then devuelve = Mid(devuelve, 1, 500) & " ....."
            devuelve = "Clientes sin mail: " & vbCrLf & devuelve & "¿Continuar?"
            If MsgBox(devuelve, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
        
        
        If Not PrepararCarpetasEnvioMail Then Exit Sub
            
        
        PonerTamnyosMail True
        frmPpal.visible = False
        'Voy arriesgar.
        'Confio en que no envien por mail mas de 32000 facturas (un integer)
        Label4(22).Caption = "Preparando datos"
        Me.PBMail.Max = CInt(NumRegElim)
        Me.PBMail.Value = 0
        
        
        
        NumRegElim = 0
        If GeneracionEnvioMail() Then NumRegElim = 1
            
    
        'Si ha ido todo bien entonces numregelim=1
        If NumRegElim = 1 Then
            'Procederemos a enviarlos por mail
            If Me.optMante(0).Value Then
                '1
                cadSelect = "1"  'de maiclie2
            Else
                cadSelect = "2"  'de maiclie1
            End If
            cadSelect = "Select nomclien,maiclie" & cadSelect
            cadSelect = cadSelect & " as email,tmpnlotes.* from tmpnlotes,sclien where codusu = " & vUsu.Codigo & " and codclien=codprove"
        
            
            frmEMail.DatosEnvio = "Carta renovacion|Muchas gracias|" & Abs(chkMante(3).Value) & "|" & cadSelect & "|"
            frmEMail.Opcion = 5 'Multienvio de renovacion
            frmEMail.Show vbModal
            
            
            'Para tranquilizar las pantallas, borrar los ficheros generados
            'Confio en que no envien por mail mas de 32000 facturas (un integer)
            Label14(22).Caption = "Restaurando ...."
            Me.ProgressBar1.visible = False
            Me.Refresh
            DoEvents
            Espera 1
            PrepararCarpetasEnvioMail
            Me.ProgressBar1.visible = True
            
            
        End If
        
        
        
        
        'Es para evitar la cantidad de pantallas abriendose y cerrandose
        Me.visible = False
        PonerTamnyosMail False
        Espera 1
        Unload Me
        frmPpal.Show
    
        Screen.MousePointer = vbDefault
    
    
    End If
    
    
End Sub





Private Sub cmdAceptarNSerie_Click()
Dim campo As String
Dim cad As String
Dim nTabla As String

'    If txtCodigo(37).Text = "" Or txtCodigo(38).Text = "" Then 'And (txtCodigo(33).Text = "" Or txtCodigo(34).Text = "") Then
'        MsgBox "Debe seleccionar un cliente para Imprimir.", vbInformation
'        PonerFoco txtCodigo(37)
'        Exit Sub
'    End If
    
    InicializarVbles
    
    If Check1.Value Then
        cadNomRPT = "rRepNumSerieArt.rpt"  'Informe Numeros de Serie Articulos
        cadTitulo = "Informe por Artículo"
    Else
        cadNomRPT = "rRepNumSerie.rpt"  'Informe Numeros de Serie Articulos
        cadTitulo = "Informe Equipamiento"
    End If
    
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(37).Text <> "" Or txtCodigo(38).Text <> "" Then
        campo = Codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 37, 38, cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Del tipo de articulo
    '--------------------------------------------
    If txtCodigo(39).Text <> "" Or txtCodigo(40).Text <> "" Then
        campo = Codigo & ".codtipar}"
        'Parametro Desde/Hasta tipo de articulo
        cad = "pDHTipoArt=""Tipo Art.: "
        If Not PonerDesdeHasta(campo, "T", 39, 40, cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Del tipo de articulo
    '--------------------------------------------
    If txtCodigo(41).Text <> "" Or txtCodigo(42).Text <> "" Then
        campo = Codigo & ".codartic}"
        'Parametro Desde/Hasta tipo de articulo
        cad = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(campo, "T", 41, 42, cad) Then Exit Sub
    End If
    
    cadSelect = cadFormula
    
    '[Monica]26/09/2012: puede que el socio ya no tenga nro de uve pero sí que tengo la referencia
    '                    quito estas 2 condiciones
'    If Not AnyadirAFormula(cadFormula, "not isnull({sclien.numeruve})") Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, "{sclien.numeruve} is not null") Then Exit Sub

    '[Monica]24/10/2012: solo los registros de la sserie que no esten dados de baja
    If Check2.Value = 0 Then
        If Not AnyadirAFormula(cadFormula, "isnull({sserie.fechabaja})") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "{sserie.fechabaja} is null") Then Exit Sub
    End If
    
    
    nTabla = "sserie" ' right join sclien on sserie.codclien = sclien.codclien "
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If Not HayRegParaInforme(nTabla, cadSelect) Then Exit Sub
    
    conSubRPT = True
    
    LlamarImprimir False
    
End Sub


Private Sub cmdAceptarRepxClien_Click()
'Reparaciones por Cliente
Dim devuelve As String
Dim campo As String
Dim Tabla As String

    InicializarVbles
    
    If OpcionListado = 406 Then 'Frecuencia de reparaciones
        Tabla = "schrep"
    Else
        Tabla = "scarep"
    End If
    'David Enero 2010
    Tabla = "schrep"  'siempre va con el hco
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Desde y Hasta CLIENTE
    '---------------------------------------------
    If txtCodigo(33).Text <> "" Or txtCodigo(34).Text <> "" Then
        campo = "{" & Tabla & ".codclien}"
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 33, 34, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta DIREC/DPTO
    '-----------------------------------------------
    If txtCodigo(35).Text <> "" Or txtCodigo(36).Text <> "" Then
        campo = "{" & Tabla & ".coddirec}"
        If vParamAplic.Departamento Then
            devuelve = "pDHDpto=""Departamento: "
        Else
            devuelve = "pDHDpto=""Dirección: "
        End If
        If Not PonerDesdeHasta(campo, "N", 35, 36, devuelve) Then Exit Sub
    End If
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '----------------------------------------------
    If Trim(txtCodigo(43).Text) <> "" Or Trim(txtCodigo(44).Text) <> "" Then
        'ANTES
        'campo = "{" & tabla & ".fecentre}"
        'Marzo 2010
        'Fecha reparacion la tengo en la fechaalb
        campo = "{" & Tabla & ".fechaalb}"
        'If OpcionListado = 406 Then campo = "{" & tabla & ".fecrepar}"
        devuelve = "pDHFecha=""Fecha Rep.: "
        If Not PonerDesdeHasta(campo, "F", 43, 44, devuelve) Then Exit Sub
    End If
    
   'Comprobar si hay registros a Mostrar antes de abrir el Informe
   If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    
    If OpcionListado <> 406 Then
        cadTitulo = "Reparaciones por Cliente"
        cadNomRPT = "rRepReparxClien.rpt"
        conSubRPT = True
    Else
        cadTitulo = "Frecuencia de Reparaciones"
        cadNomRPT = "rRepFrecuencia.rpt"
        conSubRPT = True
        
        'Nº de Reparaciones, Añadirlo como parametro
        '----------------------------------------------
        cadParam = cadParam & "pNumVeces=" & txtCodigo(0).Text & "|"
        numParam = numParam + 1
        
        On Error GoTo EFrecu
        'Insertar en la tabla temporal tmpInformes el total de reparaciones para cada
        'codartic, numserie para el criterio de seleccion introducid
        devuelve = "INSERT INTO tmpinformes(codusu,nombre1,nombre2,campo1) "
        devuelve = devuelve & "SELECT " & vUsu.Codigo & ", codartic,numserie,count(numserie) as campo1 from schrep "
        devuelve = devuelve & " WHERE " & cadSelect
        devuelve = devuelve & " group by codartic,numserie"
        conn.Execute devuelve
        
        'Eliminamos de la tabla aquellos registros que no superen el nº de reparaciones introducido
        devuelve = "DELETE FROM tmpinformes where codusu=" & vUsu.Codigo & " and campo1<=" & txtCodigo(0).Text
        conn.Execute devuelve
        
        'Volver a comprobar que hay registro a mostrar para ello miramos en la
        'tabla tmpInformes que supere el nº de reparaciones a mostrar
        cadSelect = "codusu=" & vUsu.Codigo
        If Not HayRegParaInforme("tmpinformes", cadSelect) Then
            BorrarTempInformes
            Exit Sub
        End If
    End If
    
    LlamarImprimir False
    
    'Eliminar de la tabla temporal
    If OpcionListado = 406 Then BorrarTempInformes
    
EFrecu:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo nº de reparaciones.", Err.Description
End Sub


Private Sub cmdAceptarRepxDia_Click()
'Reparaciones por Dia
Dim devuelve As String
Dim param As String
Dim TotalMante As Integer
Dim RS As ADODB.Recordset
Dim fecha1 As String, fecha2 As String
Dim NomTabla As String
Dim bOk As Boolean

' ====  [17/09/2009] LAURA : Cambiar conexion CONTAB si FAZ
Dim ConexionContaOk As Boolean
Dim CambiaConta As Boolean
' ====

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    Select Case OpcionListado
        Case 63
            Codigo = "{scarep.fecentre}"
            param = "pDHFecha=""Fecha Rep.: "
            NomTabla = "scarep"
            cadNomRPT = "rRepReparxDia.rpt"
            conSubRPT = True
            cadTitulo = "Reparaciones por día"
        Case 73
            'Añadir el parametro total Mantenim. si estamos en Informe de Altas
            devuelve = "SELECT DISTINCT COUNT(*) FROM scaman "
            Set RS = New ADODB.Recordset
            RS.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                TotalMante = RS.Fields(0).Value
                cadParam = cadParam & "pTotalMante=" & TotalMante & "|"
                numParam = numParam + 1
            End If
            RS.Close
            Set RS = Nothing
            
            'Añadir el Total Mantenim. del Periodo anterior
            fecha1 = Day(txtCodigo(31).Text) & "/" & Month(txtCodigo(31).Text) & "/" & Year(txtCodigo(31).Text) - 1
            fecha2 = Day(txtCodigo(32).Text) & "/" & Month(txtCodigo(32).Text) & "/" & Year(txtCodigo(32).Text) - 1
            Codigo = "scaman.fechaini"
            devuelve = CadenaDesdeHastaBD(fecha1, fecha2, Codigo, "F")
            If devuelve <> "" And devuelve <> "Error" Then
                devuelve = "SELECT DISTINCT COUNT(*) FROM scaman WHERE " & devuelve
                Set RS = New ADODB.Recordset
                RS.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    TotalMante = RS.Fields(0).Value
                    cadParam = cadParam & "pTotalAnte=" & TotalMante & "|"
                    numParam = numParam + 1
                End If
                RS.Close
                Set RS = Nothing
            End If
            
            '================= FORMULA =========================
            Codigo = "{scaman.fechaini}"
            param = "pDHFecha=""Fecha: "
            NomTabla = "scaman"
            cadNomRPT = "rManListAltas.rpt"
            cadTitulo = "Informe Altas Mantenimientos"
        
        Case 223
            param = ""
            If Me.OptClientes Then
                Codigo = "{scafac.fecfactu}"
                NomTabla = "scafac"
            Else
                Codigo = "{scafpc.fecrecep}"
                NomTabla = "scafpc"
            End If
    End Select
   
        
    '===================================================
    '================= FORMULA =========================
    
    '== Cadena para seleccion Desde y Hasta FECHA ==
    If OpcionListado = 223 Then
        'comprobar que se han rellenado los dos campos de fecha
        'sino rellenar con fechaini o fechafin del ejercicio
        'que guardamos en vbles Orden1,Orden2
        
        'fechaini del ejercicio de la conta
        If txtCodigo(31).Text = "" Then txtCodigo(31).Text = Orden1
     
        'fecha fin del ejercicio de la conta
        If txtCodigo(32).Text = "" Then txtCodigo(32).Text = Orden2
     
        'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
        'contabilidad par ello mirar en la BD de la Conta los parámetros
        If Not ComprobarFechasConta(31) Then Exit Sub
        If Not ComprobarFechasConta(32) Then Exit Sub
    End If
    
    devuelve = CadenaDesdeHasta(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F", "Fecha Factura")
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    'Parametro D/H Fecha
    If devuelve <> "" And param <> "" Then
        cadParam = cadParam & AnyadirParametroDH(param, 31, 32) & """|"
        numParam = numParam + 1
    End If
    
    
    '## LAURA 20/06/2008
    '## Añadir frame de selec. factuar en contabilizar
    '- cadena para select en BDatos
    cadSelect = CadenaDesdeHastaBD(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F")
    
    
    '== Cadena para seleccion Desde y Hasta NºFactura ==
    If OpcionListado = 223 Then
        '- comprobar: si nº factura tienen valor tipoMov tb
        If txtCodigo(121).Text <> "" Or txtCodigo(122).Text <> "" Then
            If Me.cboTipMov.ListIndex = -1 Then
                MsgBox "Debe seleccionar el tipo de movimiento si quiere contabilizar Desde/Hasta Nº Factura.", vbInformation
                Exit Sub
            End If
            
            If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) = "" Then
                MsgBox "Debe seleccionar el tipo de movimiento si quiere contabilizar Desde/Hasta Nº Factura.", vbInformation
                Exit Sub
            End If
            
            '- añadir desde/hasta factura a cadena seleccion registros
            Codigo = "{scafac.numfactu}"
            devuelve = CadenaDesdeHasta(txtCodigo(121).Text, txtCodigo(122).Text, Codigo, "N", "Nº Factura")
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            'Parametro D/H nº factura
            If devuelve <> "" And param <> "" Then
                cadParam = cadParam & AnyadirParametroDH(param, 31, 32) & """|"
                numParam = numParam + 1
            End If
            ' añadir a la formula de bd
            devuelve = CadenaDesdeHastaBD(txtCodigo(121).Text, txtCodigo(122).Text, Codigo, "N")
            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        End If
    
                
        '- añadir tipo movimiento a cadena seleccion
        If Me.cboTipMov.ListIndex >= 0 Then
            If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) <> "" Then
                Codigo = "{scafac.codtipom}"
                devuelve = Mid(Me.cboTipMov.List(Me.cboTipMov.ListIndex), 1, 3)
                devuelve = Codigo & "=" & DBSet(devuelve, "T")
                If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
            End If
        End If
    End If

    '===================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
'    cadSelect = CadenaDesdeHastaBD(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F")
    If OpcionListado = 223 Then
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & NomTabla & ".intconta=0 "
        
        
        'Nuevo 7 Abril 08
        'Hay un parametro que permite contbilizar los tickets agrupados (NO uno a uno)
        'para ello, a partir de los FTI crearemos los FTG (tickets agrupados)
        'y los FTI NO se contabilizaran
        If Me.OptProve.Tag = "" Then
            'Contabilizacion NORMAL. Viene del MENU contabilizar
            'Comprueblo de agrupar tickets o no
            If vParamAplic.ContabilizarTicketAgrupados Then
                'Solo las de clientes
                If Me.OptClientes.Value Then cadSelect = cadSelect & " AND scafac.codtipom <> 'FTI'"
            End If
                
        Else
            'CONTABILZIACION DE LOS TICKETS AGRUPADOS
            'Añado el tipom al cad select
            cadSelect = cadSelect & " AND scafac.codtipom = 'FTG'"
        End If
    End If
    
    If Not HayRegParaInforme(NomTabla, cadSelect) Then Exit Sub
    
    
    If OpcionListado <> 223 Then
        LlamarImprimir False
    Else
    
        ' ====  [17/09/2009] LAURA : Cambiar conexion CONTAB si FAZ
        If Me.OptProve.Tag = "" Then
            If Me.OptClientes.Value Then
                devuelve = "CLI"
            Else
                devuelve = "PRO"
            End If
        Else
            devuelve = "TIK"
        End If

        CambiaConta = False
        ConexionContaOk = True
        
        If devuelve = "CLI" Then
            'CLIENTES para tipos de factura FAZ, es decir, el B
            If Me.cboTipMov.ListIndex >= 0 Then
                If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) <> "" Then
                    If Mid(Me.cboTipMov.List(Me.cboTipMov.ListIndex), 1, 3) = "FAZ" Then
                    '            If vUsu.TrabajadorB Then
                        If AbrirConexionConta(True) Then
                            CambiaConta = True
                            ConexionContaOk = True
                        Else
                            ConexionContaOk = False
                        End If
                    End If
                End If
            End If
        End If
            
        ' ====

        ' ====  [17/09/2009] LAURA : Cambiar conexion CONTAB si FAZ
        If ConexionContaOk Then
        ' ====
            '------------------------------------------------------------------------------
            '  LOG de acciones.                      5: Facturas compras
            Set LOG = New cLOG
            
            ' ====  [17/09/2009] LAURA : Cambiar conexion CONTAB si FAZ (Este bloque se pone arriba)
    '        If Me.OptProve.Tag = "" Then
    '            If Me.OptClientes.Value Then
    '                devuelve = "CLI"
    '            Else
    '                devuelve = "PRO"
    '            End If
    '        Else
    '            devuelve = "TIK"
    '        End If
            ' ====
            
            devuelve = "Contabilizar facturas " & devuelve & ":" & vbCrLf & NomTabla & vbCrLf & cadSelect
            LOG.Insertar 5, vUsu, devuelve
            Set LOG = Nothing
            '-----------------------------------------------------------------------------
            
        
            bOk = ContabilizarFacturas(NomTabla, cadSelect)
        
            TerminaBloquear
            'Eliminar la tabla TMP
            BorrarTMPFacturas
            'Desbloqueamos ya no estamos contabilizando facturas
            If Me.OptClientes.Value Then
                DesBloqueoManual ("VENCON") 'VENtas CONtabilizar
            Else
                DesBloqueoManual ("COMCON") 'COMpras CONtabilizar
            End If
            Me.FrameProgress.visible = False
            If Me.FrameTipMov.visible Then
                Me.FrameRepxDia.Height = 4400
            Else
                Me.FrameRepxDia.Height = 3500
            End If
            Me.Height = Me.FrameRepxDia.Height + 350
            Me.Refresh
            If bOk Then Unload Me
        
        ' ====  [17/09/2009] LAURA : Cambiar conexion CONTAB si FAZ
        End If
        If CambiaConta Then AbrirConexionConta False
        ' ====
    End If
End Sub



Private Sub cmdAceptarSustNSerie_Click(Index As Integer)
'Sustitucion de un Nº de Serie que este en garantía por otro nº de serie.
Dim Sql As String
Dim RS As ADODB.Recordset

    txtCodigo(81).Text = Trim(txtCodigo(81).Text)
    
    If txtCodigo(81).Text <> "" Then
        'Comprobar que el nuevo nº de serie no existe ya
        Sql = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", txtCodigo(81).Text, "T", , "codartic", Me.CadTag, "T")
        If Sql <> "" Then
            MsgBox "Ya existe ese Nº de serie.", vbExclamation
            Exit Sub
        End If
        
        On Error GoTo ESustNSerie
        conn.BeginTrans
        
        'Insertar un registro con ese nº de serie y todos los valores que tenga el
        'num serie que sustituye
        Sql = "SELECT codartic,codtipar,codclien,coddirec,tieneman,nummante,ultrepar,fingaran,codtipom,numfactu,fechavta,numalbar,numline1,codprove,numalbpr,fechacom,numline2 FROM sserie "
        Sql = Sql & " WHERE numserie=" & DBSet(NumCod, "T") & " AND codartic=" & DBSet(CadTag, "T")
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        If Not RS.EOF Then
            Sql = "(" & DBSet(txtCodigo(81).Text, "T") & ", " & DBSet(RS!codArtic, "T", "N") & "," & DBSet(RS!codTipar, "T", "N") & ","
            Sql = Sql & DBSet(RS!CodClien, "N", "S") & "," & DBSet(RS!CodDirec, "N", "S") & "," & DBSet(RS!TieneMan, "N", "S") & ","
            Sql = Sql & DBSet(RS!nummante, "T", "S") & "," & DBSet(RS!ultrepar, "F", "S") & "," & DBSet(RS!fingaran, "F", "S") & ","
            Sql = Sql & DBSet(RS!codtipom, "T", "S") & "," & DBSet(RS!NumFactu, "N", "S") & "," & DBSet(RS!FechaVta, "F", "S") & ","
            Sql = Sql & DBSet(RS!NumAlbar, "N", "S") & "," & DBSet(RS!numline1, "N", "S") & "," & DBSet(RS!codProve, "N", "S") & ","
            Sql = Sql & DBSet(RS!numalbpr, "T", "S") & "," & DBSet(RS!fechaCom, "F", "S") & "," & DBSet(RS!numline2, "N", "S") & ")"
        End If
        RS.Close
        Set RS = Nothing
        
        If Sql <> "" Then
            Sql = "INSERT INTO sserie (numserie,codartic,codtipar,codclien,coddirec,tieneman,nummante,ultrepar,fingaran,codtipom,numfactu,fechavta,numalbar,numline1,codprove,numalbpr,fechacom,numline2) VALUES " & Sql
            conn.Execute Sql
        
            'sustituir el campo numalbar del numserie viejo por 9999999
            'y poner en el campo "numsersu" en num. serie por el que se sustituye
            'limpiar campos del cliente
            Sql = "UPDATE sserie SET numalbar=9999999, numsersu=" & DBSet(txtCodigo(81).Text, "T")
            Sql = Sql & ", codclien=" & ValorNulo & ", coddirec=" & ValorNulo
            Sql = Sql & ", numfactu=" & ValorNulo
            Sql = Sql & " WHERE numserie=" & DBSet(NumCod, "T") & " AND codartic=" & DBSet(CadTag, "T")
            conn.Execute Sql
        End If
    Else
        MsgBox "Debe introducir el Nº Serie por el que se sustituye.", vbInformation
        Exit Sub
    End If

ESustNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Sustitución Nº Serie.", Err.Description
        conn.RollbackTrans
    Else
        conn.CommitTrans
        Unload Me
    End If
End Sub



Private Sub cmdAceptarTarif_Click()
Dim cadFrom As String

    InicializarVbles
   
   '========= Frame de Tarifas y Descuentos ===============================
    'Nombre fichero .rpt a Imprimir
    'Ordenar por: codtarifa, codfamia, codmarca, codartic
    Select Case OpcionListado
        Case 28: cadNomRPT = "rFacTarifasAlm.rpt"  'Listado Tarifas Articulos
        Case 29: cadNomRPT = "rFacPromociones.rpt"  'Listado Promociones
        Case 30: cadNomRPT = "rFacPreciosEsp.rpt"
        Case 245: cadNomRPT = "rFacTarifasMargen.rpt"
    End Select
    
    If Not PonerFormulaYParametrosInf28() Then Exit Sub
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If cadFormula <> "" Or (OpcionListado = 245) Then
        cadFrom = Codigo & " INNER JOIN sartic ON " & Codigo & ".codartic=sartic.codartic "
    Else
        cadFrom = Codigo
    End If
    
    'seleccionar solo los que tienen margen con error
    If OpcionListado = 245 Then
        If Me.chkMostrarErrores Then
            AnyadirAFormula cadSelect, " sartic.preciove <> sartic.preciouc + round((sartic.preciouc * if(isnull(sartic.margecom),0,sartic.margecom))/100,4)"
            AnyadirAFormula cadFormula, " {sartic.preciove} <> {sartic.preciouc} + round(({sartic.preciouc} * iif(IsNull({sartic.margecom}),0,{sartic.margecom}))/100,4)"
        End If
    End If
    
    
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    LlamarImprimir False
End Sub


Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView2
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub cmdDeselTodos_Click()
Dim I As Byte

    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = False
    Next I
End Sub

Private Sub cmdElimiaFacturas_Click()
Dim b As Boolean

'Igual hay que quitarlo


    'Proceso de borre de facturas
    If cmbEliFac.ListIndex < 0 Then Exit Sub
    
    
    
    'Tablas que voy a tener que borrar
    'Para que no se queden datos
    cadTitulo = String(60, "*") & vbCrLf & vbCrLf
    cadTitulo = cadTitulo & " Se eliminarán los datos con fecha anterior a la solicitada de: " & vbCrLf
    cadTitulo = cadTitulo & " CLIENTES: " & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "Albaranes, hco albaranes, ofertas, hco ofertas, pedidos, hco pedidos" & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "facturas, hco facturas, ventas tpv, reparaciones, hco reparaciones, produccion" & vbCrLf & vbCrLf & vbCrLf
    cadTitulo = cadTitulo & " PRVEEDORES: " & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "Albaranes, hco albaranes,  pedidos, hco pedidos, facturas, hco facturas " & vbCrLf & vbCrLf & vbCrLf
    
    Codigo = cadTitulo & "El proceso es irreversible." & vbCrLf & vbCrLf & vbCrLf & "SEGURO QUE DESEA CONTINUAR?"
    
    'Reestablecer variables
    InicializarVbles
    cadTitulo = ""
    
    If MsgBox(Codigo, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Codigo = InputBox("Password seguridad")
    Codigo = UCase(Codigo)
    If Codigo <> "ARIADNA" Then Exit Sub
    
    Label3(83).Caption = "Inicio del proceso del borre de facturas"
    Me.cmdElimiaFacturas.Enabled = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    'Conn.BeginTrans
    b = BorrarFacturas
    'Conn.RollbackTrans
    'Volvemos a dejarlo todo como estaba
    Set miRsAux = Nothing
    Orden1 = ""
    Codigo = ""
    Label3(83).Caption = ""
    Me.cmdElimiaFacturas.Enabled = True
    Screen.MousePointer = vbDefault
    
    If b Then Unload Me
End Sub

Private Sub cmdEtiqBulto_Click()
Dim I As Integer

    If Me.txtClie.Text = "" Then
        MsgBox "Ponga el cliente", vbExclamation
        Exit Sub
    End If
        
    If Val(txtBultos(1).Text) = 0 Then txtBultos(1).Text = "1"
    cadParam = "delete from tmpinformes where codusu =" & vUsu.Codigo
    conn.Execute cadParam
       
    numParam = 0
    
    Orden2 = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,nombre2,nombre3) VALUES (" & vUsu.Codigo & ","
    cadParam = "," & vParam.Codigo & ",'" & DevNombreSQL(txtNombre(10).Text) & "')"
    cadFormula = ""
    If txtBultos(7).Text <> "" Then
        'Lleva etiquetas en blanco
        For I = 1 To Val(txtBultos(7).Text)
            '           secuencia               'El cliente a blancos
            numParam = numParam + 1
            cadFormula = numParam & ",''"
            cadFormula = Orden2 & cadFormula & cadParam
            conn.Execute cadFormula
        Next I
    End If
    For I = 1 To Val(txtBultos(1).Text)
          '           secuencia               'El cliente a blancos
            numParam = numParam + 1
            cadFormula = numParam & ",'" & txtClie.Text & "'"
            cadFormula = Orden2 & cadFormula & cadParam
            conn.Execute cadFormula
            
    Next I
    cadFormula = ""
       
    'Como puede llevar saltos de linea
    Orden2 = SaltosDeLinea(txtBultos(0).Text)
    'Le pasare los datos
    cadParam = ""
    numParam = 0
    If PonerParamRPT(19, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt) Then
        Orden1 = "0"

        'Metemos los campos de direccion
        cadParam = cadParam & "Dom=""" & txtBultos(2).Text & """|"
        cadParam = cadParam & "Pob=""" & txtBultos(3).Text & """|"
        cadParam = cadParam & "Pro=""" & Trim(txtBultos(4).Text & "      " & txtBultos(5).Text) & """|"
        
        'AÑado la direccion que se ve
        cadParam = cadParam & "DireccionAlternativa=" & Orden1 & "|"
        cadParam = cadParam & "Texto= """ & Orden2 & """|"
        numParam = numParam + 2
        cadSelect = "codusu=" & vUsu.Codigo
        cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
        LlamarImprimir True
        If Me.NumCod <> "" Then Unload Me
    End If
        
End Sub

'INTENTARE METERLO DENTRO DE OTRO PROC

'Abril 2010
'En una columna de tmpinforme voy a grabar el dto para la familia
'De moemnto pong la UNO a piñon
'Veremos si hay que pedir datos o no. De momento esta a piñon

Private Sub cmdEtiqEstanteria_Click()
'Estadistica margen ventas por articulo
Dim campo As String
Dim param As String
Dim Tabla As String
Dim RS As ADODB.Recordset
Dim Li As Collection
Dim I As Integer
Dim Dto As Currency
Dim Precio As Currency
Dim Codfamia As Integer

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    cadParam = cadParam & "|pImprimeBarras=""" & Abs(Me.chkImprimeCodigoBarras.Value) & """|"
    numParam = numParam + 1
    cadParam = cadParam & "|numerodecimales=" & Me.cboDecimal.List(cboDecimal.ListIndex) & "|"
    numParam = numParam + 1
    
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H familia
    '--------------------------------------------
    If txtCodigo(94).Text <> "" Or txtCodigo(95).Text <> "" Then
        campo = "{sartic.codfamia}"
        param = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(campo, "N", 94, 95, param) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H artículo
    '--------------------------------------------
    If txtCodigo(92).Text <> "" Or txtCodigo(93).Text <> "" Then
        campo = "{sartic.codartic}"
        param = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(campo, "T", 92, 93, param) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H Fecha
    '--------------------------------------------
    If txtCodigo(123).Text <> "" Or txtCodigo(124).Text <> "" Then
        campo = "{sartic.ultfecpvp}"
        param = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(campo, "F", 123, 124, param) Then Exit Sub
    End If
    
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    Tabla = " sartic  "
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    
    
    
    'Borro tmptemporal
    Tabla = "DELETE FROM tmpinformes WHERE codusu =" & vUsu.Codigo
    conn.Execute Tabla
    
    'Añadire los tipos de IVA a esta tabla
    Tabla = "INSERT INTO tmpinformes(codusu,codigo1)  select " & vUsu.Codigo & ",codigiva from sartic"
    If cadSelect <> "" Then Tabla = Tabla & " WHERE " & cadSelect
    Tabla = Tabla & " GROUP BY codigiva"
    conn.Execute Tabla
    
    
    
    
    
    'AHora desde conta cargo los % de IVA desde la conta
    Set RS = New ADODB.Recordset
    Tabla = "Select * from tmpinformes where codusu =" & vUsu.Codigo
    RS.Open Tabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Li = New Collection
    While Not RS.EOF
        Li.Add Val(RS.Fields(1))
        RS.MoveNext
    Wend
    RS.Close
    
    
    '
    
    'Abrimos los IVAS en conta
    Tabla = "Select codigiva,porceiva from tiposiva"
    RS.Open Tabla, ConnConta, adOpenKeyset, adLockOptimistic, adCmdText
    For I = 1 To Li.Count
        Tabla = "codigiva = " & Li.item(I)
        RS.Find Tabla, , , 1
        If RS.EOF Then
            MsgBox "Tipo de IVA no encontrado en la contabilidad" & Tabla, vbExclamation
            RS.Close
            Exit Sub
        Else
            Tabla = "UPDATE tmpinformes SET porcen1 =" & TransformaComasPuntos(CStr(RS!PorceIVA))
            Tabla = Tabla & " WHERE codusu =" & vUsu.Codigo & " AND codigo1 = " & RS!codigiva
            conn.Execute Tabla
        End If
    Next I
    RS.Close
    Set Li = Nothing
    
    
    'Borramos los datos de la tabla donde iran los articulos
    Tabla = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.Codigo
    conn.Execute Tabla
    I = Me.cboDecimal.List(cboDecimal.ListIndex)
    If I = 0 Then
        Tabla = "0"
    Else
        Tabla = "#,##0." & Mid("0000", 1, I)
    End If
    frmMensajes.cadWHERE2 = Tabla
    frmMensajes.cadWHERE = cadSelect
    frmMensajes.OpcionMensaje = 15
    frmMensajes.Show vbModal
    
    'Si ha devuelto seleccionados
    Tabla = " tmpnseries   "
    cadFormula = " codusu =" & vUsu.Codigo
    
    If Not HayRegParaInforme(Tabla, cadFormula) Then Exit Sub
    
    
    'Para los articulos que hay que mostrar, si tienen dto hay que poner
    'cargalro
    If Me.chkDtoFM.Value = 1 Then
        'Cargo los dtos
        'A piñon para ALZIRA
        Tabla = "select * from sdtofm where codactiv=1 and codclien is null and codmarca is null and codfamia >=0 order by codfamia "
        RS.Open Tabla, conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        If Not RS.EOF Then
                Tabla = "SELECT tmpinformes.codusu,`sartic`.`nomartic`, `sartic`.`preciove`, `tmpinformes`.`porcen1`, `sartic`.`codartic`,codfamia,codmarca,numlinea"
                Tabla = Tabla & " FROM   ((`tmpnseries` `tmpnseries` INNER JOIN `sartic` `sartic` ON `tmpnseries`.`codartic`=`sartic`.`codartic`)"
                Tabla = Tabla & " INNER JOIN `tmpinformes` `tmpinformes` ON (`sartic`.`codigiva`=`tmpinformes`.`codigo1`)"
                Tabla = Tabla & " AND (`tmpnseries`.`codusu`=`tmpinformes`.`codusu`)) Where tmpinformes.CodUsu = " & vUsu.Codigo & " ORDER BY codfamia,codmarca"
                Set miRsAux = New ADODB.Recordset
                
                
                miRsAux.Open Tabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Codfamia = -1
                While Not miRsAux.EOF
                    
                    If Codfamia <> miRsAux!Codfamia Then
                        'Hay que buscar
                        I = 1
                    Else
                        I = 0
                    End If
                    
                    
                    If I = 1 Then
                        Codfamia = miRsAux!Codfamia
                        Dto = 0
                        RS.MoveFirst
                        Tabla = ""
                        While I = 1
                            If RS!Codfamia = Codfamia Then
                                'OK. ESte es. No muevo
                                I = 0 'salga
                                Dto = RS!dtoline1 + RS!dtoline2
                            Else
                                If RS!Codfamia > Codfamia Then RS.MoveLast
                                RS.MoveNext
                            End If
                            If RS.EOF Then I = 0
                        Wend
                    End If
                    If Not RS.EOF Then
                        'OK hay dto
                        
                        If Dto > 0 Then
                            Precio = DBLet(miRsAux!porcen1, "N")
                            Precio = (miRsAux!preciove * Precio) / 100
                            Precio = Precio + miRsAux!preciove
                            Precio = (Precio * Dto) / 100
                            
                            If Precio > 0 Then
                                Tabla = Format(Precio, FormatoCantidad)
                                
                                Tabla = "update tmpnseries set numserie = '" & Tabla & "' WHERE codusu = " & vUsu.Codigo
                                Tabla = Tabla & " AND codartic = " & DBSet(miRsAux!codArtic, "T")
                                Tabla = Tabla & " AND numlinea = " & miRsAux!numlinea
                                conn.Execute Tabla
                            End If
                        End If
                        
                    End If
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                Set miRsAux = Nothing
        End If
        
        
        RS.Close
    End If
    
    
    
    
    cadFormula = "({tmpnseries.codusu} =" & vUsu.Codigo & ")"
    
    campo = ""
    If Not PonerParamRPT(23, cadParam, numParam, campo, pImprimeDirecto, pPdfRpt) Then
        cadNomRPT = "rEtiqEsta.rpt"
    Else
        cadNomRPT = campo
    End If
    
    LlamarImprimir True
    
    BorrarTempInformes
    
    'Borramos los datos de la tabla donde iran los articulos
    Tabla = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.Codigo
    conn.Execute Tabla
    
End Sub



Private Sub cmdFactAlbaranes_Click()
    Codigo = "¿Seguro que desea continuar?"
    If MsgBox(Codigo, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    If HacerSQLListado82_83 Then
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If
End Sub

Private Sub cmdFrecuencias_Click()
Dim campo As String

    ' ---- [06/11/2009] [LAURA] : corregir informe de frecuencias
    
    InicializarVbles
    
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'Añadir parametro es departamento o direccion
    cadParam = cadParam & "|pDpto=" & IIf(vParamAplic.Departamento, 1, 0) & "|"
    numParam = numParam + 1
    

    
    '================= FORMULA =========================
    'Cadena para seleccion D/H CLIENTE
    '----------------------------------
    If txtCodigo(98).Text <> "" Or txtCodigo(99).Text <> "" Then
        campo = "{scafre.codclien}"
        If Not PonerDesdeHasta(campo, "N", 98, 99, "pDHCliente=""Cliente: ") Then Exit Sub
    End If
    
    If Not HayRegParaInforme("scafre", cadSelect) Then Exit Sub
    
    
    If Me.OptFrecResumen.Value = True Then
        cadNomRPT = "rFrecuResum.rpt" 'Informe resumen
    Else
        cadNomRPT = "rFrecuFicha.rpt" 'Ficha
    End If
    
    cadTitulo = "Frecuencias"
    
'    conSubRPT = False
    
    LlamarImprimir False
    
    ' ----

'## ANTES
'        'Le pasare los datos
'    cadParam = ""
'    numParam = 0
'
'    If PonerParamRPT(19, cadParam, numParam, cadNomRPT) Then
'        Orden1 = "0"
'       ' If Me.optDirEnvio(1).Value Then Orden1 = "1"
'
'        'AÑado la direccion que se ve
'        cadParam = cadParam & "DireccionAlternativa=" & Orden1 & "|"
'        cadParam = cadParam & "Texto= """ & Orden2 & """|"
'        numParam = numParam + 2
'        cadSelect = "codusu=" & vUsu.Codigo
'
'        LlamarImprimir
'    End If
'##
End Sub


Private Sub cmdHcoMante_Click()
    Codigo = ""
    For indCodigo = 110 To 112
        If txtCodigo(indCodigo).Text = "" Then Codigo = Codigo & "M"
        If indCodigo > 110 Then If txtNombre(indCodigo).Text = "" Then Codigo = Codigo & "M"
    Next indCodigo
    If Codigo <> "" Then
        MsgBox "Rellene correctamente todos los datos", vbExclamation
        Exit Sub
    End If
    'CUATRO CAMPOS. El primero de control
    CadenaDesdeOtroForm = "OK|" & txtCodigo(110).Text & "|" & txtNombre(111).Text & "|" & txtCodigo(112).Text & "|"
    Unload Me
End Sub

'===================================================
'===================================================
' Informe teorico mantenimientos
Private Sub cmdManteTeorico_Click()
Dim cadFrom As String
Dim campo As String, devuelve As String
Dim Codigo  As String

    InicializarVbles

    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    
        cadNomRPT = "rManListTeorico.rpt"
    
        
        cadTitulo = "Informe Mantenimientos"
        Codigo = "scaman"
    
    cadFrom = "(" & Codigo & " INNER JOIN sclien ON " & Codigo & ".codclien=sclien.codclien) "
      
    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(102).Text <> "" Or txtCodigo(103).Text <> "" Then
        campo = "{" & Codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 102, 103, devuelve) Then Exit Sub
    End If

    
    'Cadena para seleccion TIPO CONTRATO
    '--------------------------------------------
    If txtCodigo(104).Text <> "" Or txtCodigo(105).Text <> "" Then
        campo = "{" & Codigo & ".codtipco}"
        'Parametro Desde/Hasta Tipo Contrato
        devuelve = "pDHTipoCon=""Tipo Contrato: "
        If Not PonerDesdeHasta(campo, "T", 104, 105, devuelve) Then Exit Sub
    End If
       
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    
    'Si  detalla o no
    cadParam = cadParam & "Detallar=" & Abs(Me.chkMante(0).Value) & "|"
    numParam = numParam + 1

    
    LlamarImprimir False
End Sub

Private Sub cmdSelTodos_Click()
Dim I As Byte

    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = True
    Next I
End Sub


Private Sub cmdSubir_Click()
'Subir el item seleccionado del listview2 una posicion
    SubirItemList Me.ListView2
End Sub




Private Sub Form_Activate()
Dim IndiceFoco As Integer

    If PrimeraVez Then
        PrimeraVez = False
        IndiceFoco = -1
        Select Case OpcionListado
        Case 1, 2, 3, 4, 61, 20, 21, 22, 23, 24, 27, 58, 110
            '1:Listado de Marcas, 2:Almacenes Propios, 3:Tipos de Unidad
            '4:Tipos de Artículos, 6:Artículos
            '61:Motivos Pen. Rep
            '58:Proveedores, 110:Ubicaciones
             'PonerFoco txtCodigo(1)
             IndiceFoco = 1
        Case 6 '6: Informe de Articulos
            'PonerFoco txtCodigo(62)
            IndiceFoco = 62
        Case 7, 8 '7: Informe Traspaso Almacenes/Historico
                  '8: Informe Movimientos Almacen/Historico
            'PonerFoco txtCodigo(3)
            IndiceFoco = 3
        Case 9 'Informe Movimientos Artículos
            'PonerFoco txtCodigo(5)
            IndiceFoco = 5
        Case 11     '11: Listado de Articulos con componentes ' ====  [16/09/2009] LAURA
            IndiceFoco = 125
        Case 12, 13, 14, 15, 16, 17, 19
                        '12: Listado Toma de Inventario Articulos
                        '13: Listado Diferencias de Inventario Articulos
                        '14: Actualizar Diferencias de Inventario (No IMPRIME INFORME)
                        '15: Listado Articulos Inactivos
                        '16: Listado Valoracion de Stocks Inventariados
                        '17: Listado Valoración Stocks
                        '19: Inf. Stocks a una Fecha
            'PonerFoco txtCodigo(13)
            IndiceFoco = 13
        Case 18      '18: Informe Stocks MAximos y Minimos
            'PonerFoco txtCodigo(72)
            IndiceFoco = 72
        Case 28, 29, 30 '28: Informe Tarifas de Articulos
                    '29: Informe Promociones
                    '30: Informe Precios Especiales
            'PonerFoco txtCodigo(23)
            IndiceFoco = 23
        Case 31, 73 '31: Informe Ofertas
                    '73: Listado Altas Mantenimientos
            'PonerFoco txtCodigo(31)
            IndiceFoco = 31
        Case 54 'Listado Descuentos Familia/ Marca
            'PonerFoco txtCodigo(73)
            IndiceFoco = 73
            
        Case 60 '60: Informe Reparacions - Nº Series
            'PonerFoco txtCodigo(37)
            IndiceFoco = 37
        Case 63
            '63: Listado Reparaciones x día
            IndiceFoco = 31
        
        
        Case 223
            '223: Contabilizar facturas
            If Me.OptProve.Tag = "" Then
                'Contabilizacion normal clie/prov
                IndiceFoco = 31
            
            Else
                'TICKETS AGRUPADOS
                'Contabilizacion de facturas de tickets agrupadas. Lanzamos YA el proceso
                DoEvents
                cmdAceptarRepxDia_Click
                Me.Refresh
                Unload Me
                Exit Sub
            End If
        Case 246 '246: Informe margen ventas x articulo
            'PonerFoco txtCodigo(88)
            IndiceFoco = 130
        Case 64, 406 '64: Listado Reparaciones x Cliente
                     '406: List. Frecuencia de Reparaciones
            'PonerFoco txtCodigo(33)
            IndiceFoco = 33
           
        Case 82, 83
            'Marca facturar a 1
            IndiceFoco = 119
            
        ' ---- [06/11/2009] [LAURA] : corregir informe de frecuencias
        Case 96 'Informe frecuencias
            IndiceFoco = 98
        ' ----
            
        Case 309 '309:Listado precios de compra
            'PonerFoco txtCodigo(79)
            IndiceFoco = 79
        Case 407 'Sustitución Nº Serie
            'PonerFoco txtCodigo(81)
            IndiceFoco = 81
        Case 409 'List. Avisos de averias pendientes
            'PonerFoco txtCodigo(82)
            IndiceFoco = 82
        Case 95
            PonerFoco txtClie
            
        Case 99
            'PonerFoco txtCodigo(110)
            IndiceFoco = 110
        Case 247  'y Correccion de listados de precios tarias etc
             'PonerFoco txtCodigo(107)
             IndiceFoco = 107
             
        Case 120
            IndiceFoco = 59
            Me.Option1(0).Value = True
        
        
        
        End Select
        If IndiceFoco >= 0 Then PonerFoco txtCodigo(IndiceFoco)
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Load()
Dim H As Integer, W As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    
    For kCampo = 1 To 4
        Me.imgBuscar(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 0 To 36
        Me.imgBuscarG(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 39 To 79
        Me.imgBuscarG(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    
    Me.imgBuscarG(87).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    
    For kCampo = 89 To 90
        Me.imgBuscarG(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 93 To 95
        Me.imgBuscarG(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    
    Me.imgBuscarG(98).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    
    For kCampo = 0 To 12
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    For kCampo = 15 To 20
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    Me.imgFecha(109).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    
    
    
    'Ocultar todos los Frames de Formulario
    FrameListado.visible = False
    FrameInfAlmacen.visible = False
    FrameMovArtic.visible = False
    FrameInventario.visible = False
    FrameTarifas.visible = False
    FrameRepNSerie.visible = False
    FrameRepxDia.visible = False
    FrameRepxClien.visible = False
    FrameMantenimientos.visible = False
    FrameInfArticulos.visible = False
    FrameDtosFM.visible = False
    FrameRepSustNSerie.visible = False
    FrameEstMargenes.visible = False
    Me.FrameEtiqEstanteria.visible = False
    FrameBultos.visible = False
    Me.FrameFrecuencia.visible = False
    FrEliminarFacturas.visible = False
    FrameEnvioMail.visible = False
    FrameHcoMante.visible = False
    FrameAlbaranesMarcaFacturar.visible = False
    FrameServicios.visible = False
    
    Me.FrameInvArtComp.visible = False ' ====  [16/09/2009] LAURA : Listado Articulos con componentes
    
    
    CommitConexion
    
    cadTitulo = ""
    cadNomRPT = ""
    
    Select Case OpcionListado
        Case 1 To 19, 247 'Listado de ALMACEN
            ListadosAlmacen H, W
        Case 100 To 199 'Listados de ALMACEN
            ListadosAlmacen H, W
        Case 20 To 30 'Listadod de FACTURACION
            ListadosFacturacion H, W
        Case 245, 246 'Listados tarifas
            ListadosFacturacion H, W
        Case 300 To 390 'Listados de COMPRAS
            ListadosCompras H, W
        Case 407 To 490 'Listados de Reparaciones
            ListadosReparaciones H, W
    End Select
    
    
    Select Case OpcionListado
    
    'LISTADOS DE FACTURACION
    '-----------------------
        
    Case 54 '54: Listado Descuentos Familia/Marca
        H = 5450
        W = 6920
        PonerFrameVisible Me.FrameDtosFM, True, H, W
        ponerOptVisible True
        Me.Frame4.visible = False
        indFrame = 6
        
    Case 58 '58: listado Proveedores
        PonerFrameListadoVisible True, H, W
        Me.lblTitulo(1).Caption = "Listado Proveedores"
        indFrame = 1
        Codigo = "{sprove.codprove}"
        Orden1 = "{sprove.codprove}"
        Orden2 = "{sprove.nomprove}"
        
        
    'LISTADOS DE REPARACIONES
    '-------------------------
    Case 60 '60: Informe Nº Series
        
        If NumCod <> "" Then
            frmListado.txtCodigo(37).Text = Format(CLng(NumCod), "000000")
            frmListado.txtCodigo(38).Text = Format(CLng(NumCod), "000000")
            frmListado.txtNombre(37).Text = PonerNombreDeCod(frmListado.txtCodigo(37), conAri, "sclien", "nomclien", "codclien", "N")
            frmListado.txtNombre(38).Text = frmListado.txtNombre(37).Text
        End If
        
        
        H = 4995
        W = 6675
        PonerFrameVisible Me.FrameRepNSerie, True, H, W
        indFrame = 6
        Codigo = "{sserie"
        
     Case 61, 65  'Listados de Motivos Pend. Rep.
        PonerFrameListadoVisible True, H, W
        Me.lblTitulo(1).Caption = "Listado de Motivos"
        indFrame = 1
        If OpcionListado = 61 Then
            Codigo = "{smotre.codmotre}"
            Orden1 = "{smotre.codmotre}"
            Orden2 = "{smotre.nommotre}"
        Else
            Codigo = "{smotba.codmotiv}"
            Orden1 = "{smotba.codmotiv}"
            Orden2 = "{smotba.desmotiv}"
        End If
        
    Case 63, 73, 223, 224, 248
                '63: Listado Reparaciones por Día
                '73: Listado Altas Mantenimientos
                '223,224,248  Contabi facturas
                
        PonerFrameRepxDiaVisible True, H, W
        indFrame = 7
        If Me.OptProve.Tag = "" Then
            txtCodigo(31).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(32).Text = Format(Now, "dd/mm/yyyy")
        End If
        
        If OpcionListado = 223 Then
            Dim cad As String
            If vParamAplic.ContabilizarTicketAgrupados Then
                cad = "codtipom like 'FA%'"
            Else
                cad = "codtipom like 'FA%' or codtipom='FTI'"
            End If
            cad = "(" & cad & " OR codtipom = 'FRT')"
            cad = cad & " and not isnull(letraser) and trim(letraser)<>''"
            CargarCombo_TipMov Me.cboTipMov, "stipom", "codtipom", "nomtipom", cad, True
        End If
        
    Case 64, 406 'Listado Reparaciones por Cliente
                 '406: Listado Frecuencia de reparaciones
        H = 5415
        W = 6850
        PonerFrameVisible Me.FrameRepxClien, True, H, W
        indFrame = 8
        txtCodigo(43).Text = Format(Now, "dd/mm/yyyy")
        txtCodigo(44).Text = Format(Now, "dd/mm/yyyy")
        cadTitulo = "Reparaciones por Cliente"
        conSubRPT = False
        Me.Frame1.visible = (OpcionListado = 406)
        If OpcionListado = 406 Then
             cadTitulo = "Frecuencia de Reparaciones"
             Me.lblTitulo(8).Caption = "Frecuencia de Reparaciones"
             'Me.Label4(21).Caption = "Fecha Reparación:"
             txtCodigo(0).Text = "1"
        End If
        
        
        
    Case 82, 83
        
        'LIstado etiquetas estanterias
        H = Me.FrameAlbaranesMarcaFacturar.Height
        W = FrameAlbaranesMarcaFacturar.Width
        PonerFrameVisible Me.FrameAlbaranesMarcaFacturar, True, H, W
        indFrame = 82
        If OpcionListado = 82 Then
            cadTitulo = "Poner marca facturación"
            
        Else
            Label7(3).Caption = "Borre avisos cerrados"
        End If
        txtCodigo(117).visible = OpcionListado = 82
        txtCodigo(118).visible = OpcionListado = 82
        Frame7.visible = OpcionListado = 83
        conSubRPT = False
    Case 94
        'LIstado etiquetas estanterias
        H = Me.FrameEtiqEstanteria.Height
        W = FrameEtiqEstanteria.Width
        PonerFrameVisible Me.FrameEtiqEstanteria, True, H, W
        indFrame = 94
        cadTitulo = "Etiq. estanteria"
        conSubRPT = False
        cboDecimal.ListIndex = 4
        
    Case 95
        'LIstado etiquetas estanterias
        H = Me.FrameBultos.Height
        W = FrameBultos.Width
        PonerFrameVisible Me.FrameBultos, True, H, W
        indFrame = 95
        cadTitulo = "Etiq. bultos"
        conSubRPT = False
'        If vParamAplic.Departamento Then
'            optDirEnvio(1).Caption = "Departamento"
'        Else
'            optDirEnvio(1).Caption = "Dirección"
'        End If
        LimpiarTextosBultos
        Me.cmbBulto.Clear
        
        '- Traer datos del Albaran: cliente, dpto, nº bultos
        If NumCod <> "" Then PonerCamposAlbaran
        
        
        
        
    Case 96
        
        H = Me.FrameFrecuencia.Height
        W = FrameFrecuencia.Width
        PonerFrameVisible Me.FrameFrecuencia, True, H, W
        indFrame = 96
        cadTitulo = "Etiq. bultos"
        conSubRPT = False
        HabilitarTextoCliente False
        
    Case 97
        H = Me.FrEliminarFacturas.Height
        W = Me.FrEliminarFacturas.Width
        PonerFrameVisible FrEliminarFacturas, True, H, W
        indFrame = 97
        cadTitulo = "Eliminar facturas"
        conSubRPT = False
        'Textos
        '--------------------------------------------------------------------
        Label11(0).Caption = "Este proceso es irreversible." & vbCrLf & " No deberia haber nadie trabajando en esta empresa y " & vbCrLf & _
            "deberia hacer una copia de seguridad."
        
        Label11(1).Caption = ""
        CargaFechasPosibleEliminacion
        
    Case 99
        
        H = Me.FrameHcoMante.Height
        W = Me.FrameHcoMante.Width
        PonerFrameVisible FrameHcoMante, True, H, W
        indFrame = 99
        cadTitulo = "Pasar a mantenimientos anulados"
        conSubRPT = False
        txtCodigo(110).Text = Format(Now, "dd/mm/yyyy")


    Case 120
        
        'LIstado de servicios
        H = Me.FrameServicios.Height
        W = Me.FrameServicios.Width
        PonerFrameVisible Me.FrameServicios, True, H, W
        indFrame = 99
        conSubRPT = False



    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub



Private Sub Form_Unload(Cancel As Integer)
    NumCod = ""
    CadTag = ""
End Sub



Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMtoActiv_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Actividades de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoAgentes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Agentes Comerciales
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoAlPropios_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCartasOfe_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Cartas de Oferta
    txtCodigo(32).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(32).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoClientes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    If indCodigo > 0 Then
        txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
        txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        'EL 0 es para el listado de bultos
        Me.txtClie.Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
        txtClie_LostFocus
        
    End If

End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familia de Articulos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFEnvio_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Envio
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoIncid_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoMarcas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Marcas de Artículos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoMotivos_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Marcas de Artículos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoProveedor_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoSituac_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Situaciones
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoSocios_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tarifas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTarifas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tarifas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTArticulo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Artículo
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTraba_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTUnidad_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Unidad
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoUbica_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Ubicaciones de Almacen
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    imgBuscar(1).Tag = Index
    indCodigo = Index
    
    Select Case Index
    Case 1, 2 'FrameListado
        Select Case OpcionListado
            Case 1 'Listado de MARCAS
                AbrirFrmMarcas
                    
            Case 2 'Listado de ALMACENES Propios
                AbrirFrmAlmPropios
            
            Case 3  'Listado de Tipos de Unidad
                Set frmMtoTUnidad = New frmAlmTipoUnidad
                frmMtoTUnidad.DatosADevolverBusqueda = "0|1"
                frmMtoTUnidad.DeConsulta = True
                frmMtoTUnidad.Show vbModal
                Set frmMtoTUnidad = Nothing
            
            Case 4  'Listado de Tipos de Articulos
                AbrirFrmTipoArt

            Case 110 'Listado de Ubicaciones de Almacen
            
            Case 20 'Listado de Actividades de Clientes
                AbrirFrmActividades
            
            Case 21 'Listado de Zonas de Clientes
            
            
            Case 23 'Listado de Formas de Envío
                Set frmMtoFEnvio = New frmFacFormasEnvio
                frmMtoFEnvio.DatosADevolverBusqueda = "0|1"
                frmMtoFEnvio.DeConsulta = True
                frmMtoFEnvio.Show vbModal
                Set frmMtoFEnvio = Nothing
            
            Case 24 'Listado de Tarifas Venta
                AbrirFrmTarifas
            
            Case 27 'Listado de Situaciones Especiales
                Set frmMtoSituac = New frmFacSituaciones
                frmMtoSituac.DatosADevolverBusqueda = "0|1"
                frmMtoSituac.DeConsulta = True
                frmMtoSituac.Show vbModal
                Set frmMtoSituac = Nothing
                
            Case 58
                'DAVID
                indCodigo = Index
                Set frmMtoProveedor = New frmComProveedores
                frmMtoProveedor.DatosADevolverBusqueda = "0|1"
                frmMtoProveedor.Show vbModal
                Set frmMtoProveedor = Nothing
            Case 61 'Listado de Motivos Pend. Rep.
                Set frmMtoMotivos = New frmRepMotivosPend
                frmMtoMotivos.DatosADevolverBusqueda = "0|1"
                frmMtoMotivos.DeConsulta = True
                frmMtoMotivos.Show vbModal
                Set frmMtoMotivos = Nothing
        End Select
        
    Case 3, 4 'FrameInfAlmacen
            If OpcionListado = 7 Or OpcionListado = 8 Then
'            Case 7, 8 '7: Informe de Traspasos de Almacenes
                  '8: Informe de Movimientos de Almacen
                MandaBusquedaPrevia ""
            End If
    End Select
    
    PonerFoco Me.txtCodigo(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgBuscarG_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0, 1, 6, 7, 35, 36, 43, 44, 75, 76, 77, 80, 81, 93, 94  'cod. CLIENTE
            Select Case Index
                Case 0, 1: indCodigo = Index + 73
                Case 6, 7: indCodigo = Index + 27
                Case 35, 36: indCodigo = Index + 20
                Case 43, 44: indCodigo = Index + 4
                Case 49, 50: indCodigo = Index - 12
                Case 75: indCodigo = 0
                Case 76, 77, 80, 81: indCodigo = Index + 22
                Case 93, 94: indCodigo = Index + 24
            End Select
            AbrirFrmClientes
        
        Case 49, 50 ' codigo de socio
            Select Case Index
                Case 49, 50: indCodigo = Index - 12
            End Select
            AbrirFrmSocios
        
        Case 53, 54 ' codigo de socio
            indCodigo = Index + 6
            AbrirFrmSocios
        
        Case 39, 40 ' codigo de cliente
            indCodigo = Index + 18
            AbrirFrmClientes
        
        
        
        Case 2, 3, 13, 14, 19, 20, 31, 32, 57, 58, 67, 68, 73, 74 'cod. FAMILIA
            Select Case Index
                Case 2, 3: indCodigo = Index + 73
                Case 13, 14: indCodigo = Index + 3
                Case 19, 20: indCodigo = Index + 43
                Case 31, 32: indCodigo = Index - 24
                Case 57, 58: indCodigo = Index - 32
                Case 67, 68, 73, 74: indCodigo = Index + 21
            End Select
            Set frmMtoFamilia = New frmAlmFamiliaArticulo
            frmMtoFamilia.DatosADevolverBusqueda = "0|1"
            frmMtoFamilia.Show vbModal
            Set frmMtoFamilia = Nothing
            
            
        Case 90, 91, 92
            indCodigo = 22 + Index
            Set frmMtoIncid = New frmIncidencias
            frmMtoIncid.DatosADevolverBusqueda = "0|1|"
            frmMtoIncid.Show vbModal
            Set frmMtoIncid = Nothing
            
        Case 4, 5, 21, 22, 59, 60 'cod. MARCA
            Select Case Index
                Case 4, 5: indCodigo = Index + 73
                Case 21, 22: indCodigo = Index + 43
                Case 59, 60:  indCodigo = Index - 32
            End Select
            AbrirFrmMarcas
            
        Case 51, 52 ' tipo de articulo
            Select Case Index
                Case 51, 52
                    indCodigo = Index - 12
            End Select
            AbrirFrmTipoArt
            
            
        Case 8, 9 'cod. Direc/DPTO
'            Select Case Index
'                Case 8, 9:
'                Case 51, 52: indCodigo = Index - 12
'            End Select
        
            If Index = 51 Or Index = 52 Then
                'Desde hsta departamento en Numserie
                'Si no teinen el mismo cliente NO pude ver dpto
                If txtCodigo(37).Text = "" And txtCodigo(38).Text = "" Then
                        MsgBox "Ponga un cliente", vbExclamation
                    
                ElseIf (txtCodigo(37).Text <> txtCodigo(38).Text) Then
                    MsgBox "No ha puesto el mismo cliente", vbExclamation
                Else
                    indCodigo = 39 + (Index - 51)
                    MandaBusquedaPrevia "codclien = " & txtCodigo(37).Text

                End If
            End If
        Case 10, 18, 33, 34 'cod. ALMACEN
            Select Case Index
                Case 10: indCodigo = Index + 3
                Case 18: indCodigo = Index + 54
                Case 33, 34: indCodigo = Index - 22
            End Select
            AbrirFrmAlmPropios
            
        Case 11, 12, 27, 28, 29, 30, 35, 36, 61, 62, 69, 70, 71, 72, 95, 98 'cod. ARTICULO
            ' ====  [16/09/2009] LAURA : Listado Articulos con componentes (añade index 95 y 98)
            Select Case Index
                Case 11, 12: indCodigo = Index + 3
                Case 27, 28: indCodigo = Index + 43
                Case 29, 30: indCodigo = Index - 24
                Case 61, 62: indCodigo = Index - 32
                Case 69, 70, 71, 72: indCodigo = Index + 21
                Case 35, 36: indCodigo = Index + 6
                ' ====  [16/09/2009] LAURA : Listado Articulos con componentes (añade index 95 y 98)
                Case 95: indCodigo = 125
                Case 98: indCodigo = 126
                ' ====
            End Select
            Set frmMtoArticulos = New frmAlmArticulos
            frmMtoArticulos.DatosADevolverBusqueda2 = "@1@" 'Abrimos en Modo Busqueda
            frmMtoArticulos.Show vbModal
            Set frmMtoArticulos = Nothing
            
        Case 25, 26 'cod TIPO ARTICULO
            indCodigo = Index + 43
            AbrirFrmTipoArt

        Case 55, 56
            indCodigo = Index - 32
            If OpcionListado = 30 Then 'segun Informe mismo boton abre 2 distintas
               AbrirFrmClientes
            Else 'cod. TARIFA
                AbrirFrmTarifas
            End If
            
        Case 15, 16, 23, 24, 63, 64 'cod. PROVEEDOR
            Select Case Index
                Case 15, 16: indCodigo = Index + 3
                Case 23, 24: indCodigo = Index + 43
                Case 63, 64: indCodigo = Index + 16
            End Select
            Set frmMtoProveedor = New frmComProveedores
            frmMtoProveedor.DatosADevolverBusqueda = "0|1"
            frmMtoProveedor.Show vbModal
            Set frmMtoProveedor = Nothing
            
        Case 41, 42
        Case 17, 96, 97, 89 'cod. TRABAJADOR
            If Index = 89 Then
                indCodigo = 111
            Else
                indCodigo = 21
            End If
            Set frmMtoTraba = New frmAdmTrabajadores
            frmMtoTraba.DatosADevolverBusqueda = "0|1"
            frmMtoTraba.Show vbModal
            Set frmMtoTraba = Nothing
            
        Case 45, 46 'cod. AGENTE
            indCodigo = Index + 4
            Set frmMtoAgentes = New frmFacAgentesCom
            frmMtoAgentes.DatosADevolverBusqueda = "0|1"
            frmMtoAgentes.Show vbModal
            Set frmMtoAgentes = Nothing
            
        
        Case 39, 40, 53, 54 'cod. Nº CONTRATO (= nº mantenimiento)

        
        Case 84, 85, 86, 88 'RUTA DEL CLIENTE
            
        ' ---- [30/10/2009] (LAURA) : Agrupar etiquetas mantenimiento por cliente, departamento
        Case 99, 100 'COD. ACTIVIDAD
            indCodigo = Index + 28
            AbrirFrmActividades
        ' ----
        Case 87
            indCodigo = 107
            AbrirFrmTarifas
    End Select
    PonerFoco txtCodigo(indCodigo)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 0 'frameMovArtic
            indCodigo = 9
        Case 1 'frameMovArtic
            indCodigo = 10
        Case 2 'frameInventario (indFrame=4)
            indCodigo = 20
        Case 3 'frameInventario (indFrame=4)
            indCodigo = 22
        Case 4 'frameReparacionesxDia (indFrame=7)
            indCodigo = 31
        Case 5 'frameReparacionesxDia (indFrame=7)
            indCodigo = 32
        Case 6 'frameReparacionesxClien (indFrame=8)
            indCodigo = 43
        Case 7 'frameReparacionesxClien (indFrame=8)
            indCodigo = 44
        Case 8 'frameMAntenimientos
            indCodigo = 53
        Case 9 'frameMAntenimientos
            indCodigo = 54
'        Case 10 'FrameListAvisosPtes
'            indCodigo = 82
'        Case 11 'FrameListAvisosPtes
'            indCodigo = 83
        Case 13, 14
            indCodigo = Index + 102
        Case 15, 16
            indCodigo = Index + 104
        Case 17, 18
            indCodigo = Index + 106
        Case 19, 20
             indCodigo = Index + 111
        Case 109
            indCodigo = 109
        Case 10, 11
            indCodigo = Index + 45
            
   End Select
   
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)
   
   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
End Sub




Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Optclientes_Click()
    If Me.OptClientes.Value = True Then
        Label2(2).Caption = "Fecha Factura: "
    End If
    
    Me.FrameTipMov.visible = (OpcionListado = 223) And Me.OptClientes.Value = True
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub optDirEnvio_Click(Index As Integer)
    If Index = 0 Then
        txtNombre(0).Text = RecuperaValor(txtNombre(0).Tag, 1)
    Else
        txtNombre(0).Text = RecuperaValor(txtNombre(0).Tag, 2)
    End If
End Sub

Private Sub optDirEnvio_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar(1)
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptProve_Click()
    If Me.OptProve.Value = True Then
        Label2(2).Caption = "Fecha Recepción: "
    End If
    
     Me.FrameTipMov.visible = (OpcionListado = 223) And Me.OptClientes.Value = True
    
End Sub


Private Sub txtBultos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 0 Then KEYpress KeyAscii
End Sub

Private Sub txtBultos_LostFocus(Index As Integer)
    If Index = 1 Or Index = 7 Then
        'Campos NUMERICOS
        txtBultos(Index).Text = Trim(txtBultos(Index).Text)
        If txtBultos(Index).Text <> "" Then
            If Not PonerFormatoEntero(txtBultos(Index)) Then
                txtBultos(Index).Text = ""
                PonerFoco txtBultos(Index)
            End If
        End If
    End If
End Sub

Private Sub txtClie_GotFocus()
    PonerFoco txtClie
    
End Sub

Private Sub txtClie_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtClie_LostFocus()
Dim Reestablecer As Boolean
Dim CliVario As Boolean
Dim RS As ADODB.Recordset
Dim Ind As Integer
                Screen.MousePointer = vbHourglass
                txtClie.Text = Trim(txtClie.Text)
                Orden2 = ""
                CliVario = False
                If txtClie = "" Then
                    Reestablecer = True
                Else
                    If Not PonerFormatoEntero(txtClie) Then
                        Reestablecer = True
                    Else
                        cmbBulto.Clear
                        Set RS = New ADODB.Recordset
                        Codigo = "select nomclien,domclien,sclien.codpobla as cpos,sclien.pobclien,proclien,sdirec.*,clivario from sclien left join sdirec on sclien.codclien=sdirec.codclien "
                        Codigo = Codigo & " WHERE sclien.codclien =" & txtClie.Text
                        Codigo = Codigo & " order by nomdirec"
                        RS.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        Orden1 = ""
                        
                        While Not RS.EOF
                            'Meto primero la direccion de la ficha
                            If Orden1 = "" Then
                                cmbBulto.AddItem "Ppal:  " & DBLet(RS.Fields(1), "T") & " - " & DBLet(RS.Fields(3), "T")
                                txtBultos(2).Tag = DBLet(RS.Fields(1), "T") & "|"
                                txtBultos(3).Tag = DBLet(RS.Fields(3), "T") & "|"
                                txtBultos(4).Tag = DBLet(RS.Fields(2), "T") & "|"
                                txtBultos(5).Tag = DBLet(RS.Fields(4), "T") & "|"
                                txtBultos(6).Tag = "|"
                                Orden1 = "T"
                                
                                Orden2 = RS!nomclien
                                CliVario = DBLet(RS!CliVario, "N") = 1
                            End If
                            'Las direcciones alternativas
                            If Not IsNull(RS!domdirec) Then
                                'TIENE DIRECCION ALTERNATIVA
                                txtBultos(2).Tag = txtBultos(2).Tag & DBLet(RS!domdirec, "T") & "|"
                                txtBultos(3).Tag = txtBultos(3).Tag & DBLet(RS!pobdirec, "T") & "|"
                                txtBultos(4).Tag = txtBultos(4).Tag & DBLet(RS!codpobla, "T") & "|"
                                txtBultos(5).Tag = txtBultos(5).Tag & DBLet(RS!prodirec, "T") & "|"
                                txtBultos(6).Tag = txtBultos(6).Tag & "|"
'                                cmbBulto.AddItem "       " & DBLet(RS!domdirec, "T") & " - " & DBLet(RS!pobdirec, "T")
                                cmbBulto.AddItem DBLet(RS!nomdirec, "T") & ":   " & DBLet(RS!domdirec, "T") & " - " & DBLet(RS!pobdirec, "T")
                                If Me.CadTag = CStr(DBLet(RS!CodDirec, "N")) Then
                                    Ind = cmbBulto.ListCount - 1
                                End If
                            End If
                            RS.MoveNext
                        Wend   '
                        If cmbBulto.ListCount > 0 Then
                            If Ind > 0 Then
                                cmbBulto.ListIndex = Ind
                            Else
                                cmbBulto.ListIndex = 0
                            End If
                            'PonerCamposDireccionBultos 0 'Lo hace el poner a 0 el list index
                        Else
                            Reestablecer = True
                        End If
                        RS.Close
                        Set RS = Nothing

                        
                    End If
                End If
                    'La direccion
                If Reestablecer Then
                    txtClie.Text = ""
                    'Hbilitamos o no
                    cmbBulto.Clear
                    LimpiarTextosBultos
                    txtNombre(10).Text = ""
                    CliVario = False
                Else
                    
                    txtNombre(10).Text = Orden2
                End If
                HabilitarTextoCliente CliVario
                
             Screen.MousePointer = vbDefault
    
End Sub

Private Sub HabilitarTextoCliente(Habilitar As Boolean)
    If Not Habilitar Then
        txtNombre(10).BackColor = &H80000018
    Else
        txtNombre(10).BackColor = &H80000005
    End If
    txtNombre(10).Locked = Not Habilitar
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub


Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Tabla As String
Dim codCampo As String, NomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean 'Si es campo Cod-Descripcion llama a PonerNombreDeCod


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    EsNomCod = False
    If Index = 1 Or Index = 2 Then
    'el mismo frame ( y por tanto los mismos campos) se utilizan para distintos
    'informes. Según de donde llamemos código de una tabla u otra
        Select Case OpcionListado
            Case 1 'Listado MARCAS
                EsNomCod = True
                Tabla = "smarca"
                codCampo = "codmarca"
                NomCampo = "nommarca"
                TipCampo = "N"
                Formato = "0000"
                Titulo = "Marca"
                
            Case 2 'Listado ALMACENES Propios
                EsNomCod = True
                Tabla = "salmpr"
                codCampo = "codalmac"
                NomCampo = "nomalmac"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Almacen Propio"
                
            Case 3 'Listado Tipos UNIDADES
                EsNomCod = True
                Tabla = "sunida"
                codCampo = "codunida"
                NomCampo = "nomunida"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Tipo Unidad"
                
            Case 4 'Listado Tipos Artículos
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), 1, "stipar", "nomtipar", "codtipar", "Tipo de Artículo", "T")
    
            Case 110 'Listado Ubicaciones Almacen
           
            
            Case 20 'Listado ACTIVIDADES de Clientes
                EsNomCod = True
                Tabla = "sactiv"
                codCampo = "codactiv"
                NomCampo = "nomactiv"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Actividad de Cliente"
            
            
            Case 23 'Listado Formas de Envío
                EsNomCod = True
                Tabla = "senvio"
                codCampo = "codenvio"
                NomCampo = "nomenvio"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Forma de Envío"
            
            Case 24 'Listado Tarifas Venta
                EsNomCod = True
                Tabla = "starif"
                codCampo = "codlista"
                NomCampo = "nomlista"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Tarifa de Venta"
            
            Case 27 'Listado SITUACIONES Especiales
                EsNomCod = True
                Tabla = "ssitua"
                codCampo = "codsitua"
                NomCampo = "nomsitua"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Situación Especial"
            
            Case 58 'Listado PROVEEDORES
                EsNomCod = True
                Tabla = "sprove"
                codCampo = "codprove"
                NomCampo = "nomprove"
                TipCampo = "N"
                Formato = "000000"
                Titulo = "Proveedor"
            
            Case 61 'Listado MOTIVOS Pend. Rep.
                EsNomCod = True
                Tabla = "smotre"
                codCampo = "codmotre"
                NomCampo = "nommotre"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Motivos Pend. Rep."
                
            Case 65 'Listados NOTIVOS baja equipos
                EsNomCod = True
                Tabla = "smotba"
                codCampo = "codmotiv"
                NomCampo = "desmotiv"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Motivos Baja equipos"
        End Select
        
    ElseIf Index = 3 Or Index = 4 Then
         '7: Informe Traspaso Almacenes
         '8: Informe Movimientos Almacen
         txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
    Else
        Select Case Index
        Case 0, 86, 87
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoEntero txtCodigo(Index)
                If (Index = 86 Or Index = 87) Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            End If
            
        Case 5, 6, 14, 15, 29, 30, 41, 42, 70, 71, 90, 91, 92, 93, 125, 126 'Cod. ARTICULO
            ' ====  [16/09/2009] LAURA : añade index 125,126
            EsNomCod = True
            Tabla = "sartic"
            codCampo = "codartic"
            NomCampo = "nomartic"
            TipCampo = "T"
            Titulo = "Artículo"
            txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        
        Case 7, 8, 16, 17, 25, 26, 62, 63, 75, 76, 88, 89, 94, 95 'Cod. FAMILIA
            EsNomCod = True
            Tabla = "sfamia"
            codCampo = "codfamia"
            NomCampo = "nomfamia"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Familia"
        
        'FECHA Desde Hasta
        Case 9, 10, 20, 22, 31, 32, 43, 44, 53, 54, 55, 56, 82, 83, 109, 110, 115, 116, 119, 120, 123, 124, 130, 131
            If txtCodigo(Index).Text <> "" Then
                If Index = 22 And OpcionListado = 19 Then 'Este campo sera Hora y no Fecha
                    PonerFormatoHora txtCodigo(Index)
                Else
                    PonerFormatoFecha txtCodigo(Index)
                    If OpcionListado = 223 And txtCodigo(Index).Text <> "" Then
                        'Contabilizar facturas
                        If Not ComprobarFechasConta(Index) Then PonerFoco txtCodigo(Index)
                    End If
                End If
            End If
            
        Case 11, 12, 13, 72 'ALMACENES Propios
            EsNomCod = True
            Tabla = "salmpr"
            codCampo = "codalmac"
            NomCampo = "nomalmac"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Almacen Propio"
            
        Case 18, 19, 66, 67, 79, 80 'PROVEEDOR
            EsNomCod = True
            Tabla = "sprove"
            codCampo = "codprove"
            NomCampo = "nomprove"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Proveedor"
        
        Case 21, 96, 97, 111 'Cod. Operario/Trabajador
            EsNomCod = True
            Tabla = "straba"
            codCampo = "codtraba"
            NomCampo = "nomtraba"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Trabajador"
        
        Case 23, 24, 107
            EsNomCod = True
            TipCampo = "N"
            If OpcionListado = 30 Then 'Precios Especiales
                Tabla = "sclien"
                codCampo = "codclien"
                NomCampo = "nomclien"
                Formato = "000000"
                Titulo = "Cliente"
            Else   'Tarifas Precios
                Tabla = "starif"
                codCampo = "codlista"
                NomCampo = "nomlista"
                Formato = "000"
                Titulo = "Tarifa de Venta"
            End If
        
        Case 27, 28, 64, 65, 77, 78 'MARCAS
            EsNomCod = True
            Tabla = "smarca"
            codCampo = "codmarca"
            NomCampo = "nommarca"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Marca"
        
        Case 31 'Nº de Oferta
            If txtCodigo(Index).Text = "" Then Exit Sub
            codCampo = DevuelveDesdeBDNew(conAri, "scapre", "numofert", "numofert", txtCodigo(Index).Text, "N")
            If codCampo = "" Then
                MsgBox "No existe el código de Oferta: " & NumCod, vbInformation
                PonerFoco txtCodigo(Index)
            Else
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            
        Case 32, 43 'Carta de la Oferta
            EsNomCod = True
            Tabla = "scartas"
            codCampo = "codcarta"
            NomCampo = "descarta"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Cartas para Ofertas"
            
        Case 37, 38, 33, 34, 47, 48, 73, 74, 98, 99, 102, 103, 117, 118 'Cod. CLIENTE
            EsNomCod = True
            Tabla = "sclien"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Cliente"
            
        Case 57, 58 'Cliente
            EsNomCod = True
            Tabla = "scliente"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Cliente"
            
            
        Case 59, 60 'Socios
            EsNomCod = True
            Tabla = "sclien"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Socio"
            
        Case 112, 113, 114
            EsNomCod = True
            Tabla = "sincid"
            codCampo = "codincid"
            NomCampo = "nomincid"
            TipCampo = "T"
            'Formato = "0000"
            Titulo = "Incidencias"
            
        Case 35, 36  'Direcc./Dpto del Cliente
            If txtCodigo(Index).Text = "" Then
                txtNombre(Index).Text = ""
                Exit Sub
            End If
            txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            'comprobar el departamento del cliente, cuando en el campo
            'Desde/Hasta se ha seleccionado un único cliente
            If Index = 39 Or Index = 40 Then
                If txtCodigo(37).Text <> txtCodigo(38).Text Then
                    MsgBox "Seleccionar dpto/direc. solo cuando se seleccione un único cliente.", vbInformation
                    txtCodigo(Index).Text = ""
                    Exit Sub
                End If
            ElseIf Index = 35 Or Index = 36 Then
                If txtCodigo(33).Text <> txtCodigo(34).Text Then
                    MsgBox "Seleccionar dpto/direc. solo cuando se seleccione un único cliente.", vbInformation
                    txtCodigo(Index).Text = ""
                    Exit Sub
                End If
            End If
            
            'Comprobar que el cliente seleccionado tiene esa direccion o dpto
            codCampo = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", txtCodigo(Index - 2).Text, "N", , "coddirec", txtCodigo(Index).Text, "N")
            txtNombre(Index).Text = codCampo 'Nombre direc. o dpto
            If codCampo = "" Then 'No existe el dpto
                If vParamAplic.Departamento Then
                    codCampo = " el Departamento "
                Else
                    codCampo = " la Dirección "
                End If
                codCampo = "No existe" & codCampo & txtCodigo(Index).Text & " para el cliente: "
                codCampo = codCampo & txtCodigo(Index - 2).Text & " - " & txtNombre(Index - 2).Text
                MsgBox codCampo, vbInformation
                txtCodigo(Index).Text = ""
                PonerFoco txtCodigo(Index)
            End If
        
        
        Case 41, 42, 59, 60 'Nº Contrato
'            If txtCodigo(Index).Text <> "" Then
'                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
'            End If

        
        Case 49, 50 'Cod. AGENTE
            EsNomCod = True
            Tabla = "sagent"
            codCampo = "codagent"
            NomCampo = "nomagent"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Agente"
            
        Case 51, 52, 104, 105 'Tipos Contratos/MAntenimientos
            EsNomCod = True
            Tabla = "stipco"
            codCampo = "codtipco"
            NomCampo = "nomtipco"
            TipCampo = "T"
            Titulo = "Tipos de Contratos"
            
        Case 61 'Año Ejercicio
            If txtCodigo(Index).Text = "" Then Exit Sub
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBox "El Ejercicio debe ser un Año", vbInformation
                Exit Sub
            End If
        
        Case 39, 40, 68, 69 'Tipos de Articulos
            EsNomCod = True
            Tabla = "stipar"
            codCampo = "codtipar"
            NomCampo = "nomtipar"
            TipCampo = "T"
            Titulo = "Tipo de Articulo"
            
            
        Case 127, 128 'ACTIVIDADES del cliente
            EsNomCod = True
            Tabla = "sactiv"
            codCampo = "codactiv"
            NomCampo = "nomactiv"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Actividades"
            
        Case 121, 122 'Nº Factura
            If PonerFormatoEntero(txtCodigo(Index)) Then
                
                
            End If
        End Select
    End If
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, NomCampo, codCampo, Titulo, TipCampo)
            
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
            Else
                txtNombre(Index).Text = ""
            End If
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, NomCampo, codCampo, Titulo, TipCampo)
        End If
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    Conexion = conAri    'Conexión a BD: Aritaxi
    Select Case OpcionListado
        Case 7 'Traspaso de Almacenes
            cad = cad & "Nº Trasp|scatra|codtrasp|N|0000000|40·Almacen Origen|scatra|almaorig|N|000|20·Almacen Destino|scatra|almadest|N|000|20·Fecha|scatra|fechatra|F||20·"
            Tabla = "scatra"
            Titulo = "Traspaso Almacenes"
        Case 8 'Movimientos de Almacen
            cad = cad & "Nº Movim.|scamov|codmovim|N|0000000|40·Almacen|scamov|codalmac|N|000|30·Fecha|scamov|fecmovim|F||30·"
            Tabla = "scamov"
            Titulo = "Movimientos Almacen"
        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
                   '12: Inventario Articulos
                   '14:Actualizar Diferencias de Stock Inventariado
                   '16: Listado Valoracion stock inventariado
            cad = cad & "Código|sartic|codartic|T||30·Denominacion|sartic|nomartic|T||70·"
            Tabla = "sartic"
            Titulo = "Articulos"
            
            
        Case 60
            If vParamAplic.Departamento Then
                Titulo = "Dptos Cliente: "

            Else
                 Titulo = "Direc. Cliente: "
            End If
            Titulo = Titulo & txtCodigo(37).Text & " - " & txtNombre(37)
            cad = cad & "Codigo|sdirec|coddirec|N|000|15·"
            cad = cad & "Descripcion|sdirec|nomdirec|T||55·"
            Tabla = "sdirec"
    End Select
          
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        'frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = Conexion
'        frmB.vBuscaPrevia = 1
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        Select Case OpcionListado
            Case 7, 8 'Informe Traspasos Almacen
                txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
                PonerFoco txtCodigo(indCodigo)
            Case 9, 12, 13, 14, 15, 16, 17, 60 '9: Informe Movimiento Articulos
                                'Inventario Articulos
                                '14: Actualizar diferencias Stock Inventariado
                                '16: Listado Valoracion stock inventariado
                txtCodigo(indCodigo).Text = RecuperaValor(CadenaDevuelta, 1)
                txtNombre(indCodigo).Text = RecuperaValor(CadenaDevuelta, 2)
                PonerFoco txtCodigo(indCodigo)
            
            
                
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerFrameListadoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de los mantenimientos de tabla: smarca, stipar,...

    H = 4695
    W = 6555
    PonerFrameVisible Me.FrameListado, visible, H, W

    If visible = True Then
        Me.Optcodigo.Value = True
    End If
End Sub



Private Sub PonerFrameInventarioVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Inventario Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Inventario
Dim VerOpcion As Boolean
       
    If visible = True Then
        H = 6400
        W = 7995
        VerOpcion = (OpcionListado = 15) Or (OpcionListado = 16) Or (OpcionListado = 17) Or (OpcionListado = 19)
        
        If VerOpcion Then
            H = 6900
            Me.cmdAceptar(4).Top = 6360
            Me.cmdCancel(4).Top = 6360
        ElseIf OpcionListado = 13 Then
            H = 6000
            Me.cmdAceptar(4).Top = 5200
            Me.cmdCancel(4).Top = Me.cmdAceptar(4).Top
        End If
        PonerFrameVisible Me.FrameInventario, visible, H, W

                
        '======================================
        'Valorar con Precios
        If VerOpcion Then
            Me.FrameValorar.visible = VerOpcion
            Me.FrameValorar.Left = 500
            If OpcionListado = 17 Then
                Me.FrameValorar.Top = 4500
            Else
                Me.FrameValorar.Top = 5000
            End If
            Me.chkSinStock.visible = VerOpcion
        End If
        
                
                
        '====================================
        'Poner el Trabajador
        VerOpcion = (OpcionListado = 14)
        Me.Label4(7).visible = VerOpcion
        Me.imgBuscarG(17).visible = VerOpcion
        Me.txtCodigo(21).visible = VerOpcion
        Me.txtNombre(21).visible = VerOpcion
'        If VerOpcion Then txtCodigo(21).TabIndex = 47
        
        
        '======================================
        'Fecha Listados
        If OpcionListado = 15 Then '15: Listado Articulos Inactivos
            Me.Label4(5).Caption = "Fecha Inactividad"
        ElseIf OpcionListado = 19 Then
            Me.Label4(5).Caption = "Fecha Stock"
        Else
            Me.Label4(5).Caption = "Fecha Inventario"
        End If
        
        VerOpcion = (OpcionListado = 12) Or (OpcionListado = 15) Or (OpcionListado = 16) Or (OpcionListado = 19)
        Me.Label4(5).visible = VerOpcion  'campo fecha
        Me.imgFecha(2).visible = VerOpcion
        Me.txtCodigo(20).visible = VerOpcion
        'campo HAsta Fecha
        Me.Label4(8).visible = (OpcionListado = 16)
        'Si opcionlistado=19 este campo sera la hora
        Me.Label4(9).visible = (OpcionListado = 16) Or (OpcionListado = 19)
        If OpcionListado = 19 Then
            Me.Label4(9).Caption = "Hora"
            Me.Label4(9).Left = 4250
            Me.txtCodigo(22).Left = 4700
        End If
        Me.imgFecha(3).visible = (OpcionListado = 16)
        Me.txtCodigo(22).visible = (OpcionListado = 16) Or (OpcionListado = 19)
        If OpcionListado = 16 Then
            Me.Label4(8).Left = 2280
            Me.imgFecha(2).Left = 2820
            Me.txtCodigo(20).Left = 3120
            Me.Label4(9).Left = 4680
            Me.imgFecha(3).Left = 5160
            Me.txtCodigo(22).Left = 5430
'            txtCodigo(22).TabIndex = 48
        End If
        
        
        '====================================
        'Activar o no los check de Opcion:
        VerOpcion = (OpcionListado = 12) Or (OpcionListado = 13) Or (OpcionListado = 16) Or (OpcionListado = 17) Or (OpcionListado = 19) Or OpcionListado = 15
                    '12: Toma de Inventario
                    '13: Listado Diferencias stock
        
        Me.FrameOpciones.visible = VerOpcion
        Me.FrameOpciones.Top = 5000
        If OpcionListado = 13 Then
            Me.FrameOpciones.Top = 4500
            Me.FrameOpciones.BorderStyle = 0
        End If
        Me.FrameOpciones.Height = 1000

        Me.chkSaltaPag.visible = VerOpcion
        Me.chkValorado.visible = (OpcionListado = 16) Or (OpcionListado = 17)

        
        VerOpcion = (OpcionListado = 12)
        If VerOpcion Or OpcionListado = 13 Then Me.FrameOpciones.Left = 700
        Me.chkImprimeStock.visible = VerOpcion
        Me.chkImprimeStock.Top = 600
        If VerOpcion Then Me.txtCodigo(20).Text = Date
    End If
End Sub



Private Sub PonerFrameTarifasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Tarifas Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Tarifas
Dim VerOpcion As Boolean

    H = 6375
    If OpcionListado = 245 Then H = 5675
    W = 7635
    PonerFrameVisible Me.FrameTarifas, visible, H, W
    
    If visible = True Then
        '====================================
        '28: Tarifas Precios 29: Promociones
        VerOpcion = (OpcionListado = 28) Or (OpcionListado = 29)
        Me.chkSaltaPagTarif.visible = VerOpcion
        Me.Label4(12).visible = VerOpcion
        
        '====================================
        If OpcionListado = 30 Then Me.Label4(11).Caption = "Cliente"
        
        
        '245: Control margenes tarifas
        '==================================
        VerOpcion = (OpcionListado = 245)
        Me.chkMostrarErrores.visible = VerOpcion
        'Decimales
        Me.cboDecimales.visible = VerOpcion
        Label4(88).visible = VerOpcion
        If VerOpcion Then
            Me.chkMostrarErrores.Top = 4600
            Label4(88).Top = 4300
            cboDecimales.Top = 4600
            
            'no mostrar seleccion de marca D/H
            Me.Label4(13).visible = Not VerOpcion
            Me.Label3(13).visible = Not VerOpcion
            Me.Label3(14).visible = Not VerOpcion
            Me.imgBuscarG(59).visible = Not VerOpcion
            Me.imgBuscarG(60).visible = Not VerOpcion
            Me.txtCodigo(27).visible = Not VerOpcion
            Me.txtCodigo(28).visible = Not VerOpcion
            Me.txtNombre(27).visible = Not VerOpcion
            Me.txtNombre(28).visible = Not VerOpcion
            'subir seleccion Articulo D/H al sitio de la marca
            Me.Label4(14).Top = Me.Label4(13).Top
            Me.Label3(15).Top = Me.Label3(13).Top
            Me.Label3(16).Top = Me.Label3(14).Top
            Me.imgBuscarG(61).Top = Me.imgBuscarG(59).Top
            Me.imgBuscarG(62).Top = Me.imgBuscarG(60).Top
            Me.txtCodigo(29).Top = Me.txtCodigo(27).Top
            Me.txtCodigo(30).Top = Me.txtCodigo(28).Top
            Me.txtNombre(29).Top = Me.txtNombre(27).Top
            Me.txtNombre(30).Top = Me.txtNombre(28).Top
            Me.cmdAceptarTarif.Top = 4600
            Me.cmdCancel(indFrame).Top = Me.cmdAceptarTarif.Top
        End If
    End If
End Sub


Private Sub PonerFrameRepxDiaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de las Reparaciones x dia, de tabla: scarep
    

    If OpcionListado = 223 Or OpcionListado = 224 Then
        H = 4400
        W = 6100
    Else
        H = 3500
        W = 6000
    End If
    
    PonerFrameVisible Me.FrameRepxDia, visible, H, W
    
    If visible = True Then
        Me.Caption = "AriTaxi"
'        Me.FrameContab.Enabled = False
'        Me.OptClientes.Enabled = False
        Me.FrameContab.visible = (OpcionListado = 223 Or OpcionListado = 224 Or OpcionListado = 248)
        Me.FrameTipMov.visible = (OpcionListado = 223)
        Me.FrameProgress.visible = False
        
        '-- alto del boton aceptar y cancelar
        
        If OpcionListado = 223 Or OpcionListado = 224 Then
            Me.cmdAceptarRepxDia.Top = 3800
        Else
            Me.cmdAceptarRepxDia.Top = 2800
        End If
        Me.cmdCancel(7).Top = Me.cmdAceptarRepxDia.Top
        
        Select Case OpcionListado
            Case 63
                Me.lblTitulo(0).Caption = "Reparaciones por Día"
                Me.Label2(2).Caption = "Fecha Reparación:"
                Frame2.Top = 1350
            Case 73
                Me.lblTitulo(0).Caption = "Altas de Mantenimientos"
                Me.Label2(2).Caption = "Fecha Mantenimiento:"
                Frame2.Top = 1350
            Case 223, 224, 248 'Pedir datos para contabilizar facturas
                Me.lblTitulo(0).Caption = "Contabilizar Facturas"
                Me.Label2(2).Caption = "Fecha Factura:"
                Frame2.Top = 1680
                Me.FrameTipMov.Top = 2650
                
                
                Me.OptProve.Tag = ""
                If OpcionListado = 248 Then
                    Me.OptProve.Tag = "TIK"  'Son las de tickets agrupados
                    OpcionListado = 223
                End If
                If OpcionListado = 224 Then
                    Me.OptProve.Value = True
                    OpcionListado = 223
                End If
        End Select
    End If
End Sub
Private Sub ponerFrameArticulosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el informe de Articulos, de tabla: sartic
Dim b As Boolean



    'Hay una opcion mas que mostrara este frame. la 247. Correccion de de tarifas e importes en articulos
    FrameTapaINCORRECTO.visible = False
    chkMinimoCorreg.visible = False
    b = (OpcionListado = 6)
    If b Then
        Me.Label9.Caption = "Informe de Articulos"
       
        W = 8595
    Else
        If OpcionListado = 18 Then
            Me.Label9.Caption = "Informe Stocks Maximos y Minimos"
            Label4(36).Caption = "Almacén"
        Else
            'NUEVA OCPION:  247
            'Corregir tarifas y eso
            chkMinimoCorreg.visible = True
            Me.Label9.Caption = "Verificación tarifas y P.V.P."
            FrameTapaINCORRECTO.visible = True
            Label4(36).Caption = "Tarifa"
            cmbDecimales.ListIndex = 0
        End If
        W = 7395
       
    End If
    H = 6820
    
    
    PonerFrameVisible Me.FrameInfArticulos, visible, H, W
    If visible = True Then
        'visible orden campos si opcionlistado=6
        Me.FrameOrden.visible = b
        Label4(36).visible = Not b

        Me.imgBuscarG(18).visible = Not b
        Me.txtCodigo(72).visible = Not b
        Me.txtNombre(72).visible = Not b
        
        'Visible Frame stocks Max Minimos si opcionlistado=18
        Me.optStockMax.Value = True
        Me.FrameStockMaxMin.visible = OpcionListado = 18
  
        FrameSituacionArticulo.visible = OpcionListado = 6
    
    
        'REajustes.
        'El articulo NO se muestra si la opcion es 247
        b = OpcionListado <> 247
        PonerLabelsArticulosFrameVisible b
        Label4(75).visible = Not b
        cmbDecimales.visible = Not b
        Label4(90).visible = Not b
        cmbDecimales.visible = Not b
    
    End If
End Sub


Private Sub PonerLabelsArticulosFrameVisible(Si As Boolean)
    Label4(38).visible = Si
    Label3(51).visible = Si
    imgBuscarG(27).visible = Si
    txtCodigo(70).visible = Si
    txtNombre(70).visible = Si
    Label3(54).visible = Si
    imgBuscarG(28).visible = Si
    txtCodigo(71).visible = Si
    txtNombre(71).visible = Si
    chkMinimoCorreg.visible = Not Si
    
End Sub


Private Sub CargarListView()
'Carga el List View del frame: frameMovimArtic
'con los parametros de la tabla: stipom (Tipos de Movimientos)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim ItmX As ListItem

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", 800
    ListView1.ColumnHeaders.Add , , "Descripción", 2250
    
    Sql = "select * from stipom where muevesto=1"
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = RS.Fields(0).Value
        ItmX.Checked = True
        ItmX.SubItems(1) = RS.Fields(1).Value
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub



Private Sub CargarListViewOrden()
'Carga el List View del frame: frameInfArticulos
'para establecer el orden en que se van a mostrar los datos en el Informe
'Orden: Familia, MArca, Proveedor, Tipo de Articulo, Articulo
Dim ItmX As ListItem

    'Los encabezados
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Campo", 1600
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Familia"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Marca"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Proveedor"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Tipo Articulo"
End Sub


Private Function PonerFormulaYParametrosInf9() As Boolean
Dim cad As String
Dim todosMarcados As Boolean
Dim devuelve As String
Dim I As Byte

    PonerFormulaYParametrosInf9 = False
    InicializarVbles
    
    'Parametro EMPRESA
    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = 1
        
    '-- Cadena para seleccion Desde y Hasta ARTICULO
    If txtCodigo(5).Text <> "" Or txtCodigo(6).Text <> "" Then
        Codigo = "{smoval.codartic}"
        devuelve = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(Codigo, "T", 5, 6, devuelve) Then Exit Function
    End If
                    
    '-- Cadena para seleccion Desde y Hasta FAMILIA
    If txtCodigo(7).Text <> "" Or txtCodigo(8).Text <> "" Then
        Codigo = "{sartic.codfamia}"
        devuelve = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(Codigo, "N", 7, 8, devuelve) Then Exit Function
    End If
        
    '-- Cadena para seleccion Desde y Hasta ALMACEN
    If txtCodigo(11).Text <> "" Or txtCodigo(12).Text <> "" Then
        Codigo = "{smoval.codalmac}"
        devuelve = "pDHAlmacen=""Almacen: "
        If Not PonerDesdeHasta(Codigo, "N", 11, 12, devuelve) Then Exit Function
    End If
    
    
    '-- Cadena para seleccion Desde y Hasta CLIENTE/PROVEEDOR
    If txtCodigo(86).Text <> "" Or txtCodigo(87).Text <> "" Then
        Codigo = "{smoval.codigope}"
        devuelve = "pDHOperario=""Cliente/Proveedor/Trab.: "
        If Not PonerDesdeHasta(Codigo, "N", 86, 87, devuelve) Then Exit Function
    End If
    
        
'    cadSelect = QuitarCaracterACadena(cadFormula, "{")
'    cadSelect = QuitarCaracterACadena(cadSelect, "}")
        
    '=================================================
    '-- Cadena para seleccion Desde y Hasta FECHA
    If txtCodigo(9).Text <> "" Or txtCodigo(10).Text <> "" Then
        Codigo = "{smoval.fechamov}"
        devuelve = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(Codigo, "F", 9, 10, devuelve) Then Exit Function
    End If
        
    '-- seleccionar los articulos que tienen control de stock
    Codigo = "{sartic.ctrstock}=1"
    AnyadirAFormula cadFormula, Codigo
    AnyadirAFormula cadSelect, Codigo
        
        
    '-- Cadena de Seleccion TIPOS de MOVIMIENTOS
    Codigo = "{smoval.detamovi}"
    devuelve = ""
    'Si todos seleccionados no añadir la select
    todosMarcados = True
    I = 1
    While Not I > Me.ListView1.ListItems.Count And todosMarcados
        If Not Me.ListView1.ListItems(I).Checked Then todosMarcados = False
        I = I + 1
    Wend
    
    'si no estan todos seleccionados montar select de los seleccionados
    If Not todosMarcados Then
        cad = ""
        devuelve = ""
        For I = 1 To Me.ListView1.ListItems.Count
            If Me.ListView1.ListItems(I).Checked Then
                If cad = "" Then
                    cad = Me.ListView1.ListItems(I).Text
                Else
                    cad = cad & ", " & Me.ListView1.ListItems(I).Text
                End If
                If devuelve = "" Then
                    devuelve = Codigo & " = """ & Me.ListView1.ListItems(I).Text & """"
                Else
                    devuelve = devuelve & " or " & Codigo & " = """ & Me.ListView1.ListItems(I).Text & """"
                End If
            End If
        Next I

        If devuelve <> "" Then 'Hay algun movimiento marcado
            If cadFormula <> "" Then
                cadFormula = cadFormula & " AND " & "(" & devuelve & ")"
'                devuelve = QuitarCaracterACadena(devuelve, "{")
'                devuelve = QuitarCaracterACadena(devuelve, "}")
                cadSelect = cadSelect & " AND " & "(" & devuelve & ")"
                cadParam = cadParam
            Else
                cadFormula = "(" & devuelve & ")"
'                devuelve = QuitarCaracterACadena(devuelve, "{")
'                devuelve = QuitarCaracterACadena(devuelve, "}")
                cadSelect = "(" & devuelve & ")"
            End If
            cad = "pTiposMov=""Tipos Movimiento: " & cad
            cadParam = cadParam & cad & """|"
            numParam = numParam + 1
        Else 'Todos desmarcados
            cad = ""
            For I = 1 To ListView1.ListItems.Count
                If cad = "" Then
                    cad = """" & ListView1.ListItems(I).Text & """"
                Else
                    cad = cad & ", """ & ListView1.ListItems(I).Text & """"
                End If
            Next I
            devuelve = Codigo & " NOT IN [" & cad & "]"
            cad = Codigo & " NOT IN (" & cad & ")"
            cad = QuitarCaracterACadena(cad, "{")
            cad = QuitarCaracterACadena(cad, "}")
            If cadFormula = "" Then
                cadFormula = "(" & devuelve & ")"
                cadSelect = "(" & cad & ")"
            Else
                cadFormula = cadFormula & " AND " & "(" & devuelve & ")"
                cadSelect = cadSelect & " AND " & "(" & cad & ")"
            End If
        End If
    End If
    
    
    If cadFormula = "" Then
        MsgBox "Introduzca algún criterio de selección para el Informe.", vbInformation
        Exit Function
    End If
    PonerFormulaYParametrosInf9 = True
    
End Function


Private Function PonerFormulaYParametrosInf12() As Boolean
Dim cad As String, cadFrom As String
Dim devuelve As String
Dim ImprStock As String
Dim CodAux As String
Dim strValorado As String
Dim strSinStock As String
Dim bytPrecio As Byte

'    InicializarVbles
    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    cadFrom = ""
    devuelve = ""
    PonerFormulaYParametrosInf12 = False
    
    '===================================================
    '================= FORMULA =========================
    
    Select Case OpcionListado
        Case 12, 15, 16, 17, 19
            CodAux = "{salmac."
            cadFrom = "  salmac "
'        Case 15 'Listado articulos inactivos
'            CodAux = "{salmac."
'            cadFrom = "  (salmac LEFT OUTER JOIN smoval ON salmac.codartic=smoval.codartic AND salmac.codalmac=smoval.codalmac) "
'            cadFrom = "salmac"
        Case 13, 14
            CodAux = "{sinven."
            cadFrom = " sinven "
    End Select
    
    'Cadena para seleccion De ALMACEN
    '-----------------------------------
    Codigo = CodAux & "codalmac}"
    If Trim(txtCodigo(13).Text) <> "" Then _
    devuelve = Codigo & " = " & Val(txtCodigo(13).Text)
    If devuelve <> "" Then
        cadFormula = devuelve
        cad = "pAlmacen= ""Almacen: " & Format(txtCodigo(13).Text, "000") & " " & txtNombre(13).Text
        cadParam = cadParam & cad & """|"
        numParam = numParam + 1
    End If
            
    'Cadena para seleccion Desde y Hasta ARTICULOS
    '----------------------------------------------
    If txtCodigo(14).Text <> "" Or txtCodigo(15).Text <> "" Then
        Codigo = CodAux & "codartic}"
        devuelve = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(Codigo, "T", 14, 15, devuelve) Then Exit Function
    End If
    
    'Cadena para seleccion Desde y Hasta FAMILIA
    '--------------------------------------------
    If txtCodigo(16).Text <> "" Or txtCodigo(17).Text <> "" Then
        Select Case OpcionListado
            Case 12, 15, 16, 17, 19: Codigo = "{sartic.codfamia}"
            Case Else: Codigo = "{sinven.codfamia}"
        End Select
        cad = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(Codigo, "N", 16, 17, cad) Then Exit Function
    End If
    cadFrom = cadFrom & " INNER JOIN sartic ON " & CodAux & "codartic=sartic.codartic "
    'Enero 2008
    'David
    cadFormula = cadFormula & " AND {sartic.ctrstock} = 1"
    
    'Enero 2009
    'David
    'Solo saldran los articulos que esten en situacion normal o bloqueados.
    'Los caducados NO salen
    cadFormula = cadFormula & " AND {sartic.codstatu} < 2"
    
    
    
    'Cadena para seleccion Desde y Hasta PROVEEDOR
    '----------------------------------------------
    If txtCodigo(18).Text <> "" Or txtCodigo(19).Text <> "" Then
        Select Case OpcionListado
            Case 12, 15, 16, 17, 19: Codigo = "{sartic.codprove}"
            Case Else: Codigo = "{sinven.codprove}"
        End Select
        cad = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(Codigo, "N", 18, 19, cad) Then Exit Function
    End If
    

    
    'Select para MySQL
    cadSelect = QuitarCaracterACadena(cadFormula, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    cadSelect = QuitarCaracterACadena(cadSelect, "_1")
    cadFrom = QuitarCaracterACadena(cadFrom, "{")
    
    'Cadena para seleccion Desde y Hasta FECHA
    '----------------------------------------------
    If (OpcionListado = 16) Then
        If txtCodigo(20).Text <> "" Or txtCodigo(22).Text <> "" Then
            'codigo = "{salmac.codartic}"
            Codigo = CodAux & "fechainv}"
            devuelve = CadenaDesdeHasta(txtCodigo(20).Text, txtCodigo(22).Text, Codigo, "F")
    
            If devuelve = "Error" Then Exit Function
            
            If Not AnyadirAFormula(cadFormula, devuelve) Then
                Exit Function
            ElseIf devuelve <> "" Then
                cad = "pDHFecha=""Fecha: "
                If txtCodigo(20).Text <> "" Then _
                    cad = cad & "desde " & txtCodigo(20).Text
                If txtCodigo(22).Text <> "" Then _
                    cad = cad & "  hasta " & txtCodigo(22).Text
                cadParam = cadParam & cad & """|"
                numParam = numParam + 1
                'Para Comprobar si hay registros a Mostrar antes de abrir el Informe
                devuelve = "salmac.fechainv"
                devuelve = CadenaDesdeHastaBD(txtCodigo(20).Text, txtCodigo(22).Text, devuelve, "F")
                AnyadirAFormula cadSelect, devuelve
            Else
                'Si no hay fecha de inventario seleccionada coger solo
                'los articulos de los que se haya hecho inventario alguna vez
                devuelve = "not isnull({salmac.fechainv})"
                If Not AnyadirAFormula(cadFormula, devuelve) Then
                    Exit Function
                End If
                devuelve = "not isnull(salmac.fechainv)"
                AnyadirAFormula cadSelect, devuelve
            End If
        End If
    End If
    
    'Cadena de seleccion de FECHA de Inactividad
    '------------------------------------------------
    If OpcionListado = 15 Then '15: Listado de Articulos Inactivos
         If txtCodigo(20).Text <> "" Then _
            cad = "pFechaInve=""" & txtCodigo(20).Text & """"
        
        'Poner en el parametro pListaArt la lista de Articulos que no tiene
        'un registro de movimiento en la smoval con fecha posterior a la
        'fecha de inactividad
        strValorado = ListaArtActivos(cadSelect, txtCodigo(20).Text)
        cad = "pListaArtic=""" & strValorado & """|"
        cadParam = cadParam & cad
        numParam = numParam + 1
        
        'Añadir a la formula de seleccion que no sea uno de la lista
        devuelve = " not (" & CodAux & "codartic} in {@pListaArtic})"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
        
        strValorado = QuitarCaracterACadena(strValorado, "[")
        strValorado = QuitarCaracterACadena(strValorado, "]")
        devuelve = " not (salmac.codartic in (" & strValorado & "))"
        AnyadirAFormula cadSelect, devuelve
    End If
    
    'Cadena de seleccion de FECHA de Stocks a una Fecha
    '--------------------------------------------------
     If OpcionListado = 19 Then
        If txtCodigo(20).Text <> "" Then
            cad = txtCodigo(20).Text
            'Hora
            If txtCodigo(22).Text <> "" Then _
                cad = cad & "  " & txtCodigo(22).Text
                
            cadParam = cadParam & "pFechaStock=""" & cad & """|"
            numParam = numParam + 1
        End If
     End If
     
    'Cadena para Seleccion de Articulos con Stock<>0
    '------------------------------------------------
    If OpcionListado = 16 Or OpcionListado = 17 Or OpcionListado = 15 Then
        If Me.chkSinStock.Value = 0 Then
            If OpcionListado = 16 Then
                devuelve = "{salmac.stockinv}<>0"
            Else
                devuelve = CodAux & "canstock}<>0"
            End If
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
            
            devuelve = QuitarCaracterACadena(devuelve, "{")
            devuelve = QuitarCaracterACadena(devuelve, "}")
            devuelve = QuitarCaracterACadena(devuelve, "_1")
            AnyadirAFormula cadSelect, devuelve
        End If
    ElseIf OpcionListado = 19 Then
         If Me.chkSinStock.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        cadParam = cadParam & "pSinStock=" & ImprStock & "|"
        numParam = numParam + 1
    End If
    
       
    '============================================
    '============= PARAMETROS ===================
    If OpcionListado = 12 Or OpcionListado = 15 Then
        '12: Toma de Inventario
        '15: Listado Articulos Inactivos
        cadParam = cadParam & "pFechaInve=""" & txtCodigo(20).Text & """|"
        numParam = numParam + 1
    End If
    
    If OpcionListado = 12 Then
        'Parámetro Imprime Stock (Si/No)
        If Me.chkImprimeStock.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        cadParam = cadParam & "pImprimeStock=" & ImprStock & "|"
        numParam = numParam + 1
        
'        'seleccionar para inventariar los articulos que no tienen control stock
'        devuelve = " {sartic.ctrstock} = 1 "
'        AnyadirAFormula cadFormula, devuelve
'        AnyadirAFormula cadSelect, devuelve
        'Laura 03/01/07
        If Not (InStr(cadFrom, "sartic") > 0) Then
            cadFrom = cadFrom & " INNER JOIN sartic ON " & CodAux & "codartic=sartic.codartic "
        End If
    End If
    
    If OpcionListado = 12 Or OpcionListado = 13 Or OpcionListado = 15 Or OpcionListado = 16 Or OpcionListado = 17 Or OpcionListado = 19 Then
        'Parámetro Salta Pag. en Familia (Si/No)
        If Me.chkSaltaPag.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        cadParam = cadParam & "pSaltaFamilia=" & ImprStock & "|"
        numParam = numParam + 1
    End If
    
    If OpcionListado = 16 Or OpcionListado = 17 Then '16: Valoración de Stocks Inventariados
                                                     '17: Valoración Stocks
        'Parámetro Valorado
        If Me.chkValorado.Value Then
            strValorado = "True"
        Else
            strValorado = "False"
        End If
        cadParam = cadParam & "pValorado=" & strValorado & "|"
        numParam = numParam + 1
    End If
    
    If (OpcionListado = 15) Or (OpcionListado = 16) Or (OpcionListado = 17) Or (OpcionListado = 19) Then
        'Parametro Precio de Valoracion, elegir un Precio para realizar la valoracion: canstock * precio
        If Me.optPrecioMP.Value Then bytPrecio = 1
        If Me.optPrecioMA.Value Then bytPrecio = 2
        If Me.optPrecioUC.Value Then bytPrecio = 3
        If Me.optPrecioStd.Value Then bytPrecio = 4
        cadParam = cadParam & "pPrecio=" & bytPrecio & "|"
        numParam = numParam + 1
    End If
    '=====================================================================
    
       
    'comprobar si hay registros para mostrar en el Informe antes de Abrirlo
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Function
    
    If OpcionListado = 19 Then
        cadSelect = "Select count(*) FROM " & cadFrom & " WHERE " & cadSelect
        cadSelect = Replace(cadSelect, "count(*)", "*")
        DescargarDatosTMPStockFecha
        If Not CargarTMPStockFecha(cadSelect, txtCodigo(20).Text, txtCodigo(22).Text) Then Exit Function
    End If
    
    PonerFormulaYParametrosInf12 = True
End Function



Private Function PonerFormulaYParametrosInf28() As Boolean
'Informes de Descuentos y Tarifas
Dim cad As String
Dim cadCodigo As String

    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = 1
    
    PonerFormulaYParametrosInf28 = False
    
    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion Desde y Hasta TARIFA o D/H CLIENTE
    '--------------------------------------------
     If txtCodigo(23).Text <> "" Or txtCodigo(24).Text <> "" Then
        If OpcionListado = 30 Then 'Precios Especiales Cliente
            cadCodigo = Codigo & ".codclien}"
            cad = "pDHCliente=""Cliente: "
        Else
            cadCodigo = Codigo & ".codlista}"
            cad = "pDHTarifa=""Tarifa: "
        End If
        If Not PonerDesdeHasta(cadCodigo, "N", 23, 24, cad) Then Exit Function
    End If
            
            
    'Cadena para seleccion Desde y Hasta FAMILIA
    '--------------------------------------------
    If txtCodigo(25).Text <> "" Or txtCodigo(26).Text <> "" Then
        cadCodigo = "{sartic.codfamia}"
        cad = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(cadCodigo, "N", 25, 26, cad) Then Exit Function
    End If
    
    If OpcionListado <> 245 Then
        'Cadena para seleccion Desde y Hasta MARCA
        '--------------------------------------------
        If txtCodigo(27).Text <> "" Or txtCodigo(28).Text <> "" Then
            cadCodigo = "{sartic.codmarca}"
            cad = "pDHMarca=""Marca: "
            If Not PonerDesdeHasta(cadCodigo, "N", 27, 28, cad) Then Exit Function
        End If
    End If
            
    'Cadena para seleccion Desde y Hasta ARTICULOS
    '----------------------------------------------
    If txtCodigo(29).Text <> "" Or txtCodigo(30).Text <> "" Then
        cadCodigo = Codigo & ".codartic}"
        cad = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(cadCodigo, "T", 29, 30, cad) Then Exit Function
    End If
 
 
    '=====================================================================
    '====   PARAMETROS
    If (OpcionListado = 28 Or OpcionListado = 29) Then
        'Parámetro Salta Pag. en Familia (Si/No)
        If Me.chkSaltaPagTarif.Value = 1 Then
            cad = "True"
        Else
           cad = "False"
        End If
        cadParam = cadParam & "pSaltaFamilia=" & cad & "|"
        numParam = numParam + 1
    End If
       
    If OpcionListado = 245 Then
        'Parámetro mostrar solo tarifas con errores (Si/No)
        cad = Abs(Val(Me.chkMostrarErrores.Value))
        cadParam = cadParam & "Suprimr=" & cad & "|"
        numParam = numParam + 1
        'Decimales
        If cboDecimales.ListIndex < 0 Then
            MsgBox "Seleccione decimales", vbExclamation
            Exit Function
        End If
        cad = (cboDecimales.ItemData(Me.cboDecimales.ListIndex))
        cadParam = cadParam & "Decimales=" & cad & "|"
        numParam = numParam + 1
    End If
       
    PonerFormulaYParametrosInf28 = True
End Function


Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    
     If txtCodigo(indD).Text <> "" Then
        cad = cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        cad = cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then cad = cad & " - " & txtNombre(indH).Text
    End If
    
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function InsertarInventario() As Boolean
'Inserta en la Tabla:sinven los articulos seleccionados para realizar Inventario
'Inserta en la Tabla Hist.: shinve los datos que habia de inventario
'Además Actualiza la Tabla:salmac los campos:fechainv, horainve, statusin
Dim Sql As String, ADonde As String
Dim RS As ADODB.Recordset
Dim hora As Date

On Error GoTo EInventario:
   
'   If CrearTmpInventario(cadSelect) Then
   

        'Aqui empieza transaccion
        conn.BeginTrans
    
          
    
'        'Insertar en la tabla de Histórico: shinve
'        'Pasamos al Hist. los datos que habia antes de hacer nuevo inventario
'        ADonde = "Insertando datos en Histórico. Tabla: shinve"
'        SQL = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
'        SQL = SQL & " SELECT salmac.codartic, salmac.codalmac, salmac.fechainv,salmac.horainve,salmac.stockinv "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
'        'si no se ha inventariado antes no lo pasamos al historico
'        SQL = SQL & " AND not isnull(salmac.fechainv) "
'        Conn.Execute SQL
'
        
        'Insertar en la tabla de Histórico: shinve
        'Pasamos al Hist. los datos que habia antes de hacer nuevo inventario
        ADonde = "Insertando datos en Histórico. Tabla: shinve"
        Sql = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
        Sql = Sql & " SELECT codartic,codalmac,fechainv,horainve,stockinv "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
        Sql = Sql & " FROM tmpInven "
'        SQL = SQL & " WHERE " & cadFormula
        'si no se ha inventariado antes no lo pasamos al historico
        Sql = Sql & " WHERE not isnull(fechainv) "
        '--- Laura 03/01/2006
        Sql = Sql & " AND fechainv<>'0000-00-00' AND date(horainve)<>'0000-00-00' "
        '---
        conn.Execute Sql
        
        
        
        
        
        hora = Format(txtCodigo(20).Text & " " & Time, "yyyy-mm-dd hh:mm:ss")
        
'        'Insertamos en la Tabla sinven
'        ADonde = "Insertando datos en Inventario Real. Tabla: sinven"
'        SQL = "INSERT INTO sinven (codartic, codalmac, codfamia, codprove, fechainv, horainve, existenc) "
'        SQL = SQL & "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove," & DBSet(txtCodigo(20).Text, "F") & " as fechainv," & DBSet(hora, "FH") & " as horainve, 0 as existenc "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
'        'Insertamos los articulos que tiene control de stock
'        SQL = SQL & " AND sartic.ctrstock=1"
'        Conn.Execute SQL

        'Insertamos en la Tabla sinven
        ADonde = "Insertando datos en Inventario Real. Tabla: sinven"
        Sql = "INSERT INTO sinven (codartic, codalmac, codfamia, codprove, fechainv, horainve, existenc) "
        Sql = Sql & "SELECT codartic, codalmac, codfamia, codprove," & DBSet(txtCodigo(20).Text, "F") & " as fechainv," & DBSet(hora, "FH") & " as horainve, 0 as existenc "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
        Sql = Sql & " FROM tmpInven "
'        SQL = SQL & " WHERE " & cadFormula
        'Insertamos los articulos que tiene control de stock
'        SQL = SQL & " AND sartic.ctrstock=1"
        conn.Execute Sql


        
        
'        SQL = "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove "
'        SQL = SQL & "FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
        
        Sql = "SELECT codartic, codalmac, codfamia, codprove "
        Sql = Sql & " FROM tmpInven "
    
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RS.EOF
            
        
        
    '        'Insertamos en la Tabla sinven
    '        ADonde = "Insertando datos en Inventario Real. Tabla: sinven"
    '        SQL = "INSERT INTO sinven (codartic, codalmac, codfamia, codprove, fechainv, horainve, existenc) "
    '        SQL = SQL & " VALUES (" & DBSet(Rs.Fields(0).Value, "T") & ", " & Rs.Fields(1).Value & ", "
    '        SQL = SQL & Rs.Fields(2).Value & ", " & Rs.Fields(3).Value & ", '"
    '        'SQL = SQL & Format(txtCodigo(20).Text, "yyyy-mm-dd") & "', '" & hora & "', " & rs.Fields(2).Value & ")"
    '        SQL = SQL & Format(txtCodigo(20).Text, "yyyy-mm-dd") & "', '" & Format(hora, "yyyy-mm-dd hh:mm:ss") & "', 0)"
    '        Conn.Execute SQL
            
            
            
            
            
            
            
            'Actualizamos la tabla salmac ponemos statusin=1 para indicar que se
            'esta realizando inventario y bloquear los articulos para que no se puedan
            'realizar movimientos, traspasos, etc.
            'Actualizamos la Tabla: salmac los campos: fechainv, horainve
            ADonde = "Actualizando datos en Articulos x Almacen"
            Sql = "UPDATE salmac SET fechainv='" & Format(txtCodigo(20).Text, "yyyy-mm-dd") & "', "
            Sql = Sql & " horainve='" & Format(hora, "yyyy-mm-dd hh:mm:ss") & "', " & "statusin=1 , stockinv=0"
            Sql = Sql & " WHERE codartic=" & DBSet(RS.Fields(0).Value, "T") & " AND "
            Sql = Sql & "codalmac=" & RS.Fields(1).Value
            conn.Execute Sql
            RS.MoveNext
        Wend
    
        RS.Close
        Set RS = Nothing
'    Else
'        Exit Function
'    End If
    
EInventario:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
          Sql = "Insertando Datos de Inventario." & vbCrLf & "--------------------------------------" & vbCrLf
          Sql = Sql & ADonde
          MuestraError Err.Number, Sql, Err.Description
        conn.RollbackTrans
        InsertarInventario = False
    Else
        InsertarInventario = True
        conn.CommitTrans
    End If
End Function


Private Function CrearTmpInventario(cadFormula As String) As Boolean
Dim Sql As String
Dim b As Boolean

    On Error GoTo ECrearInv
    
    b = False
    Sql = "CREATE TEMPORARY TABLE tmpInven ( "
    Sql = Sql & "codartic varchar(16) NOT NULL default '0', "
    Sql = Sql & "codalmac smallint(3) unsigned NOT NULL default '0', "
    Sql = Sql & "codfamia smallint(4) unsigned NOT NULL default '0', "
    Sql = Sql & "codprove int(6) unsigned NOT NULL default '0', "
    Sql = Sql & "fechainv date NOT NULL default '0000-00-00', "
    Sql = Sql & "horainve datetime NOT NULL default '0000-00-00 00:00:00', "
    Sql = Sql & "stockinv decimal(12,2) NOT NULL default '0.00')"
    conn.Execute Sql
    b = True
    
    
    'Seleccionar todos los registros que vamos a inventariar, insertarlos en la TMP
    'y trabajar con estos valores
    Sql = "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove,salmac.fechainv,salmac.horainve,salmac.stockinv  "
    Sql = Sql & "FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
    Sql = Sql & " WHERE " & cadFormula
    Sql = Sql & " AND sartic.ctrstock=1"

    Sql = " INSERT INTO tmpInven " & Sql
    conn.Execute Sql
    
    
    
ECrearInv:
    If Err.Number <> 0 Then
        Sql = " DROP TABLE IF EXISTS tmpInven;"
        conn.Execute Sql
        b = False
        'Err.Clear
        MuestraError Err.Number, "Crear temporal inventario.", Err.Description
    End If
    CrearTmpInventario = b
End Function






Private Function ActualizarInventario() As Boolean
'-----------------------------------------------------------------
'* Modifica en la Tabla: salmac los campos: cansotck, fechainv, horainve,statusin de los articulos seleccionados
'y les asigna los valores de los campos: existenc, fechainv, horainve, false de la tabla: sinven
'* Elimina de la Tabla: sinven los registros seleccinados para actualizar
'* Inserta Movimientos de Articulos en la Tabla: smoval
'-------------------------------------------------------------------
Dim Sql As String, ADonde As String
Dim RS As ADODB.Recordset
Dim DevStock As String
Dim CanStock As Long, Diferencia As Long
Dim vTipoMov As CTiposMov
'Dim CodTipoMov As String * 3
Dim NumMovim As Long, numlinea As Long
Dim LetraSerie As String * 1
Dim CadValues As String
Dim bol As Boolean
    
    On Error Resume Next
    
    'Obtener Registros de la Tabla:sinven de los que se va a actualizar el Stock
    Sql = "SELECT sinven.* "
    
    'DAVID ENERO 2008
    'SQL = SQL & " FROM sinven "
    Sql = Sql & " FROM sinven  INNER JOIN sartic ON sinven.codartic=sartic.codartic"
    
    Sql = Sql & " WHERE " & cadFormula
    

    bol = True
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        bol = False
        ActualizarInventario = False
        MsgBox "No existen Registros en la Tabla: sinven para Actualizar Inventario.", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    
    'Obtener el contador para los movimientos del Almacen que se esta inventariando
    'A cada registro de la tabla sinven se le asignará un numero de linea.
    '----------------------------------------------------------------------------
    Set vTipoMov = New CTiposMov
'    CodTipoMov = "REG"
    If vTipoMov.Leer("DFI") Then  'Se han cargado correctamente los valores de la clase
        'Obtener el contador para el codigo de Movimiento
        LetraSerie = vTipoMov.LetraSerie
        NumMovim = vTipoMov.ConseguirContador("DFI")
        numlinea = 1
        bol = True
    Else
        bol = False
    End If
    
    If Not bol Then
        Set vTipoMov = Nothing
        Exit Function
    End If
    
   
    On Error GoTo EActualizarInven:
    'Aqui empieza la transaccion
    conn.BeginTrans
    
    While Not RS.EOF And bol 'Para cada registro de la tabla sinven
    
        'Introducir Movimiento de Entrada/Salida si hay diferencia entre el
        'Stock del Sistema y el Stock Real Inventariado.
        '------------------------------------------------------------------
        ADonde = "Introduciendo Movimiento de Entrada/Salida. Tabla: smoval."
        DevStock = DevuelveDesdeBDNew(conAri, "salmac", "canstock", "codartic", RS!codArtic, "T", , "codalmac", RS!codAlmac, "N")
        If DevStock <> "" Then
            CanStock = CLng(DevStock)
            Diferencia = RS!existenc - CanStock
            If Diferencia <> 0 Then 'Insertar Movimiento de Entrada/Salida en Almacen
                CadValues = DBSet(RS!codArtic, "T") & ", " & RS!codAlmac & ", '" & Format(RS!fechainv, "yyyy-mm-dd") & "', '"
                CadValues = CadValues & Format(RS!horainve, "yyyy-mm-dd hh:mm:ss") & "', "
                bol = InsertarMovimArticulos(CadValues, RS!codArtic, Diferencia, LetraSerie, NumMovim, numlinea)
                numlinea = numlinea + 1
            Else
                bol = True
            End If
        Else
            bol = False
        End If
        
        'Actualizamos la Tabla: salmac
        '           salmac.canstock := existencia Real
        '           salmac.statusin := false (desbloqueamos los articulos )
        '---------------------------------------
        If bol Then
            ADonde = "Actualizando Stock de Articulos en Almacen. Tabla: salmac."
            Sql = "UPDATE salmac SET canstock=" & DBSet(RS!existenc, "N") & ", statusin=0"
            Sql = Sql & " WHERE codartic=" & DBSet(RS!codArtic, "T") & " AND codalmac=" & RS!codAlmac
            conn.Execute Sql
        End If

        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    If bol Then
'        'Pasamos la tabla de inventario real sinven al historico: shinve
'        'antes de eliminarla
'        ADonde = "Pasando registros de Inventario al Histórico: shinve."
'        SQL = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
'        SQL = SQL & "SELECT codartic,codalmac,fechainv,horainve,existenc "
'        SQL = SQL & " FROM sinven WHERE " & cadFormula
'        Conn.Execute SQL
    
        'Eliminamos los registros seleccionados de la Tabla: sinven
        '----------------------------------------------------------
        ADonde = "Eliminando registros de Inventario. Tabla: sinven."
       ' SQL = "DELETE FROM sinven "
  
        'DAVID ENERO 2008
        Sql = "DELETE sinven.* FROM sinven  INNER JOIN sartic ON"
        Sql = Sql & " sinven.codartic=sartic.codartic WHERE " & cadFormula
        conn.Execute Sql
        
        
        'Incrementamos el contador para el Tipo De Movimiento
        '-----------------------------------------------------
        ADonde = "Actualizando el contador ."
        'bol = vTipoMov.IncrementarContador(
        vTipoMov.IncrementarContador ("DFI")
    End If
    Set vTipoMov = Nothing
        
EActualizarInven:
    If Err.Number <> 0 Or Not bol Then
         'Hay error , almacenamos y salimos
          Sql = "Actualizar Inventario." & vbCrLf & "----------------------------" & vbCrLf
          Sql = Sql & ADonde
          MuestraError Err.Number, Sql, Err.Description
          conn.RollbackTrans
          ActualizarInventario = False
          Set vTipoMov = Nothing
    Else
        ActualizarInventario = True
        conn.CommitTrans
    End If
End Function


Private Function InsertarMovimArticulos(CadValues As String, codArtic As String, Cantidad As Long, LetraSerie As String, NumMovim As Long, numlinea As Long) As Boolean
Dim vImporte As Single, vPrecioVenta As String
Dim tipoMov As Byte
Dim Sql As String
On Error Resume Next
         
        'Obtener el precio de venta del articulo
         vPrecioVenta = DevuelveDesdeBDNew(conAri, "sartic", "preciomp", "codartic", codArtic, "T")
        If vPrecioVenta <> "" Then
            vImporte = Cantidad * CSng(vPrecioVenta)
        Else
            vImporte = 0
        End If
        
        'Tipo de Movimiento de Almacen
        If Cantidad > 0 Then 'Insertar Movimiento de Entrada en Almacen
            tipoMov = 1
        ElseIf Cantidad < 0 Then 'Insertar Movimiento de Salida en Almacen
            tipoMov = 0
        End If
        
        Sql = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
        Sql = Sql & " VALUES (" & CadValues & tipoMov & ", '" & "DFI" & "', " & DBSet(Cantidad, "N") & ", " & DBSet(vImporte, "N") & ", " & Val(txtCodigo(21).Text) & ", '"
        Sql = Sql & LetraSerie & "', " & NumMovim & ", " & numlinea & ")"
        conn.Execute Sql
        
        If Err.Number <> 0 Then
             'Hay error , almacenamos y salimos
            InsertarMovimArticulos = False
        Else
            InsertarMovimArticulos = True
        End If
    
End Function


Private Function ValidarCamposInventario() As Boolean
'Comprobar que los campos requeridos tienen valor antes de abrir el listado
Dim b As Boolean

        b = True
        '- campo almacen debe tener valor
        If Trim(txtCodigo(13).Text) = "" Then
             MsgBox "El campo Almacen debe tener valor.", vbInformation
             PonerFoco txtCodigo(13)
             b = False
        End If
    
        '- fecha de inventario debe tener valor
        If b Then
            If (OpcionListado = 12 Or OpcionListado = 15 Or OpcionListado = 19) And Trim(txtCodigo(20).Text) = "" Then
                 MsgBox "El campo Fecha debe tener valor.", vbInformation
                 PonerFoco txtCodigo(20)
                 b = False
            End If
        End If
        
        'informe 19: stocks a una fecha
        'la fecha tiene que ser < a fecha hoy
        If OpcionListado = 19 And txtCodigo(20).Text <> "" Then
            If Not CDate(txtCodigo(20).Text) < CDate(Format(Now, "dd/mm/yyyy")) Then
                MsgBox "La fecha stock tiene que ser anterior a la fecha de hoy.", vbInformation
                PonerFoco txtCodigo(20)
                b = False
            End If
        End If
        
        ValidarCamposInventario = b
End Function



Private Function ListaArtActivos(cadWHERE As String, FechaIn As String) As String
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Lista As String
'Devuelve una cadena con la concatenacion de todos los articulos que
'no debe seleccionar ya que si tienen movimientos con fecha posterior
'a FechaIn.
'ej:    "[""00000004"", ""00000033""]"

    Lista = "["
    
    Sql = "SELECT distinct smoval.codartic from smoval "
    If InStr(cadWHERE, "sartic") > 0 Then Sql = Sql & " INNER JOIN sartic ON smoval.codartic=sartic.codartic "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "salmac", "smoval")
    If cadWHERE <> "" Then Sql = Sql & " AND "
    Sql = Sql & " fechamov>='" & Format(FechaIn, FormatoFecha) & "' "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
'        lista = lista & """" & RS.Fields(0).Value & """"
        Lista = Lista & DBSet(RS.Fields(0).Value, "T")
        RS.MoveNext
        If Not RS.EOF Then Lista = Lista & ", "
    Wend
    Lista = Lista & "]"
    ListaArtActivos = Lista
    RS.Close
    Set RS = Nothing
End Function



Private Sub ActualizarImprimir()
Dim I As Long
Dim Desde As Long, Hasta As Long
Dim Sql As String

    Select Case OpcionListado
    Case 7  'TRASPASO ALMACEN
        If frmVisReport.EstaImpreso = True Then
        'Desde=-1 si estamos en Historico ya que aqui no se introducen valores Desde/Hasta
            If Trim(txtCodigo(3).Text) <> "" Then Desde = CLng(txtCodigo(3).Text)
            If Trim(txtCodigo(4).Text) <> "" Then Hasta = CLng(txtCodigo(4).Text)
            For I = Desde To Hasta
                Sql = "UPDATE scatra SET situacio=1" 'Impreso
                Sql = Sql & " WHERE codtrasp=" & I
                conn.Execute Sql
            Next I
        End If
        
    Case 8  'MOVIMIENTO ALMACEN
        If frmVisReport.EstaImpreso = True Then
           'Desde=-1 si estamos en Historico ya que aqui no se introducen valores Desde/Hasta
           If Trim(txtCodigo(3).Text) <> "" Then Desde = CLng(txtCodigo(3).Text)
           If Trim(txtCodigo(4).Text) <> "" Then Hasta = CLng(txtCodigo(4).Text)
           For I = Desde To Hasta
                Sql = "UPDATE scamov SET situacio=1"
                Sql = Sql & " WHERE codmovim=" & I
                conn.Execute Sql
           Next I
        End If
    End Select
End Sub


Private Sub CargarComboTipoList()
'### Combo Tipo Facturación
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'1-Equipos, 2-Pagos, 3-Importes Contrato

    Me.cboTipoList.Clear
    cboTipoList.AddItem "Equipos"
    cboTipoList.ItemData(cboTipoList.NewIndex) = 1

    cboTipoList.AddItem "Pagos"
    cboTipoList.ItemData(cboTipoList.NewIndex) = 2

    cboTipoList.AddItem "Importes Contrato"
    cboTipoList.ItemData(cboTipoList.NewIndex) = 3

End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
    pPdfRpt = ""
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If Tipo <> "F" Then
        'Fecha para Crystal Report
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, cad) Then Exit Function
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


Private Sub LlamarImprimir(PonerNombrePDF As Boolean)
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .NombrePDF = ""
        If PonerNombrePDF Then .NombrePDF = pPdfRpt
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub


Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim NomCampo As String

    campo = "pGroup" & numGrupo & "="
    NomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "Familia"
            cadParam = cadParam & campo & "{sartic.codfamia}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""FAMILIA: "" & " & " totext({sartic.codfamia},""0000"") & " & """  """ & " & {sfamia.nomfamia}" & "|"
            Else
                cadParam = cadParam & NomCampo & " totext({sartic.codfamia},""0000"") & " & """ """ & " & {sfamia.nomfamia}" & "|"
            End If
            numParam = numParam + 1
        Case "Marca"
            cadParam = cadParam & campo & "{sartic.codmarca}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""MARCA: "" & " & " totext({sartic.codmarca},""0000"") & " & """  """ & " & {smarca.nommarca}" & "|"
            Else
                cadParam = cadParam & NomCampo & " totext({sartic.codmarca},""0000"") & " & """ """ & " & {smarca.nommarca}" & "|"
            End If
            numParam = numParam + 1
        Case "Proveedor"
            cadParam = cadParam & campo & "{sartic.codprove}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""PROVEEDOR: "" & " & " totext({sartic.codprove},""000000"") & " & """  """ & " & {sprove.nomprove}" & "|"
            Else
                cadParam = cadParam & NomCampo & " totext({sartic.codprove},""000000"") & " & """ """ & " & {sprove.nomprove}" & "|"
            End If
            numParam = numParam + 1
            PonerGrupo = numGrupo
        Case "Tipo Articulo"
            cadParam = cadParam & campo & "{sartic.codtipar}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""TIPO ARTICULO: "" & " & " {sartic.codtipar} & " & """  """ & " & {stipar.nomtipar}" & "|"
            Else
                cadParam = cadParam & NomCampo & " {sartic.codtipar} & " & """ """ & " & {stipar.nomtipar}" & "|"
            End If
            numParam = numParam + 1
    End Select

'Case "Familia"
'            cadParam = cadParam & "pGroup1=" & "{sartic.codfamia}" & "|"
'            cadParam = cadParam & "pGroup1Name= ""FAMILIA: "" & " & " totext({sartic.codfamia},""0000"") & " & """  """ & " & {sfamia.nomfamia}" & "|"
'            numParam = numParam + 1
'            Select Case ListView2.ListItems(2).Text
'                Case "Marca"
'                    cadParam = cadParam & "pGroup2=" & "{sartic.codmarca}" & "|"
'                    cadParam = cadParam & "pGroup2Name= ""MARCA: "" & " & " totext({sartic.codmarca},""0000"") & " & """  """ & " & {smarca.nommarca}" & "|"
'                    numParam = numParam + 1
'                    If ListView2.ListItems(3).Text = "Proveedor" Then
'                        Opcion = 1
'                    Else
'                        Opcion = 2
'                    End If
'                Case "Proveedor"
'                Case "Tipo Articulo"
'            End Select
End Function



Private Sub AbrirFrmActividades(Optional indice As Integer)
    Set frmMtoActiv = New frmFacActividades
    frmMtoActiv.DatosADevolverBusqueda = "0|1|"
    frmMtoActiv.DeConsulta = True
    frmMtoActiv.Show vbModal
    Set frmMtoActiv = Nothing
End Sub



Private Sub AbrirFrmMarcas()
    Set frmMtoMarcas = New frmAlmMarcas
    frmMtoMarcas.DatosADevolverBusqueda = "0|1"
    frmMtoMarcas.DeConsulta = True
    frmMtoMarcas.Show vbModal
    Set frmMtoMarcas = Nothing
End Sub


Private Sub AbrirFrmAlmPropios()
    Set frmMtoAlPropios = New frmAlmAlPropios
    frmMtoAlPropios.DatosADevolverBusqueda = "0|1"
    frmMtoAlPropios.DeConsulta = True
    frmMtoAlPropios.Show vbModal
    Set frmMtoAlPropios = Nothing
End Sub




Private Sub AbrirFrmTarifas()
'tarifas venta
    Set frmMtoTarifas = New frmFacTarifas
    frmMtoTarifas.DatosADevolverBusqueda = "0|1"
    frmMtoTarifas.Show vbModal
    Set frmMtoTarifas = Nothing
End Sub


Private Sub AbrirFrmTipoArt()
'Tipos de Articulos
    Set frmMtoTArticulo = New frmAlmTipoArticulo
    frmMtoTArticulo.DatosADevolverBusqueda = "0|1"
    frmMtoTArticulo.DeConsulta = True
    frmMtoTArticulo.Show vbModal
    Set frmMtoTArticulo = Nothing
End Sub

Private Sub AbrirFrmClientes()
'Clientes
    Set frmMtoClientes = New frmFacClientes
    frmMtoClientes.DatosADevolverBusqueda = "0|1"
    frmMtoClientes.Show vbModal
    Set frmMtoClientes = Nothing
End Sub

Private Sub AbrirFrmSocios()
'Socios
    Set frmMtoSocios = New frmGesSocios
    frmMtoSocios.DatosADevolverBusqueda = "0|1"
    frmMtoSocios.Show vbModal
    Set frmMtoSocios = Nothing
End Sub



Private Function ComprobarFechasConta(Ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim cad As String
Dim RS As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtCodigo(Ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set RS = New ADODB.Recordset
        RS.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not RS.EOF Then
            FechaIni = DBLet(RS!FechaIni, "F")
            '## LAURA 19/06/2008
'            FechaFin = DBLet(RS!FechaFin, "F") + 365
'            FechaFin = DateAdd("d", 365, DBLet(RS!FechaFin, "F"))
            FechaFin = DateAdd("yyyy", 1, DBLet(RS!FechaFin, "F"))
            '##
            
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(Ind).Text, FechaFin) Then
                 cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 cad = cad & "    Desde: " & FechaIni & vbCrLf
                 cad = cad & "    Hasta: " & FechaFin
                 MsgBox cad, vbExclamation
                 txtCodigo(Ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        RS.Close
        Set RS = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function



Private Function ContabilizarFacturas(cadTabla As String, cadWHERE As String) As Boolean
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste2 As Byte

        '0.- Si devuelve la funcion el 0 habra CC sin confgurar en trabaja
        '1.- Todos los CC son el mismo
        '2.- Mas de un CC. Hay que agrupar

    ContabilizarFacturas = False

    If cadTabla = "scafac" Then
        Sql = "VENCON" 'contabilizar facturas de venta
    ElseIf cadTabla = "scafpc" Then
        Sql = "COMCON" 'contabilizar facturas de compra
    End If

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2
     If txtCodigo(31).Text = "" Then
        txtCodigo(31).Text = vEmpresa.FechaIni  'fechaini del ejercicio de la conta
     End If

     If txtCodigo(32).Text = "" Then
        txtCodigo(32).Text = vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
     End If


     
     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los parámetros
     If Not ComprobarFechasConta(32) Then Exit Function
     
     

    'La comprobacion solo lo hago para facturas nuestras, ya que mas adelante
    'el programa hara cdate(text1(31) cuando contabilice las facturas y dara error de tipos
    If cadTabla = "scafac" Then
        If Me.txtCodigo(31).Text = "" Then
            MsgBox "Fecha inicio incorrecta", vbExclamation
            Exit Function
        End If
    End If
    
    
    
    'comprobar si existen en Aritaxi facturas anteriores al periodo solicitado
    'sin contabilizar.
    If Me.txtCodigo(31).Text <> "" Then 'anteriores a fechadesde
        Sql = "SELECT COUNT(*) FROM " & cadTabla
        If cadTabla = "scafac" Then
            Sql = Sql & " WHERE fecfactu <"
        ElseIf cadTabla = "scafpc" Then
            Sql = Sql & " WHERE fecrecep <"
        End If
        Sql = Sql & DBSet(txtCodigo(31), "F") & " AND intconta=0 "
        
        
        'Si contabiliza tickets agrupados
        'SOLO PARA CLIENTES, obviamente
        If cadTabla = "scafac" Then
            If OptProve.Tag = "" Then
                If vParamAplic.ContabilizarTicketAgrupados Then Sql = Sql & " AND codtipom <>'FTI' "
            Else
                Sql = Sql & " AND scafac.codtipom  = 'FTG' "
            End If
            
            '## LAURA 20/06/2008
            If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) <> "" Then
                Sql = Sql & " AND scafac.codtipom = " & DBSet(Mid(Me.cboTipMov.List(Me.cboTipMov.ListIndex), 1, 3), "T")
            End If
        End If
        
        If RegistrosAListar(Sql) > 0 Then
            If MsgBox("Hay Facturas anteriores sin contabilizar. " & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                Exit Function
            End If
        End If
    End If
    
    
'    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    If Not BloqueaRegistro(cadTabla, cadWhere) Then
'        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
'    Me.lblProgess(0).Caption = "Comprobaciones: "
'    CargarProgres Me.ProgressBar1, 100
        
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    b = CrearTMPFacturas(cadTabla, cadWHERE)
    If Not b Then Exit Function
            
            
    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    TerminaBloquear
    Sql = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla
    If cadTabla = "scafac" Then
        Sql = Sql & ".codtipom=tmpFactu.codtipom AND "
    Else
        Sql = Sql & ".codprove=tmpFactu.codprove AND "
    End If
    Sql = Sql & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    If Not BloqueaRegistro(Sql, cadWHERE) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
            
            
    '---- Preparamos la pantalla de Contabilizar
    'Visualizar la barra de Progreso
    
    
    
    If Me.FrameTipMov.visible Then
        Me.FrameRepxDia.Height = 6100
        Me.FrameProgress.Top = 4400
    Else
        Me.FrameRepxDia.Height = 5100
        Me.FrameProgress.Top = 3350
    End If
    Me.Height = Me.FrameRepxDia.Height
    Me.FrameProgress.visible = True
    Me.Refresh
            
    Me.lblProgess(0).Caption = "Comprobaciones: "
    CargarProgres Me.ProgressBar1, 100
        
        
    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Aritaxi
    '--------------------------------------------------------------------------
    IncrementarProgres Me.ProgressBar1, 10
    If cadTabla = "scafac" Then
        Me.lblProgess(1).Caption = "Comprobando letras de serie ..."
        b = ComprobarLetraSerie(cadTabla)
    End If
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not b Then Exit Function
    
    
    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    If cadTabla = "scafac" Then
        Me.lblProgess(1).Caption = "Comprobando Nº Facturas en contabilidad ..."
        Sql = "anofaccl>=" & Year(txtCodigo(31).Text) & " AND anofaccl<= " & Year(txtCodigo(32).Text)
        b = ComprobarNumFacturas_new(cadTabla, Sql)
    End If
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Function
    
    
    'comprobar que todas las CUENTAS de los distintos clientes que vamos a
    'contabilizar existen en la Conta: sclien.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    b = ComprobarCtaContable_new(cadTabla, 1, True)
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Function
    
    
    'comprobar que todas las CUENTAS de venta de la familia de los articulos que vamos a
    'contabilizar existen en la Conta: sfamia.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    If cadTabla = "scafac" Then
        Me.lblProgess(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    Else
        Me.lblProgess(1).Caption = "Comprobando Cuentas Ctbles Compras en contabilidad ..."
    End If
    b = ComprobarCtaContable_new(cadTabla, 2)
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Function
    
    
    
    If Me.OptProve.Tag <> "" Then
        'TIKETS. Voy a comprobar las cuentas de la familia
        Me.lblProgess(1).Caption = "Comprobando Cuentas Ctbles tickets ..."
        Me.lblProgess(1).Refresh
        
        
    End If
    
    
    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: scafac.codigiv1,codigiv2,codigiv3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    b = ComprobarTiposIVA(cadTabla)
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not b Then Exit Function
    
    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de sfamia.ctaventa empiezan por el digito
    'de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    
    If vEmpresa.TieneAnalitica Then
       Me.lblProgess(1).Caption = "Comprobando Contabilidad Analítica ..."
       b = ComprobarCtaContable_new(cadTabla, 3)
       If Not b Then Exit Function
       
       '(si tiene analítica requiere un centro de coste para insertar en conta.linfact)
       b = cadTabla = "scafac"
       CCoste2 = ComprobarCCoste(cadWHERE, b)
       If CCoste2 = 0 Then Exit Function 'Error comprobando CCs
       
    Else
        'No tiene analitica, NO agrupamos por codtraba
        CCoste2 = 0
    End If
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    
    If Me.OptProve.Tag <> "" Then
        Me.lblProgess(1).Caption = "Comprobando Ctas facmilias TICKETS ..."   'FTG
        b = ComprobarCtaContable_new(cadTabla, 4)
        If Not b Then Exit Function
    End If
    
    
    'Comprobamos, si es factura proveedore, que si el tipoprove = 3 (REA)
    'entonces tiene que existir el paremetro aplicacion codret
    If cadTabla = "scafpc" Then
        If vParamAplic.CtaReten = "" Then
            Sql = "SELECT COUNT(*) FROM scafpc,sprove WHERE scafpc.codprove = sprove.codprove and tipprove=3"
            If cadWHERE <> "" Then Sql = Sql & " AND " & cadWHERE
            If RegistrosAListar(Sql) > 0 Then
                MsgBox "Existen facturas proveedor con cta. retencion y no esta configurada", vbExclamation
                Exit Function
            End If
        
        
            'Neuvo 29Mayo 2008
            ' Cualquier factura puede llevar retencion. Necesito que la cuenta de retencion este configurada
            Sql = "SELECT COUNT(*) FROM scafpc  WHERE  tiporet=0 and impret<>0"
            If cadWHERE <> "" Then Sql = Sql & " AND " & cadWHERE
            If RegistrosAListar(Sql) > 0 Then
                MsgBox "Existen facturas proveedor con retencion y no esta configurada", vbExclamation
                Exit Function
            End If
         End If
        
    End If
    
    Me.lblProgess(1).Caption = "Fechas contabilizacion"
    Me.lblProgess(1).Refresh
    b = NuevasComprobacionesContabilizacion(cadTabla = "scafpc", cadWHERE)
    If Not b Then Exit Function
    
    
    
    
    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgess(0).Caption = "Contabilizar Facturas: "
    CargarProgres Me.ProgressBar1, 10
    Me.lblProgess(1).Caption = "Insertando Facturas en Contabilidad..."
       
    
    
    
    
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar facturas: " & vbCrLf & cadTabla & vbCrLf & cadWHERE
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    
    '---- Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTabla)
    
    '---- Pasar las Facturas a la Contabilidad
    b = PasarFacturasAContab(cadTabla, CCoste2)
    
    
    
    '---- Mostrar ListView de posibles errores (si hay)
    If Not b Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensajes.OpcionMensaje = 10
            frmMensajes.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        'Para la facturacion de TICKTS agrupada NO mostramos el mensaje de OK
        If Me.OptProve.Tag = "" Then
            If cadTabla = "scafac" Then MsgBox "El proceso ha finalizado correctamente.", vbInformation
        End If
    End If
    
    'Este bien o mal, si son proveedores abriremos el listado
    'Imprimimiremos un listado de contabilizacion de facturas
    '------------------------------------------------------
    If cadTabla <> "scafac" Then
        If NumRegistros("Select count(*) from tmpinformes where codusu = " & vUsu.Codigo) > 0 Then
            InicializarVbles
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
            
            cadParam = "|pDHFecha=""" & vUsu.Nombre & "   Hora: " & Format(Now, "hh:mm") & """|"
            numParam = numParam + 1
            cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
            cadNomRPT = "rContabPRO.rpt"
            conSubRPT = False
            cadTitulo = "Listado contabilizacion FRAPRO"
            
            LlamarImprimir True
        End If
    End If
    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact
    ContabilizarFacturas = True
End Function

'Ccoste
'   0: No tendra analitica
'   1: Solo hay un CC que tratar. NO agruparemos por trabajador
'   2: Mas de un CC. Agruparemos por trabajador
Private Function PasarFacturasAContab(cadTabla As String, miCC As Byte) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim b As Boolean
Dim I As Integer
Dim NumFactu As Integer
Dim Codigo1 As String
Dim ContabilizacionAgrupadaTickets As Boolean
'ENERO 2009
Dim cContaFra As cContabilizarFacturas


    On Error GoTo EPasarFac

    PasarFacturasAContab = False
    
    
    'Si escontailizacion de facturas de tickets agrupados
    ContabilizacionAgrupadaTickets = False
    If Me.OptProve.Tag <> "" Then ContabilizacionAgrupadaTickets = True
    
    '---- Obtener el total de Facturas a Insertar en la contabilidad
    Sql = "SELECT count(*) "
    Sql = Sql & " FROM " & cadTabla & " INNER JOIN tmpFactu "
    If cadTabla = "scafac" Then
        Codigo1 = "codtipom"
    Else
        Codigo1 = "codprove"
    End If
    Sql = Sql & " ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1
    Sql = Sql & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        NumFactu = RS.Fields(0)
    Else
        NumFactu = 0
    End If
    RS.Close
    Set RS = Nothing


    'Enero 2009
    '------------------------------------------------------------
    Set cContaFra = New cContabilizarFacturas
    
    If Not cContaFra.EstablecerValoresInciales(ConnConta) Then
        'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
        ' obviamente, no va a contabilizar las FRAS
        Sql = "Si continua, las facturas se insertaran en el registro, pero no serán contabilizadas" & vbCrLf
        Sql = Sql & "en este momento. Deberán ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
        Sql = Sql & Space(50) & "¿Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    
    
    
    

    'Modificacion 20 Abril 2008
    '-----------------------------------------------------------
    ' Mosrtaremos para cada factura de PROVEEDOR
    ' que numregis le ha asignado
    Sql = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Sql

    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If NumFactu > 0 Then
    
        Set RS = New ADODB.Recordset
    
        CargarProgres Me.ProgressBar1, NumFactu
        
        
        'PreComproabacion de los asientos
        If cContaFra.RealizarContabilizacion Then
            Sql = "Select min(fecfactu) from tmpfactu"
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                If Not cContaFra.PreComprobacionNumeroAsiento(RS.Fields(0), NumFactu) Then
                    
                    'Para que la ventana siguiente muestr bien el error
                    Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) VALUES ("
                    Sql = Sql & "'',0,'" & Format(RS.Fields(0), FormatoFecha) & "','Error contadores')"
                    
                    conn.Execute Sql
                    RS.Close
                    Err.Raise 6, , "Comprobacion numeros asiento"
                End If
            End If
            RS.Close
        End If
        
        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        Sql = "SELECT * "
        Sql = Sql & " FROM tmpFactu "
            

        RS.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
        I = 1

        b = True
   
   
   
   
   
   
   
        
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not RS.EOF
        
            'Segun sea cli o pro
            If cadTabla = "scafac" Then
                Sql = cadTabla & "." & Codigo1 & "=" & DBSet(RS.Fields(0), "T") & " AND scafac.numfactu=" & RS!NumFactu
                Sql = Sql & " and scafac.fecfactu=" & DBSet(RS!FecFactu, "F")
                If PasarFacturaSOC(Sql, miCC, ContabilizacionAgrupadaTickets, cContaFra) = False And b Then b = False
            Else
                Sql = cadTabla & "." & Codigo1 & "=" & DBSet(RS.Fields(0), "N") & " and scafpc.numfactu=" & DBSet(RS!NumFactu, "T")
                Sql = Sql & " and scafpc.fecfactu=" & DBSet(RS!FecFactu, "F")
                If PasarFacturaProv(Sql, miCC, Orden2, cContaFra) = False And b Then b = False
            End If
            
            '---- Laura 26/10/2006
            'Al pasar cada factura al hacer el commit desbloqueamos los registros
            'que teniamos bloqueados y los volvemos a bloquear
            'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
            Sql = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
            If Not BloqueaRegistro(Sql, cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu") Then
'                MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'                Screen.MousePointer = vbDefault
'                Exit Sub
            End If
            '----
            
            IncrementarProgres Me.ProgressBar1, 1
            Me.lblProgess(1).Caption = "Insertando Facturas en Contabilidad...   (" & I & " de " & NumFactu & ")"
            Me.Refresh
            I = I + 1
            RS.MoveNext   'Siguiente factura
        Wend
        
        'Veremos si ha dado error la contabilizacion de factiras
        If cContaFra.TieneErrores Then cContaFra.MuestraErroresContabilizacion
        
        
        RS.Close
        Set RS = Nothing
    End If
    
EPasarFac:
    If Err.Number <> 0 Then b = False
    Set cContaFra = Nothing
    If b Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function



Private Sub ListadosAlmacen(H As Integer, W As Integer)
    'LISTADOS DE ALMACENES
    '---------------------
    Select Case OpcionListado
        Case 1   'Listados de Marcas
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado de Marcas"
            indFrame = 1
            Codigo = "{smarca.codmarca}"
            Orden1 = "{smarca.codmarca}"
            Orden2 = "{smarca.nommarca}"
            cadTitulo = "Listado Marcas"
            cadNomRPT = "rAlmMarcas.rpt"
            conSubRPT = False
            
        Case 2   'Listado de Almacenes Propios
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado de Almacenes"
            indFrame = 1
            Codigo = "{salmpr.codalmac}"
            Orden1 = "{salmpr.codalmac}"
            Orden2 = "{salmpr.nomalmac}"
            cadTitulo = "Listado Almacenes Propios"
            cadNomRPT = "rAlmAPropios.rpt"
            conSubRPT = False
            
        Case 3   'Listado de Tipos de Unidad
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tipos de Unidad"
            indFrame = 1
            Codigo = "{sunida.codunida}"
            Orden1 = "{sunida.codunida}"
            Orden2 = "{sunida.nomunida}"
            cadTitulo = "Listado Tipos de Unidad"
            cadNomRPT = "rAlmTUnidad.rpt"
            conSubRPT = False
            
        Case 4   'Listado de Tipos de Artículos
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tipos de Artículos"
            indFrame = 1
            Codigo = "{stipar.codtipar}"
            Orden1 = "{stipar.codtipar}"
            Orden2 = "{stipar.nomtipar}"
            txtCodigo(1).Tag = CadTag
            txtCodigo(2).Tag = CadTag
            cadTitulo = "Listado Tipos de Artículos"
            cadNomRPT = "rAlmTArticulo.rpt"
            conSubRPT = False
            
        Case 6    'Listado de Artículo
            ponerFrameArticulosVisible True, H, W
            CargarListViewOrden
            Codigo = "{sartic"
            indFrame = 11
            cadTitulo = "Listado de Artículos"
            
            
        Case 110   'Listados Ubicaciones Almacen
            
        Case 18, 247 'Informe Stocks Maximos y Minimos   'OPCION: 247 es este tb
            ponerFrameArticulosVisible True, H, W
            Codigo = "{salmac"
            indFrame = 11
            cmbProduccion.ListIndex = 0
            cmbProduccion.visible = vParamAplic.Produccion
            Label4(90).visible = vParamAplic.Produccion
            
        Case 7, 8 '7: Informe de Traspasos de Almacen
                  '8: Informe de Movimientos de Almacen
            If OpcionListado = 7 Then
                Me.lblTitulo(2).Caption = "Informe Traspaso de Almacen"
                Me.Label2(1).Caption = "Nº Traspaso"
                Codigo = "{scatra.codtrasp}"
            Else
                Me.lblTitulo(2).Caption = "Informe Movimientos de Almacen"
                Me.Label2(1).Caption = "Nº Movimiento"
                Codigo = "{scamov.codmovim}"
            End If
            H = 3495
            W = 5835
            PonerFrameVisible Me.FrameInfAlmacen, True, H, W
            indFrame = 2
            If NumCod <> "" Then
                txtCodigo(3).Text = NumCod
                txtCodigo(4).Text = NumCod
            End If
            
        Case 9 'Informe Movimiento Artículos
            W = 10700
            H = 5775
            PonerFrameVisible Me.FrameMovArtic, True, H, W
            indFrame = 3
            Codigo = "{smoval.codartic}"
            cadTitulo = "Informe Movimientos Articulos"
            conSubRPT = True
            CargarListView
            
        ' ====  [16/09/2009] LAURA : Listado Articulos con componentes
        Case 11
            W = Me.FrameInvArtComp.Width
            H = Me.FrameInvArtComp.Height
            PonerFrameVisible Me.FrameInvArtComp, True, H, W
            Codigo = "{sartic.codartic}"
            cadTitulo = "Listado Artículos con Componentes"
        ' ====
            
        Case 12 '12: Listado Toma de Inventario Articulos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.chkImprimeStock.visible = True
            Me.lbltituloInven.Caption = "Listado Toma de Inventario Articulos"
            cadTitulo = "Toma Inventario Articulos"
            'codigo = "{salmac.codalmac}"
            
        Case 13 '13: Listado Diferencias de Inventario Articulos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Diferencias de Inventario Articulos"
            'codigo = "{sinven.codalmac}"
            cadTitulo = "Diferencias Inventario Articulos"
            
        Case 14 '14: Actualizar Direfencias Inventario (NO IMPRIME INFORME)
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Actualizar Diferencias de Inventario de Articulos"
            Me.Caption = "Inventario de Articulos"
            
        Case 15 '15: Listado de Articulos Inactivos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Articulos Inactivos"
            cadTitulo = "Listado Articulos Inactivos"
    
        Case 16 '16 .- Listado Valoracion de Stocks Inventariados
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Valoración Stocks Inventariados"
            cadTitulo = "Listado Valoración Stocks Inventariados"
            
        Case 17 '17 .- Listado Valoración Stocks
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Valoración Stocks"
            cadTitulo = "Listado Valoración Stocks"
            
        Case 19 '19 .- Inf. Stocks a una Fecha
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Informe Stocks a una Fecha"
            cadTitulo = "Stocks a una Fecha"
    End Select
End Sub



Private Sub ListadosFacturacion(H As Integer, W As Integer)
    Select Case OpcionListado
        Case 20    'Listado de Actividades de Clientes
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Actividades de Clientes"
            indFrame = 1
            Codigo = "{sactiv.codactiv}"
            Orden1 = "{sactiv.codactiv}"
            Orden2 = "{sactiv.nomactiv}"
            cadTitulo = "Listado Actividades de Clientes"
            cadNomRPT = "rFacActividades.rpt"
            
        
            
        Case 23     'Listado de Tipos de Formas de Envío
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Formas de Envío"
            indFrame = 1
            Codigo = "{senvio.codenvio}"
            Orden1 = "{senvio.codenvio}"
            Orden2 = "{senvio.nomenvio}"
            cadTitulo = "Listado Formas de Envio"
            cadNomRPT = "rFacEnvio.rpt"
            
        Case 24    'Tarifas Venta
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tarifas Venta"
            indFrame = 1
            Codigo = "{starif.codlista}"
            Orden1 = "{starif.codlista}"
            Orden2 = "{starif.nomlista}"
            cadTitulo = "Listado Tarifas Venta"
            cadNomRPT = "rFacTarifasVen.rpt"
            
        Case 27     'Situaciones Especiales
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Situaciones Especiales"
            indFrame = 1
            Codigo = "{ssitua.codsitua}"
            Orden1 = "{ssitua.codsitua}"
            Orden2 = "{ssitua.nomsitua}"
            cadTitulo = "Listado Situaciones Especiales"
            cadNomRPT = "rFacSituaciones.rpt"
            
        Case 28    '28: Informe de Tarifas de Precios
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Tarifas de Artículos"
            Codigo = "{slista"
            indFrame = 5
            cadTitulo = "Listado Tarifas Articulos"
            
        Case 29  '29: Informe Promociones
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Promociones Tarifas"
            Codigo = "{spromo"
            indFrame = 5
            cadTitulo = "Listado Promociones de Tarifas"
            
        Case 30 '30: Informe Precios Especiales
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Precios Especiales Artículos"
            Codigo = "{sprees"
            indFrame = 5
            cadTitulo = "Listado Precios Especiales"
            
        Case 245, 247 '245: Informe control margenes tarifas
            indFrame = 5
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Control Margenes de Tarifas"
            Codigo = "{slista"
            cadTitulo = "Listado Control Margenes Tarifas"
            cboDecimales.ListIndex = 4
        Case 246 '246: Informe margen ventas x articulo
            indFrame = 15
            H = 5300
            W = 7820
            PonerFrameVisible Me.FrameEstMargenes, True, H, W
            cadTitulo = "Listado Margen ventas por artículo"
    End Select
End Sub


Private Sub ListadosCompras(H As Integer, W As Integer)
'=============================================
'==== Listados de COMPRAS

    Select Case OpcionListado
        Case 309 '309: Listado precios de compra
            H = 4450
            W = 6920
            PonerFrameVisible Me.FrameDtosFM, True, H, W
            ponerOptVisible False
            Me.Frame4.visible = True
            Me.Frame4.Top = 840
            Me.Frame5.visible = False
            Me.Frame6.visible = False
            Me.cmdAceptarDtosFM.Top = 3500
            Me.cmdCancel(12).Top = Me.cmdAceptarDtosFM.Top
            indFrame = 6
    End Select
End Sub



Private Sub ListadosReparaciones(H As Integer, W As Integer)
'=============================================
'==== Listados de REPARACIONES

    Select Case OpcionListado
        Case 407 'Sustitución Num. serie
            H = 3700
            W = 5720
            PonerFrameVisible Me.FrameRepSustNSerie, True, H, W
            Me.lblNumSerie(0).Caption = "Nº Serie:   " & NumCod
            Me.lblNumSerie(1).Caption = "Artículo:   " & Me.CadTag
            Me.Caption = "Numeros de Serie"
            indFrame = 13
            
    End Select
End Sub




'---------------------------------------------------
'Para los bultos
Private Sub LimpiarTextosBultos()
Dim I As Integer
    For I = 2 To 6
        Me.txtBultos(I).Text = ""
        Me.txtBultos(I).Tag = ""
    Next I
End Sub



Private Sub PonerCamposDireccionBultos(indice As Integer)
Dim I As Integer

    'El indice mara el listindex del combo, por lo tanto sera indice + 1
    For I = 2 To 6
        Me.txtBultos(I).Text = RecuperaValor(Me.txtBultos(I).Tag, indice + 1)
    Next I
End Sub


Private Sub PonerCamposAlbaran()
'Informe Etiquetas Bultos
'si en NumCod se ha pasado el nº de un Albaran cargar por defectos valores
Dim Sql As String
Dim RS As ADODB.Recordset

    On Error GoTo ErrAlb
    
    '1) -- Buscar en la tabla de ALBARANES: PED -> ALV
    Sql = "SELECT codclien,coddirec, sum(numbultos) as totBultos"
    Sql = Sql & " FROM scaalb c INNER JOIN slialb l ON c.numalbar=l.numalbar and c.codtipom=l.codtipom"
    Sql = Sql & " WHERE c.numalbar=" & NumCod & " and c.codtipom='ALV'"
    Sql = Sql & " GROUP by c.numalbar,c.codtipom"
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        txtClie.Text = RS!CodClien
    
        CadTag = DBLet(RS!CodDirec, "T")
        
        txtBultos(1).Text = DBLet(RS!totbultos, "N")
        
        txtClie_LostFocus
    End If
    
    RS.Close
    Set RS = Nothing
    
    '2) Buscar en la tabla de FACTURAS PED -> FAV
    If txtClie.Text = "" Then
         'Comprobar en FACTURAS: x si se pasa de PED -> FAC
        Sql = "SELECT codclien,coddirec, sum(numbultos) as totBultos "
        Sql = Sql & " FROM (scafac c INNER JOIN scafac1 a ON c.numfactu=a.numfactu and c.codtipom=a.codtipom and c.fecfactu=a.fecfactu)"
        Sql = Sql & " INNER JOIN slifac l ON a.numfactu=l.numfactu and a.codtipom=l.codtipom and a.fecfactu=l.fecfactu and a.numalbar=l.numalbar and a.codtipoa=l.codtipoa"
        Sql = Sql & " WHERE a.numalbar=" & NumCod & " and a.codtipoa='ALV'"
        Sql = Sql & " GROUP BY a.numalbar,a.codtipoa"
        
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        If Not RS.EOF Then
            txtClie.Text = RS!CodClien
        
            CadTag = DBLet(RS!CodDirec, "T")
            
            txtBultos(1).Text = DBLet(RS!totbultos, "N")
            
            txtClie_LostFocus
        End If
        RS.Close
        Set RS = Nothing
    End If
    
    Exit Sub
    
ErrAlb:
    MuestraError Err.Number, "Poner campos Albaran.", Err.Description
End Sub



'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'   Borre de facturas
'
'
'   Borraremos las tablas de facturas , albaranes, hcos....
'
Private Sub CargaFechasPosibleEliminacion()
Dim F As Date
Dim F2 As Date
    Set miRsAux = New ADODB.Recordset
    cmbEliFac.Clear
    Codigo = "select min(fecfactu) from scafac"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    F2 = DateAdd("yyyy", -5, CDate("01/01/" & Year(Now)))

    Codigo = F2
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then Codigo = miRsAux.Fields(0)
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    Codigo = "31/12/" & Year(CDate(Codigo))
    
    While CDate(Codigo) < F2
        
        cmbEliFac.AddItem "     " & Format(CDate(Codigo), "dd/mm/yyyy")
        Codigo = CStr(DateAdd("yyyy", 1, CDate(Codigo)))
    
    Wend
    If cmbEliFac.ListCount > 0 Then cmbEliFac.ListIndex = 0
End Sub

Private Function BorrarFacturas() As Boolean
Dim FechaBorre As Date



    On Error GoTo EBorraFac
    BorrarFacturas = False
    
    FechaBorre = CDate(Trim(Me.cmbEliFac.List(cmbEliFac.ListIndex)))
    
    'Compruebo si estaban todas las facturas contabilizadas
    '------------------------------------------------------
    Codigo = "Select count(*) from scafac where fecfactu<='" & Format(FechaBorre, FormatoFecha) & "' and intconta = 0"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Existen " & NumRegElim & " facturas sin contabilizar en esas fechas", vbExclamation
        Exit Function
    End If
    
        
    'lo mismo para proeedores
    Codigo = "Select count(*) from scafpc where fecfactu<='" & Format(FechaBorre, FormatoFecha) & "' and intconta = 0"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Existen " & NumRegElim & " facturas de proveedores sin contabilizar en esas fechas", vbExclamation
        Exit Function
    End If
        
        
        
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 1, vUsu, "Borre facturas: " & Format(FechaBorre, "dd/mm/yyyy")
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    
    '   Lo dicho. LAS TABLAS son las indcadas above (jeje arriba)
    '   La fecha la manda fecfactu
    Codigo = "slifac|scafac1|svenci|srecom|scafac|"
    For NumRegElim = 1 To 5
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Tabla CLI: " & Orden1
        Orden1 = "DELETE FROM " & Orden1 & " WHERE  fecfactu<='" & Format(FechaBorre, FormatoFecha) & "'"
        Me.Refresh
        DoEvents
        conn.Execute Orden1
    Next NumRegElim
    
    '---------------------------------------------------------------------------------
    'Albarananes CLIENTES.
    '--
    Codigo = "scaalb|schalb|slialb|slhalb|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Albaranes(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fechaalb<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE codtipom = '"
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!codtipom & "'  AND numalbar = " & miRsAux!NumAlbar
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Borramos las cabceeras
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Albaranes(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fechaalb<='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
    Next
    DoEvents
    
    
    '---------------------------------------------------------------------------------
    'Pedidos CLIENTES.
    '--
    Codigo = "scaped|schped|sliped|slhped|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedcl<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numpedcl = "
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!Numpedcl
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Cabce
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedcl<='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
    Next
    DoEvents
    
    
    '---------------------------------------------------------------------------------
    'ofertas CLIENTES.
    '--
    Codigo = "scapre|schpre|slipre|slhpre|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Ofertas(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecofert<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numofert = "
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!NumOfert
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Ofertas(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecofert <='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
        
    Next
    DoEvents
    
    
    Codigo = "scarep|schrep|slirep|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Reparaciones: " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecrepar<='" & Format(FechaBorre, FormatoFecha) & "'"
        If NumRegElim = 1 Then
            'Lineas de reparacion solo hay en scarep
            'En shrep no hay lineas
            miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
            Orden1 = "DELETE FROM " & Orden1 & " WHERE numrepar = "
            While Not miRsAux.EOF
                conn.Execute Orden1 & miRsAux!numrepar
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
        End If
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Ofertas(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecrepar <='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
        
    Next
    DoEvents
    
    
    'TPV
    Label3(83).Caption = "TPV"
    Label3(83).Refresh
    Orden1 = " WHERE  fecventa <='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute "DELETE FROM sliven " & Orden1
    conn.Execute "DELETE FROM scaven " & Orden1

    
    'PRODUCCION
    Label3(83).Caption = "Produccion"
    Label3(83).Refresh
    Orden1 = "Select * from sordprod WHERE  feccreacion<='" & Format(FechaBorre, FormatoFecha) & "'"
    miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Orden1 = "DELETE FROM sliordpr WHERE codigo = "
    While Not miRsAux.EOF
        conn.Execute Orden1 & miRsAux!Codigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    Orden1 = "DELETE from sordprod WHERE  feccreacion <='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute Orden1

    Me.Refresh
    DoEvents
    
    '---------------------------------------------------------------------------------
    'Facturas proveedor
    '--
    Codigo = "slifpc|scafpa|scafpc|"
    For NumRegElim = 1 To 3
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Tabla PRO: " & Orden1
        Orden1 = "DELETE FROM " & Orden1 & " WHERE  fecfactu<='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
    Next NumRegElim
    
    
    
    
    Codigo = "slhalp|slialp|scaalp|schalp|"
    For NumRegElim = 1 To 4
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Albaranes prov: " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1 & " WHERE  fechaalb<='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
        
    Next
    DoEvents
    
    
    
    
    '-----------------------------------------------
    'Pedidos proveedor
    '--
    Codigo = "scappr|schppr|slippr|slhppr|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos prov.(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedpr<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numpedpr = "
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!numpedpr
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos prov.(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedpr <='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
        
    Next
    Me.Refresh
    DoEvents
    
    'slhmov slhtra
    Label3(83).Caption = "Hco movimientos"
    Label3(83).Refresh
    Orden1 = "DELETE FROM slhmov WHERE  fecmovim<='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute Orden1
    
    Label3(83).Caption = "Hco traspasos"
    Label3(83).Refresh
    Orden1 = "DELETE FROM slhtra WHERE  fechamov<='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute Orden1
    
    
    'Ahora me cargo los movimientos en la smoval
    Label3(83).Caption = "Movimientos"
    Label3(83).Refresh
    Orden1 = "DELETE FROM smoval WHERE  fechamov<='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute Orden1
    
    'Inventario
    Label3(83).Caption = "Hco inventario"
    Label3(83).Refresh
    Orden1 = "DELETE FROM shinve WHERE  fechainv<='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute Orden1
    
    
    BorrarFacturas = True
    Exit Function
EBorraFac:
    MuestraError Err.Number
End Function


'Envio -EMAIL

Private Sub PonerTamnyosMail(peque As Boolean)
    If peque Then
        Me.Height = Me.FrameEnvioMail.Height + 60
        Me.Width = Me.FrameEnvioMail.Width
    Else
        Me.Height = Me.FrameMantenimientos.Height
        Me.Width = Me.FrameMantenimientos.Width
    End If
    Me.Height = Me.Height + 420
    Me.Width = Me.Width + 120
    Me.FrameEnvioMail.visible = peque
    DoEvents
    Me.Refresh
End Sub





Private Function GeneracionEnvioMail() As Boolean
Dim m As CParamRpt

    On Error GoTo EGeneracionEnvioMail
    GeneracionEnvioMail = False

    Set m = New CParamRpt
    If m.Leer(21) = 1 Then
        Set m = Nothing
        Exit Function
    End If
    
    cadSelect = "Select * from tmpnlotes where codusu =" & vUsu.Codigo & " ORDER BY codalmac,numalbar,codprove"
    miRsAux.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    While Not miRsAux.EOF
        
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    
        Label14(22).Caption = "Mantenimiento: " & miRsAux!codArtic & " Cliente: " & miRsAux!codProve
        Label14(22).Refresh
        
'
        cadFormula = "({scaman.nummante}='" & miRsAux!codArtic & "') "
        cadFormula = cadFormula & " AND ({scaman.codclien}=" & miRsAux!codProve & ") "


        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = True
            .NombreRPT = m.Documento
            .Opcion = 78  'Carta renovacion manteniientos
            .Titulo = ""
            .Show vbModal
        End With
    
                    
        'Subo el progress bar
        Label14(22).Caption = "Generando PDF"
        Label14(22).Refresh
        Me.PBMail.Value = Me.PBMail.Value + 1
        If (Me.PBMail.Value Mod 25) = 24 Then
            Me.Refresh
            DoEvents
            Espera 1
        End If
        Me.Refresh
        DoEvents
        
        
        
        'FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & RS!NumAlbar & Format(RS!codProve, "0000000") & Format(RS!codArtic, "0000000") & Format(RS!FechaAlb, "yymmdd") & ".pdf"
        FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & Format(miRsAux!codProve, "0000000") & ".pdf"
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set m = Nothing
    GeneracionEnvioMail = True
    Exit Function
EGeneracionEnvioMail:
       MuestraError Err.Number
End Function



Private Function HacerSQLListado82_83() As Boolean
    
On Error GoTo EHacerSQLListado82_83
    
    
    HacerSQLListado82_83 = False
    InicializarVbles


    If OpcionListado = 82 Then
        'Hacer UPDATE de scaalb
        Codigo = "UPDATE scaalb set factursn = 1 "
        If NumCod <> "" Then cadSelect = " codtipom ='" & NumCod & "'"
        
        cadParam = "fechaalb"
        cadFormula = CadenaDesdeHastaBD(txtCodigo(117).Text, txtCodigo(118).Text, "codclien", "N")
        If cadFormula <> "" Then
            If cadSelect <> "" Then cadSelect = cadSelect & " AND "
            cadSelect = cadSelect & cadFormula
        End If
        

    Else
        'Hacer borrar avisos
        Codigo = "DELETE FROM scaavi"
        cadSelect = " situacio = 3"
        cadParam = "fechaavi"
    End If
    
    cadFormula = CadenaDesdeHastaBD(txtCodigo(119).Text, txtCodigo(120).Text, cadParam, "F")
    If cadFormula <> "" Then
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & cadFormula
    End If
    
    If cadSelect <> "" Then cadSelect = " WHERE " & cadSelect
    Codigo = Codigo & cadSelect
    conn.Execute Codigo
    
    If OpcionListado = 83 Then MsgBox "Proceso finalizado", vbExclamation
    
    HacerSQLListado82_83 = True
    Exit Function
EHacerSQLListado82_83:
    MuestraError Err.Number
End Function







Private Function NuevasComprobacionesContabilizacion(Proveedores As Boolean, ByVal Sql As String) As Boolean
Dim RT As ADODB.Recordset
Dim C As String
Dim F As Date
Dim Fin As Boolean
Dim ComprobacionFechaMenor As Boolean
Dim cControlFra As CControlFacturaContab
    
    On Error GoTo ENuevasComprobacionesContabilizacion
    NuevasComprobacionesContabilizacion = False
    
    
    
    Set cControlFra = New CControlFacturaContab
        'Tenemos que comprobar la fecha factura
    Set RT = New ADODB.Recordset
    ComprobacionFechaMenor = False

    If Proveedores Then
        C = "select fecrecep from scafpc WHERE " & Sql
        C = C & " GROUP BY fecrecep ORDER BY fecrecep"
    Else
        C = "Select fecfactu from scafac WHERE " & Sql
        C = C & " GROUP BY fecfactu ORDER BY fecfactu"
    End If
    
    
    RT.Open C, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Fin = False
    While Not Fin
        F = RT.Fields(0)
        C = cControlFra.FechaCorrectaContabilizazion(ConnConta, F)
        If C <> "" Then
            Fin = True
        Else
            C = cControlFra.FechaCorrectaIVA(ConnConta, F)
            If C <> "" Then
                Fin = True
            Else
                If Proveedores Then
                    'Solo compruebo una vez
                    If Not ComprobacionFechaMenor Then
                        If cControlFra.FechaRecepMenorQueProveedor(ConnConta, F) Then C = "Factura contabilizada con fecha de recepción menor"
                            
                        ComprobacionFechaMenor = True
                    End If
                End If
            End If
        End If
        RT.MoveNext
        If Not Fin Then Fin = RT.EOF
    Wend
    RT.Close
    
    If C <> "" Then
        C = C & "(" & F & ")"
        MsgBox C, vbExclamation
    Else
        NuevasComprobacionesContabilizacion = True
    End If
    
    
ENuevasComprobacionesContabilizacion:
    If Err.Number <> 0 Then MuestraError Err.Number, "Nueva Comprobacion Contabilizacion"
    Set RT = Nothing
    Set cControlFra = Nothing
End Function

Private Sub ponerOptVisible(Vis As Boolean)

        Me.optFrDto(0).visible = Vis
        Me.optFrDto(1).visible = Vis
        Me.optFrDto(2).visible = Vis
        Me.optFrDto(3).visible = Vis
End Sub
