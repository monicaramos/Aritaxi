VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLiqListReten 
   Caption         =   "Informes"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin VB.Frame FrameRecibosReten 
      Height          =   5835
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   7515
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
         Index           =   6
         Left            =   2310
         TabIndex        =   11
         Top             =   4290
         Width           =   765
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
         Left            =   3090
         TabIndex        =   45
         Top             =   4290
         Width           =   3945
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
         Index           =   96
         Left            =   2310
         MaxLength       =   10
         TabIndex        =   8
         Top             =   2760
         Width           =   1215
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
         Index           =   4
         Left            =   2310
         MaxLength       =   10
         TabIndex        =   9
         Top             =   3270
         Width           =   1215
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
         Index           =   82
         Left            =   2310
         TabIndex        =   4
         Tag             =   "Num vehiculo|N|N|||shilla|codclien|000000|S|"
         Top             =   1230
         Width           =   945
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
         Index           =   82
         Left            =   3270
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1230
         Width           =   3765
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
         Index           =   14
         Left            =   6030
         TabIndex        =   14
         Top             =   4980
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepRecibos 
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
         Left            =   4980
         TabIndex        =   13
         Top             =   4980
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
         Index           =   83
         Left            =   2310
         TabIndex        =   5
         Tag             =   "Num vehiculo|N|N|||shilla|codclien|000000|S|"
         Top             =   1620
         Width           =   945
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
         Index           =   83
         Left            =   3270
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1620
         Width           =   3765
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
         Index           =   102
         Left            =   2310
         TabIndex        =   6
         Top             =   2190
         Width           =   1245
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
         Index           =   103
         Left            =   5040
         TabIndex        =   7
         Top             =   2190
         Width           =   1245
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
         Index           =   61
         Left            =   2310
         TabIndex        =   10
         Top             =   3780
         Width           =   765
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
         Index           =   61
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   3780
         Width           =   3945
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   255
         Left            =   570
         TabIndex        =   48
         Top             =   4710
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label3 
         Caption         =   "Cargando tabla temporal..."
         Height          =   195
         Index           =   1
         Left            =   570
         TabIndex        =   47
         Top             =   5010
         Visible         =   0   'False
         Width           =   3075
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
         Left            =   570
         TabIndex        =   46
         Top             =   4290
         Width           =   1125
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   5
         Left            =   2040
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta"
         Top             =   4290
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   2040
         Top             =   3270
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recibo"
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
         Index           =   48
         Left            =   570
         TabIndex        =   23
         Top             =   3300
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
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
         Index           =   47
         Left            =   570
         TabIndex        =   22
         Top             =   2760
         Width           =   885
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
         Index           =   34
         Left            =   1320
         TabIndex        =   21
         Top             =   1590
         Width           =   690
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
         Index           =   17
         Left            =   570
         TabIndex        =   20
         Top             =   990
         Width           =   555
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   38
         Left            =   2055
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Recibos de Retenciones"
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
         Left            =   570
         TabIndex        =   19
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
         Index           =   88
         Left            =   1320
         TabIndex        =   18
         Top             =   1230
         Width           =   585
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   39
         Left            =   2055
         ToolTipText     =   "Buscar socio"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   2040
         Top             =   2190
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   4740
         Top             =   2190
         Width           =   240
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
         Index           =   87
         Left            =   570
         TabIndex        =   17
         Top             =   1950
         Width           =   630
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
         Index           =   86
         Left            =   3870
         TabIndex        =   16
         Top             =   2190
         Width           =   600
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
         Index           =   85
         Left            =   1320
         TabIndex        =   15
         Top             =   2190
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
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
         Index           =   50
         Left            =   570
         TabIndex        =   12
         Top             =   3780
         Width           =   1290
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   8
         Left            =   2025
         ToolTipText     =   "Buscar f.pago"
         Top             =   3780
         Width           =   240
      End
   End
   Begin VB.Frame FrameListado 
      Height          =   5865
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   7515
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
         Left            =   5670
         TabIndex        =   42
         Top             =   4110
         Visible         =   0   'False
         Width           =   1005
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
         Left            =   5670
         TabIndex        =   41
         Top             =   3660
         Visible         =   0   'False
         Width           =   1005
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
         Left            =   5610
         TabIndex        =   33
         Top             =   4860
         Width           =   1135
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
         Left            =   4380
         TabIndex        =   32
         Top             =   4860
         Width           =   1135
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
         Index           =   86
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   30
         Top             =   2925
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
         Index           =   85
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   29
         Top             =   2520
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
         Index           =   1
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   28
         Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
         Top             =   1770
         Width           =   855
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1770
         Width           =   3765
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
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   27
         Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
         Top             =   1365
         Width           =   855
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1365
         Width           =   3765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Resumen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   750
         TabIndex        =   25
         Top             =   3750
         Width           =   2265
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   5370
         Picture         =   "frmLiqListReten.frx":0000
         ToolTipText     =   "Buscar fecha"
         Top             =   3690
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Importe Base"
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
         Height          =   255
         Index           =   0
         Left            =   4500
         TabIndex        =   44
         Top             =   4110
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Informe"
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
         Height          =   255
         Index           =   0
         Left            =   4500
         TabIndex        =   43
         Top             =   3360
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label10 
         Caption         =   "Retenciones Servicios a Crédito"
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
         Left            =   510
         TabIndex        =   40
         Top             =   390
         Width           =   5655
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   1410
         Picture         =   "frmLiqListReten.frx":008B
         Top             =   2940
         Width           =   240
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
         Left            =   750
         TabIndex        =   39
         Top             =   2970
         Width           =   570
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
         Left            =   750
         TabIndex        =   38
         Top             =   2550
         Width           =   600
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
         Left            =   510
         TabIndex        =   37
         Top             =   2160
         Width           =   630
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   1410
         Picture         =   "frmLiqListReten.frx":0116
         Top             =   2520
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
         Left            =   750
         TabIndex        =   36
         Top             =   1770
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   1
         Left            =   1410
         Picture         =   "frmLiqListReten.frx":01A1
         Top             =   1770
         Width           =   240
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
         Left            =   750
         TabIndex        =   35
         Top             =   1365
         Width           =   600
      End
      Begin VB.Label Label9 
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
         Index           =   2
         Left            =   510
         TabIndex        =   34
         Top             =   1020
         Width           =   555
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   0
         Left            =   1410
         Picture         =   "frmLiqListReten.frx":02A3
         Top             =   1365
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmLiqListReten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OpcionListado As Integer

Private Const IdPrograma = 206

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
Private WithEvents frmMtoV As frmGesSocios ' socios
Attribute frmMtoV.VB_VarHelpID = -1

Public WithEvents frmFP As frmFacFormasPago ' formas de pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios ' bancos propios
Attribute frmMtoBancosPro.VB_VarHelpID = -1

' Importes para la Grabacion de Cabecera de Facturas de Socio
Dim TotalFac As Currency
Dim TotalLiq As Currency
Dim BaseImpo As Currency
Dim BaseReten As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim vPorcIva As String
Dim PorceIVA As Currency

Dim iva As String
Dim porIva As Currency
Dim NomArtic As String


Dim tipoMov As String
Dim codSocio As String

Dim kCampo As Integer

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Sub cmdAceptar_Click()
Dim Codigo As String
Dim FecFac As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


    If Not DatosOk Then Exit Sub
    
    InicializarVbles
    
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 1
   
    'Desde/Hasta numero de V
    '---------------------------------------------
    If txtcodigo(0).Text <> "" Or txtcodigo(1).Text <> "" Then
        Codigo = "{" & Tabla & ".codsocio}"
        If Not PonerDesdeHasta(Codigo, "N", 0, 1, "pDHSocio=""") Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtcodigo(85).Text <> "" Or txtcodigo(86).Text <> "" Then
        Codigo = "{" & Tabla & ".fecfactu}"
        If Not PonerDesdeHasta(Codigo, "F", 85, 86, "pDHFecha=""") Then Exit Sub
    End If
    
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub

    If Check1.Value Then
        cadNombreRPT = "rListRetencionesRes.rpt"
    Else
        cadNombreRPT = "rListRetenciones.rpt"
    End If
    
    cadTitulo = "Retenciones Servicios de Crédito"

    cadParam = cadParam & "pFecFac= """ & txtcodigo(3).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pTitulo= ""Retenciones Servicios de Crédito""|"
    numParam = numParam + 1
    cadParam = cadParam & "pBase=" & TransformaComasPuntos(ImporteSinFormato(txtcodigo(2).Text)) & "|"
    numParam = numParam + 1
    
    ConSubInforme = False
    
    LlamarImprimir False
        
    cmdCancelar_Click

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

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdAcepRecibos_Click()
Dim Codigo As String
Dim FecFac As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


    If Not DatosOk Then Exit Sub
    
    InicializarVbles
    
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 1
   
    'Desde/Hasta numero de V
    '---------------------------------------------
    If txtcodigo(82).Text <> "" Or txtcodigo(83).Text <> "" Then
        Codigo = "{" & Tabla & ".codsocio}"
        If Not PonerDesdeHasta(Codigo, "N", 82, 83, "pDHSocio=""") Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtcodigo(102).Text <> "" Or txtcodigo(103).Text <> "" Then
        Codigo = "{" & Tabla & ".fecfactu}"
        If Not PonerDesdeHasta(Codigo, "F", 102, 103, "pDHFecha=""") Then Exit Sub
    End If
    
'    If Not AnyadirAFormula(cadFormula, "{sreten.tiporeten} = 0") Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, "{sreten.tiporeten} = 0") Then Exit Sub
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub

    If CargarTablaTemporal(Tabla, cadSelect) Then
        If Not HayRegParaInforme("tmpinformes", "codusu= " & vUsu.Codigo) Then Exit Sub

        '[Monica]19/02/2018: Entra Cordoba
            '[Monica]19/11/2018: entra Sevilla
        If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3 Then

            cadNombreRPT = "rRecRetenciones.rpt"
            cadTitulo = "Recibos Retenciones Servicios de Crédito"
        
            cadParam = cadParam & "pFecFac= """ & txtcodigo(3).Text & """|"
            numParam = numParam + 1
            cadParam = cadParam & "pTitulo= ""Retenciones Servicios de Crédito""|"
            numParam = numParam + 1
            cadParam = cadParam & "pBase=" & TransformaComasPuntos(ImporteSinFormato(txtcodigo(2).Text)) & "|"
            numParam = numParam + 1
            
            ConSubInforme = False
            
            ' llamamos a la impresion de recibo
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            LlamarImprimir False
            
            If MsgBox("¿Impresion correcta para actualizar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                If ActualizarRegistros Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancelar_Click
                End If
            End If
        Else
            GenerarFacturaRetenciones
            BotonReimprimir
            cmdCancelar_Click
        End If
    End If

End Sub

Private Sub BotonReimprimir()
Dim Sql As String
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    InicializarVbles
    
    
        
        indRPT = 12 'Facturas Clientes
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, False, pPdfRpt) Then Exit Sub
    
    
        cadFormula = "{scafac.codtipom} = 'FAV'  and "
        cadFormula = cadFormula & "{scafac.fecfactu}= Date(" & Year(CDate(txtcodigo(4))) & "," & Month(CDate(txtcodigo(4).Text)) & "," & Day(CDate(txtcodigo(4).Text)) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
    
        devuelve = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", "FAV", "T")
     
     
         With frmImprimir
                'Nuevo. Febrero 2010
'                .outClaveNombreArchiv = devuelve & Format(Me.data1.Recordset!NumFactu, "000")
'                .outCodigoCliProv = Me.data1.Recordset!codSocio
'                .outTipoDocumento = 100
                
                .FormulaSeleccion = cadFormula
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                .NombrePDF = pPdfRpt
                .NombreRPT = pPdfRpt
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = 53 'OpcionListado
                .Titulo = ""
                .Show vbModal
        End With
    
    
    
    

End Sub






Private Function ActualizarRegistros() As Boolean
Dim Sql As String
Dim SQL2 As String
Dim Rs As ADODB.Recordset
Dim Sql2Values As String
Dim fac As CFacturaCom
Dim b As Boolean
Dim Socio As Long
Dim FormatSocio As String
Dim cuenta As String
Dim vSocio As CSocio
Dim MenError As String
Dim Mens As String


    On Error GoTo eActualizarRegistros
    
    ActualizarRegistros = False
        
    Label3(1).visible = True
    Label3(1).Caption = "Insertando registros..."
    DoEvents
        
    Screen.MousePointer = vbHourglass

    conn.BeginTrans
    ConnConta.BeginTrans
    
    Sql = "select * from tmpinformes where codusu = " & vUsu.Codigo
    Sql = Sql & " order by codigo1, importe1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    SQL2 = "insert into sreten (codsocio, numeruve, fecfactu, numfactu, impreten, tiporeten, desdefec, hastafec) values "
    b = True
    While Not Rs.EOF And b
        Sql2Values = "(" & DBSet(Rs!Codigo1, "N") & "," & DBSet(Rs!Importe1, "N") & "," & DBSet(txtcodigo(4).Text, "F") & ","
        Sql2Values = Sql2Values & "0," & DBSet(Rs!importe2 * (-1), "N") & ",1," & DBSet(txtcodigo(102).Text, "F") & ","
        Sql2Values = Sql2Values & DBSet(txtcodigo(103).Text, "F") & ")"
        
        conn.Execute SQL2 & Sql2Values

'desde aqui
        Set fac = New CFacturaCom
    
        fac.TotalFac = DBLet(Rs!importe2, "N")
        fac.FecFactu = txtcodigo(4).Text
        fac.NumFactu = "R-" & Format(Rs!Codigo1, "00000") & Format(Rs!Importe1, "00000")
        
        fac.Proveedor = DBLet(Rs!Codigo1, "N")
        fac.NombreProv = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Rs!Codigo1, "T")
        fac.DomicilioProv = DevuelveDesdeBD(conAri, "domclien", "sclien", "codclien", Rs!Codigo1, "T")
        fac.CPostalProv = DevuelveDesdeBD(conAri, "codpobla", "sclien", "codclien", Rs!Codigo1, "T")
        fac.PoblacionProv = DevuelveDesdeBD(conAri, "pobclien", "sclien", "codclien", Rs!Codigo1, "T")
        fac.ProvinciaProv = DevuelveDesdeBD(conAri, "proclien", "sclien", "codclien", Rs!Codigo1, "T")
        fac.NIFProv = DevuelveDesdeBD(conAri, "nifclien", "sclien", "codclien", Rs!Codigo1, "T")
        fac.ForPago = txtcodigo(61).Text
        
        'Cuenta Prevista de Cobro de las Facturas
        fac.BancoPr = txtcodigo(6).Text
        fac.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", fac.BancoPr, "N")
        'cuenta contable de proveedor
        'comprobamos q la cuenta contable exista en contabilidad
        Socio = DBLet(Rs!Codigo1, "N")
        FormatSocio = String((vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior), "0")
        cuenta = Trim(vParamAplic.Raiz_Cta_Reten_Soc & Format(Socio, FormatSocio))
        Sql = ""
        Sql = DevuelveDesdeBD(conConta, "nommacta", "cuentas", "codmacta", cuenta, "T")
        If Sql = "" Then
            MsgBox "La cuenta contable del socio: " & Format(Socio, "000000") & " no existe.", vbExclamation
            conn.RollbackTrans
            ConnConta.RollbackTrans
            Screen.MousePointer = vbDefault
            Me.Label3(1).visible = False
            DoEvents
            Exit Function
        End If
        fac.CtaProve = cuenta
        
        '[Monica]añadido no se cargaba la ccc del socio en tesoreria
        Set vSocio = New CSocio
        If vSocio.LeerDatos(CStr(Socio)) Then
            '[Monica]22/11/2013: iban
            fac.CCC_Iban = vSocio.Iban
            fac.CCC_Entidad = vSocio.Banco
            fac.CCC_Oficina = vSocio.Sucursal
            fac.CCC_CC = vSocio.DigControl
            fac.CCC_CTa = vSocio.CuentaBan
        End If
        Set vSocio = Nothing
      
        MenError = "Error al pasar a tesoreria"
        '[Monica]26/01/2012: cambiamos el parametro opcional para que imprima en texto de csb otra cosa
        fac.Proveedor = Year(CDate(txtcodigo(103).Text))
        b = fac.InsertarEnTesoreria(MenError, True) ' true = indicamos que venimos de pago de retenciones
        
        Set fac = Nothing
'hasta aqui
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing

    conn.CommitTrans
    ConnConta.CommitTrans
    
    ActualizarRegistros = True
    
    Screen.MousePointer = vbDefault
    Me.Label3(1).visible = False
    DoEvents
    
    Exit Function

eActualizarRegistros:
    If Err.Number <> 0 Or Not b Then
        Mens = ""
        If Not b Then Mens = Mens & MenError
        MuestraError Err.Number, "Actualizar registros", Mens & vbCrLf & Err.Description
        conn.RollbackTrans
        ConnConta.RollbackTrans
        Me.Label3(1).visible = False
        DoEvents
        Screen.MousePointer = vbDefault
    End If
End Function


Private Function CargarTablaTemporal(Tabla As String, cadSelect As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim SqlValues As String
Dim Importe As Currency
    
    On Error GoTo eCargarTablaTemporal

    CargarTablaTemporal = False
    
    Me.Label3(1).visible = True
    Label3(1).Caption = "Cargando tabla temporal..."
    DoEvents
    
    Screen.MousePointer = vbHourglass
    
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "select codsocio, numeruve, sum(if(impreten is null,0,impreten)) as Importe from sreten "
    If cadSelect <> "" Then Sql = Sql & " where " & cadSelect
    Sql = Sql & " group by 1 "
    Sql = Sql & " order by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SqlValues = ""
    While Not Rs.EOF
        Importe = DBLet(Rs!Importe, "N") - ComprobarCero(txtcodigo(96).Text)
        
        If Importe > 0 Then
            SqlValues = SqlValues & "(" & vUsu.Codigo & "," & DBLet(Rs!codSocio, "N") & "," & DBLet(Rs!NumerUve, "N") & "," & DBSet(Importe, "N") & "," & DBSet(txtcodigo(103).Text, "F") & "),"
        End If
    
        Rs.MoveNext
    Wend
    
    If SqlValues <> "" Then
        Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, fecha1) values "
        Sql = Sql & Mid(SqlValues, 1, Len(SqlValues) - 1)
        
        conn.Execute Sql
    
    End If
    
    
    
    
    Set Rs = Nothing

    CargarTablaTemporal = True
    Screen.MousePointer = vbDefault
    Me.Label3(1).visible = False
    DoEvents
    Exit Function
    
eCargarTablaTemporal:
    Screen.MousePointer = vbDefault
    Me.Label3(1).visible = False
    DoEvents
    MuestraError Err.Number, "Cargando Tabla Temporal", Err.Description
End Function


Private Sub Form_Activate()
    cadFormula = ""
    numParam = 0
    cadParam = ""


    PonerFoco txtcodigo(0)

End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer

    'Icono del form
    Me.Icon = frmppal.Icon
    
    Me.FrameListado.visible = False
    Me.FrameRecibosReten.visible = False
    
    For kCampo = 0 To 1
        Me.imgBuscarOfer(kCampo).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 38 To 39
        Me.imgBuscarOfer(kCampo).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    Me.imgBuscarOfer(5).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgBuscarOfer(8).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    Me.imgFecha(0).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Me.imgFecha(13).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Me.imgFecha(14).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Me.imgFecha(21).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Me.imgFecha(23).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Me.imgFecha(24).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    
    
    
    Select Case OpcionListado
        Case 0
            PonerFrameListadoVisible True, H, W

            Tabla = "sreten"
        
            txtcodigo(3).Text = Format(Now, "dd/mm/yyyy")
            txtcodigo(2).Text = "0,00"
            
        Case 1
            PonerFrameRecibosRetenVisible True, H, W

            Tabla = "sreten"
            
            '[Monica]28/10/2016: cambiamos labels
            If vParamAplic.Cooperativa = 1 Then
                Me.Label4(48).Caption = "Fec.Factura"
                Me.Label13.Caption = "Facturas de Retención"
                
                'no metemos el importe
                Me.Label4(47).Enabled = False
                Me.Label4(47).visible = False
                Me.txtcodigo(96).Enabled = False
                Me.txtcodigo(96).visible = False
                
            End If
        
        
    End Select

End Sub

Private Sub PonerFrameListadoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de los mantenimientos de tabla: smarca, stipar,...

    H = 6405
    W = 7515
    PonerFrameVisible Me.FrameListado, visible, H, W

End Sub

Private Sub PonerFrameRecibosRetenVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de los mantenimientos de tabla: smarca, stipar,...

    H = 6405
    W = 7515
    PonerFrameVisible Me.FrameRecibosReten, visible, H, W

End Sub



Private Function DatosOk() As Boolean
Dim encontrado As String
Dim Codigo As String

    DatosOk = True
    
    Select Case OpcionListado
        Case 0 ' listado de retenciones
                If txtcodigo(2).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente un Importe Base.", vbExclamation
                    DatosOk = False
                    Exit Function
                End If
                If txtcodigo(3).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la fecha de listado.", vbExclamation
                    DatosOk = False
                    Exit Function
                End If
                
            
        Case 1 ' Impresion de recibos de retenciones
            '[Monica]19/02/2018: Entra Cordoba
                '[Monica]19/11/2018: Entra Sevilla
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3 Then
                'fecha de recibo
                If txtcodigo(4).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de Recibo.", vbExclamation
                    PonerFoco txtcodigo(4)
                    DatosOk = False
                    Exit Function
                End If
            Else
                If txtcodigo(4).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la fecha de factura.", vbExclamation
                    DatosOk = False
                    Exit Function
                End If
                If vParamAplic.ArtRetenciones = "" Then
                    MsgBox "No está configurado el artículo de retenciones en parámetros. Revise", vbExclamation
                    DatosOk = False
                    Exit Function
                Else
                    'busco el iva del articulo
                    iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArtRetenciones, "T")
                    If iva = "" Then
                        MsgBox "El artículo de retenciones no tiene asignado el iva. Revise.", vbExclamation
                        DatosOk = False
                        Exit Function
                    Else
                        'busco el nombre del articulo
                        NomArtic = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtRetenciones, "T")
                    End If
                End If
            
            
            End If
            'forma de pago
            If txtcodigo(61).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la Forma de Pago.", vbExclamation
                PonerFoco txtcodigo(61)
                DatosOk = False
                Exit Function
            End If
            'cuenta de pago
            If txtcodigo(6).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la Cuenta de Pago.", vbExclamation
                PonerFoco txtcodigo(6)
                DatosOk = False
                Exit Function
            End If
            
            '[Monica]01/10/2012: obligamos a meter el desde/hasta fecha
            If txtcodigo(102).Text = "" Or txtcodigo(103).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la Fecha Desde / Hasta.", vbExclamation
                PonerFoco txtcodigo(102)
                DatosOk = False
                Exit Function
            End If
            
    End Select
        
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
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoV_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 0, 1 ' Vsocio
            indCodigo = Index
            
            Set frmMtoV = New frmGesSocios
            frmMtoV.DatosADevolverBusqueda = "0|1|"
            frmMtoV.Show vbModal
            Set frmMtoV = Nothing
        
        Case 38, 39 ' socio
            indCodigo = Index + 44
            
            Set frmMtoV = New frmGesSocios
            frmMtoV.DatosADevolverBusqueda = "0|1|"
            frmMtoV.Show vbModal
            Set frmMtoV = Nothing
        
        Case 8 ' forma de pago
            indCodigo = Index + 53
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0|1|"
            frmFP.Show vbModal
            Set frmFP = Nothing
        
        Case 5 ' cuenta de pago
            indCodigo = Index + 1
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
'Dim cerrar As Boolean
'
'    KEYpressGnral KeyAscii, 2, cerrar
'    If cerrar Then Unload Me
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If

End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Tabla As String
Dim codCampo As String, NomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean
Dim encontrado As String


    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    Select Case Index
        Case 85, 86, 102, 103 'FECHA Desde Hasta
            PonerFormatoFecha txtcodigo(Index)
            
        Case 0, 1 'V Socio
            PonerFormatoEntero txtcodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sclien", "nomclien", "codclien", "N")
            
        Case 82, 83 'Socio
            PonerFormatoEntero txtcodigo(Index)
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "sclien", "nomclien", "codclien", "N")
            
            
        Case 2 ' importe base
            PonerFormatoDecimal txtcodigo(Index), 3
            
        Case 3, 4 ' fecha de listado
            PonerFormatoFecha txtcodigo(Index)
            
        Case 96 ' importe
            PonerFormatoDecimal txtcodigo(Index), 3
            
        Case 6 ' cta de banco
            If txtcodigo(Index).Text <> "" Then
                encontrado = DevuelveDesdeBD(conAri, "nombanpr", "sbanpr", "codbanpr", txtcodigo(Index).Text, "T")
                If encontrado = "" Then
                    MsgBox "El banco introducido no existe", vbExclamation
                    PonerFoco txtcodigo(Index)
                Else
                    txtnombre(Index).Text = encontrado
                End If
            End If
        
        Case 61 ' forma de pago
            If txtcodigo(Index).Text <> "" Then
                If Not IsNumeric(txtcodigo(Index).Text) Then
                    MsgBox "La forma de pago debe ser numérica.", vbExclamation
                    PonerFoco txtcodigo(Index)
                    Exit Sub
                End If
                txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
                encontrado = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", txtcodigo(Index).Text, "T")
                If encontrado = "" Then
                    MsgBox "La forma de pago introducida no existe.", vbExclamation
                    PonerFoco txtcodigo(Index)
                Else
                    txtnombre(Index).Text = encontrado
                End If
            End If
        
        
        
    End Select
End Sub

Private Function GenerarFacturaRetenciones() As Boolean
Dim vC As CTiposMov
Dim fac As CFactura
Dim Cad As String
Dim Sql As String
Dim totfactu As Currency
Dim BaseImp As Currency
Dim base0 As Currency
Dim base1 As Currency
Dim base2 As Currency
Dim base4 As Currency
Dim ImpIVA As Currency
Dim cli As CCliente
Dim b As Boolean
Dim CADENA As String
Dim LetraSer As String
Dim ForPago As Integer
Dim FecFactu As Date
Dim NumFactu As Long
Dim codtipom As String
Dim Cantidad As Currency
Dim total As Currency
Dim I As Currency
Dim J As Integer
Dim SqlArt As String
Dim RsArt As ADODB.Recordset
Dim SQL2 As String
Dim Sql2Values As String

Dim cad1 As String
Dim CodTraba As String
Dim almac As String
Dim Prove As String

Dim Rs As ADODB.Recordset
Dim vSQL As String

    On Error GoTo EGenerarFacturas
    
    ' vamos a protegerlo con transacciones
    conn.BeginTrans
    
    
    'guardo el contador inicial por si falla para volver a guardarlo
    Set miRsAux = New ADODB.Recordset
    codtipom = "FAV"
    
    'valores grales para todos los socios
    porIva = CCur(DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T"))
    LetraSer = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", codtipom, "T")
    ForPago = txtcodigo(61).Text
    CodTraba = DevuelveDesdeBD(conAri, "codtraba", "straba", "login", vUsu.Login, "T")
    If CodTraba = "" Then CodTraba = DevuelveValor("select min(codtraba) from straba")
    'busco el minimo almacen y el minimo proveedor
    Sql = "select min(codalmac) from salmpr"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        almac = miRsAux.Fields(0)
    End If
        
    miRsAux.Close
        
    Sql = "select min(codprove) from sprove"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        Prove = miRsAux.Fields(0)
    End If
    
    
    Set miRsAux = Nothing
    
    PB1.visible = True
    
    Cantidad = 0
    
    Set Rs = New ADODB.Recordset
    
    vSQL = "select * from tmpinformes inner join sclien on tmpinformes.codigo1 = sclien.codclien where codusu = " & vUsu.Codigo & " order by codigo1 "
    
    total = TotalRegistrosConsulta(vSQL)
    
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = True
    
    'inicializamos cadenas
    Cad = ""
    While Not Rs.EOF
        Cantidad = Cantidad + 1
        PB1.Value = Cantidad * 100 / total
        Set vC = New CTiposMov
        Set cli = New CCliente
        Set fac = New CFactura

        If vC.TipoMovimiento <> codtipom Then
            If Not vC.Leer(codtipom) Then
                Data1.Recordset.Close
                If NumRegElim > 0 Then MsgBox "Se han generado " & NumRegElim & " factura(s) antes del error", vbExclamation
                Exit Function
            End If
        End If
        vC.IncrementarContador (vC.TipoMovimiento)
        
        totfactu = Rs!importe2
        
        BaseImp = Round(totfactu / (1 + (porIva / 100)), 2)
        
        ' insertamos en la tabla de retenciones
        SQL2 = "insert into sreten (codsocio, numeruve, fecfactu, numfactu, impreten, tiporeten, desdefec, hastafec) values "
        Sql2Values = "(" & DBSet(Rs!Codigo1, "N") & "," & DBSet(Rs!Importe1, "N") & "," & DBSet(txtcodigo(4).Text, "F") & ","
        Sql2Values = Sql2Values & DBSet(vC.Contador, "N") & "," & DBSet(Rs!importe2 * (-1), "N") & ",1," & DBSet(txtcodigo(102).Text, "F") & ","
        Sql2Values = Sql2Values & DBSet(txtcodigo(103).Text, "F") & ")"
        
        conn.Execute SQL2 & Sql2Values
            

        DoEvents
        fac.BaseImp = BaseImp
        fac.BrutoFac = BaseImp
        ImpIVA = totfactu - BaseImp
        fac.TotalFac = totfactu
        fac.codtipom = codtipom
        FecFactu = txtcodigo(4).Text
        fac.FecFactu = FecFactu
        fac.LetraSerie = LetraSer
        NumFactu = vC.Contador
        fac.NumFactu = NumFactu
'        fac.CuentaPrev = Text1(7).Text
        fac.ForPago = ForPago
        
        fac.BancoPr = txtcodigo(6)
        fac.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", fac.BancoPr, "N")
        
        fac.Agente = vParamAplic.PorDefecto_Agente
        
        'datos del cliente
        fac.Cliente = Rs!CodClien
        cli.Nombre = Rs!nomclien
        fac.NombreClien = Rs!nomclien
        cli.Domicilio = Rs!domclien
        fac.DomicilioClien = Rs!domclien
        cli.CPostal = Rs!codpobla
        fac.CPostal = Rs!codpobla
        cli.Poblacion = Rs!pobclien
        fac.Poblacion = Rs!pobclien
        cli.Provincia = Rs!proclien
        fac.Provincia = Rs!proclien
        cli.NIF = Rs!nifClien
        fac.NIF = Rs!nifClien
        
        '[Monica]22/11/2013:iban
        fac.Iban = Rs!Iban
        fac.Banco = Rs!codbanco
        fac.Sucursal = Rs!codsucur
        fac.DigControl = Rs!digcontr
        fac.CuentaBan = Rs!cuentaba
    
        'scafac
        Cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "'," & fac.Cliente & ","
        Cad = Cad & DBSet(cli.Nombre, "T") & "," & DBSet(cli.Domicilio, "T") & "," & DBSet(cli.CPostal, "T") & ","
        Cad = Cad & DBSet(cli.Poblacion, "T") & "," & DBSet(cli.Provincia, "T") & "," & DBSet(cli.NIF, "T") & "," & vParamAplic.PorDefecto_Agente
        Cad = Cad & "," & fac.ForPago & ",0,0," & TransformaComasPuntos(CStr(BaseImp)) & ",0,0," & TransformaComasPuntos(CStr(BaseImp)) & "," & iva
        Cad = Cad & "," & TransformaComasPuntos(CStr(porIva)) & "," & TransformaComasPuntos(CStr(ImpIVA)) & "," & TransformaComasPuntos(CStr(totfactu)) & ",0,NULL,"
        Cad = Cad & DBSet(Rs!codbanco, "N", "S") & "," & DBSet(Rs!codsucur, "N", "S") & "," & DBSet(Rs!digcontr, "T", "S") & "," & DBSet(Rs!cuentaba, "T", "S") & "," & DBSet(Rs!Iban, "T") & ")"
        Sql = "INSERT INTO scafac (codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,"
        Sql = Sql & "nifclien,codagent,codforpa,dtoppago,dtognral,brutofac,impdtopp,impdtogr,baseimp1,codigiv1,porciva1,"
        Sql = Sql & "imporiv1,totalfac,intconta,coddirec, codbanco, codsucur, digcontr, cuentaba, iban) VALUES ("
        Sql = Sql & Cad
        If Not ejecutar(Sql, False) Then
            vC.DevolverContador vC.TipoMovimiento, vC.Contador
            Exit Function
        Else
            'scafac1
            If cadFormula = "" Then
                cadFormula = "{scafac.numfactu}=" & NumFactu
            Else
                cadFormula = cadFormula & " or {scafac.numfactu}=" & NumFactu
            End If
            Cad = ""
            Cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0,'"
            Cad = Cad & Format(FecFactu, FormatoFecha) & "'," & vParamAplic.PorDefecto_Envio & "," & CodTraba
            Cad = Cad & "," & CodTraba & ",NULL,NULL,NULL,NULL,NULL,NULL"
    
            Sql = "INSERT INTO scafac1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,"
            Sql = Sql & "codenvio,codtraba,codtrab2,observa1,observa2,observa3,observa4,observa5,codtrab1) VALUES ("
            Sql = Sql & Cad & ")"
            conn.Execute Sql
            'slifac
            Cad = ""

            Cad = DBSet(codtipom, "T") & "," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','ALV',0,1," & almac & ","
            Cad = Cad & DBSet(vParamAplic.ArtRetenciones, "T") & "," & DBSet(NomArtic, "T") & ",1," & TransformaComasPuntos(CStr(BaseImp)) & ","
            Cad = Cad & TransformaComasPuntos(CStr(BaseImp)) & "," & TransformaComasPuntos(CStr(BaseImp)) & "," & TransformaComasPuntos(CStr(BaseImp)) & ","
            Cad = Cad & TransformaComasPuntos(CStr(BaseImp)) & ",0,0,'M'," & Prove & "," & TransformaComasPuntos(CStr(BaseImp)) & "," & ValorNulo & ",1)"
            Sql = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,"
            Sql = Sql & "numbultos,precioar,precioiv,preciomp,preciost,preciouc,dtoline1,dtoline2,origpre,codprovex,importel,ampliaci,cantidad) VALUES ("
            Sql = Sql & Cad
            conn.Execute Sql
        
        End If
        
        Set vC = Nothing
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    PB1.visible = False

EGenerarFacturas:
    If Err.Number <> 0 Or Not b Then
        GenerarFacturaRetenciones = False
        conn.RollbackTrans
        MsgBox "Error al generar facturas: " & Err.Description
    Else
        MsgBox "Proceso realizado correctamente.", vbExclamation
        GenerarFacturaRetenciones = True
        conn.CommitTrans
    End If
End Function


