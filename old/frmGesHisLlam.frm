VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGesHisLlam 
   Caption         =   "Histórico de Llamadas."
   ClientHeight    =   9630
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   1335
      Left            =   240
      TabIndex        =   89
      Top             =   7440
      Width           =   9495
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   39
         Left            =   2040
         MaxLength       =   200
         TabIndex        =   44
         Tag             =   "Observaciones 2|T|S|||shilla|observa2|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   960
         Width           =   7335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   38
         Left            =   2040
         MaxLength       =   200
         TabIndex        =   43
         Tag             =   "Observaciones Cliente|T|S|||shilla|observa1|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   600
         Width           =   7335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   37
         Left            =   2040
         MaxLength       =   60
         TabIndex        =   42
         Tag             =   "Observaciones|T|S|||shilla|observac2|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   240
         Width           =   7335
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1800
         Picture         =   "frmGesHisLlam.frx":0000
         Tag             =   "-1"
         ToolTipText     =   "Ver Observaciones"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1800
         Picture         =   "frmGesHisLlam.frx":0102
         Tag             =   "-1"
         ToolTipText     =   "Ver Observaciones"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1800
         Picture         =   "frmGesHisLlam.frx":0204
         Tag             =   "-1"
         ToolTipText     =   "Ver Observaciones"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones II:"
         Height          =   255
         Index           =   35
         Left            =   120
         TabIndex        =   92
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones Cliente:"
         Height          =   255
         Index           =   34
         Left            =   120
         TabIndex        =   91
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones:"
         Height          =   255
         Index           =   33
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "IMPORTES FACTURADOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   3135
      Left            =   5160
      TabIndex        =   72
      Top             =   4320
      Width           =   4575
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   36
         Left            =   3480
         TabIndex        =   41
         Tag             =   "Imp.Propina|N|S|||shilla|imppropi|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   35
         Left            =   3480
         TabIndex        =   40
         Tag             =   "Imp.Peaje|N|S|||shilla|imppeaje|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   34
         Left            =   3480
         TabIndex        =   39
         Tag             =   "Suplemento|N|S|||shilla|suplemen|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   33
         Left            =   3480
         TabIndex        =   38
         Tag             =   "Distancia|N|S|||shilla|distanci|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   32
         Left            =   3480
         TabIndex        =   37
         Tag             =   "Ext.Venta|N|S|||shilla|extventa|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   31
         Left            =   1320
         TabIndex        =   36
         Tag             =   "Ext.Compra|N|S|||shilla|extcompr|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   30
         Left            =   1320
         TabIndex        =   35
         Tag             =   "Imp.Venta|N|S|||shilla|impventa|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   29
         Left            =   1320
         TabIndex        =   34
         Tag             =   "Imp.Compra|N|S|||shilla|impcompr|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   28
         Left            =   1320
         TabIndex        =   33
         Tag             =   "Imp.TX|N|S|||shilla|importtx|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Propina:"
         Height          =   255
         Index           =   32
         Left            =   2400
         TabIndex        =   88
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Peaje:"
         Height          =   255
         Index           =   31
         Left            =   2400
         TabIndex        =   87
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Suplemento:"
         Height          =   255
         Index           =   30
         Left            =   2400
         TabIndex        =   86
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Distancia:"
         Height          =   255
         Index           =   29
         Left            =   2400
         TabIndex        =   85
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ext.Venta:"
         Height          =   255
         Index           =   26
         Left            =   2400
         TabIndex        =   82
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ext.Compra:"
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   81
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. Venta:"
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   80
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. Compra:"
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   79
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. TX:"
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   78
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3135
      Left            =   240
      TabIndex        =   71
      Top             =   4320
      Width           =   4815
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   27
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   32
         Tag             =   "Operador Despa.|T|S|||shilla|opedespa|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   26
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   31
         Tag             =   "Operador Reserva|T|S|||shilla|opereser|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   25
         Left            =   2670
         MaxLength       =   8
         TabIndex        =   30
         Tag             =   "Hora Final|H|S|||shilla|horfinal|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   24
         Left            =   2670
         MaxLength       =   8
         TabIndex        =   28
         Tag             =   "Hora Ocupado|H|S|||shilla|horocupa|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   23
         Left            =   2670
         MaxLength       =   8
         TabIndex        =   26
         Tag             =   "Hora Llegada|H|S|||shilla|horllega|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   22
         Left            =   2670
         MaxLength       =   8
         TabIndex        =   24
         Tag             =   "Hora Aviso|H|S|||shilla|horaviso|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   21
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "Fecha Final|F|S|||shilla|fecfinal|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   20
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "Fecha ocupado|F|S|||shilla|fecocupa|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   19
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   25
         Tag             =   "Fecha Llegada|F|S|||shilla|fecllega|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   18
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "Fecha Aviso|F|S|||shilla|fecaviso|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   17
         Left            =   2670
         MaxLength       =   8
         TabIndex        =   22
         Tag             =   "Hora Reserva|H|S|||shilla|horreser|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   11
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "Fecha Reserva|F|S|||shilla|fecreser|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Ope.Des:"
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   84
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Ope.Res:"
         Height          =   255
         Index           =   27
         Left            =   240
         TabIndex        =   83
         Top             =   2280
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1140
         Picture         =   "frmGesHisLlam.frx":0306
         ToolTipText     =   "Buscar fecha"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   1140
         Picture         =   "frmGesHisLlam.frx":0391
         ToolTipText     =   "Buscar fecha"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1140
         Picture         =   "frmGesHisLlam.frx":041C
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1140
         Picture         =   "frmGesHisLlam.frx":04A7
         ToolTipText     =   "Buscar fecha"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Finalizado:"
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   77
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Ocupado:"
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   76
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Llegada:"
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   75
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Aviso:"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   74
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Reserva:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   73
         Top             =   240
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1140
         Picture         =   "frmGesHisLlam.frx":0532
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
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
      Height          =   375
      Left            =   7200
      TabIndex        =   45
      Top             =   9000
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8520
      TabIndex        =   46
      Top             =   9000
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8520
      TabIndex        =   47
      Top             =   9000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "LOCALIZACION DEL SERVICIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   240
      TabIndex        =   49
      Top             =   600
      Width           =   9495
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   41
         Left            =   2970
         MaxLength       =   10
         TabIndex        =   95
         Tag             =   "Puerta|T|S|||shilla|puerllama|||"
         Text            =   "Text"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   40
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   94
         Tag             =   "Numero|T|S|||shilla|numllama|||"
         Text            =   "Text"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   16
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   13
         Tag             =   "Nombre|T|S|||shilla|nomclien|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   15
         Left            =   6360
         MaxLength       =   14
         TabIndex        =   12
         Tag             =   "Autorización|T|S|||shilla|codautor|||"
         Text            =   "Text"
         Top             =   1830
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   14
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   11
         Tag             =   "Usuario|T|S|||shilla|codusuar|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   13
         Left            =   3360
         MaxLength       =   6
         TabIndex        =   9
         Tag             =   "Codigo cliente|N|S|||shilla|codclien|000000||"
         Text            =   "999999"
         Top             =   1440
         Width           =   960
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   12
         Left            =   1080
         TabIndex        =   8
         Tag             =   "Telefono|T|S|||shilla|telefono|||"
         Text            =   "1234567890"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Taxitronic"
         Height          =   645
         Left            =   5100
         TabIndex        =   65
         Top             =   690
         Width           =   3735
         Begin VB.CheckBox Check1 
            Caption         =   "Validado"
            Height          =   375
            Index           =   2
            Left            =   2520
            TabIndex        =   7
            Tag             =   "Validado|N|S|||shilla|validado|||"
            Top             =   210
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Abonado"
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   6
            Tag             =   "Abonado|N|S|||shilla|abonados|||"
            Top             =   210
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Facturado"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Tag             =   "Facturado|N|S|||shilla|facturad|||"
            Top             =   210
            Width           =   1095
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "Tipo servicio|N|S|0|1|shilla|tipservi|0||"
         Top             =   2190
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4320
         TabIndex        =   62
         Text            =   "Text2"
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   10
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Matricula|T|S|||shilla|matricul|||"
         Text            =   "Text"
         Top             =   2910
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   6360
         MaxLength       =   7
         TabIndex        =   10
         Tag             =   "Identificacion|T|S|||shilla|idservic|||"
         Text            =   "Text"
         Top             =   1470
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   3360
         MaxLength       =   5
         TabIndex        =   4
         Tag             =   "Num vehiculo|N|N|||shilla|numeruve|00000|S|"
         Text            =   "Text"
         Top             =   840
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "Hora|H|N|||shilla|hora|hh:mm:ss|S|"
         Text            =   "99:99:99"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "Licencia|T|S|||shilla|licencia|||"
         Text            =   "Text"
         Top             =   2550
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   6120
         MaxLength       =   35
         TabIndex        =   20
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   18
         Tag             =   "Población|T|S|||shilla|ciudadre|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   17
         Text            =   "Text"
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   15
         Tag             =   "Domicilio|T|S|||shilla|dirllama|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha|F|N|||shilla|fecha|dd/mm/yyyy|S|"
         Text            =   "99/99/9999"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   3360
         MaxLength       =   5
         TabIndex        =   2
         Tag             =   "Codigo socio|N|N|||shilla|codsocio|00000|S|"
         Text            =   "Text"
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Nº/Puerta:"
         Height          =   255
         Index           =   36
         Left            =   240
         TabIndex        =   93
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   70
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Autorización:"
         Height          =   255
         Index           =   16
         Left            =   5160
         TabIndex        =   69
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario:"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   68
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   800
         Picture         =   "frmGesHisLlam.frx":05BD
         ToolTipText     =   "Buscar fecha"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   3000
         Picture         =   "frmGesHisLlam.frx":0648
         Tag             =   "-1"
         ToolTipText     =   "Buscar Socio"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   255
         Index           =   14
         Left            =   2400
         TabIndex        =   67
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono:"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   66
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Hora:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   64
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   3000
         Picture         =   "frmGesHisLlam.frx":074A
         Tag             =   "-1"
         ToolTipText     =   "Buscar cliente"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Matricula:"
         Height          =   255
         Index           =   12
         Left            =   5160
         TabIndex        =   61
         Top             =   2910
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Identificación:"
         Height          =   255
         Index           =   11
         Left            =   5160
         TabIndex        =   60
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Vehiculo:"
         Height          =   255
         Index           =   10
         Left            =   2400
         TabIndex        =   59
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de servicio:"
         Height          =   255
         Index           =   9
         Left            =   5160
         TabIndex        =   58
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Licencia:"
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   57
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia:"
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   56
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Población:"
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   55
         Top             =   3240
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmGesHisLlam.frx":084C
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "CP:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   54
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   53
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Socio:"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   51
         Top             =   360
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   48
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   240
      TabIndex        =   50
      Top             =   8760
      Width           =   3975
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label"
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
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Menu mnopciones 
      Caption         =   "&Opciones"
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
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnsalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmGesHisLlam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
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


Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cadB As String
Dim cad As String

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
                    PosicionarData
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
Error1:
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        PonerFoco Text1(1)
    End If
End Sub
Private Function DatosOk() As Boolean

DatosOk = False
If Modo = 4 Then
    DatosOk = True
    Exit Function
End If
'CODIGO DE socio
If Text1(0).Text = "" Then
    MsgBox "Debe introducir el código de socio.", vbExclamation
    PonerFoco Text1(0)
    Exit Function
ElseIf Not IsNumeric(Text1(0).Text) Then
        MsgBox "El código de socio debe ser numérico.", vbExclamation
        PonerFoco Text1(0)
        Exit Function
End If

'Fecha
If Text1(1).Text = "" Then
    MsgBox "Debe introducir la fecha de la llamada.", vbExclamation
    PonerFoco Text1(1)
    Exit Function
ElseIf Not IsDate(Text1(1).Text) Then
    MsgBox "La fecha no tiene formato correcto.", vbExclamation
    PonerFoco Text1(1)
    Exit Function
End If

'Hora
If Text1(7).Text = "" Then
    MsgBox "Debe introducir la hora de la llamada.", vbExclamation
    PonerFoco Text1(7)
    Exit Function
ElseIf Not IsDate(Text1(7).Text) Then
    MsgBox "La hora no tiene formato correcto.", vbExclamation
    PonerFoco Text1(7)
    Exit Function
End If

'numero de vehiculo
If Text1(8).Text = "" Then
    MsgBox "Debe introducir el número de vehiculo.", vbExclamation
    PonerFoco Text1(8)
    Exit Function
ElseIf Not IsNumeric(Text1(8).Text) Then
        MsgBox "El número de vehiculo debe ser numérico.", vbExclamation
        PonerFoco Text1(8)
        Exit Function
End If
DatosOk = True
    
End Function

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
        If Adodc1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Adodc1.Recordset.Fields(0) & "|"
        cad = cad & Adodc1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
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

Private Sub Form_Load()
    'Icono del form
    Me.Icon = frmPpal.Icon
    

    'ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 12
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        '.Buttons(9).Image = 10 'Lineas
        .Buttons(9).Image = 16 'Imprmir
        .Buttons(10).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
      
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    PrimeraVez = True
    CargarCombo
    
    '## A mano
    NombreTabla = "shilla"
    Ordenacion = " ORDER BY fecha,hora"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Adodc1.ConnectionString = conn
    Adodc1.RecordSource = "Select * from " & NombreTabla & " where numeruve=-1"
    Adodc1.Refresh
    
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
    End If
    
End Sub
Private Sub CargarCombo()
    Combo1.AddItem "Normal"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Crédito"
    Combo1.ItemData(Combo1.NewIndex) = 1
    
End Sub
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim B As Boolean
On Error GoTo EPonerModo


    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    
    Modo = Kmodo
    PonerIndicador lblIndicador(0), Modo
    
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    B = (Modo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Adodc1.Recordset.EOF Then
        If Adodc1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    '---------------------------------------------
    B = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    Combo1.Enabled = B
    Check1(0).Enabled = B
    Check1(1).Enabled = B
    Check1(2).Enabled = B
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = B
    Next i
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = B
    Next i
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    '-----------------------------
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub
Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub
Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean
Dim i As Byte

    B = (Modo = 2 Or Modo = 5 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = B
    
    B = (Modo = 2 Or Modo = 5)
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    
    
    '------------------------------------------
    B = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not B
End Sub
Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador(0).Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
        cadB = Aux
        Aux = Format(ValorDevueltoFormGrid(Text1(7), CadenaDevuelta, 2), FormatoHora)
        cadB = cadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 3)
        cadB = cadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(Text1(8), CadenaDevuelta, 4)
        cadB = cadB & " AND " & Aux
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
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
Dim Indice As Byte
Dim devuelve As String
    
    Indice = 3
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'Poblacion
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)
    'provincia
    Text1(Indice + 2).Text = devuelve

End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte
Dim Observaciones As String

Select Case Index
    Case 2, 4, 5 'observaciones
                If Index = 2 Then
                    Indice = 37
                    If Not Adodc1.Recordset.EOF Then
                        If Not IsNull(Adodc1.Recordset!observac2) Then
                            Observaciones = Adodc1.Recordset!observac2
                        Else
                            Observaciones = ""
                        End If
                    End If
                ElseIf Index = 4 Then
                    Indice = 38
                    If Not Adodc1.Recordset.EOF Then
                        If Not IsNull(Adodc1.Recordset!observa1) Then
                            Observaciones = Adodc1.Recordset!observa1
                        Else
                            Observaciones = ""
                            
                        End If
                    End If
                Else
                    Indice = 39
                    If Not Adodc1.Recordset.EOF Then
                        If Not IsNull(Adodc1.Recordset!observa2) Then
                            Observaciones = Adodc1.Recordset!observa2
                        Else
                            Observaciones = ""
                        End If
                    End If
                End If
                If Modo = 3 Or Modo = 4 Then
                    CadenaDesdeOtroForm = Text1(Indice).Text
                Else
                    CadenaDesdeOtroForm = ""
                    If Not Adodc1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Observaciones, "T")
                End If
                frmFacClienteObser.Modificar = Modo >= 3
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 3 Or Modo = 4 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(Indice).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
            
    Case 0 'población
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            If Index = 9 Then
                Indice = 4
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
Dim Indice As Byte
    Select Case Index
        Case 0
            Indice = 1
        Case 1
            Indice = 11
        Case 2
            Indice = 18
        Case 3
            Indice = 19
        Case 4
            Indice = 20
        Case 5
            Indice = 21
    End Select
    Set frmCal = New frmCal
    If Text1(Indice).Text <> "" Then PonerFormatoFecha Text1(Indice)
    frmCal.Fecha = Now
    If Text1(Indice).Text <> "" Then frmCal.Fecha = CDate(Text1(Indice).Text)
    Screen.MousePointer = vbDefault
    frmCal.Show vbModal
    If IsDate(Fecha) Then
        If Fecha <> "0:00:00 " Then Text1(Indice) = Fecha
    End If
    Set frmCal = Nothing
    PonerFoco Text1(Indice)
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
        If Text1(Index).Text <> "" Then
            Text1(Index).Text = Format(Text1(Index).Text, "00000")
            encontrado = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(Index).Text, "T")
            If encontrado = "" Then
                MsgBox "El código de socio introducido no existe.", vbExclamation
                PonerFoco Text1(Index)
            Else
                Text2(1).Text = encontrado
            End If
        End If
    Case 8
        If Text1(Index).Text <> "" Then Text1(Index).Text = Format(Text1(Index).Text, "00000")
    Case 13 'cliente
        If Text1(Index).Text <> "" Then
            Text1(Index).Text = Format(Text1(Index).Text, "000000")
            encontrado = DevuelveDesdeBD(conAri, "nomclien", "scliente", "codclien", Text1(Index).Text, "T")
            If encontrado = "" Then
                MsgBox "El código de cliente introducido no existe.", vbExclamation
                PonerFoco Text1(Index)
            Else
                Text1(16).Text = encontrado
            End If
        End If
    Case 3 'CP
        If Text1(Index) <> "" Then
            'Poblacion
            Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, encontrado)
            'provincia
            Text1(Index + 2).Text = encontrado
        End If
    Case 28 To 36
        PonerFormatoDecimal Text1(Index), 6
    Case 7, 17, 22 To 25
        If Text1(Index).Text <> "" Then PonerFormatoHora Text1(Index)
    Case 1, 11, 18 To 21
        If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
End Select
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
        Case 1  'Buscar
           mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos
            
        Case 5  'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7  'Borrar
            mnEliminar_Click
'        Case 9
'            mnLineas_Click
        Case 9  'imprimir
            
        Case 10  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
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
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    PonerFoco Text1(1)
End Sub
Private Sub mnModificar_Click()
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub
Private Sub BotonModificar()
'Prepara el Form para Modificar
Dim DeVarios As Boolean
Dim SQL As String
On Error GoTo EModificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    If Combo1.Text = "0" Then Text1(13).Enabled = False
    imgFecha(0).Enabled = False
    imgBuscar(3).Enabled = False
    If Combo1.Text = "0" Then
        imgBuscar(1).Enabled = False
    Else
        imgBuscar(1).Enabled = True
    End If
    PonerFoco Text1(12)
   
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
Dim SQL As String

On Error GoTo EEliminar

msg = "Esta seguro que desea eliminar la llamada del día:" & Text1(1).Text & "?"
If MsgBox(msg, vbYesNo) = vbYes Then
    NumRegElim = Adodc1.Recordset.AbsolutePosition
    SQL = "Delete from shilla where fecha='" & Format(Text1(1).Text, FormatoFecha) & "' and hora='" & Format(Text1(7).Text, FormatoHora)
    SQL = SQL & "' and codsocio=" & Text1(0).Text & " and numeruve=" & Text1(8).Text
    conn.Execute SQL
End If

If SituarDataTrasEliminar(Adodc1, NumRegElim) Then
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
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Adodc1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Adodc1.Recordset.EOF Then Exit Sub
    DesplazamientoData Adodc1, Index
    PonerCampos
End Sub
Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Adodc1.RecordSource = CadenaConsulta
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(0)
            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Adodc1.Recordset.MoveFirst
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

    
    If Adodc1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Adodc1
    
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
    
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador(0).Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    '##A mano
    cad = ""
    cad = cad & ParaGrid(Text1(1), 14, "Fecha")
    cad = cad & ParaGrid(Text1(7), 14, "Hora")
    cad = cad & ParaGrid(Text1(0), 14, "Socio")
    cad = cad & ParaGrid(Text1(8), 14, "Vehiculo")

    tabla = "shilla"
    Titulo = "Histórico"
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|2|3|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(fecha='" & Format(Text1(1).Text, FormatoFecha) & "' and hora='" & Format(Text1(7).Text, FormatoHora)
    cad = cad & "' and numeruve=" & Text1(8).Text & ")"
    If SituarDataMULTI(Adodc1, cad, Indicador) Then
       PonerModo 2
       lblIndicador(0).Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

