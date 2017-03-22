VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGesHisLlam 
   Caption         =   "Histórico de Llamadas."
   ClientHeight    =   10125
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10125
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   1335
      Left            =   240
      TabIndex        =   90
      Top             =   7800
      Width           =   9495
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   39
         Left            =   2040
         MaxLength       =   200
         TabIndex        =   45
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
         TabIndex        =   44
         Tag             =   "Observaciones Cliente|T|S|||shilla|observa1|||"
         Text            =   $"frmGesHisLlam.frx":0000
         Top             =   600
         Width           =   7335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   37
         Left            =   2040
         MaxLength       =   60
         TabIndex        =   43
         Tag             =   "Observaciones|T|S|||shilla|observac2|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   240
         Width           =   7335
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1800
         Picture         =   "frmGesHisLlam.frx":00CC
         Tag             =   "-1"
         ToolTipText     =   "Ver Observaciones"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1800
         Picture         =   "frmGesHisLlam.frx":01CE
         Tag             =   "-1"
         ToolTipText     =   "Ver Observaciones"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1800
         Picture         =   "frmGesHisLlam.frx":02D0
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
         TabIndex        =   93
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones Cliente:"
         Height          =   255
         Index           =   34
         Left            =   120
         TabIndex        =   92
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones:"
         Height          =   255
         Index           =   33
         Left            =   120
         TabIndex        =   91
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
      ForeColor       =   &H00972E0B&
      Height          =   3135
      Left            =   5160
      TabIndex        =   73
      Top             =   4680
      Width           =   4575
      Begin VB.CheckBox Check1 
         Caption         =   "Facturado Cliente"
         Height          =   375
         Index           =   4
         Left            =   2400
         TabIndex        =   42
         Tag             =   "Facturado Cliente|N|N|0|1|shilla|facturadocliente|||"
         Top             =   2400
         Width           =   1845
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Liquidado Socio"
         Height          =   375
         Index           =   3
         Left            =   270
         TabIndex        =   41
         Tag             =   "Liquidado Socio|N|N|0|1|shilla|liquidadosocio|||"
         Top             =   2400
         Width           =   1905
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   36
         Left            =   3390
         TabIndex        =   40
         Tag             =   "Imp.Propina|N|S|||shilla|imppropi|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1770
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   35
         Left            =   3390
         TabIndex        =   39
         Tag             =   "Imp.Peaje|N|S|||shilla|imppeaje|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1410
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   34
         Left            =   3390
         TabIndex        =   38
         Tag             =   "Suplemento|N|S|||shilla|suplemen|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1050
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   33
         Left            =   3390
         TabIndex        =   37
         Tag             =   "Distancia|N|S|||shilla|distanci|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   690
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   32
         Left            =   3390
         TabIndex        =   36
         Tag             =   "Ext.Venta|N|S|||shilla|extventa|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   330
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   31
         Left            =   1320
         TabIndex        =   35
         Tag             =   "Ext.Compra|N|S|||shilla|extcompr|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1410
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   30
         Left            =   1320
         TabIndex        =   34
         Tag             =   "Imp.Venta|N|S|||shilla|impventa|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1050
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   29
         Left            =   1320
         TabIndex        =   33
         Tag             =   "Imp.Compra|N|S|||shilla|impcompr|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   690
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   28
         Left            =   1320
         TabIndex        =   32
         Tag             =   "Imp.TX|N|S|||shilla|importtx|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Propina:"
         Height          =   255
         Index           =   32
         Left            =   2400
         TabIndex        =   89
         Top             =   1770
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Peaje:"
         Height          =   255
         Index           =   31
         Left            =   2400
         TabIndex        =   88
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Suplemento:"
         Height          =   255
         Index           =   30
         Left            =   2400
         TabIndex        =   87
         Top             =   1050
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Distancia:"
         Height          =   255
         Index           =   29
         Left            =   2400
         TabIndex        =   86
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ext.Venta:"
         Height          =   255
         Index           =   26
         Left            =   2400
         TabIndex        =   83
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ext.Compra:"
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   82
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. Venta:"
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   81
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. Compra:"
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   80
         Top             =   690
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. TX:"
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   79
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3135
      Left            =   240
      TabIndex        =   72
      Top             =   4680
      Width           =   4875
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   27
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   31
         Tag             =   "Operador Despa.|T|S|||shilla|opedespa|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2640
         Width           =   3675
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   26
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   30
         Tag             =   "Operador Reserva|T|S|||shilla|opereser|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2280
         Width           =   3675
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   25
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   29
         Tag             =   "Hora Final|H|S|||shilla|horfinal|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   1770
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   24
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   27
         Tag             =   "Hora Ocupado|H|S|||shilla|horocupa|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   1410
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   23
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   25
         Tag             =   "Hora Llegada|H|S|||shilla|horllega|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   22
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   23
         Tag             =   "Hora Aviso|H|S|||shilla|horaviso|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   690
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   21
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   28
         Tag             =   "Fecha Final|F|S|||shilla|fecfinal|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   1770
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   20
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   26
         Tag             =   "Fecha ocupado|F|S|||shilla|fecocupa|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   1410
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   19
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   24
         Tag             =   "Fecha Llegada|F|S|||shilla|fecllega|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   1050
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   18
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   22
         Tag             =   "Fecha Aviso|F|S|||shilla|fecaviso|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   17
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   21
         Tag             =   "Hora Reserva|H|S|||shilla|horreser|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   330
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   11
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "Fecha Reserva|F|S|||shilla|fecreser|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Ope.Des:"
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   85
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Ope.Res:"
         Height          =   255
         Index           =   27
         Left            =   240
         TabIndex        =   84
         Top             =   2280
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1080
         Picture         =   "frmGesHisLlam.frx":03D2
         ToolTipText     =   "Buscar fecha"
         Top             =   1770
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   1080
         Picture         =   "frmGesHisLlam.frx":045D
         ToolTipText     =   "Buscar fecha"
         Top             =   1410
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1080
         Picture         =   "frmGesHisLlam.frx":04E8
         ToolTipText     =   "Buscar fecha"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1080
         Picture         =   "frmGesHisLlam.frx":0573
         ToolTipText     =   "Buscar fecha"
         Top             =   690
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Finalizado:"
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   78
         Top             =   1770
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Ocupado:"
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   77
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Llegada:"
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   76
         Top             =   1050
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Aviso:"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   75
         Top             =   690
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Reserva:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   74
         Top             =   330
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmGesHisLlam.frx":05FE
         ToolTipText     =   "Buscar fecha"
         Top             =   330
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
      TabIndex        =   46
      Top             =   9420
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8520
      TabIndex        =   47
      Top             =   9420
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8520
      TabIndex        =   48
      Top             =   9420
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
      ForeColor       =   &H00972E0B&
      Height          =   3975
      Left            =   240
      TabIndex        =   50
      Top             =   600
      Width           =   9495
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   42
         Left            =   1050
         MaxLength       =   80
         TabIndex        =   99
         Tag             =   "Destino|T|S|||shilla|destino|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   3600
         Width           =   3705
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   41
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Puerta|T|S|||shilla|puerllama|||"
         Text            =   "ABCDEFGHIJ"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   40
         Left            =   1050
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Numero|T|S|||shilla|numllama|||"
         Text            =   "ABCDEFGHIJ"
         Top             =   2520
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   16
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "Nombre|T|S|||shilla|nomclien|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   15
         Left            =   6360
         MaxLength       =   14
         TabIndex        =   16
         Tag             =   "Autorización|T|S|||shilla|codautor|||"
         Text            =   "Text"
         Top             =   1470
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   14
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Usuario|T|S|||shilla|codusuar|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   13
         Left            =   3840
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Codigo cliente|N|S|||shilla|codclien|000000||"
         Text            =   "999999"
         Top             =   1080
         Width           =   960
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   12
         Left            =   1050
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Telefono|T|S|||shilla|telefono|||"
         Text            =   "1234567890"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Taxitronic"
         ForeColor       =   &H00972E0B&
         Height          =   645
         Left            =   5130
         TabIndex        =   66
         Top             =   3000
         Width           =   4185
         Begin VB.CheckBox Check1 
            Caption         =   "Facturado"
            Height          =   375
            Index           =   0
            Left            =   210
            TabIndex        =   97
            Tag             =   "Facturado|N|S|||shilla|facturad|||"
            Top             =   210
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Abonado"
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   96
            Tag             =   "Abonado|N|S|||shilla|abonados|||"
            Top             =   210
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Validado"
            Height          =   375
            Index           =   2
            Left            =   2910
            TabIndex        =   95
            Tag             =   "Validado|N|S|||shilla|validado|||"
            Top             =   210
            Width           =   1095
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "Tipo servicio|N|S|0|1|shilla|tipservi|0||"
         Top             =   1830
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4860
         TabIndex        =   63
         Text            =   "Text2"
         Top             =   510
         Width           =   4305
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   10
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Matricula|T|S|||shilla|matricul|||"
         Text            =   "Text"
         Top             =   2550
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   6360
         MaxLength       =   6
         TabIndex        =   15
         Tag             =   "Identificacion|T|S|||shilla|idservic|||"
         Text            =   "Text"
         Top             =   1110
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   2250
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "Num vehiculo|N|N|||shilla|numeruve|0000|S|"
         Text            =   "Text"
         Top             =   510
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   1410
         MaxLength       =   8
         TabIndex        =   2
         Tag             =   "Hora|H|N|||shilla|hora|hh:mm:ss|S|"
         Text            =   "99:99:99"
         Top             =   510
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Licencia|T|S|||shilla|licencia|||"
         Text            =   "Text"
         Top             =   2190
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   1050
         MaxLength       =   35
         TabIndex        =   14
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   13
         Tag             =   "Población|T|S|||shilla|ciudadre|||"
         Text            =   "ABCDEFGHIJKLMNÑOPQRSTUVWXYZABC"
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   3960
         MaxLength       =   6
         TabIndex        =   12
         Text            =   "HHHHHH"
         Top             =   2520
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "Domicilio|T|S|||shilla|dirllama|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   210
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha|F|N|||shilla|fecha|dd/mm/yyyy|S|"
         Text            =   "99/99/9999"
         Top             =   510
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   3870
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Codigo socio|N|N|||shilla|codsocio|00000||"
         Text            =   "Text"
         Top             =   510
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Destino:"
         Height          =   255
         Index           =   37
         Left            =   210
         TabIndex        =   98
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº/Puerta:"
         Height          =   255
         Index           =   36
         Left            =   210
         TabIndex        =   94
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   17
         Left            =   210
         TabIndex        =   71
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Autorización:"
         Height          =   255
         Index           =   16
         Left            =   5160
         TabIndex        =   70
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario:"
         Height          =   255
         Index           =   15
         Left            =   210
         TabIndex        =   69
         Top             =   1440
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1020
         Picture         =   "frmGesHisLlam.frx":0689
         ToolTipText     =   "Buscar fecha"
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   4500
         Picture         =   "frmGesHisLlam.frx":0714
         Tag             =   "-1"
         ToolTipText     =   "Buscar Socio"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   255
         Index           =   14
         Left            =   2850
         TabIndex        =   68
         Top             =   1110
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono:"
         Height          =   255
         Index           =   13
         Left            =   210
         TabIndex        =   67
         Top             =   1110
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Hora:"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   65
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   64
         Top             =   285
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   3510
         Picture         =   "frmGesHisLlam.frx":0816
         Tag             =   "-1"
         ToolTipText     =   "Buscar cliente"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Matricula:"
         Height          =   255
         Index           =   12
         Left            =   5160
         TabIndex        =   62
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Identificación:"
         Height          =   255
         Index           =   11
         Left            =   5160
         TabIndex        =   61
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Vehículo:"
         Height          =   255
         Index           =   10
         Left            =   2250
         TabIndex        =   60
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de servicio:"
         Height          =   255
         Index           =   9
         Left            =   5160
         TabIndex        =   59
         Top             =   1830
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Licencia:"
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   58
         Top             =   2190
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia:"
         Height          =   255
         Index           =   7
         Left            =   210
         TabIndex        =   57
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Población:"
         Height          =   255
         Index           =   6
         Left            =   210
         TabIndex        =   56
         Top             =   2880
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   3690
         Picture         =   "frmGesHisLlam.frx":0918
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   2550
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "CP:"
         Height          =   255
         Index           =   5
         Left            =   3360
         TabIndex        =   55
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio:"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   54
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Socio:"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   52
         Top             =   285
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
         TabIndex        =   49
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   240
      TabIndex        =   51
      Top             =   9210
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
         Left            =   120
         TabIndex        =   53
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

Public FechaServ As String
Public HoraServ As String
Public NumerUve As String

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

'variables control del log
Dim Socio As Currency
Dim Tfno As String
Dim clien As String
Dim Ident As String
Dim Autor As String
Dim TipoSer As Byte
Dim Licencia As String
Dim Matricula As String
Dim TxFactu As Byte
Dim TxAbo As Byte
Dim TxVali As Byte
Dim ImpCompra As Currency
Dim ImpVta As Currency
Dim LiqSoc As Byte
Dim FacCli As Byte
Dim Usuario As String
Dim Nombre As String
Dim Domicilio As String
Dim Destino As String
Dim numero As String
Dim Puerta As String
Dim Poblacion As String
Dim Obser As String
Dim Obs1 As String
Dim Obs2 As String


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
Private BuscaChekc As String


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
Dim i As Long
Dim cadB As String
Dim cad As String
Dim Indicador As String
Dim cad1 As String

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
                     '[Monica]04/02/2015: guardo en el slog los campos que me han cambiado
                    cad1 = ""
                    'Socio
                    If Text1(0).Text <> Socio Then cad1 = cad1 & "Socio: " & Format(Socio, "000000") & " a " & Text1(0).Text & ";"
                    'Liquidado socio
                    If Check1(3).Value <> LiqSoc Then cad1 = cad1 & "LiqSoc: " & LiqSoc & " a " & Check1(3).Value & ";"
                    'Facturado Cliente
                    If Check1(4).Value <> FacCli Then cad1 = cad1 & "FacCli: " & FacCli & " a " & Check1(4).Value & ";"
                    'Taxitronic facturado
                    If Check1(0).Value <> TxFactu Then cad1 = cad1 & "TxFra: " & TxFactu & " a " & Check1(0).Value & ";"
                    'Taxitronic abonado
                    If Check1(1).Value <> TxAbo Then cad1 = cad1 & "TxAbo: " & TxAbo & " a " & Check1(1).Value & ";"
                    'Taxitronic validado
                    If Check1(2).Value <> TxVali Then cad1 = cad1 & "TxVal: " & TxVali & " a " & Check1(2).Value & ";"
                    'Tfno
                    If Trim(Text1(12).Text) <> Trim(Tfno) Then cad1 = cad1 & "Tfno: " & Trim(Tfno) & " a " & Trim(Text1(12).Text) & ";"
                    'Codclien
                    If Text1(13).Text <> clien Then cad1 = cad1 & "Cli: " & clien & " a " & Text1(13).Text & ";"
                    'Usuario
                    If Trim(Text1(14).Text) <> Trim(Usuario) Then cad1 = cad1 & "Usu: " & Trim(Usuario) & " a " & Trim(Text1(14).Text) & ";"
                    'Identificacion
                    If Trim(Text1(9).Text) <> Trim(Ident) Then cad1 = cad1 & "Iden: " & Trim(Ident) & " a " & Trim(Text1(9).Text) & ";"
                    'Autorizacion
                    If Trim(Text1(15).Text) <> Trim(Autor) Then cad1 = cad1 & "Autor: " & Trim(Autor) & " a " & Trim(Text1(15).Text) & ";"
                    'Tipo de Servicio
                    If Combo1.ListIndex <> TipoSer Then cad1 = cad1 & "TSer: " & TipoSer & " a " & Combo1.ListIndex & ";"
                    'Licencia
                    If Trim(Text1(6).Text) <> Trim(Licencia) Then cad1 = cad1 & "Lic: " & Trim(Licencia) & " a " & Trim(Text1(6).Text) & ";"
                    'Matricula
                    If Trim(Text1(10).Text) <> Trim(Matricula) Then cad1 = cad1 & "Mat: " & Trim(Matricula) & " a " & Trim(Text1(10).Text) & ";"
                    'Importe de Compra
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(29).Text))) <> ImpCompra Then cad1 = cad1 & "ImpCom: " & ImpCompra & " a " & Text1(29).Text & ";"
                    'Importe de Venta
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(30).Text))) <> ImpVta Then cad1 = cad1 & "ImpVta: " & ImpVta & " a " & Text1(30).Text & ";"
                    'Nombre
                    If Trim(Text1(16).Text) <> Trim(Nombre) Then cad1 = cad1 & "Nombre: " & Trim(Nombre) & " a " & Trim(Text1(16).Text) & ";"
                    'Domicilio
                    If Trim(Text1(2).Text) <> Trim(Domicilio) Then cad1 = cad1 & "Dom: " & Trim(Domicilio) & " a " & Trim(Text1(2).Text) & ";"
                    'Numero
                    If Trim(Text1(40).Text) <> Trim(numero) Then cad1 = cad1 & "Nro: " & Trim(numero) & " a " & Trim(Text1(40).Text) & ";"
                    'Puerta
                    If Trim(Text1(41).Text) <> Trim(Puerta) Then cad1 = cad1 & "Pta: " & Trim(Puerta) & " a " & Trim(Text1(41).Text) & ";"
                    'Poblacion
                    If Trim(Text1(4).Text) <> Trim(Poblacion) Then cad1 = cad1 & "Pob: " & Trim(Poblacion) & " a " & Trim(Text1(4).Text) & ";"
                    'Destino
                    If Trim(Text1(42).Text) <> Trim(Destino) Then cad1 = cad1 & "Des: " & Trim(Destino) & " a " & Trim(Text1(42).Text) & ";"
                    
                    'Observaciones
                    If Trim(Text1(37).Text) <> Trim(Obser) Then cad1 = cad1 & "Obs: " & Trim(Obser) & " a " & Trim(Text1(37).Text) & ";"
                    'Observaciones 1
                    If Trim(Text1(38).Text) <> Trim(Obs1) Then cad1 = cad1 & "Obs1: " & Trim(Obs1) & " a " & Trim(Text1(38).Text) & ";"
                    'Observaciones 2
                    If Trim(Text1(39).Text) <> Trim(Obs2) Then cad1 = cad1 & "Obs2: " & Trim(Obs2) & " a " & Trim(Text1(39).Text) & ";"
                    
                    
                    
                                       
                    Set LOG = New cLOG
                    LOG.Insertar 9, vUsu, "Llamada modificada: " & Text1(1).Text & " " & Text1(7).Text & " " & Text1(8).Text & " " & cad1 & vbCrLf
                    Set LOG = Nothing
                
                    TerminaBloquear
                    
                    '[Monica] 02/06/2011: tras modificar volvemos al registro correspondiente
                    NumRegElim = Adodc1.Recordset.AbsolutePosition
                    Me.Adodc1.Refresh
                    If SituarDataPosicion(Adodc1, NumRegElim, "") Then
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
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    cadB1 = ObtenerBusqueda(Me, True, BuscaChekc)
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

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If PrimeraVez Then
        PrimeraVez = False
        If FechaServ <> "" Then
            If Me.Adodc1.Recordset.EOF Then
                PonerCadenaBusqueda
            Else
                PonerCampos
            End If
        End If
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
    
    NombreTabla = "shilla"
    CadenaConsulta = "Select * from " & NombreTabla
    
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
    Adodc1.ConnectionString = conn
    Adodc1.RecordSource = CadenaConsulta ' "Select * from " & NombreTabla & " where numeruve=-1"
    Adodc1.Refresh
    
    If FechaServ = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
        End If
    Else
        If Adodc1.Recordset.EOF Then
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
Dim i As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo

    For i = 0 To Text1.Count - 1
        Text1(i).BackColor = vbWhite
    Next i

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    BuscaChekc = ""
    
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
    If Not Adodc1.Recordset.EOF Then
        If Adodc1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    BloquearCmb Combo1, (Modo = 0 Or Modo = 2)
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    Combo1.Enabled = b
    For i = 0 To 2
        Check1(i).Enabled = b
    Next i
    
'[Monica]04/02/2015: dejamos modificar liquidado y facturado, antes solo podian consultar
    For i = 3 To 4
        Check1(i).Enabled = b '(Modo = 1)
    Next i
    
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    
    ' No hay icono para las observaciones de 60 de longitud maxima
    Me.imgBuscar(2).Enabled = False
    Me.imgBuscar(2).visible = False
    
    Me.imgBuscar(4).Enabled = (Modo > 0)
    Me.imgBuscar(4).visible = (Modo > 0)
    Me.imgBuscar(5).Enabled = (Modo > 0)
    Me.imgBuscar(5).visible = (Modo > 0)
    
    
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = b
    Next i
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    '-----------------------------
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
Dim i As Byte

    b = (Modo = 2 Or Modo = 5 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    
    b = (Modo = 2 Or Modo = 5)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    'imprimir
    Toolbar1.Buttons(9).Enabled = b
    
    
    
    
    '------------------------------------------
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
End Sub
Private Sub LimpiarCampos()
Dim i As Integer

On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    Me.Combo1.ListIndex = -1
    
    For i = 0 To Check1.Count - 1
        Check1(i).Value = 0
    Next i
    
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
Dim indice As Byte
Dim devuelve As String
    
    indice = 3
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'Poblacion
    Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve)
    'provincia
    Text1(indice + 2).Text = devuelve

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
            CadenaDesdeOtroForm = Text1(38).Text
        Else
            CadenaDesdeOtroForm = ""
            If Not Adodc1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Adodc1.Recordset!observa1, "T")
        End If
        frmFacClienteObser.Modificar = Modo >= 3
        frmFacClienteObser.Text1 = CadenaDesdeOtroForm
        frmFacClienteObser.Show vbModal
        'Llevara DOS VALORES.
        'Si modifica y el texto
        If Modo = 3 Or Modo = 4 Then
            If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(38).Text = Mid(CadenaDesdeOtroForm, 3)
        End If
        CadenaDesdeOtroForm = ""

    
    Case 5 ' observaciones
        If Modo = 3 Or Modo = 4 Then
            CadenaDesdeOtroForm = Text1(39).Text
        Else
            CadenaDesdeOtroForm = ""
            If Not Adodc1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Adodc1.Recordset!observa2, "T")
        End If
        frmFacClienteObser.Modificar = Modo >= 3
        frmFacClienteObser.Text1 = CadenaDesdeOtroForm
        frmFacClienteObser.Show vbModal
        'Llevara DOS VALORES.
        'Si modifica y el texto
        If Modo = 3 Or Modo = 4 Then
            If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(39).Text = Mid(CadenaDesdeOtroForm, 3)
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
        If Text1(Index).Text <> "" Then Text1(Index).Text = Format(Text1(Index).Text, "0000")
    Case 13 'cliente
        If Modo = 1 Then Exit Sub
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
            printNou
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
Dim Sql As String
On Error GoTo EModificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    '[Monica]04/02/2015: cargamos las variables para el log
    Socio = Text1(0).Text
    Tfno = Text1(12).Text
    clien = Text1(13).Text
    Ident = Text1(9).Text
    Autor = Text1(15).Text
    TipoSer = Combo1.ListIndex
    Licencia = Text1(6).Text
    Matricula = Text1(10).Text
    TxFactu = Check1(0).Value
    TxAbo = Check1(1).Value
    TxVali = Check1(2).Value
    ImpCompra = Text1(29).Text
    ImpVta = Text1(30).Text
    LiqSoc = Check1(3).Value
    FacCli = Check1(4).Value
    Usuario = Text1(14).Text
    Nombre = Text1(16).Text
    Domicilio = Text1(2).Text
    Destino = Text1(42).Text
    numero = Text1(40).Text
    Puerta = Text1(41).Text
    Poblacion = Text1(4).Text
    Obser = Text1(37).Text
    Obs1 = Text1(38).Text
    Obs2 = Text1(39).Text
    
    
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
    NumRegElim = Adodc1.Recordset.AbsolutePosition
    Sql = "Delete from shilla where fecha='" & Format(Text1(1).Text, FormatoFecha) & "' and hora='" & Format(Text1(7).Text, FormatoHora)
    Sql = Sql & "' and codsocio=" & Text1(0).Text & " and numeruve=" & Text1(8).Text
    conn.Execute Sql
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
    lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    '##A mano
    cad = ""
    cad = cad & ParaGrid(Text1(1), 14, "Fecha")
    cad = cad & ParaGrid(Text1(7), 14, "Hora")
    cad = cad & ParaGrid(Text1(0), 14, "Socio")
    cad = cad & ParaGrid(Text1(8), 14, "Vehiculo")

    Tabla = "shilla"
    Titulo = "Histórico"
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
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

    If Not Adodc1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Me.Adodc1, vWhere, Indicador) Then
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
    Sql = Sql & "' and numeruve=" & Text1(8).Text & ")"
    
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
        .cadRegActua = POS2SF(Adodc1, Me)
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
Dim cad As String
Dim RS As ADODB.Recordset

    On Error Resume Next

    cad = ""
'    If Me.DesdeFichaCliente Then
        '
    cad = " WHERE fecha=" & DBSet(FechaServ, "F") & " AND hora= " & DBSet(HoraServ, "H") & " AND numeruve=" & DBSet(NumerUve, "N")
        
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
    ObtenerSelFactura = cad
End Function

