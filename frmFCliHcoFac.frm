VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFCliHcoFac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas de Servicios de Clientes"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   14460
   Icon            =   "frmFCliHcoFac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFCliHcoFac.frx":000C
   ScaleHeight     =   6840
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   9
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   143
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6440
      Visible         =   0   'False
      Width           =   6045
   End
   Begin VB.Frame Frame2 
      Height          =   710
      Left            =   120
      TabIndex        =   129
      Top             =   400
      Width           =   14175
      Begin VB.CheckBox Check1 
         Caption         =   "FacturaE"
         Height          =   375
         Index           =   1
         Left            =   12690
         TabIndex        =   6
         Tag             =   "FacturaE|N|N|0|1|scafaccli|exportada||N|"
         Top             =   210
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   176
         Top             =   300
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   8010
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Nombre Cliente|T|N|||scafaccli|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   240
         Width           =   4350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   7125
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Cod. Cliente|N|N|0|999999|scafaccli|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1290
         TabIndex        =   1
         Tag             =   "Tipo Factura|T|N|||scafaccli|codtipom||S|"
         Text            =   "Text3"
         Top             =   315
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   2670
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||scafaccli|fecfactu|dd/mm/yyyy|S|"
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Factura|N|N|||scafaccli|numfactu|0000000|S|"
         Text            =   "Text1 7"
         Top             =   315
         Width           =   980
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   3
         Tag             =   "Contabilizado|N|N|0|1|scafaccli|intconta||N|"
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   133
         Top             =   240
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6855
         ToolTipText     =   "Buscar cliente"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fact."
         Height          =   255
         Index           =   29
         Left            =   2670
         TabIndex        =   132
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   131
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Fact."
         Height          =   255
         Index           =   27
         Left            =   1320
         TabIndex        =   130
         Top             =   120
         Width           =   795
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   10080
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9240
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4800
      Left            =   120
      TabIndex        =   59
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1095
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   8467
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmFCliHcoFac.frx":0A0E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(30)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(11)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(25)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(26)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1(23)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1(17)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1(16)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FrameCliente"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FrameFactura"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmFCliHcoFac.frx":0A2A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameObserva"
      Tab(1).Control(1)=   "cmdaux"
      Tab(1).Control(2)=   "txtAux3(2)"
      Tab(1).Control(3)=   "txtAux3(1)"
      Tab(1).Control(4)=   "txtAux3(0)"
      Tab(1).Control(5)=   "cmdObserva"
      Tab(1).Control(6)=   "Text2(0)"
      Tab(1).Control(7)=   "Text3(0)"
      Tab(1).Control(8)=   "DataGrid2"
      Tab(1).Control(9)=   "txtAux(11)"
      Tab(1).Control(10)=   "txtAux(10)"
      Tab(1).Control(11)=   "txtAux(9)"
      Tab(1).Control(12)=   "txtAux(5)"
      Tab(1).Control(13)=   "txtAux(3)"
      Tab(1).Control(14)=   "txtAux(1)"
      Tab(1).Control(15)=   "txtAux(0)"
      Tab(1).Control(16)=   "txtAux(4)"
      Tab(1).Control(17)=   "txtAux(6)"
      Tab(1).Control(18)=   "txtAux(7)"
      Tab(1).Control(19)=   "txtAux(8)"
      Tab(1).Control(20)=   "txtAux(2)"
      Tab(1).Control(21)=   "DataGrid1"
      Tab(1).Control(22)=   "Text3(6)"
      Tab(1).Control(23)=   "Text3(1)"
      Tab(1).Control(24)=   "Text2(3)"
      Tab(1).Control(25)=   "Text3(3)"
      Tab(1).Control(26)=   "Text2(1)"
      Tab(1).Control(27)=   "Text2(2)"
      Tab(1).Control(28)=   "Text3(2)"
      Tab(1).Control(29)=   "Text3(8)"
      Tab(1).Control(30)=   "Text3(7)"
      Tab(1).Control(31)=   "Text3(5)"
      Tab(1).Control(32)=   "Text3(4)"
      Tab(1).Control(33)=   "Text3(14)"
      Tab(1).Control(34)=   "Text3(15)"
      Tab(1).Control(35)=   "imgBuscar(6)"
      Tab(1).Control(36)=   "imgBuscar(9)"
      Tab(1).Control(37)=   "imgBuscar(8)"
      Tab(1).Control(38)=   "Label1(40)"
      Tab(1).Control(39)=   "Label1(22)"
      Tab(1).Control(40)=   "Label1(18)"
      Tab(1).Control(41)=   "Label1(6)"
      Tab(1).Control(42)=   "Label1(2)"
      Tab(1).Control(43)=   "Label1(21)"
      Tab(1).Control(44)=   "Label1(24)"
      Tab(1).Control(45)=   "Label1(23)"
      Tab(1).Control(46)=   "Label1(9)"
      Tab(1).Control(47)=   "imgBuscar(7)"
      Tab(1).ControlCount=   48
      TabCaption(2)   =   "Servicios"
      TabPicture(2)   =   "frmFCliHcoFac.frx":0A46
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtAux2(1)"
      Tab(2).Control(1)=   "txtAux2(0)"
      Tab(2).Control(2)=   "txtAux4(13)"
      Tab(2).Control(3)=   "txtAux4(5)"
      Tab(2).Control(4)=   "txtAux4(6)"
      Tab(2).Control(5)=   "txtAux4(7)"
      Tab(2).Control(6)=   "txtAux4(8)"
      Tab(2).Control(7)=   "txtAux4(9)"
      Tab(2).Control(8)=   "txtAux4(10)"
      Tab(2).Control(9)=   "txtAux4(11)"
      Tab(2).Control(10)=   "txtAux4(2)"
      Tab(2).Control(11)=   "txtAux4(1)"
      Tab(2).Control(12)=   "txtAux4(0)"
      Tab(2).Control(13)=   "txtAux4(3)"
      Tab(2).Control(14)=   "txtAux4(4)"
      Tab(2).Control(15)=   "txtAux4(12)"
      Tab(2).Control(16)=   "DataGrid3"
      Tab(2).Control(17)=   "Data4"
      Tab(2).Control(18)=   "Label1(50)"
      Tab(2).Control(19)=   "Label1(48)"
      Tab(2).ControlCount=   20
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   345
         Index           =   1
         Left            =   -68670
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   174
         Top             =   4320
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   -74730
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   169
         Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
         Top             =   4320
         Visible         =   0   'False
         Width           =   5865
      End
      Begin VB.Frame FrameObserva 
         Caption         =   "Observaciones"
         ForeColor       =   &H00972E0B&
         Height          =   1815
         Left            =   -71070
         TabIndex        =   105
         Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
         Top             =   720
         Width           =   10005
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   13
            Left            =   480
            MaxLength       =   80
            TabIndex        =   110
            Tag             =   "Observación 5|T|S|||scafaccli1|observa5||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1440
            Width           =   9180
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   12
            Left            =   480
            MaxLength       =   80
            TabIndex        =   109
            Tag             =   "Observación 4|T|S|||scafaccli1|observa4||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1140
            Width           =   9180
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   11
            Left            =   480
            MaxLength       =   80
            TabIndex        =   108
            Tag             =   "Observación 3|T|S|||scafaccli1|observa3||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   840
            Width           =   9180
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   10
            Left            =   480
            MaxLength       =   80
            TabIndex        =   107
            Tag             =   "Observación 2|T|S|||scafaccli1|observa2||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   540
            Width           =   9180
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   9
            Left            =   480
            MaxLength       =   80
            TabIndex        =   106
            Tag             =   "Observación 1|T|S|||scafaccli1|observa1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   240
            Width           =   9180
         End
      End
      Begin VB.CommandButton cmdaux 
         Caption         =   "+"
         Height          =   320
         Left            =   -65520
         TabIndex        =   123
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   13
         Left            =   -63300
         MaxLength       =   15
         TabIndex        =   160
         Tag             =   "Nombre socio |T|N|||scafaccli_serv|codsocio|||"
         Text            =   "socio"
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   -70410
         MaxLength       =   15
         TabIndex        =   159
         Tag             =   "Socio |N|N|||scafaccli_serv|codsocio|0000||"
         Text            =   "socio"
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   -69450
         MaxLength       =   15
         TabIndex        =   158
         Tag             =   "Uve|N|N|||scafaccli_serv|numeruve|0000|S|"
         Text            =   "Uve"
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   -68520
         MaxLength       =   15
         TabIndex        =   157
         Tag             =   "Direccion|T|S|||scafaccli_serv|dirllama|||"
         Text            =   "direccion"
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   -67560
         MaxLength       =   15
         TabIndex        =   156
         Tag             =   "Numllama|T|S|||scafaccli_serv|numllama|||"
         Text            =   "numllama"
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   -66660
         MaxLength       =   15
         TabIndex        =   155
         Tag             =   "Puerllama|T|S|||scafaccli_serv|puerllama|||"
         Text            =   "puerllama"
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   10
         Left            =   -65730
         MaxLength       =   15
         TabIndex        =   154
         Tag             =   "Identificacion|T|S|||scafaccli_serv|idservic|||"
         Text            =   "ciudad"
         Top             =   2040
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   11
         Left            =   -65010
         MaxLength       =   15
         TabIndex        =   153
         Tag             =   "Linea |T|S|||scafaccli_serv|telefono|||"
         Text            =   "tfno"
         Top             =   2040
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtAux4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -73290
         MaxLength       =   30
         TabIndex        =   152
         Tag             =   "Fecha Factura|F|N|||scafaccli_serv|fecfactu|dd/mm/yyyy|S|"
         Text            =   "fecfactu"
         Top             =   2040
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtAux4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -74220
         MaxLength       =   15
         TabIndex        =   151
         Tag             =   "NºFactura |N|N|||scafaccli_serv|numfactu|0000000|S|"
         Text            =   "numfactu"
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74940
         MaxLength       =   7
         TabIndex        =   150
         Tag             =   "Tipo Movimiento|T|N|||scafaccli_serv|codtipom||S|"
         Text            =   "codtipoa"
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   -72600
         MaxLength       =   15
         TabIndex        =   149
         Tag             =   "Fecha|F|N|||scafaccli_serv|fecha|dd/mm/yyyy|N|"
         Text            =   "fecha"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   -71370
         MaxLength       =   15
         TabIndex        =   148
         Tag             =   "Linea |H|N|||scafaccli_serv|hora|hh:mm:ss||"
         Text            =   "hora"
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   12
         Left            =   -64260
         MaxLength       =   15
         TabIndex        =   147
         Tag             =   "Importe |N|N|||scafaccli_serv|impventa|###,##0.00|N|"
         Text            =   "importe"
         Top             =   2040
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   128
         Tag             =   "Fecha Albaran|F|N|||scafaccli1|fechaalb|dd/mm/yyyy|N|"
         Text            =   "fecalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73920
         MaxLength       =   15
         TabIndex        =   127
         Tag             =   "Nº Albaran|N|N|||scafaccli1|numalbar|0000000|N|"
         Text            =   "numalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74640
         MaxLength       =   7
         TabIndex        =   126
         Tag             =   "Tipo Albaran|T|N|||scafaccli1|codtipoa||N|"
         Text            =   "codtipoa"
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdObserva 
         Height          =   375
         Left            =   -71040
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   870
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   -68070
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   93
         Text            =   "Text2"
         Top             =   390
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   -68790
         MaxLength       =   30
         TabIndex        =   50
         Tag             =   "Trabajador Albaran|N|N|0|9999|scafaccli1|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   390
         Width           =   660
      End
      Begin VB.Frame FrameFactura 
         Height          =   1980
         Left            =   150
         TabIndex        =   82
         Top             =   2730
         Width           =   12975
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   2070
            MaxLength       =   15
            TabIndex        =   26
            Tag             =   "Imp. Dto Gn|N|N|||scafaccli|impdtogr|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   330
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   5550
            MaxLength       =   15
            TabIndex        =   28
            Tag             =   "Imp.Suplidos|N|S|||scafaccli|suplidos|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   345
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   45
            Tag             =   "% IVA 1|N|S|0|99.90|scafaccli|porciva3re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   43
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   46
            Tag             =   "Importe IVA 1|N|S|||scafaccli|imporiv3re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   39
            Tag             =   "% IVA 1|N|S|0|99.90|scafaccli|porciva2re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   41
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   40
            Tag             =   "Importe IVA 1|N|S|||scafaccli|imporiv2re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   33
            Tag             =   "% IVA 1|N|S|0|99.90|scafaccli|porciva1re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   39
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   34
            Tag             =   "Importe IVA 1|N|S|||scafaccli|imporiv1re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   38
            Left            =   9720
            MaxLength       =   15
            TabIndex        =   47
            Tag             =   "Total Factura|N|N|||scafaccli|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   37
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   44
            Tag             =   "Importe IVA 3|N|S|||scafaccli|imporiv3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   31
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   43
            Tag             =   "% IVA 3|N|S|0|99.90|scafaccli|porciva3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   28
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   41
            Tag             =   "Cod. IVA 3|N|S|0|9999|scafaccli|codigiv3|0000|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   42
            Tag             =   "Base Imponible 3|N|S|||scafaccli|baseimp3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   38
            Tag             =   "Importe IVA 2|N|S|||scafaccli|imporiv2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   37
            Tag             =   "% IVA 2|N|S|0|99.90|scafaccli|porciva2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   27
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   35
            Tag             =   "Cod. IVA 2|N|S|0|9999|scafaccli|codigiv2|0000|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   36
            Tag             =   "Base Imponible 2 |N|S|||scafaccli|baseimp2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   35
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   32
            Tag             =   "Importe IVA 1|N|N|||scafaccli|imporiv1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   29
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   31
            Tag             =   "% IVA 1|N|S|0|99.90|scafaccli|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   26
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   29
            Tag             =   "Cod. IVA 1|N|S|0|9999|scafaccli|codigiv1|0000|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   30
            Tag             =   "Base Imponible 1|N|N|||scafaccli|baseimp1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   25
            Left            =   3900
            MaxLength       =   15
            TabIndex        =   27
            Text            =   "Text1 7"
            Top             =   330
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   240
            MaxLength       =   15
            TabIndex        =   25
            Tag             =   "Imp.Bruto|N|N|||scafaccli|brutofac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   330
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   1890
            TabIndex        =   173
            Top             =   330
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   12
            Left            =   2070
            TabIndex        =   172
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Suplidos"
            Height          =   255
            Index           =   49
            Left            =   5580
            TabIndex        =   171
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Importe RE"
            Height          =   195
            Index           =   44
            Left            =   7560
            TabIndex        =   138
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
            Height          =   255
            Index           =   43
            Left            =   6720
            TabIndex        =   137
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
            Height          =   195
            Index           =   37
            Left            =   5520
            TabIndex        =   136
            Top             =   720
            Width           =   825
         End
         Begin VB.Line Line1 
            X1              =   2280
            X2              =   2280
            Y1              =   960
            Y2              =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "Desglose IVA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   42
            Left            =   960
            TabIndex        =   125
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   4560
            TabIndex        =   124
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL FACTURA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   39
            Left            =   9720
            TabIndex        =   89
            Top             =   1320
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   9360
            TabIndex        =   88
            Top             =   1560
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   36
            Left            =   11880
            TabIndex        =   87
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base impo. IVA"
            Height          =   255
            Index           =   33
            Left            =   3120
            TabIndex        =   86
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   3630
            TabIndex        =   85
            Top             =   330
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   14
            Left            =   3900
            TabIndex        =   84
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   83
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Cliente"
         ForeColor       =   &H00972E0B&
         Height          =   2295
         Left            =   150
         TabIndex        =   61
         Top             =   420
         Width           =   12975
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   49
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   178
            Text            =   "Text2"
            Top             =   1290
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   49
            Left            =   6885
            MaxLength       =   4
            TabIndex        =   18
            Tag             =   "Banco Propio|N|S|||scafaccli|codbanpr|0000|N|"
            Text            =   "Text1"
            Top             =   1290
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   48
            Left            =   5430
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "IBAN|T|S|||scafaccli|iban||N|"
            Text            =   "Text1 7"
            Top             =   1845
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   46
            Left            =   9780
            MaxLength       =   10
            TabIndex        =   24
            Tag             =   "Servicios|N|S|||scafaccli|numservi|###,##0|N|"
            Text            =   "Text1 7"
            Top             =   1860
            Width           =   1065
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   11400
            MaxLength       =   10
            TabIndex        =   145
            Tag             =   "Aportacion|N|S|||scafaccli|portes|#,##0.00|N|"
            Text            =   "Portes"
            Top             =   1890
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   45
            Left            =   11040
            MaxLength       =   10
            TabIndex        =   140
            Tag             =   "Aportacion|N|S|||scafaccli|aportacion|#,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1890
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   21
            Left            =   8070
            MaxLength       =   10
            TabIndex        =   23
            Tag             =   "Cuenta Bancaria|T|S|||scafaccli|cuentaba|0000000000|N|"
            Text            =   "Text1 7"
            Top             =   1845
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   20
            Left            =   7590
            MaxLength       =   2
            TabIndex        =   22
            Tag             =   "Digito Control|T|S|||scafaccli|digcontr|00|N|"
            Text            =   "Text1 7"
            Top             =   1845
            Width           =   405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   19
            Left            =   6870
            MaxLength       =   4
            TabIndex        =   21
            Tag             =   "Sucursal|N|S|0|9999|scafaccli|codsucur|0000|N|"
            Text            =   "Text1 7"
            Top             =   1845
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   18
            Left            =   6150
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "Banco|N|S|0|9999|scafaccli|codbanco|0000|N|"
            Text            =   "Text1 7"
            Top             =   1845
            Width           =   645
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   16
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   13
            Tag             =   "Refere. Cliente|T|S|||scafac1|referenc|||"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   1740
            Width           =   1725
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   15
            Tag             =   "Direccion/Dpto.|T|S|||scafaccli|nomdirec||N|"
            Text            =   "Text1"
            Top             =   285
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   14
            Tag             =   "Direccion/Dpto.|N|S|0|999|scafaccli|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   285
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1125
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||scafaccli|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1350
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1125
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "CPostal|T|N|||scafaccli|codpobla||N|"
            Text            =   "Text15"
            Top             =   990
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1755
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Población|T|N|||scafaccli|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   990
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3195
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "teléfono Cliente|T|S|||scafaccli|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   285
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1125
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "NIF Cliente|T|N|||scafaccli|nifclien||N|"
            Text            =   "123456789"
            Top             =   285
            Width           =   1110
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   6885
            MaxLength       =   4
            TabIndex        =   16
            Tag             =   "Cod. Agente|N|N|0|9999|scafaccli|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   615
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   14
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   65
            Text            =   "Text2"
            Top             =   615
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   17
            Tag             =   "Forma de Pago|N|N|0|999|scafaccli|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   960
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   15
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   63
            Text            =   "Text2"
            Top             =   960
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1125
            MaxLength       =   35
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||scafaccli|domclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   645
            Width           =   4030
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   6600
            ToolTipText     =   "Buscar banco propio"
            Top             =   1290
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Banco Propio"
            Height          =   255
            Index           =   52
            Left            =   5460
            TabIndex        =   179
            Top             =   1290
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
            Height          =   255
            Index           =   51
            Left            =   5430
            TabIndex        =   177
            Top             =   1650
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "SERVICIOS"
            Height          =   255
            Index           =   47
            Left            =   9780
            TabIndex        =   168
            Top             =   1650
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Aportación"
            Height          =   255
            Index           =   45
            Left            =   11040
            TabIndex        =   141
            Top             =   1650
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6600
            ToolTipText     =   "Buscar agente"
            Top             =   615
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta"
            Height          =   255
            Index           =   8
            Left            =   8070
            TabIndex        =   81
            Top             =   1650
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "DC"
            Height          =   255
            Index           =   5
            Left            =   7590
            TabIndex        =   80
            Top             =   1650
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   4
            Left            =   6870
            TabIndex        =   79
            Top             =   1650
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Banco"
            Height          =   255
            Index           =   3
            Left            =   6150
            TabIndex        =   78
            Top             =   1650
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   72
            Top             =   1740
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   855
            ToolTipText     =   "Buscar población"
            Top             =   1005
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc./Dpto"
            Height          =   255
            Index           =   1
            Left            =   5460
            TabIndex        =   71
            Top             =   285
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   3
            Left            =   6600
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   285
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   70
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   69
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   2445
            TabIndex        =   68
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   67
            Top             =   285
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   855
            ToolTipText     =   "Buscar cliente varios"
            Top             =   300
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   255
            Index           =   34
            Left            =   5460
            TabIndex        =   66
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5460
            TabIndex        =   64
            Top             =   960
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   6600
            ToolTipText     =   "Buscar forma de pago"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   62
            Top             =   645
            Width           =   735
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmFCliHcoFac.frx":0A62
         Height          =   1950
         Left            =   -74760
         TabIndex        =   90
         Top             =   570
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   3440
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Albaranes de la Factura"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmFCliHcoFac.frx":0A77
         Height          =   3600
         Left            =   -74760
         TabIndex        =   146
         Top             =   405
         Width           =   13410
         _ExtentX        =   23654
         _ExtentY        =   6350
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Servicios de la Factura"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Data4 
         Height          =   330
         Left            =   -74940
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   11
         Left            =   -70920
         MaxLength       =   9
         TabIndex        =   144
         Tag             =   "Nº Bultos|N|N|0||slifaccli|numbultos|#,###,##0|N|"
         Text            =   "numbultos"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   10
         Left            =   -65280
         MaxLength       =   15
         TabIndex        =   139
         Tag             =   "Nº Lote|T|S|||slifaccli|numlote||N|"
         Text            =   "NLote"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   -66120
         MaxLength       =   30
         TabIndex        =   122
         Tag             =   "Cod. Proveedor|N|N|||slifaccli|codprovex|0||"
         Text            =   "prove"
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   -69360
         MaxLength       =   5
         TabIndex        =   117
         Text            =   "origp"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   -71760
         MaxLength       =   12
         TabIndex        =   115
         Tag             =   "Cantidad|N|N|0||slifaccli|cantidad|#,###,###,##0.00|N|"
         Text            =   "cantidad"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73680
         MaxLength       =   12
         TabIndex        =   113
         Tag             =   "Art.|T|N|||slifaccli|codartic||N|"
         Text            =   "codartic"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74640
         MaxLength       =   12
         TabIndex        =   112
         Tag             =   "Almacen|N|N|0|999|slifaccli|codalmac|000|N|"
         Text            =   "almacen"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   -70200
         MaxLength       =   12
         TabIndex        =   116
         Tag             =   "Precio|N|N|0|999999.0000|slifaccli|precioar|###,##0.0000|N|"
         Text            =   "Precio"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   -68640
         MaxLength       =   5
         TabIndex        =   118
         Tag             =   "Dto 1|N|N|0|99.90|slifaccli|dtoline1|#0.00|N|"
         Text            =   "Dto1"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   -67920
         MaxLength       =   30
         TabIndex        =   119
         Tag             =   "Dto 2|N|N|0|99.90|slifaccli|dtolinea|#0.00|N|"
         Text            =   "Dto2"
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   -67320
         MaxLength       =   12
         TabIndex        =   121
         Tag             =   "Importe|N|N|0||slifaccli|importel|#,###,###,##0.00|N|"
         Text            =   "Importe"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -72840
         MaxLength       =   12
         TabIndex        =   114
         Tag             =   "Nombre Art.|T|N|||slifaccli|nomartic||N|"
         Text            =   "nomartic"
         Top             =   4320
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFCliHcoFac.frx":0A8C
         Height          =   2025
         Left            =   -74760
         TabIndex        =   77
         Top             =   2640
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   3572
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   -70350
         MaxLength       =   7
         TabIndex        =   95
         Tag             =   "Nº Pedido|N|S|||scafac1|numpedcl|0000000|N|"
         Text            =   "Text1 7"
         Top             =   3030
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   -70440
         MaxLength       =   30
         TabIndex        =   51
         Tag             =   "Trabajador pedido|N|S|0|9999|scafac1|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   3810
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   -68550
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   104
         Text            =   "Text2"
         Top             =   3930
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   -69270
         MaxLength       =   30
         TabIndex        =   53
         Tag             =   "Cod. Envío|N|N|0|999|scafac1|codenvio|000|N|"
         Text            =   "Text1"
         Top             =   3930
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   -68550
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   92
         Text            =   "Text2"
         Top             =   2805
         Width           =   3525
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   -68550
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   91
         Text            =   "Text2"
         Top             =   3375
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -69270
         MaxLength       =   30
         TabIndex        =   52
         Tag             =   "Preparador materia|N|N|0|9999|scafac1|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   3375
         Width           =   660
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   -62970
         MaxLength       =   10
         TabIndex        =   94
         Tag             =   "Semana Entrega|N|S|||scafac1|sementre||N|"
         Top             =   3150
         Width           =   705
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   -64290
         MaxLength       =   10
         TabIndex        =   96
         Tag             =   "Fecha Pedido|F|S|||scafac1|fecpedcl|dd/mm/yyyy|N|"
         Top             =   3150
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   -63540
         MaxLength       =   10
         TabIndex        =   97
         Tag             =   "Fecha Oferta|F|S|||scafac1|fecofert|dd/mm/yyyy|N|"
         Top             =   3540
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   -64740
         MaxLength       =   7
         TabIndex        =   98
         Tag             =   "Nº Oferta|N|S|||scafac1|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   3540
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   14
         Left            =   -63690
         MaxLength       =   7
         TabIndex        =   134
         Tag             =   "Nº Venta|N|S|||scafac1|numventa|0000000|N|"
         Text            =   "Text1 7"
         Top             =   3990
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   15
         Left            =   -64890
         MaxLength       =   7
         TabIndex        =   135
         Tag             =   "Nº Terminal|N|S|||scafac1|numtermi||N|"
         Text            =   "Text1 7"
         Top             =   3990
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   16
         Left            =   11475
         MaxLength       =   5
         TabIndex        =   161
         Tag             =   "Descuento P.Pago|N|N|0|99.90|scafaccli|dtoppago|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   1200
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   17
         Left            =   11475
         MaxLength       =   5
         TabIndex        =   162
         Tag             =   "Descuento General|N|N|0|99.90|scafaccli|dtognral|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   1560
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   165
         Tag             =   "Imp. Dto PP|N|N|||scafaccli|impdtopp|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3495
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Observación 2"
         Height          =   255
         Index           =   50
         Left            =   -68640
         TabIndex        =   175
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto"
         Height          =   255
         Index           =   48
         Left            =   -74700
         TabIndex        =   170
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   -69120
         ToolTipText     =   "Buscar trabajador"
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -69240
         ToolTipText     =   "Buscar forma de envio"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -69240
         ToolTipText     =   "Buscar trabajador"
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
         Height          =   255
         Index           =   40
         Left            =   -64080
         TabIndex        =   103
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
         Height          =   255
         Index           =   22
         Left            =   -62640
         TabIndex        =   102
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         Height          =   255
         Index           =   18
         Left            =   -63480
         TabIndex        =   101
         Top             =   2775
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   6
         Left            =   -64320
         TabIndex        =   100
         Top             =   2775
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Sem. Entrega"
         Height          =   255
         Index           =   2
         Left            =   -62280
         TabIndex        =   99
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Albaran"
         Height          =   255
         Index           =   21
         Left            =   -70590
         TabIndex        =   76
         Top             =   435
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo  Envío"
         Height          =   195
         Index           =   24
         Left            =   -70680
         TabIndex        =   75
         Top             =   3840
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Prepar. Material"
         Height          =   255
         Index           =   23
         Left            =   -70680
         TabIndex        =   74
         Top             =   3360
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Pedido"
         Height          =   255
         Index           =   9
         Left            =   -70680
         TabIndex        =   73
         Top             =   2760
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   -69240
         ToolTipText     =   "Buscar trabajador"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Dto. Gral"
         Height          =   255
         Index           =   26
         Left            =   10785
         TabIndex        =   164
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Dto. P.P"
         Height          =   255
         Index           =   25
         Left            =   10770
         TabIndex        =   163
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. Dto PP"
         Height          =   255
         Index           =   11
         Left            =   2580
         TabIndex        =   167
         Top             =   3300
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   2220
         TabIndex        =   166
         Top             =   3420
         Width           =   135
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   16
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   250
      TabIndex        =   120
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6060
      Visible         =   0   'False
      Width           =   7125
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   55
      Top             =   5895
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   56
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   13050
      TabIndex        =   49
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11760
      TabIndex        =   48
      Top             =   6000
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            Object.Width           =   1e-4
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
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Factura"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir albarán"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   58
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   13050
      TabIndex        =   54
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3000
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
      Height          =   255
      Index           =   46
      Left            =   2400
      TabIndex        =   142
      Top             =   6435
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliación Línea"
      Height          =   255
      Index           =   35
      Left            =   2400
      TabIndex        =   60
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnImprimirAlbaran 
         Caption         =   "Imprimir &albarán"
         Enabled         =   0   'False
         Shortcut        =   ^A
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFCliHcoFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento

Public publicidad As Boolean
Public DesdeFichaCliente As Boolean

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmC As frmFacClientes 'Form Mto Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCV As frmFacClientesV  'Form Mto Clientes Varios
Attribute frmCV.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmA As frmFacAgentesCom   'Form Mto Agentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmFE As frmFacFormasEnvio  'Form Formas de Envio
Attribute frmFE.VB_VarHelpID = -1
Private WithEvents frmP As frmComProveedores
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios
Attribute frmMtoBancosPro.VB_VarHelpID = -1


Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Private UnaVez As Boolean
Private BuscaChekc As String

Dim cadB1 As String


Dim ImpIv1 As Currency
Dim ImpIv2 As Currency
Dim ImpIv3 As Currency
Dim BImp1 As Currency
Dim BImp2 As Currency
Dim BImp3 As Currency
Dim TotFac As Currency
Dim FPago As Currency
Dim BPr As Currency
Dim Serv As Currency

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then CheckCadenaBusqueda Check1(Index), BuscaChekc
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cad1 As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 4  'MODIFICAR
            If DatosOk Then
               If ModificarFactura Then
               
                     '[Monica]06/03/2012: guardo en el slog los campos que me han cambiado
                    cad1 = ""
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(15).Text))) <> FPago Then cad1 = cad1 & "FPago: " & FPago & " a " & Text1(12).Text & ";"
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(46).Text))) <> Serv Then cad1 = cad1 & "Serv.: " & Serv & " a " & Text1(22).Text & ";"
                    
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(49).Text))) <> BPr Then cad1 = cad1 & "B.Pr.: " & BPr & " a " & Text1(49).Text & ";"
                    
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(32).Text))) <> BImp1 Then cad1 = cad1 & "B.Imp1: " & BImp1 & " a " & Text1(32).Text & ";"
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(33).Text))) <> BImp2 Then cad1 = cad1 & "B.Imp2: " & BImp2 & " a " & Text1(33).Text & ";"
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(34).Text))) <> BImp3 Then cad1 = cad1 & "B.Imp3: " & BImp3 & " a " & Text1(34).Text & ";"
                    
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(35).Text))) <> ImpIv1 Then cad1 = cad1 & "Imp.Iv1: " & ImpIv1 & " a " & Text1(35).Text & ";"
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(36).Text))) <> ImpIv2 Then cad1 = cad1 & "Imp.Iv2: " & ImpIv2 & " a " & Text1(36).Text & ";"
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(37).Text))) <> ImpIv3 Then cad1 = cad1 & "Imp.Iv3: " & ImpIv3 & " a " & Text1(37).Text & ";"
                    
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(38).Text))) <> TotFac Then cad1 = cad1 & "Tot.Fac.: " & TotFac & " a " & Text1(38).Text & ";"
                                        
                    Set LOG = New cLOG
                    LOG.Insertar 8, vUsu, "Factura Serv.Cliente modificada: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & " " & cad1 & vbCrLf
                    Set LOG = Nothing
               
                    Espera 0.2
                    TerminaBloquear
                    PosicionarData
                    FormatoDatosTotales
                    i = Data3.Recordset.AbsolutePosition
                    PonerCamposLineas
                    SituarDataPosicion Data3, CLng(i), ""
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            If ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                
                        'INSERTA LOG
                        '-------------------------------------------------
'                        Set LOG = New cLOG
'                        BuscaChekc = "     Alb : " & Data2.Recordset!NumAlbar & "   Linea: " & Data2.Recordset!numlinea
'                        BuscaChekc = "Modificar linea: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & BuscaChekc
'                        LOG.Insertar 8, vUsu, BuscaChekc
'                        Set LOG = Nothing
'                        BuscaChekc = ""
                
                     '[Monica]06/03/2012: guardo en el slog los campos que me han cambiado
                    cad1 = ""
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(15).Text))) <> FPago Then cad1 = cad1 & "FPago: " & FPago & " a " & Text1(12).Text & ";"
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(46).Text))) <> Serv Then cad1 = cad1 & "Serv.: " & Serv & " a " & Text1(22).Text & ";"
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(49).Text))) <> BPr Then cad1 = cad1 & "BPr.: " & BPr & " a " & Text1(49).Text & ";"
                    
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(32).Text))) <> BImp1 Then cad1 = cad1 & "B.Imp1: " & BImp1 & " a " & Text1(32).Text & ";"
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(33).Text))) <> BImp2 Then cad1 = cad1 & "B.Imp2: " & BImp2 & " a " & Text1(33).Text & ";"
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(34).Text))) <> BImp3 Then cad1 = cad1 & "B.Imp3: " & BImp3 & " a " & Text1(34).Text & ";"
                    
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(35).Text))) <> ImpIv1 Then cad1 = cad1 & "Imp.Iv1: " & ImpIv1 & " a " & Text1(35).Text & ";"
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(36).Text))) <> ImpIv2 Then cad1 = cad1 & "Imp.Iv2: " & ImpIv2 & " a " & Text1(36).Text & ";"
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(37).Text))) <> ImpIv3 Then cad1 = cad1 & "Imp.Iv3: " & ImpIv3 & " a " & Text1(37).Text & ";"
                    
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(38).Text))) <> TotFac Then cad1 = cad1 & "Tot.Fac.: " & TotFac & " a " & Text1(38).Text & ";"
                                        
                                       
                    Set LOG = New cLOG
                    LOG.Insertar 8, vUsu, "Factura modificada: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & " " & cad1 & vbCrLf
                    Set LOG = Nothing
               
                
                    TerminaBloquear
                    CargaGrid DataGrid1, Data2, True
                    CargaGrid DataGrid3, Data4, True
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
            
                    LLamaLineas Modo, 0, "DataGrid1"
                    LLamaLineas Modo, 0, "DataGrid3"
                    PosicionarData
                Else
                    TerminaBloquear
                End If
                Me.DataGrid1.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click()
    Set frmP = New frmComProveedores
    frmP.DatosADevolverBusqueda = "0|1|"
    frmP.Show vbModal
    Set frmP = Nothing

End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            BloquearTxt Text2(16), True
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            DataGrid2.Enabled = True
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid1"
            LLamaLineas Modo, 0, "DataGrid3"
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    End Select
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    ModificaLineas = 1
    'Buscar
    If Modo <> 1 Then
        BuscaChekc = ""
        LimpiarCampos
        
        Select Case Me.Combo1.ListIndex
            Case 0
                Text1(1).Text = "FAC"
            Case 1
                Text1(1).Text = "FRN"
        End Select
        
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
        'poner los txtaux para buscar por lineas de albaran
        anc = DataGrid2.Top
        If DataGrid2.Row < 0 Then
            anc = anc + 440
        Else
            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid2"
        
        'poner los txtaux para buscar por lineas de albaran
        anc = DataGrid3.Top
        If DataGrid3.Row < 0 Then
            anc = anc + 440
        Else
            anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid3"
        
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        Text1(1).Text = "FAC"

        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        lblIndicador.Caption = "Preparando bus."
        lblIndicador.Refresh
        LimpiarCampos
        LimpiarDataGrids
        DoEvents
        
        CadenaConsulta = "Select scafaccli.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE scafaccli.codtipom='FAC'"
        lblIndicador.Caption = "Obteniendo reg."
        lblIndicador.Refresh
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean
Dim EnTesoreria  As String
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada2(EnTesoreria) Then
        TerminaBloquear
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    Me.SSTab1.Tab = 0
    
    PonerFocoChk Me.Check1(0)
        
    'Inserto en slog
    
    Set LOG = New cLOG
    If EnTesoreria <> "" Then EnTesoreria = "Tesoreria: " & vbCrLf & EnTesoreria
    EnTesoreria = Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & vbCrLf & EnTesoreria
    EnTesoreria = "Pulsa mod factura: " & EnTesoreria
    LOG.Insertar 8, vUsu, EnTesoreria
    Set LOG = Nothing
    Espera 0.3
    '
    
    '[Monica]6/03/2012: para guardarlo en el slog
    ImpIv1 = CCur(ImporteSinFormato(ComprobarCero(Text1(35).Text)))
    ImpIv2 = CCur(ImporteSinFormato(ComprobarCero(Text1(36).Text)))
    ImpIv3 = CCur(ImporteSinFormato(ComprobarCero(Text1(37).Text)))
    BImp1 = CCur(ImporteSinFormato(ComprobarCero(Text1(32).Text)))
    BImp2 = CCur(ImporteSinFormato(ComprobarCero(Text1(33).Text)))
    BImp3 = CCur(ImporteSinFormato(ComprobarCero(Text1(34).Text)))
    TotFac = CCur(ImporteSinFormato(ComprobarCero(Text1(38).Text)))
    FPago = CCur(ImporteSinFormato(ComprobarCero(Text1(15).Text)))
    Serv = CCur(ImporteSinFormato(ComprobarCero(Text1(46).Text)))
    BPr = ComprobarCero(Text1(49).Text)
    
    
    'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = EsClienteVarios(Text1(4).Text)
    BloquearDatosCliente (DeVarios)
    
    PonerFoco Text1(15)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte
Dim EstaEnTesoreria As String
    On Error GoTo EModificarLinea


     'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada3(EstaEnTesoreria) Then
        TerminaBloquear
        Exit Sub
    End If

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then  '1= Insertar
        TerminaBloquear
        Exit Sub
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " AND codtipoa='" & Data3.Recordset.Fields!codtipoa & "' AND numalbar=" & Data3.Recordset.Fields!NumAlbar
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If

    'INSERTA LOG
    '-------------------------------------------------
    Set LOG = New cLOG
    If EstaEnTesoreria <> "" Then EstaEnTesoreria = "Tesoreria: " & EstaEnTesoreria
    EstaEnTesoreria = "     Alb : " & Data2.Recordset!NumAlbar & "   Linea: " & Data2.Recordset!numlinea & vbCrLf & EstaEnTesoreria
    EstaEnTesoreria = "Pulsa mod linea: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & EstaEnTesoreria
    LOG.Insertar 8, vUsu, EstaEnTesoreria
    Set LOG = Nothing

    If DataGrid3.Bookmark < DataGrid3.FirstRow Or DataGrid3.Bookmark > (DataGrid3.FirstRow + DataGrid3.VisibleRows - 1) Then
        J = DataGrid3.Bookmark - DataGrid3.FirstRow
        DataGrid3.Scroll 0, J
        DataGrid3.Refresh
    End If
    
'    anc = ObtenerAlto(Me.DataGrid1)
    anc = DataGrid3.Top
    If DataGrid3.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 10
    End If

    txtAux4(12).Text = DataGrid3.Columns(14).Text
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid3"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
'    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    PonerFoco txtAux4(12)
    Me.DataGrid2.Enabled = False
    Me.DataGrid1.Enabled = False

EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Sub BotonModificarLineaTele()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte
Dim EstaEnTesoreria As String
    On Error GoTo EModificarLinea


     'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada2(EstaEnTesoreria) Then
        TerminaBloquear
        Exit Sub
    End If

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then  '1= Insertar
        TerminaBloquear
        Exit Sub
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " AND codtipoa='" & Data3.Recordset.Fields!codtipoa & "' AND numalbar=" & Data3.Recordset.Fields!NumAlbar
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If

    'INSERTA LOG
    '-------------------------------------------------
'    Set LOG = New cLOG
'    If EstaEnTesoreria <> "" Then EstaEnTesoreria = "Tesoreria: " & EstaEnTesoreria
'    EstaEnTesoreria = "     Alb : " & Data2.Recordset!NumAlbar & "   Linea: " & Data2.Recordset!numlinea & vbCrLf & EstaEnTesoreria
'    EstaEnTesoreria = "Pulsa mod linea: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & EstaEnTesoreria
'    LOG.Insertar 8, vUsu, EstaEnTesoreria
'    Set LOG = Nothing

    '[Monica]6/03/2012: para guardarlo en el slog
    ImpIv1 = ComprobarCero(Text1(35).Text)
    ImpIv2 = ComprobarCero(Text1(36).Text)
    ImpIv3 = ComprobarCero(Text1(37).Text)
    BImp1 = ComprobarCero(Text1(32).Text)
    BImp2 = ComprobarCero(Text1(33).Text)
    BImp3 = ComprobarCero(Text1(34).Text)
    TotFac = ComprobarCero(Text1(38).Text)
    FPago = ComprobarCero(Text1(15).Text)
    Serv = ComprobarCero(Text1(46).Text)
    BPr = ComprobarCero(Text1(49).Text)



    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        J = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, J
        DataGrid1.Refresh
    End If
    
'    anc = ObtenerAlto(Me.DataGrid1)
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 10
    End If

    txtAux(3).Text = DataGrid1.Columns(9).Text
    txtAux(4).Text = DataGrid1.Columns(11).Text
    txtAux(8).Text = DataGrid1.Columns(15).Text
    
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(4), False
    PonerFoco txtAux(4)
    Me.DataGrid2.Enabled = False
    Me.DataGrid1.Enabled = False

EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub





Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean

    Select Case grid
        Case "DataGrid1"
            DeseleccionaGrid Me.DataGrid1
            'PonerModo xModo + 1
    
            b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
    
            For jj = 0 To txtAux.Count - 1
                If jj = 3 Or jj = 4 Or jj = 8 Then
                    txtAux(jj).Height = DataGrid1.RowHeight
                    txtAux(jj).Top = alto
                    txtAux(jj).visible = b
                End If
            Next jj
'            cmdAux.Top = alto
'            cmdAux.visible = False

        Case "DataGrid2"
            DeseleccionaGrid Me.DataGrid2
            b = (xModo = 1)
             For jj = 0 To txtAux3.Count - 1
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto
                txtAux3(jj).visible = b
            Next jj
            
        Case "DataGrid3"
            DeseleccionaGrid Me.DataGrid3
            b = (xModo = 1)
            For jj = 3 To 11
                txtAux4(jj).Height = DataGrid3.RowHeight
                txtAux4(jj).Top = alto
                txtAux4(jj).visible = b
            Next jj
            b = (xModo = 1) Or (ModificaLineas = 2)
            txtAux4(12).Height = DataGrid3.RowHeight
            txtAux4(12).Top = alto
            txtAux4(12).visible = b
            
            
    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafaccli)
' y los registros correspondientes de las tablas cab. albaranes (scafaccli1)
' y las lineas de la factura (slifaccli)
Dim cad As String
Dim EstaEnTesoreria As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada3(EstaEnTesoreria) Then Exit Sub
    
    cad = "Cabecera de Facturas." & vbCrLf
    cad = cad & "-------------------------------------      " & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar la Factura:            "
    cad = cad & vbCrLf & "Tipo:  " & Text1(1).Text
    cad = cad & vbCrLf & "Nº Fact.:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(2).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
'        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
'        NumPedElim = Data1.Recordset.Fields(1).Value
        CodTipoMov = Text1(1).Text
        
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
        
            
            Set LOG = New cLOG
            LOG.Insertar 8, vUsu, "Factura eliminada: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & vbCrLf & EstaEnTesoreria
            Set LOG = Nothing
        
            If SituarDataTrasEliminar(Data1, NumRegElim) Then
                PonerCampos
            Else
                LimpiarCampos
                'Poner los grid sin apuntar a nada
                LimpiarDataGrids
                PonerModo 0
            End If
        End If
'        'Devolvemos contador, si no estamos actualizando
'        Set vTipoMov = New CTiposMov
'        vTipoMov.DevolverContador CodTipoMov, NumPedElim
'        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


Private Sub cmdObserva_Click()
'    If Modo <> 2 And Modo <> 4 And Modo <> 1 Then Exit Sub
'    If Me.FrameObserva.visible = False Then
'        Me.DataGrid1.visible = False
'        Me.FrameObserva.visible = True
'        Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(18).Picture
''        CargarICO Me.cmdObserva, "volver.ico"
'        Me.cmdObserva.ToolTipText = "volver lineas albaran"
'        BloqueaText3
'    Else
'        Me.DataGrid1.visible = True
'        Me.FrameObserva.visible = False
'        Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(41).Picture
''        CargarICO Me.cmdObserva, "message.ico"
'        Me.cmdObserva.ToolTipText = "ver observaciones albaran"
'    End If
'    SSTab1_Click 0
End Sub

Private Sub BloqueaText3()
Dim i As Byte
Dim b As Boolean

    'bloquear los Text3 que son las lineas de scafaccli1
    'B = Modo <> 4 And Modo <> 1
    b = Modo <> 1
    For i = 0 To 3
        BloquearTxt Text3(i), b
    Next i
    BloquearTxt Text3(16), b
    
    
'    If Me.FrameObserva.visible Then
        For i = 9 To 13
            BloquearTxt Text3(i), (Modo <> 4 And Modo <> 1)
        Next i
'    End If
    
    b = Modo <> 1
    For i = 4 To 8
        BloquearTxt Text3(i), b
    Next i
    'datos venta TPV
    BloquearTxt Text3(14), True
    BloquearTxt Text3(15), True
 
End Sub

Private Sub BloqueaText4()
Dim i As Byte
Dim b As Boolean
    'TaxiVIP
    If vParamAplic.Cooperativa = 1 Then
    
        'bloquear los Text3 que son las lineas de scafaccli1
        For i = 3 To 12
            BloquearTxt txtAux4(i), Not (Modo = 1)
            txtAux4(i).visible = (Modo = 1)
        Next i
        BloquearTxt txtAux4(12), Not ((Modo = 5) Or Modo = 1)  'B And Modo <> 4
        txtAux4(12).visible = ((Modo = 5 And ModificaLineas = 2) Or Modo = 1)
    
    Else
    'Teletaxi
        
        BloquearTxt txtAux(4), Not ((Modo = 5) Or Modo = 1)  'B And Modo <> 4
        txtAux(4).Enabled = ((Modo = 5 And ModificaLineas = 2) Or Modo = 1)
    
    End If
End Sub




Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid2.Enabled = True
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus()
    Text1(1).Text = Mid(Combo1.List(Combo1.ListIndex), 1, 3)
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Ayuda de Etiqueta de precio de salida de la Función de Precios
On Error Resume Next

    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        Me.DataGrid1.ToolTipText = ""
        If X > 7790 And X < 8170 Then
            Select Case DataGrid1.Columns(11).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoción"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Artículo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Artículo"
                Case "M": Me.DataGrid1.ToolTipText = "M: Manual"
'                Case Else
'                    Me.DataGrid1.ToolTipText = ""
            End Select
'        Else
'            Me.DataGrid1.ToolTipText = ""
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    DataGrid1RowColChange
End Sub

Private Sub DataGrid1RowColChange()
    On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        Text2(16).Text = DBLet(Data2.Recordset.Fields!Ampliaci, "T")
        
        If ModificaLineas <> 1 Then
            If vEmpresa.TieneAnalitica Then
                '- centro de coste
                ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
                Me.txtAux(9).Text = DBLet(Data2.Recordset!CodCCost, "T")
                Me.txtAux2(9).Text = PonerNombreCCoste(Me.txtAux(9))
            Else
                txtAux2(9).Text = DBLet(Data2.Recordset.Fields!nomprove, "T")
            End If
        End If
    Else
        Text2(16).Text = ""
        txtAux2(9).Text = ""
    End If
    
    Exit Sub

Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If Not Data3.Recordset.EOF Then
        'Trabajador Albaran
        Text3(0).Text = Data3.Recordset.Fields!CodTraba
        Text3_LostFocus (0)
        'Trabajador pedido
        Text3(1).Text = DBLet(Data3.Recordset.Fields!codtrab1, "T")
        Text3_LostFocus (1)
        'Trab. Prepara Material
        Text3(2).Text = Data3.Recordset.Fields!codtrab2
        Text3_LostFocus (2)
        Text3(3).Text = Data3.Recordset.Fields!CodEnvio
        Text3_LostFocus (3)
        
        'oferta
        Text3(4).Text = DBLet(Data3.Recordset.Fields!NumOfert, "N")
        If Text3(4).Text <> "0" Then
            FormateaCampo Text3(4)
        Else
            Text3(4).Text = ""
        End If
        Text3(5).Text = DBLet(Data3.Recordset.Fields!fecofert, "F")
        'pedido
        Text3(6).Text = DBLet(Data3.Recordset.Fields!Numpedcl, "N")
        If Text3(6).Text <> "0" Then
            FormateaCampo Text3(6)
        Else
            Text3(6).Text = ""
        End If
        Text3(7).Text = DBLet(Data3.Recordset.Fields!fecpedcl, "F")
        If Text3(7).Text <> "" Then FormateaCampo Text3(7)
        Text3(8).Text = DBLet(Data3.Recordset.Fields!sementre, "N")
        If Text3(8).Text = "0" Then Text3(8).Text = ""
        'venta
        Text3(15).Text = DBLet(Data3.Recordset.Fields!NumTermi, "N")
        Text3(14).Text = DBLet(Data3.Recordset.Fields!NumVenta, "N")
        FormateaCampo Text3(14)
'        If Text3(14).Text = "0" Then Text3(14).Text = ""
'        If Text3(15).Text = "0" Then Text3(15).Text = ""
        
        'Observaciones
        Text3(9).Text = DBLet(Data3.Recordset.Fields!observa1, "T")
        Text3(10).Text = DBLet(Data3.Recordset.Fields!observa2, "T")
        Text3(11).Text = DBLet(Data3.Recordset.Fields!observa3, "T")
        Text3(12).Text = DBLet(Data3.Recordset.Fields!observa4, "T")
        Text3(13).Text = DBLet(Data3.Recordset.Fields!observa5, "T")
        
        
        Text3(16).Text = DBLet(Data3.Recordset.Fields!referenc, "T")
        
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, True
    Else
        For i = 0 To Text3.Count - 1
            Text3(i).Text = ""
        Next i
        For i = 0 To 3
            Text2(i).Text = ""
        Next i
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, False
    End If
    
End Sub


Private Sub DataGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1

    If Not Data4.Recordset.EOF Then
        txtAux2(0).Text = DBLet(Data4.Recordset.Fields!observac2, "T")
        txtAux2(1).Text = DBLet(Data4.Recordset.Fields!observa1, "T")
    Else
        txtAux2(0).Text = ""
        txtAux2(1).Text = ""
    End If
    
    Exit Sub

Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If UnaVez Then
        UnaVez = False
        If hcoCodMovim <> "" Then
            If Data1.Recordset.EOF Then
                PonerCadenaBusqueda
            Else
                PonerCampos
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()

    UnaVez = True
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnPrimero = 15
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(9).Image = 10 'Mto Lineas
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(11).Image = 40 'Imprimir albaran
        .Buttons(13).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    Me.SSTab1.Tab = 0
    
    cadB1 = "scafaccli.codtipom = 'FAC'"
    
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
    
    
    'cargar icono de observaciones de los albaranes de factura
    Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(41).Picture
'    CargarICO Me.cmdObserva, "message.ico"
'    Me.FrameObserva.visible = False
'    Me.cmdObserva.ToolTipText = "ver observaciones albaran"

    VieneDeBuscar = False
    
    'Comprobar si es Departamento o Direccion
    If vParamAplic.Departamento Then
        Me.Label1(1).Caption = "Dpto."
    Else
        Me.Label1(1).Caption = "Direc."
    End If
        
        
    Me.Label1(45).visible = vParamAplic.ctaAportacion <> ""
    Text1(45).visible = vParamAplic.ctaAportacion <> ""
        
        
    If vEmpresa.TieneAnalitica Then
        txtAux(9).Tag = "Cod. centro coste|T|S|||slifaccli|codccost|||"
        Label1(46).Caption = "Centro coste"
    Else
        txtAux(9).Tag = "Cod. Proveedor|N|N|||slifaccli|codprovex|0||"
        Label1(46).Caption = "Proveedor"
    End If
        
        
    '## A mano
    NombreTabla = "scafaccli"
    NomTablaLineas = "slifaccli" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafaccli.codtipom, scafaccli.numfactu, scafaccli.fecfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Dim T1 As Single
    T1 = Timer
    
    CadenaConsulta = "Select * from " & NombreTabla
    
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        'como tenemos aqui el nº de albaran, buscar a que factura corresponde
        'en la scafaccli1
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura & " and (scafaccli.codtipom = 'FAC' or scafaccli.codtipom = 'FRN' or scafaccli.codtipom = 'FVC')"
'        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        'CadenaConsulta = CadenaConsulta & " where numfactu=-1"
        'Cambio sugerido por Msoler
        'mariela 01/07/2010
        CadenaConsulta = CadenaConsulta & " WHERE codtipom is null and numfactu is null and fecfactu is null and scafaccli.codtipom = 'FAC'"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    
    'ANTE
    'If hcoCodMovim <> "" Then Data1.Refresh
    Data1.Refresh
    

    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            BotonBuscar
        End If
'        CargaGrid DataGrid1, Data2, False
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PrimeraVez = False
    Else
        If Data1.Recordset.EOF Then
            PonerModo 0
        Else
            PonerModo 2
        End If
    End If

    ' dependiendo de cooperativa
    Label1(48).visible = (vParamAplic.Cooperativa = 1)
    txtAux2(0).visible = (vParamAplic.Cooperativa = 1)
    Label1(50).visible = (vParamAplic.Cooperativa = 1)
    txtAux2(1).visible = (vParamAplic.Cooperativa = 1)



End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1(0).Value = 0
    Me.Check1(1).Value = 0
    Me.Combo1.ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub




Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Agentes
Dim indice As Byte
    indice = 14
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod agente
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom agente
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
            cadB = cadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        Else 'Llama desde Prismatico Direcciones/Departamentos
            Text1(12).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            Text1(13).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Clien
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim indice As Byte
Dim devuelve As String

    indice = 9
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve)  'Poblacion
    'provincia
    Text1(indice + 2).Text = devuelve
End Sub


Private Sub frmCV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes Varios
Dim indice As Byte

    indice = 6
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosClienteVario (Text1(indice).Text)
End Sub


Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
'Formas de Envio
Dim indice As Byte
    indice = 29
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Envio
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Envio
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim indice As Byte
    indice = 15
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(9).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(9).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim indice As Byte
    indice = Val(Me.imgBuscar(3).Tag)
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    Text1(49).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(49).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            PonerFoco Text1(4)
            Set frmC = New frmFacClientes
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
            indice = 5
            PonerFoco Text1(indice)
            
        Case 1 'NIF para cliente de Varios
            Set frmCV = New frmFacClientesV
            frmCV.DatosADevolverBusqueda = "0"
            frmCV.Show vbModal
            Set frmCV = Nothing
            indice = 6
            PonerFoco Text1(indice)
            
        Case 2 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            indice = 9
            VieneDeBuscar = True
            PonerFoco Text1(indice)
        
        Case 3 'Cod. Direc.
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera = False
                MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text)
                indice = 12
             End If
             PonerFoco Text1(indice)
             
        Case 4 'Agente
            indice = 14
            PonerFoco Text1(indice)
            Set frmA = New frmFacAgentesCom
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
            
         Case 5 'Forma de Pago
            indice = 15
            PonerFoco Text1(indice)
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 6, 7, 8 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            indice = Index - 6
            Me.imgBuscar(3).Tag = indice
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            PonerFoco Text3(indice)
       
        Case 9 'Cod Envio
            indice = 3
            PonerFoco Text3(indice)
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0"
            frmFE.Show vbModal
            Set frmFE = Nothing
            PonerFoco Text3(indice)
        
        Case 10 ' banco propio
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
        
        
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If Data1.Recordset!codtipom = "FTI" Then 'ticket de venta del TPV
        BotonImprimirTicket
    ElseIf Data1.Recordset!codtipom = "FPC" Then 'factura de publicidad
        BotonImprimirPublicidad
    ElseIf Data1.Recordset!codtipom = "FRN" Then ' factura rectificativa de servicios
        BotonImprimirRectificativa
    Else

        If CInt(DBLet(Data3.Recordset!NumTermi, "N")) > 0 Then
            'Es factura del TPV
            BotonImprimir 63
        Else
            'Impresion normal
            BotonImprimir (53) '53: Informe de Facturas
        End If

    End If
End Sub

Private Sub BotonImprimirPublicidad()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
        cadFormula = "({scafaccli.codtipom}= '" & Data1.Recordset!codtipom & "' and {scafaccli.numfactu}= "
        cadFormula = cadFormula & Data1.Recordset!NumFactu & " and {scafaccli.fecfactu}= Date(" & Year(Data1.Recordset!FecFactu) & "," & Month(Data1.Recordset!FecFactu) & "," & Day(Data1.Recordset!FecFactu) & "))"
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        cadParam = cadParam & "pDuplicado= 1|"
        numParam = 2
        
        With frmImprimir
        .Titulo = "Impresión de Facturas de publicidad"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "47", "N")
    '------ > Listado 47 = rFacPubli.rpt
        .Opcion = 101
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub


Private Sub BotonImprimirRectificativa()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
        cadFormula = "({scafaccli.codtipom}= '" & Data1.Recordset!codtipom & "' and {scafaccli.numfactu}= "
        cadFormula = cadFormula & Data1.Recordset!NumFactu & " and {scafaccli.fecfactu}= Date(" & Year(Data1.Recordset!FecFactu) & "," & Month(Data1.Recordset!FecFactu) & "," & Day(Data1.Recordset!FecFactu) & "))"
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        cadParam = cadParam & "pDuplicado= 1|"
        numParam = 2
        
        With frmImprimir
        .Titulo = "Impresión de Facturas Rectificativa"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "54", "N")
    '------ > Listado 47 = rFacPubli.rpt
        .Opcion = 101
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub


Private Sub mnImprimirAlbaran_Click()
Dim Seguir As Boolean
Dim TipoA As String
    If Me.Data1.Recordset.EOF Then Exit Sub
    If Me.Data3.Recordset.EOF Then Exit Sub
    
    
    'Albaranes que no se pueden montar
    Seguir = False
    If Not IsNull(Data3.Recordset!codtipoa) Then
        If Data3.Recordset!codtipoa <> "" Then
            TipoA = CStr(Data3.Recordset!codtipoa)
            If TipoA = "FTI" Or TipoA = "ALM" Then
                Seguir = False
            Else
                Seguir = True
            End If
        End If
    End If
    If Not Seguir Then
        MsgBox "No se puede imprimir el albaran seleccionado", vbExclamation
        Exit Sub
    End If
    
    
    
    If Val(Data3.Recordset!NumAlbar) = 0 Then
        MsgBox "No se puede imprimir el albaran seleccionado", vbExclamation
        Exit Sub
    End If
    
    
    If Data2.Recordset.EOF Then
        MsgBox "Albaran no tiene lineas", vbExclamation
        Exit Sub
    End If
    
    ImprimirAlbaran 1
    
    
End Sub

Private Sub mnLineas_Click()
    If vParamAplic.Cooperativa = 1 Then
        BotonMtoLineas 2, "Facturas"
    Else
        BotonMtoLineas 1, "Facturas"
    End If
End Sub


Private Sub mnModificar_Click()

    If vUsu.Nivel > 0 Then
        MsgBox "No tiene permiso para realizar la accion", vbExclamation
        Exit Sub
    End If

    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafaccli
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafaccli1
            If BloqueaAlbxFac Then
                If BloqueaLineasFac Then
                    If vParamAplic.Cooperativa = 1 Then
                        BotonModificarLinea
                    Else
                        BotonModificarLineaTele
                    End If
                End If
            End If
        End If
         
    Else   'Modificar Pedido
        'bloquea la tabla cabecera de factura: scafaccli
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafaccli1
            If BloqueaAlbxFac Then BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim Sql As String

    On Error GoTo EBloqueaAlb
    
    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM scafaccli1 "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea todas las lineas de la factura
Dim Sql As String

    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM slifaccli "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaLineasFac = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasFac = False
End Function


Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub


Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    Me.Label1(35).visible = Me.SSTab1.Tab = 1 And Me.DataGrid1.visible
    Me.Text2(16).visible = Me.SSTab1.Tab = 1 And Me.DataGrid1.visible
    Me.Label1(46).visible = (Modo = 5) And Me.DataGrid1.visible And Me.SSTab1.Tab = 1 And (vEmpresa.TieneAnalitica)
    Me.txtAux2(9).visible = (Modo = 5) And Me.DataGrid1.visible And Me.SSTab1.Tab = 1 And (vEmpresa.TieneAnalitica)
    
End Sub

Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    If Index = 1 And Modo = 1 Then
        SendKeys "{tab}"
        Exit Sub
    End If
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 2 'Fecha factura
                If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 3, 27, 28 'Cod Vendedor
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")

        Case 4 'Cod. Cliente
            If Modo = 1 Then 'Modo=1 Busqueda
                '-- Laura 12/01/2007
                'Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, NombreTabla, "nomclien")
                Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "scliente", "nomclien")
                '--
            Else
                PonerDatosCliente (Text1(Index).Text)
            End If
        
        Case 6 'NIF
            If Not EsDeVarios Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(6).Text = Data1.Recordset!nifClien Then Exit Sub
            End If
            PonerDatosClienteVario (Text1(Index).Text)
        
        Case 9 'Cod. Postal
             If Text1(Index).Locked Then Exit Sub
             If Text1(Index).Text = "" Then
                Text1(Index + 1).Text = ""
                Text1(Index + 2).Text = ""
                Exit Sub
             End If
             If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                 Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                 Text1(Index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
        
        Case 12 'Cod. Direc
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(Text1(Index)) Then
                'Comprobar que el cliente seleccionada tiene esa direccion
                If PonerDptoEnCliente Then
                    'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                    devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                    If devuelve <> "" And Text1(Index).Locked = False Then
                        devuelve = "El cliente tiene Mantenimientos."
                        MsgBox devuelve, vbInformation
                    End If
                Else
                    PonerFoco Text1(Index)
                End If
            Else
                Text1(Index + 1).Text = ""
            End If
            
        Case 14 'Cod. Agente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 15 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 16, 17 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then   'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 18 To 21 'banco, sucursal
            PonerFormatoEntero Text1(Index)
        Case 29 'Cod envio
'            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "senvio", "nomenvio")

        Case 22, 32, 33, 34, 35, 36, 37, 38, 39, 41, 43
            PonerFormatoDecimal Text1(Index), 3
        
        Case 29, 30, 31, 40, 42, 44
            PonerFormatoDecimal Text1(Index), 7
            
        Case 49 ' banco propio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sbanpr", "nombanpr")
            Else
                Text2(Index).Text = ""
            End If
        
            
    End Select
    
    CalcularDatosFactura2 Index
    
End Sub



Private Sub CalcularDatosFactura2(indice As Integer)

    If (Modo = 4 Or Modo = 5) And ((indice >= 26 And indice <= 37)) Then
        Dim PorceIVA As Currency
        Dim BaseImpo As Currency
        Dim TotalFac As Currency
        Dim TotalFac1 As Currency
        Dim TotalFac2 As Currency
        Dim TotalFac3 As Currency
        Dim ImpoIva As Currency

        If (indice = 26 Or indice = 27 Or indice = 28 Or indice = 32 Or indice = 33 Or indice = 34) And Text1(indice).Text <> "" Then
            If indice = 26 Or indice = 32 Then
                BaseImpo = CCur(ImporteSinFormato(Text1(32).Text))
    
                Text1(29).Text = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", Text1(26).Text, "N")
                PorceIVA = 0
                If Text1(29).Text <> "" Then PorceIVA = CCur(Text1(29).Text)
                ImpoIva = Round2(BaseImpo * PorceIVA / 100, 2)
                TotalFac1 = BaseImpo + ImpoIva
    
                Text1(32).Text = Format(BaseImpo, "#,###,###,##0.00")
                Text1(29).Text = Format(PorceIVA, "#0.00")
                Text1(35).Text = Format(ImpoIva, "#,###,###,##0.00")
            End If
            If indice = 27 Or indice = 33 Then
                BaseImpo = CCur(ImporteSinFormato(Text1(33).Text))
    
                Text1(30).Text = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", Text1(27).Text, "N")
                PorceIVA = 0
                If Text1(30).Text <> "" Then PorceIVA = CCur(Text1(30).Text)
                ImpoIva = Round2(BaseImpo * PorceIVA / 100, 2)
                TotalFac1 = BaseImpo + ImpoIva
    
                Text1(33).Text = Format(BaseImpo, "#,###,###,##0.00")
                Text1(30).Text = Format(PorceIVA, "#0.00")
                Text1(36).Text = Format(ImpoIva, "#,###,###,##0.00")
            End If
            If indice = 28 Or indice = 34 Then
                BaseImpo = CCur(ImporteSinFormato(Text1(34).Text))
    
                Text1(31).Text = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", Text1(28).Text, "N")
                PorceIVA = 0
                If Text1(31).Text <> "" Then PorceIVA = CCur(Text1(31).Text)
                ImpoIva = Round2(BaseImpo * PorceIVA / 100, 2)
                TotalFac1 = BaseImpo + ImpoIva
    
                Text1(34).Text = Format(BaseImpo, "#,###,###,##0.00")
                Text1(31).Text = Format(PorceIVA, "#0.00")
                Text1(37).Text = Format(ImpoIva, "#,###,###,##0.00")
            End If
        End If
       
       
        BaseImpo = CCur(ImporteSinFormato(ComprobarCero(Text1(32).Text))) + _
                   CCur(ImporteSinFormato(ComprobarCero(Text1(33).Text))) + _
                   CCur(ImporteSinFormato(ComprobarCero(Text1(34).Text)))
        
        Text1(22).Text = Format(BaseImpo, "#,###,###,##0.00")
        Text1(25).Text = Text1(22).Text
       
        TotalFac = CCur(ImporteSinFormato(ComprobarCero(Text1(32).Text))) + _
                   CCur(ImporteSinFormato(ComprobarCero(Text1(33).Text))) + _
                   CCur(ImporteSinFormato(ComprobarCero(Text1(34).Text))) + _
                   CCur(ImporteSinFormato(ComprobarCero(Text1(35).Text))) + _
                   CCur(ImporteSinFormato(ComprobarCero(Text1(36).Text))) + _
                   CCur(ImporteSinFormato(ComprobarCero(Text1(37).Text)))
                       
        Text1(38).Text = Format(TotalFac, "#,###,###,##0.00")

    End If

End Sub

Private Sub HacerBusqueda()
Dim cadB As String
Dim cadAux As String
    
    '--- Laura 12/01/2007
    cadAux = Text1(5).Text
    If Text1(4).Text <> "" Then Text1(5).Text = ""
    
    '---
    
    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    
    
    '--- David.  No se pq referencia NO lleva tag. Si han puesto algo lo paso a la cadena de busqueda
    
    
    
    '--- Laura 12/01/2007
    Text1(5).Text = cadAux
    '---
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
'        cadB = cadB & " and scafaccli.codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los FAV
        CadenaConsulta = "select scafaccli.* from " & NombreTabla & " INNER JOIN scafaccli1 ON scafaccli.codtipom=scafaccli1.codtipom AND scafaccli.numfactu=scafaccli1.numfactu AND scafaccli.fecfactu=scafaccli1.fecfactu "
'        CadenaConsulta = CadenaConsulta & " INNER JOIN scafaccli_serv ON scafaccli.codtipom = scafaccli_serv.codtipom and scafaccli.numfactu = scafaccli_serv.numfactu and scafaccli.fecfactu = scafaccli_serv.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB & " GROUP BY scafaccli.codtipom,scafaccli.numfactu,scafaccli.fecfactu " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    If EsCabecera Then
        cad = cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
        cad = cad & ParaGrid(Text1(0), 15, "Nº Factura")
        cad = cad & ParaGrid(Text1(2), 15, "Fecha Fac.")
        cad = cad & ParaGrid(Text1(4), 10, "Cliente")
        cad = cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
      If publicidad Then
        Tabla = NombreTabla & " INNER JOIN scafaccli1 ON scafaccli.codtipom='FPC' and scafaccli.codtipom=scafaccli1.codtipom AND scafaccli.numfactu=scafaccli1.numfactu AND scafaccli.fecfactu=scafaccli1.fecfactu "
      Else
        Tabla = NombreTabla & " INNER JOIN scafaccli1 ON scafaccli.codtipom=scafaccli1.codtipom AND scafaccli.numfactu=scafaccli1.numfactu AND scafaccli.fecfactu=scafaccli1.fecfactu "
      End If
        'CadenaConsulta = "select scafaccli.* from " & NombreTabla & " INNER JOIN scafaccli1 ON scafaccli.codtipom=scafaccli1.codtipom AND scafaccli.numfactu=scafacli1.numfactu AND scafaccli.fecfactu=scafaccli1.fecfactu "
        'CadenaConsulta = CadenaConsulta & " WHERE " & cadB & " GROUP BY scafaccli.codtipom,scafaccli.numfactu,scafaccli.fecfactu " & Ordenacion
        
        Titulo = "Facturas"
        devuelve = "0|1|2|"
    Else
        If vParamAplic.Departamento Then
            Titulo = "Dptos Cliente: "
            Desc = "Dpto."
        Else
            Titulo = "Direc. Cliente: "
            Desc = "Direc."
        End If
        Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N||15·"
        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35·"
        Tabla = "sdirec"
        devuelve = "0|1|"
    End If
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri  'Conexión a BD: Aritaxi
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If EsCabecera Then
            PonerCadenaBusqueda
            Text1(0).Text = Format(Text1(0).Text, "0000000")
        End If
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
        lblIndicador.Caption = ""
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        
        LLamaLineas Modo, 0, "DataGrid2"
        LLamaLineas Modo, 0, "DataGrid3"
        PonerCampos
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCamposLineas()
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafaccli1 de la factura seleccionada
Dim b As Boolean
Dim b2 As Boolean

    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafaccli1
    CargaGrid DataGrid2, Data3, True
    CargaGrid DataGrid3, Data4, True
    DataGrid1RowColChange
    
    'Comprobar si el albaran de la factura viene de una venta de ticket del TPV
    b = False
    b2 = False
    If Not Data3.Recordset.EOF Then
        If Not IsNull(Data3.Recordset!NumVenta) Then
            b = True
            If Data3.Recordset!codtipom = "FAV" And Data3.Recordset!codtipoa <> "FTI" Then b2 = True
        End If
    End If
    
    'Visualizar los campos de Oferta y Pedido si es una Factura q no es de venta TPV
    'o visulaizar numventa, numtermi si es una Factura de venta del TPV
    Label1(6).Caption = "Nº Pedido"
    Label1(18).Caption = "Fecha Pedido"
    If b Then
        If b2 Then
            Label1(6).Caption = "Nº Ticket"
            Label1(18).Caption = "Fecha Ticket"
        End If
        Label1(40).Caption = "Nº Terminal"
        Label1(22).Caption = "Nº Venta"
    Else
        Label1(40).Caption = "Nª Oferta"
        Label1(22).Caption = "Fecha Oferta"
    End If
    'sem. entrega
    Label1(2).visible = Not (b And b2)
    Text3(8).visible = Not (b And b2)
    'OFERTA
    Text3(4).visible = Not b
    Text3(5).visible = Not b
    'VENTA
    Text3(14).visible = b
    Text3(15).visible = b
    
    
    'Poner la referencia del cliente
  '  If Not data3.Recordset.EOF Then Text1(3).Text = DBLet(data3.Recordset.Fields!referenc, "T")
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Poner la base imponible (impbruto - dtoppago - dtognral
    BrutoFac = CSng(Text1(22).Text) - CSng(Text1(23).Text) - CSng(Text1(24).Text)
    Text1(25).Text = Format(BrutoFac, FormatoImporte)
    
    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    Text1_LostFocus (12) 'direc./dpto
    Text1_LostFocus (14) 'agente
    Text1_LostFocus (15) 'forma de pago
    Text1_LostFocus (49) 'banco propio
    Modo = 2
    
    PonerCamposLineas '
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    For i = 0 To txtAux.Count - 1
        Text1(i).BackColor = vbWhite
    Next i


    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
    ActualizarToolbar Modo, Kmodo
    Text1(3).visible = False  'SIEMPRE VISIBLE FALSE
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Modo = 2 Then
        If Not Data1.Recordset.EOF Then
            If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
          
    Me.Combo1.visible = (Modo = 1)
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
    Me.Check1(0).Enabled = (Modo = 1)
    Me.Check1(1).Enabled = (Modo = 1)
    
    b = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(3), b 'referencia
    BloquearTxt Text1(4), b 'cliente
    BloquearTxt Text1(12), b 'direccion
    BloquearTxt Text1(14), b 'agente
    BloquearTxt Text1(13), b 'direccion / departamento
    
    For i = 18 To 21
        BloquearTxt Text1(i), b And Modo <> 4
    Next i
    
    BloquearTxt Text1(48), b And Modo <> 4
    
    BloquearTxt Text1(22), b
    
    'Importes siempre bloqueados, excepto para busquedas. ivas y aportacion tb bloqueado
    For i = 26 To 37                            '[Monica]02/03/2012: dejamos modificar totales de la factura si es Teletaxi
        BloquearTxt Text1(i), (Modo <> 1) And Not (Modo = 4 And vParamAplic.Cooperativa = 0)
    Next i
    
    For i = 29 To 34
        BloquearTxt Text1(i), (Modo <> 1)
    Next i
    For i = 38 To 45
        BloquearTxt Text1(i), (Modo <> 1)
    Next i
    
    'Campo B.Imp y Imp. IVA siempre en azul
    BloquearTxt Text1(25), True
    Text1(25).BackColor = &HFFFFC0
    
    BloquearTxt Text1(23), True
    Text1(23).BackColor = &HFFFFC0
    BloquearTxt Text1(24), True
    Text1(24).BackColor = &HFFFFC0
    BloquearTxt Text1(47), True
    Text1(47).BackColor = &HFFFFC0
    
    If Modo <> 1 And Not (Modo = 4 And vParamAplic.Cooperativa = 0) Then
        Text1(35).BackColor = &HFFFFC0
        Text1(36).BackColor = &HFFFFC0
        Text1(37).BackColor = &HFFFFC0
'    Text1(38).BackColor = &HC0C0FF    'Total factura
        Text1(38).BackColor = &HC0FFC0
    End If
    
    
    
    'bloquear los Text3 que son las lineas de scafaccli1
    BloqueaText3
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 3 To 4
         BloquearTxt txtAux(i), (Modo <> 5) And vParamAplic.Cooperativa = 0
    Next i
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los albaranes de factura)
    For i = 0 To txtAux3.Count - 1
        BloquearTxt txtAux3(i), True '(Modo <> 1)
    Next i
    
    'ampliacion linea
    b = Me.DataGrid1.visible And Me.SSTab1.Tab = 1
    'Modo Linea de Albaranes
    'Me.Label1(35).visible = B
    'Me.Text2(16).visible = B
    Me.Label1(35).visible = b
    Me.Text2(16).visible = b
    
    BloquearTxt Text2(16), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)
    'nombre Proveedor
    Me.Label1(46).visible = (Modo = 5) And b
    Me.txtAux2(9).visible = (Modo = 5) And b


    'bloquear los txtaux4 que son las lineas de slifaccli
    BloqueaText4
    

    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    For i = 0 To 5
        Me.imgBuscar(i).Enabled = (Modo = 1)
        If i = 5 Then Me.imgBuscar(i).Enabled = (Modo = 1 Or Modo = 4)

    Next i
    For i = 6 To 9
        Me.imgBuscar(i).Enabled = False 'B And (Modo <> 1)
    Next i
    
    Me.imgBuscar(1).visible = False
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean

    On Error GoTo EDatosOK

    DatosOk = False
    
    ComprobarDatosTotales
    
    'comprobamos datos OK de la tabla scafaccli
    b = CompForm(Me, 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim i As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For i = 0 To txtAux4.Count - 1
        If i = 12 Then
            If txtAux4(i).Text = "" Then
                MsgBox "El campo " & txtAux4(i).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux4(i)
                Exit Function
            End If
        End If
    Next i
            
            
'    'PRoveedor
'    If txtAux(9).Text <> "" And txtAux2(9).Text = "" Then
'        MsgBox "Codigo proveedor incorrecto", vbExclamation
'        PonerFoco txtAux(9)
'        B = False
'        Exit Function
'    End If
'
    RecalcularImportes txtAux4(12), False




    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub


Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    If Modo = 1 Then Exit Sub
    Select Case Index
        Case 0, 1, 2 'trabajador
            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "straba", "nomtraba", "codtraba", "Cod. Trabajador", "N")
        Case 3 'cod. envio
            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "senvio", "nomenvio", "codenvio", "Cod. Envio", "N")
            If Screen.ActiveControl.TabIndex <> 27 Then PonerFocoBtn Me.cmdAceptar
        Case 13 'observa 5
            PonerFocoBtn Me.cmdAceptar
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnBuscar_Click  'Buscar
        Case 2: BotonVerTodos  'Todos
            

        Case 5: mnModificar_Click  'Modificar
        Case 6: mnEliminar_Click  'Borrar
        
        Case 9: mnLineas_Click  'Lineas
        Case 10: mnImprimir_Click 'Imprimir Albaran
        
        Case 11: mnImprimirAlbaran_Click
            
        Case 13: mnSalir_Click    'Salir
            
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub ActualizarToolbar(Modo As Byte, Kmodo As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner

    If (Modo = 5) And (Kmodo <> 5) Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 3
'        Toolbar1.Buttons(5).ToolTipText = "Nuevo Albaran"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 4
        Toolbar1.Buttons(5).ToolTipText = "Modificar Factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 5
        Toolbar1.Buttons(6).ToolTipText = "Eliminar Factura"
    End If
    If Kmodo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 12
'        Toolbar1.Buttons(5).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 13
        Toolbar1.Buttons(5).ToolTipText = "Modificar linea factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 14
        Toolbar1.Buttons(6).ToolTipText = "Eliminar linea factura"
    End If
End Sub
    
    
Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim Sql As String
Dim vWhere As String
Dim b As Boolean

    On Error GoTo EModificarLinea

    ModificarLinea = False
    If Data2.Recordset.EOF Then Exit Function
    
    vWhere = ObtenerWhereCP(True)
    If vParamAplic.Cooperativa = 1 Then
        vWhere = vWhere & " AND numlinea=" & Data4.Recordset.Fields!numlinea
    Else
        vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
    End If
    If vParamAplic.Cooperativa = 1 Then
        If DatosOkLinea() Then
            Sql = "UPDATE scafaccli_serv SET "
            Sql = Sql & "impventa = " & DBSet(txtAux4(12).Text, "N")
            Sql = Sql & vWhere
        End If
    Else
        Sql = "UPDATE slifaccli SET "
        Sql = Sql & "precioar = " & DBSet(txtAux(4).Text, "N")
        Sql = Sql & ", cantidad = " & DBSet(txtAux(3).Text, "N")
        Sql = Sql & ", importel = " & DBSet(txtAux(8).Text, "N")
        Sql = Sql & ", ampliaci = " & DBSet(Text2(16).Text, "T")
        Sql = Sql & vWhere & " and numalbar = " & Data2.Recordset.Fields!NumAlbar
        Sql = Sql & " and numlinea = " & DBSet(Data2.Recordset.Fields!numlinea, "N")
    End If
    
    If Sql <> "" Then
        'actualizar la factura y vencimientos
        b = ModificarFactura(Sql)
        
        ModificarLinea = b
    End If
    
EModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description
        b = False
    End If
    ModificarLinea = b
End Function


Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
    On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not b
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid1"
            Opcion = 1
        
        Case "DataGrid2"
            Opcion = 2
        
        Case "DataGrid3"
            Opcion = 3
    End Select
    
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
     b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
     vDataGrid.Enabled = Not b
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGrid

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
            'SQL = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea,
            'codalmac, codartic, nomartic, ampliaci, cantidad,numbultos, precioar, origpre, dtoline1, dtoline2, importel "
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(0)|T|Alm.|500|;S|txtAux(1)|T|Artículo|1600|;S|txtAux(2)|T|Nombre Art.|5050|;"
            tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|900|;N||||0|;S|txtAux(4)|T|Precio|1200|;S|txtAux(5)|T|OP|400|;S|txtAux(6)|T|Dto 1|600|;S|txtAux(7)|T|Dto 2|600|;S|txtAux(8)|T|Importe|2240|;"
            'TRAZA
'            tots = tots & "S|txtAux(9)|T|Prov.|750|;S|cmdaux|B|||;S|txtAux(10)|T|Nom. prov.|2000|;"
            If vEmpresa.TieneAnalitica Then
                'codprove,nomprove, codccost
'                tots = tots & "N||||0|;S|cmdaux|B|||;N||||0|;S|txtAux(9)|T|CCoste|750|;"
                tots = tots & "N||||0|;N||||0|;S|txtAux(9)|T|CCoste|750|;"

            Else
'                tots = tots & "N||||0|;S|cmdaux|B|||;N||||0|;N||||0|;"
                tots = tots & "N||||0|;N||||0|;N||||0|;"

            End If
            'numlote
            tots = tots & "N||||0|;"
            
            
            arregla tots, DataGrid1, Me
            DataGrid1.Columns(9).Alignment = dbgRight
            DataGrid1.Columns(10).Alignment = dbgRight
            DataGrid1.Columns(12).Alignment = dbgCenter
            DataGrid1.Columns(13).Alignment = dbgRight
            DataGrid1.Columns(14).Alignment = dbgRight
            DataGrid1.Columns(15).Alignment = dbgRight
                       
         Case "DataGrid2" 'albaranes x articulo
'             SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb,"
             'numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa  "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(0)|T|Tipo|600|;S|txtAux3(1)|T|Albaran|1100|;S|txtAux3(2)|T|Fecha|1200|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            arregla tots, DataGrid2, Me
                     
            DataGrid2_RowColChange 1, 1
            
         Case "DataGrid3" ' llamadas de la factura
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux4(3)|T|Fecha|1000|;S|txtAux4(4)|T|Hora|700|;S|txtAux4(5)|T|Socio|600|;"
            tots = tots & "S|txtAux4(6)|T|Uve|600|;S|txtAux4(13)|T|Nombre|3100|;"
            tots = tots & "S|txtAux4(7)|T|Direccion|1500|;S|txtAux4(8)|T|Número|1000|;S|txtAux4(9)|T|Puerta|1100|;S|txtAux4(11)|T|Tfno|1100|;"
            tots = tots & "S|txtAux4(10)|T|Identif.|800|;S|txtAux4(12)|T|Importe|1300|;N||||0|;N||||0|;"
            
            arregla tots, DataGrid3, Me
                     
'            DataGrid3_RowColChange 1, 1
            
            
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub TxtAux_Change(Index As Integer)
    If Index = 6 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
        txtAux(5).Text = "M"
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'Precio
             'Tipo 2: Decimal(10,4)
             If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 2
                txtAux(8).Text = txtAux(4).Text
             End If
            
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            If Index = 7 Then PonerFoco Me.Text2(16)
            
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
        Case 9
              txtAux(9).Text = Trim(txtAux(9).Text)
'              txtAux(10).Tag = ""
              If txtAux(9).Text <> "" Then
                    If Not IsNumeric(txtAux(9).Text) Then
                        MsgBox "Campo proveedor debe ser numérico", vbExclamation
                    Else
                        txtAux2(Index).Text = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtAux(9).Text)
                        If txtAux2(Index).Text = "" Then
                            MsgBox "No existe proveedor: " & txtAux(9).Text, vbExclamation
                            txtAux(9).Text = ""
                            PonerFoco txtAux(9)
                        End If
                    End If
                End If
'                txtAux(10).Text = txtAux(10).Tag
'                txtAux(10).Tag = ""
                
    End Select
    
    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
        If txtAux(1).Text = "" Then Exit Sub
'        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(8), 1
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
    
    Me.SSTab1.Tab = numTab
    If Me.DataGrid1.visible Then
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "La factura no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = cad
    End If
    If vUsu.Nivel >= 1 Then
        MsgBox "No tiene permiso", vbExclamation
        Exit Sub
    End If
    
    
    ModificaLineas = 0
    PonerModo 5
    PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim SQL2 As String
Dim vTipoMov As CTiposMov
Dim cContaFra As cContabilizarFacturas
    
    On Error GoTo FinEliminar

    b = False
    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
    ConnConta.BeginTrans
        
    'Eliminar en las tablas de la Contabilidad
    '------------------------------------------
    LEtra = ObtenerLetraSerie(Data1.Recordset!codtipom)
    
    Set cContaFra = New cContabilizarFacturas
    
    
    If LEtra <> "" Then
'        SQL = " numserie='" & LEtra & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
'        SQL = SQL & " AND anofaccl=" & Year(Data1.Recordset.Fields!FecFactu)
'
'        'Lineas
'        ConnConta.Execute "Delete from linfact WHERE " & SQL
'
'        'cabecera
'        ConnConta.Execute "Delete from cabfact WHERE " & SQL
'
        'cobros
        Sql = " numserie='" & LEtra & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
        Sql = Sql & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
        ConnConta.Execute "Delete from scobro WHERE " & Sql
        b = True
    Else
        b = False
    End If

    'Eliminar en tablas de factura de Aritaxi
    '------------------------------------------
    If b Then
        Sql = " " & ObtenerWhereCP(True)
    
        'Lineas de facturas (slifaccli)
        conn.Execute "Delete from slifaccli " & Sql
    
        'Lineas de cabeceras de albaranes de la factura
        conn.Execute "Delete from scafaccli1 " & Sql
        
        ' antes de eliminar los servicios de la factura hemos de desmarcarlos
        SQL2 = "update shilla, scafaccli_serv set facturadocliente = 0, facturad = 0 where shilla.fecha = scafaccli_serv.fecha "
        SQL2 = SQL2 & " and shilla.hora = scafaccli_serv.hora and shilla.numeruve = scafaccli_serv.numeruve "
        SQL2 = SQL2 & " and (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from scafaccli " & Sql
        SQL2 = SQL2 & ")"
        
        conn.Execute SQL2
        
        'Lineas de cabeceras de albaranes de la factura
        conn.Execute "Delete from scafaccli_serv " & Sql
        
        'Eliminar los vencimientos
        conn.Execute "Delete from svencicli " & Sql
        
        'Cabecera de facturas (scafaccli)
        conn.Execute "Delete from " & NombreTabla & Sql
        
        'Decrementar contador si borramos la ult. factura
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador Data1.Recordset!codtipom, Val(Text1(0).Text)
        Set vTipoMov = Nothing
    End If
    
'    b = True
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        ConnConta.RollbackTrans
        Eliminar = False
    Else
        If LEtra <> "" Then
            'Preparao para eliminar
            If cContaFra.EstablecerValoresInciales(ConnConta) Then
                Sql = CStr(Data1.Recordset!FecFactu)
                cContaFra.FijarNumeroFactura CLng(Data1.Recordset!NumFactu), Year(Data1.Recordset!FecFactu), LEtra
            End If
        End If
        
        'De ARIGES
        conn.CommitTrans
        ConnConta.CommitTrans
'
'        If cContaFra.RealizarContabilizacion Then
'            ConnConta.BeginTrans
'            'YA HE FIJADO LOS VALORES. En sql tengo la fecha factura
'            If cContaFra.EliminarFRACLIcontab(True, CDate(SQL)) Then
'                ConnConta.CommitTrans
'            Else
'                ConnConta.RollbackTrans
'            End If
'        End If
        Set cContaFra = Nothing
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid2, Data3, False
    CargaGrid DataGrid1, Data2, False
    CargaGrid DataGrid3, Data4, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
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
    
    Sql = " codtipom= '" & Text1(1).Text & "' and numfactu= " & Val(Text1(0).Text) & " and fecfactu='" & Format(Text1(2).Text, FormatoFecha) & "' "
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    Select Case Opcion
        Case 1
            Sql = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,numbultos, precioar, origpre, dtoline1, dtoline2, importel ,codprovex, nomprove,codccost,numlote"
            Sql = Sql & " FROM slifaccli left join sprove on codprovex=codprove " 'lineas de factura
        Case 2
            Sql = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb, numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa  "
            Sql = Sql & " FROM scafaccli1 " 'cabeceras albaranes de la factura
        Case 3
            Sql = "SELECT codtipom,numfactu,fecfactu,numlinea, fecha, hora, scafaccli_serv.codsocio, scafaccli_serv.numeruve, sclien.nomclien,  scafaccli_serv.dirllama, scafaccli_serv.numllama, scafaccli_serv.puerllama, scafaccli_serv.telefono, scafaccli_serv.idservic, scafaccli_serv.impventa, scafaccli_serv.observac2, scafaccli_serv.observa1  "
            Sql = Sql & " FROM scafaccli_serv inner join sclien on scafaccli_serv.codsocio = sclien.codclien " 'servicios de la factura
        
    End Select
    
    If enlaza Then
        Sql = Sql & " " & ObtenerWhereCP(True)
        If Opcion = 1 Then Sql = Sql & " AND numalbar=" & Data3.Recordset.Fields!NumAlbar
    Else
        'aNTES
        'SQL = SQL & " WHERE numfactu = -1 "
        'AHORA     Cambio sugerido por mangel para acelerar la entrada
        If Opcion <= 2 Then
            Sql = Sql & " WHERE codtipom is null and numfactu is null and fecfactu is null and codtipoa is null and numalbar is null "
        Else
            Sql = Sql & " WHERE codtipom is null and numfactu is null and fecfactu is null "
        End If
        If Opcion = 1 Then Sql = Sql & " AND numlinea is null"
    End If
    If Opcion <= 2 Then
        Sql = Sql & " ORDER BY codtipom, numfactu, fecfactu,numalbar "
    Else
        Sql = Sql & " ORDER BY codtipom, numfactu, fecfactu,numlinea "
    End If
    If Opcion = 1 Then Sql = Sql & ", numlinea "
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Modificar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(6).Enabled = (Modo = 2)
        Me.mnEliminar.Enabled = (Modo = 2)
            
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(9).Enabled = b
        Me.mnLineas.Enabled = b
        'Imprimir
        Toolbar1.Buttons(10).Enabled = b
        Me.mnImprimir.Enabled = b
        Toolbar1.Buttons(11).Enabled = b
        mnImprimirAlbaran.Enabled = b
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub



Private Sub PonerDatosCliente(CodClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim Observaciones As String
Dim b As Boolean

    On Error GoTo EPonerDatos
    
    If CodClien = "" Then
        LimpiarDatosCliente
        Exit Sub
    End If
    
    
    Set vCliente = New CCliente
    'si se ha modificado el cliente volver a cargar los datos
    If vCliente.Existe(CodClien) Then
        If vCliente.LeerDatos(CodClien) Then
            'si el cliente esta bloqueado salimos
            If vCliente.ClienteBloqueado Then
                If Modo = 3 Then
                    b = True
                ElseIf Modo = 4 Then
                     If (Val(Text1(4).Text) <> Val(Data1.Recordset!CodClien)) Then b = True
                End If
                If b Then
                    LimpiarDatosCliente
                    Set vCliente = Nothing
                    Exit Sub
                End If
            End If
            
'            EsDeVarios = vCliente.EsClienteVarios(Text1(4).Text)
            EsDeVarios = vCliente.DeVarios
            BloquearDatosCliente (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el cliente no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!CodClien) Then
                    Set vCliente = Nothing
                    Exit Sub
                End If
            End If
        
        
'            If Actualizar = False And EsDeVarios = False Then Exit Sub
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = Format(vCliente.Codigo, "000000")
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vCliente.Nombre  'Nom clien
                Text1(8).Text = vCliente.Domicilio
                Text1(9).Text = vCliente.CPostal
                Text1(10).Text = vCliente.Poblacion
                Text1(11).Text = vCliente.Provincia
                Text1(6).Text = vCliente.NIF
                Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            End If
            
            'insertar
            If Modo = 3 Then Text1(15).Text = vCliente.ForPago

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
            End If
                
            'cuenta bancaria
            Text1(48).Text = vCliente.Iban
            
            Text1(18).Text = vCliente.Banco
            FormateaCampo Text1(18)
            Text1(19).Text = vCliente.Sucursal
            FormateaCampo Text1(19)
            Text1(20).Text = vCliente.DigControl
            Text1(21).Text = vCliente.CuentaBan
            
            'Comprobar si el cliente tiene cobros pendientes
            ComprobarCobrosCliente2 CodClien, Text1(1).Text
        End If
    Else
        LimpiarDatosCliente
    End If
    Set vCliente = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub


Private Sub PonerDatosClienteVario(nifClien As String)
Dim vCliente As CCliente
Dim b As Boolean
   
    If nifClien = "" Then Exit Sub
   
    Set vCliente = New CCliente
    b = vCliente.LeerDatosCliVario(nifClien)
    Text1(5).Text = vCliente.Nombre  'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
    If Not b Then PonerFoco Text1(6)
    Set vCliente = Nothing
End Sub



Private Sub LimpiarDatosCliente()
Dim i As Byte
    
    For i = 4 To 13
        Text1(i).Text = ""
    Next i
    If (Modo = 3 Or Modo = 4) Then PonerFoco Text1(4)
End Sub
    
    
Private Sub BotonImprimir(OpcionListado As Byte)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim ImprimeDirecto As Boolean


    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Factura para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 52 'Facturas Clientes
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, ImprimeDirecto, pPdfRpt) Then Exit Sub
      
      
      
'    'PUNTO VERDE
'    '--------------------------------------------------------------------------
'    If vParamAplic.ArtReciclado <> "" Then
'        cadParam = cadParam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
'        numParam = numParam + 1
'    End If
      
    'Nombre fichero .rpt a Imprimir
    If Not ImprimeDirecto Then frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Factura
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Cod Tipo Movimiento
        devuelve = "{" & NombreTabla & ".codtipom}='" & Text1(1).Text & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        'Nº Factura
        devuelve = "{" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        cadSelect = cadFormula
        
        'Fecha Factura
        devuelve = "{" & NombreTabla & ".fecfactu}= Date(" & Year(Text1(2).Text) & "," & Month(Text1(2).Text) & "," & Day(Text1(2).Text) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Fecha Factura en cadSelect
        devuelve = "{" & NombreTabla & ".fecfactu}= '" & Format(Text1(2).Text, FormatoFecha) & "'"
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    End If
   
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
     
     If ImprimeDirecto Then
        'Imrpime directo
        If MsgBox("Imprimir la factura?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        ImprimirDirectoFact cadSelect
     Else
     
        '[Monica]31/03/2014: en el caso de teletaxi pedimos si imprime o no detalle
        If vParamAplic.Cooperativa = 0 And Text1(1).Text = "FAC" Then
            If MsgBox("¿ Desea imprimir el detalle de servicios ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                cadParam = cadParam & "pDetalle=0|"
            Else
                cadParam = cadParam & "pDetalle=1|"
            End If
            numParam = numParam + 1
        End If
        'hasta aquí
        
        devuelve = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
     
     
        With frmImprimir
                'Nuevo. Febrero 2010
                .outClaveNombreArchiv = devuelve & Format(Text1(0).Text, "000")
                .outCodigoCliProv = Text1(4).Text
                .outTipoDocumento = 2
                
                .FormulaSeleccion = cadFormula
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                .NombrePDF = pPdfRpt
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = OpcionListado
                .Titulo = "Facturas de Cliente"
                .Show vbModal
        End With
    End If
End Sub



Private Sub BotonImprimirTicket()
Dim MIPATH As String
Dim cadImpresion As String, Sql As String
Dim NomImpre As String
Dim NomImpTi As String
Dim bImpre As Boolean

    cadImpresion = "{scafaccli.codtipom}='" & Text1(1).Text & "' and {scafaccli.numfactu}=" & Text1(0).Text
    Sql = cadImpresion & " and {scafaccli.fecfactu}=" & DBSet(Text1(2).Text, "F")
    cadImpresion = cadImpresion & " and {scafaccli.fecfactu}=Date(" & Year(CDate(Text1(2).Text)) & "," & Month(CDate(Text1(2).Text)) & "," & Day(CDate(Text1(2).Text)) & ")"
    
    If Not HayRegParaInforme("scafaccli", Sql) Then Exit Sub
    
'    'Obtener que terminal es
'     'Terminal con el que trabajaremos, leemos el nombre del ordenador
'    SQL = ComputerName 'Nombre PC conectado por Terminal Server / local
'    SQL = DevuelveDesdeBDNew(conAri, "spatpvt", "numtermi", "nombrepc", SQL, "T")
'    If Not IsNumeric(SQL) Then
'        MsgBox "No se ha podido establecer la impresora de ticket." & vbCrLf & "Debe configurar primero los parámetros del TPV.", vbExclamation
'    Else
'        bImpre = True
'    End If
'
'    If bImpre Then
'         'Establecemos la impresora de ticket
'         NomImpTi = NombreImpresoraTicket(CInt(SQL))
'         If NomImpTi <> "" Then
'            If Printer.DeviceName <> NomImpTi Then
'                'guardamos la impresora que habia
'                NomImpre = Printer.DeviceName
'                'establecemos la de ticket
'                EstablecerImpresora NomImpTi
'            End If
'        End If
'    End If


    


    MIPATH = App.Path & "\Informes\"
'    cadImpresion = cadImpresion & " and {scafaccli.fecfactu}=Date(" & Year(RSVenta!fecventa) & "," & Month(RSVenta!fecventa) & "," & Day(RSVenta!fecventa) & ")"
    With frmVisReport
        .FormulaSeleccion = cadImpresion
        .SoloImprimir = False
        .OtrosParametros = ""
        .NumeroParametros = 0
        .MostrarTree = False
        .Informe = MIPATH & "rTPVTicket.rpt"
        .ConSubInforme = False
        .Opcion = 93
        .ExportarPDF = False
        .Show vbModal
   End With
   
'   If bImpre Then
'        'volver la impresora a la predeterminada
'        EstablecerImpresora NomImpre
'   End If
   
End Sub




Private Function ModificaAlbxFac() As Boolean
Dim Sql As String
Dim b As Boolean
    
    On Error GoTo EModificaAlb
    
    ModificaAlbxFac = False
    'comprobar datos OK de la scafaccli1
     b = CompForm(Me, 2) 'Comprobar formato datos ok de la cabecera alb: opcion=2
    If Not b Then Exit Function
    
    Sql = "UPDATE scafaccli1 SET codenvio=" & Text3(3).Text & ", "
    Sql = Sql & "codtraba=" & Text3(0).Text & ", "
    Sql = Sql & "codtrab1=" & DBSet(Text3(1).Text, "N", "S") & ", " 'Trab. pedido
    Sql = Sql & "codtrab2=" & Text3(2).Text & ", " 'Trab. Prep. Material
    Sql = Sql & "referenc=" & DBSet(Text3(16).Text, "T", "S") 'referencia cliente
    If Me.FrameObserva.visible Then
        Sql = Sql & ", observa1=" & DBSet(Text3(9).Text, "T")
        Sql = Sql & ", observa2=" & DBSet(Text3(10).Text, "T")
        Sql = Sql & ", observa3=" & DBSet(Text3(11).Text, "T")
        Sql = Sql & ", observa4=" & DBSet(Text3(12).Text, "T")
        Sql = Sql & ", observa5=" & DBSet(Text3(13).Text, "T")
    End If
    Sql = Sql & ObtenerWhereCP(True)
    Sql = Sql & " AND codtipoa='" & Data3.Recordset.Fields!codtipoa & "' AND numalbar=" & Data3.Recordset.Fields!NumAlbar
    conn.Execute Sql
    ModificaAlbxFac = True
    
EModificaAlb:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Albaranes de factura", Err.Description
End Function


Private Function ModificarFactura(Optional sqlLineas As String) As Boolean
'si se ha modificado la linea de slifaccli, añadir a la transaccion la modificación de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim Sql As String, LEtra As String
Dim vFactura As CFactura
Dim recalcular As Boolean
Dim ImporteServ As Currency
Dim ImpGastos As Currency
Dim Sql4 As String

    On Error GoTo EModFact

    
    'Comprobar si hay que recalcular la factura
    recalcular = False
    If sqlLineas <> "" Then
        'comprobamos si se ha modificado la linea del albaran (precio y descuentos)
        recalcular = True
    ElseIf CInt(Data1.Recordset!codforpa) <> CInt(Text1(15).Text) Then
        'si se ha cambiado la forma de pago
        recalcular = True
    ElseIf CSng(Data1.Recordset!DtoPPago) <> CSng(DBSet(Text1(16).Text, "N")) Then
        'si se ha cambiado el dto ppago
        recalcular = True
    ElseIf CSng(Data1.Recordset!DtoGnral) <> CSng(DBSet(Text1(17).Text, "N")) Then
        'si se ha cambiado el descuento general
        recalcular = True
    ElseIf CInt(Data1.Recordset!CodClien) <> CInt(Text1(4).Text) Then
        'si se ha cambiado el cliente (bonificara o no)
        recalcular = True
    ElseIf CSng(Data1.Recordset!TotalFac) <> CSng(Text1(38).Text) Then
        recalcular = True
    ElseIf CInt(DBLet(Data1.Recordset!codbanpr, "N")) <> CInt(ComprobarCero(Text1(49).Text)) Then
        recalcular = True
    End If
    
    
    bol = True
    conn.BeginTrans
    ConnConta.BeginTrans
    
    If recalcular Then
        If sqlLineas <> "" Then
            'actualizar el importe de la linea modificada
            MenError = "Modificando lineas de Factura."
            conn.Execute sqlLineas
        End If
        
        'recalcular las bases imponibles x IVA
        MenError = "Recalcular importes IVA"
        
        If vParamAplic.Cooperativa = 1 Then
            bol = CalcularDatosFactura()
        Else
            If sqlLineas <> "" Then
'                If Data2.Recordset!codArtic = vParamAplic.ArticServ Then
'
'                    Sql4 = "select sum(importel) from slifaccli where codtipom = " & DBSet(Text1(1).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and fecfactu = " & DBSet(Text1(2).Text, "F")
'                    Sql4 = Sql4 & " and codartic = " & DBSet(vParamAplic.ArticServ, "T")
'
'                    Text1(32).Text = DevuelveValor(Sql4) 'txtAux(4).Text
'                    CalcularDatosFactura2 32
'                Else
'                    Sql4 = "select sum(importel) from slifaccli where codtipom = " & DBSet(Text1(1).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and fecfactu = " & DBSet(Text1(2).Text, "F")
'                    Sql4 = Sql4 & " and codartic <> " & DBSet(vParamAplic.ArticServ, "T")
'
'                    Text1(33).Text = DevuelveValor(Sql4) 'txtAux(4).Text
'
'                    CalcularDatosFactura2 32
'                End If
                bol = CalcularDatosFactura()
            End If
        End If
    End If
    
    If bol Then
'        ComprobarDatosTotales
        
        'modificamos la scafaccli
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario(Me, 1)
        
        If bol Then
            'Si es cliente de varios actualizar datos cliente en tabla:sclvar
            MenError = "Modificando datos cliente varios"
            bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
        End If
        
        If bol Then
            MenError = "Modificando albaranes de factura"
            'modificar la tabla: scafaccli1
            bol = ModificaAlbxFac
            
            If bol And recalcular Then 'si se ha modificado la factura
                MenError = "Actualizando en Tesoreria"
                'borrar los vencimientos de aritaxi.svenci
                'y eliminar de tesoreria conta.scobros los registros de la factura(si existen en Tesoreria)
                
                'Eliminar los vencimientos
                '----------------------------------------
                Sql = ObtenerWhereCP(True)
                conn.Execute "Delete from svencicli " & Sql
                
                'Eliminar de Tesoreria
                '----------------------------------------
'                SQL = ObtenerLetraSerie(Text1(1).Text)
'                SQL = "SELECT COUNT(*) FROM scobro WHERE numserie='" & SQL & "' and codfaccl=" & Text1(0).Text
'                SQL = SQL & " AND fecfaccl=" & DBSet(Text1(2).Text, "F")
'
'                If RegistrosAListar(SQL, conConta) Then
                    'antes de Eliminar en las tablas de la Contabilidad
                Set vFactura = New CFactura
                If vFactura.LeerDatosFACcli(Text1(1).Text, Text1(0).Text, Text1(2).Text) Then
                
                Else
                  bol = False
                End If
              
                If bol Then
                    'Eliminar de la scobro
                    Sql = " numserie='" & vFactura.LetraSerie & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
                    Sql = Sql & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                    ConnConta.Execute "Delete from scobro WHERE " & Sql
                    bol = True

                    'Volvemos a Insertar los Vencimientos de la Factura. Tabla: svenci
                    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
                    If bol Then
                        vFactura.Agente = Text1(14).Text
                        vFactura.ForPago = Text1(15).Text
                        vFactura.Cliente = Text1(4).Text
                        vFactura.NombreClien = Text1(5).Text
                        vFactura.DomicilioClien = Text1(8).Text
                        vFactura.CPostal = Text1(9).Text
                        vFactura.Poblacion = Text1(10).Text
                        vFactura.Provincia = Text1(11).Text
                        vFactura.NIF = Text1(6).Text
                        
                        If Text1(49).Text <> "" Then vFactura.CuentaPrev = DevuelveValor("select codmacta from sbanpr where codbanpr = " & DBSet(Text1(49).Text, "N"))
                        
                        bol = vFactura.InsertarEnTesoreriaFACcli("", MenError)
                    End If
                End If
                Set vFactura = Nothing
            End If
        End If
    End If

EModFact:
     If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ConnConta.CommitTrans
        ModificarFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        ModificarFactura = False
        MenError = "Actualizando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
    End If
End Function



Private Function CalcularDatosFactura() As Boolean
Dim i As Integer
Dim fac As CFacturaCli
Dim FacOK As Boolean
Dim iva As String
Dim vDevuelve As String
Dim porIvaServ As Currency
Dim porIvaGtos As Currency
Dim BaseivaGtos As Currency
Dim ImporteServ As Currency
Dim ImpGastos As Currency
Dim ImpivaGtos As Currency
Dim BaseivaServ As Currency
Dim ImpivaServ As Currency
Dim Sql As String

    If vParamAplic.Cooperativa = 1 Then
    
        Sql = "select sum(impventa) from scafaccli_serv "
        Sql = Sql & " where scafaccli_serv.codtipom = " & DBSet(Text1(1).Text, "T")
        Sql = Sql & " and scafaccli_serv.numfactu = " & DBSet(Text1(0).Text, "N")
        Sql = Sql & " and scafaccli_serv.fecfactu = " & DBSet(Text1(2).Text, "F") & ""
    
    Else
    
        Sql = "select sum(importel) from slifaccli "
        Sql = Sql & " where slifaccli.codtipom = " & DBSet(Text1(1).Text, "T")
        Sql = Sql & " and slifaccli.numfactu = " & DBSet(Text1(0).Text, "N")
        Sql = Sql & " and slifaccli.fecfactu = " & DBSet(Text1(2).Text, "F") & ""
        Sql = Sql & " and slifaccli.codartic = " & DBSet(vParamAplic.ArticServ, "T")
    
    
    End If
    
    ImporteServ = DevuelveValor(Sql)
    
    Sql = "select sum(importel) from slifaccli "
    Sql = Sql & " where slifaccli.codtipom = " & DBSet(Text1(1).Text, "T")
    Sql = Sql & " and slifaccli.numfactu = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and slifaccli.fecfactu = " & DBSet(Text1(2).Text, "F") & ""
    Sql = Sql & " and slifaccli.codartic = " & DBSet(vParamAplic.ArtGastosAdmon, "T")
    
    ImpGastos = DevuelveValor(Sql)
    
    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 22 To 38
         Text1(i).Text = ""
    Next i
    
    Set fac = New CFacturaCli
    fac.DtoPPago = CCur(ComprobarCero(Text1(16).Text))
    fac.DtoGnral = CCur(ComprobarCero(Text1(17).Text))
    fac.Cliente = Text1(4).Text
    
    If vParamAplic.Cooperativa = 0 Then
    
        Sql = "scafaccli.codtipom = " & DBSet(Text1(1).Text, "T")
        Sql = Sql & " and scafaccli.numfactu = " & DBSet(Text1(0).Text, "N")
        Sql = Sql & " and scafaccli.fecfactu = " & DBSet(Text1(2).Text, "F") & ""

    
        fac.CalcularDatosFactura True, Sql, "scafaccli", "slifaccli", CDate(Text1(2).Text) <= vParamAplic.FecCambioIva
    
    Else
        ' calculo de bases iva de SERVICIOS
        iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArticServ, "T")
        vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T")
        porIvaServ = 0
        If vDevuelve <> "" Then porIvaServ = CCur(vDevuelve)
        fac.TipoIVA1 = iva
        
    '[Monica]07/11/2012: aqui la linea no tiene incluido el iva
    '    BaseivaServ = Round2((ImporteServ + 0) / (1 + (porIvaServ / 100)), 2)
    '    ImpivaServ = Round2(ImporteServ - BaseivaServ, 2)
        BaseivaServ = ImporteServ
        ImpivaServ = Round2(BaseivaServ * porIvaServ / 100, 2)
        
        ' calculo de base iva de GASTOS ADMON
        iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArtGastosAdmon, "T")
        vDevuelve = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", CStr(iva), "T")
        porIvaGtos = 0
        If vDevuelve <> "" Then porIvaGtos = CCur(vDevuelve)
        
        BaseivaGtos = ImpGastos
        ImpivaGtos = Round2(BaseivaGtos * porIvaGtos / 100, 2)
        
        ' Asignamos los importes a la factura
        If BaseivaGtos <> 0 Then
            If iva = fac.TipoIVA1 Then
                fac.TipoIVA1 = iva
                fac.BaseIVA1 = BaseivaGtos
                fac.PorceIVA1 = porIvaGtos
                fac.ImpIVA1 = ImpivaGtos
            Else
                fac.TipoIVA2 = iva
                fac.BaseIVA2 = BaseivaGtos
                fac.PorceIVA2 = porIvaGtos
                fac.ImpIVA2 = ImpivaGtos
            End If
        End If
        'el tipo de iva 1 esta asignado cuando se busca en tiposiva de la conta
        fac.PorceIVA1 = porIvaServ
        fac.BaseIVA1 = fac.BaseIVA1 + BaseivaServ
        fac.ImpIVA1 = fac.ImpIVA1 + ImpivaServ
        
        fac.BaseImp = BaseivaServ + BaseivaGtos
        fac.BrutoFac = fac.BaseImp
        fac.TotalFac = BaseivaServ + ImpivaServ + BaseivaGtos + ImpivaGtos
        
        fac.codtipom = "FAC"
        
        fac.FecFactu = Text1(2).Text
        fac.LetraSerie = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", fac.codtipom, "T")
        
        
    '    Sql = "update slifaccli set importel = " & DBSet(BaseivaServ, "N")
    '    Sql = Sql & ", precioar = " & DBSet(BaseivaServ, "N")
    '    Sql = Sql & " where slifaccli.codtipom = " & DBSet(Text1(1).Text, "T")
    '    Sql = Sql & " and slifaccli.numfactu = " & DBSet(Text1(0).Text, "N")
    '    Sql = Sql & " and slifaccli.fecfactu = " & DBSet(Text1(2).Text, "F")
    '    Sql = Sql & " and slifaccli.codartic = " & DBSet(vParamAplic.ArticServ, "T")
    '
    '    conn.Execute Sql
        
    End If
        
    FacOK = True
    
    Text1(22).Text = fac.BrutoFac
    Text1(23).Text = fac.ImpPPago
    Text1(24).Text = fac.ImpGnral
    Text1(25).Text = fac.BaseImp
    Text1(26).Text = QuitarCero(fac.TipoIVA1)
    Text1(27).Text = QuitarCero(fac.TipoIVA2)
    Text1(28).Text = QuitarCero(fac.TipoIVA3)
    Text1(29).Text = fac.PorceIVA1
    Text1(30).Text = fac.PorceIVA2
    Text1(31).Text = fac.PorceIVA3
    Text1(32).Text = fac.BaseIVA1
    Text1(33).Text = fac.BaseIVA2
    Text1(34).Text = fac.BaseIVA3
    Text1(35).Text = fac.ImpIVA1
    Text1(36).Text = fac.ImpIVA2
    Text1(37).Text = fac.ImpIVA3
    Text1(38).Text = fac.TotalFac
    
    FormatoDatosTotales

    Set fac = Nothing
    CalcularDatosFactura = FacOK

End Function

Private Sub FormatoDatosTotales()
Dim i As Byte

    For i = 22 To 25
        Text1(i).Text = QuitarCero(Text1(i).Text)
        Text1(i).Text = Format(Text1(i).Text, FormatoImporte)
    Next i
    
    'Desglose B.Imponible por IVA
    For i = 32 To 34
        If Text1(i).Text <> "" Then
             If CSng(Text1(i).Text) = 0 And Text1(i - 6).Text = "" Then
                Text1(i).Text = QuitarCero(Text1(i).Text)
                Text1(i - 3).Text = QuitarCero(Text1(i - 3).Text)
                Text1(i - 6).Text = QuitarCero(Text1(i - 6).Text)
                Text1(i + 3).Text = QuitarCero(Text1(i).Text)
            Else
                Text1(i).Text = Format(Text1(i).Text, FormatoImporte)
                Text1(i - 3) = Format(Text1(i - 3).Text, FormatoDescuento)
    '            Text3(i - 6) = Format(Text3(i - 6).Text, "000")
                Text1(i + 3).Text = Format(Text1(i + 3).Text, FormatoImporte)
            End If
        End If
    Next i
End Sub



Private Sub ComprobarDatosTotales()
Dim i As Byte

    For i = 22 To 25
        Text1(i).Text = ComprobarCero(Text1(i).Text)
    Next i
End Sub


Private Function FactContabilizada2(ByRef EstaEnTesoreria As String) As Boolean
Dim LEtra As String, numasien As String
    
    On Error GoTo EContab
    
    
    
    
    'NO deberia poder modificar fras anteriors a fecha inicio ejercicio
    
    
    'Cojo la letra de serie
    LEtra = ObtenerLetraSerie(Text1(1).Text)
    
    'Primero comprobaremos que esta el cobro en contabilidad
    EstaEnTesoreria = ""
    If Not ComprobarCobroArimoney(EstaEnTesoreria, LEtra, CLng(Text1(0).Text), CDate(Text1(2).Text)) Then
        FactContabilizada2 = True
        Exit Function
    End If




    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1(0).Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
      
        If LEtra <> "" Then
            numasien = DevuelveDesdeBDNew(conConta, "cabfact", "numasien", "numserie", LEtra, "T", , "codfaccl", Text1(0).Text, "N", "anofaccl", Year(Text1(2).Text), "N")
            If Val(ComprobarCero(numasien)) <> 0 Then
'                FactContabilizada = True
'                MsgBox "La factura esta contabilizada y no se puede modificar.", vbInformation
'                Exit Function
                
            Else
                numasien = ""
            End If
            
            
            
            
        Else
'            MsgBox "Las factura de venta no tienen asignada una letra de serie", vbInformation
            numasien = ""
        End If
        
        LEtra = "La factura esta en la contabilidad"
        If numasien <> "" Then LEtra = LEtra & vbCrLf & "Nº asiento: " & numasien
        LEtra = LEtra & vbCrLf & vbCrLf & "¿Continuar?"
        
        numasien = String(50, "*") & vbCrLf
        numasien = numasien & numasien & vbCrLf & vbCrLf
        LEtra = numasien & LEtra & vbCrLf & vbCrLf & numasien
        If MsgBox(LEtra, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            FactContabilizada2 = False
        Else
            FactContabilizada2 = True
        End If
    Else
        FactContabilizada2 = False
    End If
    
    
    
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function

Private Function FactContabilizada3(ByRef EstaEnTesoreria As String) As Boolean
Dim LEtra As String, numasien As String
    
    On Error GoTo EContab
    
    'NO deberia poder modificar fras anteriors a fecha inicio ejercicio
    
    
    'Cojo la letra de serie
    LEtra = ObtenerLetraSerie(Text1(1).Text)
    
    'Primero comprobaremos que esta el cobro en contabilidad
    EstaEnTesoreria = ""
    If Not ComprobarCobroArimoney(EstaEnTesoreria, LEtra, CLng(Text1(0).Text), CDate(Text1(2).Text)) Then
        FactContabilizada3 = True
        Exit Function
    End If

    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1(0).Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        LEtra = "La factura esta en la contabilidad. No se permite modificar ni eliminar."
        MsgBox LEtra, vbInformation
      
        FactContabilizada3 = True
        Exit Function
    Else
    
        FactContabilizada3 = False
        Exit Function
    End If

EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function



Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
End Sub


Private Sub BloquearDatosCliente(bol As Boolean)
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol
        Me.imgBuscar(1).Enabled = bol
        Me.imgBuscar(2).Enabled = bol
        
        For i = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(i), Not bol
        Next i
    End If
End Sub


Private Function ActualizarClienteVarios(clien As String, NIF As String) As Boolean
Dim vCliente As CCliente
    
    On Error GoTo EActualizarCV

    ActualizarClienteVarios = False
    
    Set vCliente = New CCliente
    If EsClienteVarios(clien) Then
        vCliente.NIF = NIF
        vCliente.Nombre = Text1(5).Text
        vCliente.Domicilio = Text1(8).Text
        vCliente.CPostal = Text1(9).Text
        vCliente.Poblacion = Text1(10).Text
        vCliente.Provincia = Text1(11).Text
        vCliente.TfnoClien = Text1(7).Text
        vCliente.ActualizarClienteV (NIF)
    End If
    Set vCliente = Nothing
    
    ActualizarClienteVarios = True
    
EActualizarCV:
    If Err.Number <> 0 Then
        ActualizarClienteVarios = False
    Else
        ActualizarClienteVarios = True
    End If
End Function



Private Function ObtenerSelFactura() As String
Dim cad As String
Dim RS As ADODB.Recordset

    On Error Resume Next

    cad = ""
    If Me.DesdeFichaCliente Then
        '
        cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
        
    Else
        'Tengo YA el codigo de la factura
                '******************************************************
                'laura: esto se puede comentar, ya no hay movimiento FTI en la smoval
                If hcoCodTipoM = "FTI" Then
                    'no hay albaran directamente va a factura de ticket
                    
                    'ver si lo encontramos como factura: codtipom, numfactu,fecfactu
                    cad = "SELECT COUNT(*) FROM scafaccli "
                    cad = cad & " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    If RegistrosAListar(cad) > 0 Then
                        cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    Else
                        cad = ""
                    End If
                Else
                    If hcoCodTipoM = "FAM" Then
                        cad = "  WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    End If
                End If
                '******************************************************
                    
                If cad = "" Then
                    'En la smoval estaba e mov. de ALbaran
                    cad = "SELECT codtipom,numfactu,fecfactu FROM scafaccli1 "
                    cad = cad & " WHERE codtipoa=" & DBSet(hcoCodTipoM, "T") & " AND numalbar=" & hcoCodMovim & " AND fechaalb=" & DBSet(hcoFechaMov, "F")
                    
                    Set RS = New ADODB.Recordset
                    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not RS.EOF Then 'where para la factura
                        cad = " WHERE codtipom='" & RS!codtipom & "' AND numfactu= " & RS!NumFactu & " AND fecfactu=" & DBSet(RS!FecFactu, "F")
                    Else
                        cad = " WHERE numfactu=-1"
                    End If
                    RS.Close
                    Set RS = Nothing
                End If
    
    End If
    ObtenerSelFactura = cad
End Function



Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.Codigo = Text1(4).Text
    'si existe el departamento para el cliente
    If vClien.DptoCliente(Text1(12).Text, NomDpto) Then
        Text1(13).Text = NomDpto
        PonerDptoEnCliente = True
    Else
        PonerDptoEnCliente = False
    End If
    Set vClien = Nothing
End Function





Private Sub ImprimirAlbaran(OpcionListado As Byte)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String



    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 42
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, False, pPdfRpt) Then Exit Sub
      
      
'--[Monica]arituclo de reciclado pasa a ser reutilizado como articulo de gastos de administracion
'    'PUNTO VERDE
'    '--------------------------------------------------------------------------
'    If vParamAplic.ArtReciclado <> "" Then
'        cadParam = cadParam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
'        numParam = numParam + 1
'    End If
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Factura
    '---------------------------------------------------
    
        'Cod Tipo Movimiento
        devuelve = "{" & NombreTabla & ".codtipom}='" & Text1(1).Text & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        'Nº Factura
        devuelve = "{" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        
        
        'cODTIPOA
        devuelve = "{scafaccli1.codtipoa}=" & DBSet(Data3.Recordset!codtipoa, "T")
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        'Numalbar
'        devuelve = "{scafaccli1.numalbar}=" & DBSet(Data3.Recordset!NumAlbar, "N")
'        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        
        
        cadSelect = cadFormula
        'Fecha Factura
        devuelve = "{" & NombreTabla & ".fecfactu}= Date(" & Year(Text1(2).Text) & "," & Month(Text1(2).Text) & "," & Day(Text1(2).Text) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Fecha Factura en cadSelect
        devuelve = "{" & NombreTabla & ".fecfactu}= '" & Format(Text1(2).Text, FormatoFecha) & "'"
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
   
        'If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
        '=========================================================================
        'ipo de IVA
        'que se aplica a ese cliente
        devuelve = DevuelveDesdeBDNew(conAri, "scliente", "tipoiva", "codclien", Text1(4).Text, "N")
        If devuelve <> "" Then
            cadParam = cadParam & "pTipoIVA= " & devuelve & "|"
            numParam = numParam + 1
        End If
         
     
     
        
        devuelve = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
     
     
         With frmImprimir
                'Nuevo. Febrero 2010
                '.outClaveNombreArchiv = devuelve & Format(Text1(0).Text, "000")
                '.outCodigoCliProv = Text1(4).Text
                '.outTipoDocumento = 2
                
                .outClaveNombreArchiv = ""
                .outCodigoCliProv = 0
                .outTipoDocumento = 0
                
                
                
                .FormulaSeleccion = cadFormula
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                .NombrePDF = pPdfRpt
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = 45
                .Titulo = "Albarán facturado"
                .Show vbModal
        End With
    
End Sub


'En vTesoreria pondremos como estaba el recibo
'Es decir. El  msgbox que pondra al final lo guardo en esta variable
Private Function ComprobarCobroArimoney(vTesoreria As String, LEtra As String, Codfaccl As Long, Fecha As Date) As Boolean
Dim vR As ADODB.Recordset
Dim cad As String


On Error GoTo EComprobarCobroArimoney
    ComprobarCobroArimoney = False
    Set vR = New ADODB.Recordset
    cad = "Select * from scobro where numserie='" & LEtra & "'"
    cad = cad & " AND codfaccl =" & Codfaccl
    cad = cad & " AND fecfaccl =" & DBSet(Fecha, "F")
    
    '
    vTesoreria = ""
    vR.Open cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    If vR.EOF Then
        vTesoreria = "NO se ha encotrado ningun vencimiento en la tesoreria"
    Else
        While Not vR.EOF
            cad = ""
            If DBLet(vR!codrem, "T") <> "" Then
                cad = "El cobro asociado a la factura esta remesado(" & vR!codrem & ")"
            Else
                If DBLet(vR!recedocu, "N") = 1 Then
                    cad = "Documento recibido"
                Else
                    If DBLet(vR!Estacaja, "N") = 1 Then
                        cad = "Cobrado por caja"
                    Else
                        If DBLet(vR!transfer, "N") = 1 Then
                            cad = "Esta en una transferencia"
                        Else
                           If DBLet(vR!impcobro, "N") > 0 Then cad = "Esta parcialmente cobrado: " & vR!impcobro
                        
                            
                                    'Si hubeira que poner mas coas iria aqui
                        End If 'transfer
                    End If 'estacaja
                End If 'recdedocu
            End If 'remesado
            If cad <> "" Then vTesoreria = vTesoreria & "Vto: " & vR!numorden & "      " & cad & vbCrLf
            vR.MoveNext
        Wend
    End If
    vR.Close
    
    
    
    If vTesoreria <> "" Then
        cad = vTesoreria & vbCrLf & vbCrLf
        If vUsu.Nivel > 1 Then
            MsgBox cad, vbExclamation
        Else
            cad = cad & "¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then ComprobarCobroArimoney = True
        End If
    Else
        ComprobarCobroArimoney = True
    End If
            
EComprobarCobroArimoney:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set vR = Nothing
End Function


Private Sub TxtAux4_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux4(Index)
End Sub

Private Sub TxtAux4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub TxtAux4_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub TxtAux4_LostFocus(Index As Integer)

    'Quitar espacios en blanco
    'If Not PerderFocoGnralLineas(txtAux4(Index), ModificaLineas) Then Exit Sub
    If Not PerderFocoGnralLineas(txtAux4(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 12 'Importe Linea
            If PonerFormatoDecimal(txtAux4(Index), 3) Then   'Tipo 3: Decimal(10,2)
                cmdAceptar.SetFocus
            End If
    End Select
    
End Sub

Private Sub CargaCombo()
Dim RS As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    Combo1.Clear
    
    '[Monica]11/02/2011: todo tipo de facturas excepto las de liquidacion,publicidad y cuotas de socio
    '                    y las facturas de cliente FAC y FPC
    Sql = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom in ('FAC','FRN','FVC')"
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Sql = RS!nomtipom
        Sql = Replace(Sql, "Factura", "")
        Combo1.AddItem RS!codtipom & "-" & Sql
        Combo1.ItemData(Combo1.NewIndex) = i
        i = i + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub

