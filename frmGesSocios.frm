VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGesSocios 
   Caption         =   "Socios."
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   4200
      Top             =   6510
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   5490
      Top             =   6510
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   41
      Top             =   1500
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   7
      Tab             =   4
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   -2147483624
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "frmGesSocios.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "imgFecha(1)"
      Tab(0).Control(1)=   "Label1(14)"
      Tab(0).Control(2)=   "imgFecha(0)"
      Tab(0).Control(3)=   "Label1(13)"
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(5)=   "ImgMail(1)"
      Tab(0).Control(6)=   "imgBuscar(2)"
      Tab(0).Control(7)=   "Label1(11)"
      Tab(0).Control(8)=   "Label1(10)"
      Tab(0).Control(9)=   "Label1(9)"
      Tab(0).Control(10)=   "Label1(8)"
      Tab(0).Control(11)=   "Label1(7)"
      Tab(0).Control(12)=   "Label1(6)"
      Tab(0).Control(13)=   "imgBuscar(0)"
      Tab(0).Control(14)=   "Label1(5)"
      Tab(0).Control(15)=   "Label1(3)"
      Tab(0).Control(16)=   "Label1(19)"
      Tab(0).Control(17)=   "Label1(20)"
      Tab(0).Control(18)=   "Label14"
      Tab(0).Control(19)=   "imgDoc(1)"
      Tab(0).Control(20)=   "imgDoc(0)"
      Tab(0).Control(21)=   "Label2"
      Tab(0).Control(22)=   "Label1(21)"
      Tab(0).Control(23)=   "Text1(23)"
      Tab(0).Control(24)=   "Frame3(1)"
      Tab(0).Control(25)=   "Check1(0)"
      Tab(0).Control(26)=   "Text1(17)"
      Tab(0).Control(27)=   "Text1(13)"
      Tab(0).Control(28)=   "Text1(12)"
      Tab(0).Control(29)=   "Text1(10)"
      Tab(0).Control(30)=   "Text1(9)"
      Tab(0).Control(31)=   "Text1(8)"
      Tab(0).Control(32)=   "Text1(6)"
      Tab(0).Control(33)=   "Text1(5)"
      Tab(0).Control(34)=   "Text1(4)"
      Tab(0).Control(35)=   "Text1(3)"
      Tab(0).Control(36)=   "Text1(7)"
      Tab(0).Control(37)=   "Text1(22)"
      Tab(0).Control(38)=   "Text1(21)"
      Tab(0).Control(39)=   "Text1(20)"
      Tab(0).Control(40)=   "Text1(19)"
      Tab(0).Control(41)=   "Text1(25)"
      Tab(0).Control(42)=   "Text1(27)"
      Tab(0).Control(43)=   "Check1(1)"
      Tab(0).Control(44)=   "Text1(28)"
      Tab(0).Control(45)=   "Check1(2)"
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "Equipamiento"
      TabPicture(1)   =   "frmGesSocios.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LabelCRM"
      Tab(1).Control(1)=   "lwCRM"
      Tab(1).Control(2)=   "cmdAccCRM(2)"
      Tab(1).Control(3)=   "cmdAccCRM(1)"
      Tab(1).Control(4)=   "cmdAccCRM(0)"
      Tab(1).Control(5)=   "Frame3(2)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Choferes"
      TabPicture(2)   =   "frmGesSocios.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid1"
      Tab(2).Control(1)=   "txtAux1(0)"
      Tab(2).Control(2)=   "txtAux1(1)"
      Tab(2).Control(3)=   "txtAux1(2)"
      Tab(2).Control(4)=   "txtAux1(3)"
      Tab(2).Control(5)=   "txtAux1(4)"
      Tab(2).Control(6)=   "cmdAux(0)"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Publicidad"
      TabPicture(3)   =   "frmGesSocios.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DataGrid2"
      Tab(3).Control(1)=   "txtAux2(0)"
      Tab(3).Control(2)=   "txtAux2(1)"
      Tab(3).Control(3)=   "txtAux2(2)"
      Tab(3).Control(4)=   "txtAux2(3)"
      Tab(3).Control(5)=   "cmdAux1"
      Tab(3).Control(6)=   "txtAux2(4)"
      Tab(3).Control(7)=   "cmdAux(1)"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Cuotas"
      TabPicture(4)   =   "frmGesSocios.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "DataGrid4"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Text3"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtAux3(0)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "txtAux3(1)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txtAux3(2)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cmdAux(2)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "Documentos"
      TabPicture(5)   =   "frmGesSocios.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "LabelDoc"
      Tab(5).Control(1)=   "imgFecha(3)"
      Tab(5).Control(2)=   "Label3"
      Tab(5).Control(3)=   "Frame3(0)"
      Tab(5).Control(4)=   "lw1"
      Tab(5).Control(5)=   "Toolbar2"
      Tab(5).Control(6)=   "Text1(26)"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "Contadores"
      TabPicture(6)   =   "frmGesSocios.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "DataGrid3"
      Tab(6).Control(1)=   "txtAux4(2)"
      Tab(6).Control(2)=   "txtAux4(1)"
      Tab(6).Control(3)=   "txtAux4(0)"
      Tab(6).ControlCount=   4
      Begin VB.CheckBox Check1 
         Caption         =   "Facturación Elec."
         Height          =   375
         Index           =   2
         Left            =   -65580
         TabIndex        =   18
         Tag             =   "Facturacion Electrónica|N|N|0|1|sclien|facturae|||"
         Top             =   1920
         Width           =   1545
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   28
         Left            =   -67050
         MaxLength       =   4
         TabIndex        =   19
         Tag             =   "IBAN|T|S|||sclien|iban|||"
         Text            =   "9999"
         Top             =   2340
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Es Contado"
         Height          =   375
         Index           =   1
         Left            =   -65580
         TabIndex        =   17
         Tag             =   "Facturado|N|N|0|1|sclien|escontado|||"
         Top             =   1560
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   27
         Left            =   -67080
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Importe a Cuenta|N|S|||sclien|impacuenta|###,##0.00||"
         Top             =   1950
         Width           =   1155
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74100
         MaxLength       =   16
         TabIndex        =   103
         Tag             =   "Tipom Mov.|T|N|||sclien_contadores|codtipom|||"
         Text            =   "tipo"
         Top             =   4320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73020
         TabIndex        =   104
         Text            =   "nomtipom"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtAux4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -70950
         MaxLength       =   10
         TabIndex        =   105
         Tag             =   "Contador|N|N|||sclien_contadores|contador|0000000||"
         Text            =   "Contador"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdAux 
         Caption         =   "+"
         Height          =   315
         Index           =   2
         Left            =   2430
         TabIndex        =   98
         ToolTipText     =   "Buscar artículo"
         Top             =   3570
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   4770
         MaxLength       =   10
         TabIndex        =   97
         Text            =   "Importe"
         Top             =   3570
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   2670
         TabIndex        =   96
         Text            =   "nomartic"
         Top             =   3570
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   780
         MaxLength       =   16
         TabIndex        =   95
         Tag             =   "Artículo|T|N|||sclien_cuotas|codartic|||"
         Text            =   "artic"
         Top             =   3570
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   4530
         Width           =   2100
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   26
         Left            =   -65550
         TabIndex        =   90
         Text            =   "Text4"
         Top             =   1410
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4215
         Index           =   2
         Left            =   -74850
         TabIndex        =   83
         Top             =   390
         Width           =   615
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   390
            Left            =   30
            TabIndex        =   84
            Top             =   300
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Appearance      =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   11
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Equipamientos"
                  Object.Tag             =   "0"
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Llamadas"
                  Object.Tag             =   "1"
                  Style           =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Correo electronico"
                  Object.Tag             =   "2"
                  Style           =   2
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Cobros"
                  Object.Tag             =   "3"
                  Style           =   2
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Observaciones departamento"
                  Object.Tag             =   "4"
                  Style           =   2
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Descuento familia/Marca"
                  Object.Tag             =   "5"
                  Style           =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   0
         Left            =   -65520
         Picture         =   "frmGesSocios.frx":00C4
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Acciones CRM"
         Top             =   330
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   1
         Left            =   -64470
         Picture         =   "frmGesSocios.frx":0AC6
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Impresion CRM"
         Top             =   330
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   2
         Left            =   -65040
         Picture         =   "frmGesSocios.frx":1050
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Eliminar"
         Top             =   330
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   25
         Left            =   -67080
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "Sumplemento Cuota|N|S|||sclien|suplecuota|###,##0.00||"
         Top             =   1560
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   19
         Left            =   -66510
         MaxLength       =   4
         TabIndex        =   20
         Tag             =   "Codigo Banco|N|N|||sclien|codbanco|0000||"
         Text            =   "9999"
         Top             =   2340
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   20
         Left            =   -65970
         MaxLength       =   4
         TabIndex        =   21
         Tag             =   "Codigo Sucursal|N|N|||sclien|codsucur|0000||"
         Text            =   "9999"
         Top             =   2340
         Width           =   465
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   21
         Left            =   -65460
         MaxLength       =   2
         TabIndex        =   22
         Tag             =   "Digito Control|T|N|||sclien|digcontr|00||"
         Text            =   "99"
         Top             =   2340
         Width           =   285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   22
         Left            =   -65130
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "Cuenta Banco|T|N|||sclien|cuentaba|0000000000||"
         Text            =   "9999999999"
         Top             =   2340
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   -67080
         MaxLength       =   9
         TabIndex        =   11
         Tag             =   "CIF|T|N|||sclien|nifclien|||"
         Text            =   "Text"
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   -73410
         MaxLength       =   35
         TabIndex        =   4
         Tag             =   "Domicilio|T|N|||sclien|domclien|||"
         Text            =   "Text"
         Top             =   360
         Width           =   4905
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   -73410
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "CP|T|N|||sclien|codpobla|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   -71490
         MaxLength       =   35
         TabIndex        =   6
         Tag             =   "Población|T|N|||sclien|pobclien|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   780
         Width           =   2985
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   -73410
         MaxLength       =   35
         TabIndex        =   7
         Tag             =   "Provincia|T|N|||sclien|proclien|||"
         Text            =   "Text"
         Top             =   1200
         Width           =   4905
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   -73410
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Telefono|T|S|||sclien|telclie1|||"
         Text            =   "963577679"
         Top             =   1620
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   -69990
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Movil|T|S|||sclien|movclien|||"
         Text            =   "Text"
         Top             =   1590
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   10
         Left            =   -73410
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "Mail|T|S|||sclien|maiclie1|||"
         Text            =   "Text"
         Top             =   2010
         Width           =   4905
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   12
         Left            =   -67080
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Fecha Alta|F|N|||sclien|fechaalt||dd/mm/yyyy|"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   780
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   13
         Left            =   -67080
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Fecha Baja|F|S|||sclien|fechabaj||dd/mm/yyyy|"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1170
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Height          =   1485
         Index           =   17
         Left            =   -68280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Tag             =   "Observaciones|T|S|||sclien|observac|||"
         Text            =   "frmGesSocios.frx":1A52
         Top             =   3120
         Width           =   4095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Es Socio"
         Height          =   375
         Index           =   0
         Left            =   -65580
         TabIndex        =   16
         Tag             =   "Facturado|N|N|0|1|sclien|essocio|||"
         Top             =   1230
         Width           =   1005
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos del Vehículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2145
         Index           =   1
         Left            =   -74730
         TabIndex        =   56
         Top             =   2490
         Width           =   6315
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   24
            Left            =   4890
            MaxLength       =   9
            TabIndex        =   26
            Tag             =   "Poliza|T|S|||sclien|numpoliza|||"
            Text            =   "Text"
            Top             =   780
            Width           =   1305
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2280
            TabIndex        =   63
            Text            =   "Text2"
            Top             =   360
            Width           =   3915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   11
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   29
            Tag             =   "Codigo Situación|N|N|||sclien|codsitua|00||"
            Text            =   "Tex"
            Top             =   1620
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   1320
            MaxLength       =   9
            TabIndex        =   27
            Tag             =   "Licencia|N|S|||sclien|licencia|00000000||"
            Text            =   "Text"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   60
            Text            =   "Text2"
            Top             =   1620
            Width           =   4155
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   18
            Left            =   4890
            MaxLength       =   10
            TabIndex        =   28
            Tag             =   "Fecha Situación|F|S|||sclien|fechasit||dd/mm/yyyy|"
            Text            =   "99/99/9999"
            Top             =   1200
            Width           =   1305
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   15
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   25
            Tag             =   "Matricula|T|S|||sclien|matricul|||"
            Text            =   "Text"
            Top             =   780
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   24
            Tag             =   "Codigo Coche|N|N|||sclien|codcoche|0000||"
            Text            =   "Text"
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label1 
            Caption         =   "Número de Póliza:"
            Height          =   255
            Index           =   23
            Left            =   3210
            TabIndex        =   64
            Top             =   780
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1020
            Picture         =   "frmGesSocios.frx":1A57
            Tag             =   "-1"
            ToolTipText     =   "Buscar vehiculo"
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1020
            Picture         =   "frmGesSocios.frx":1B59
            Tag             =   "-1"
            ToolTipText     =   "Buscar situación"
            Top             =   1620
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Situación:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   62
            Top             =   1620
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Licencia:"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   61
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   4530
            Picture         =   "frmGesSocios.frx":1C5B
            ToolTipText     =   "Buscar fecha"
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Situación:"
            Height          =   255
            Index           =   18
            Left            =   3210
            TabIndex        =   59
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Matrícula:"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   58
            Top             =   780
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Vehículo"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   645
         End
      End
      Begin VB.CommandButton cmdAux 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   -73440
         TabIndex        =   48
         ToolTipText     =   "Buscar cliente"
         Top             =   2520
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   -70560
         TabIndex        =   52
         Text            =   "hasta"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cmdAux1 
         Height          =   315
         ItemData        =   "frmGesSocios.frx":1CE6
         Left            =   -67320
         List            =   "frmGesSocios.frx":1CF0
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   2520
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   -71520
         MaxLength       =   10
         TabIndex        =   51
         Text            =   "desde"
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   -72480
         MaxLength       =   10
         TabIndex        =   50
         Tag             =   "Importe|N|N|||sclien_publicidad|importes|||"
         Text            =   "importe"
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   -73440
         TabIndex        =   55
         Text            =   "nomcliente"
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   -74640
         MaxLength       =   6
         TabIndex        =   49
         Tag             =   "Cliente|N|N|||sclien_publicidad|codclien|||"
         Text            =   "Codclien"
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdAux 
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   -73800
         TabIndex        =   43
         ToolTipText     =   "Buscar chofer"
         Top             =   2460
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   -69240
         MaxLength       =   40
         TabIndex        =   47
         Tag             =   "Observaciones|T|S|||ssocio_chofer|observac|||"
         Text            =   "Observaciones"
         Top             =   2460
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   -70320
         MaxLength       =   10
         TabIndex        =   46
         Tag             =   "Fecha Baja|F|S|||ssocio_chofer|fechabaj|||"
         Text            =   "FEcBaja"
         Top             =   2460
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -71640
         MaxLength       =   10
         TabIndex        =   45
         Tag             =   "Fecha Alta|F|S|||ssocio_chofer|fechaalt|||"
         Text            =   "FecAlta"
         Top             =   2460
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73590
         TabIndex        =   42
         Text            =   "nomchofe"
         Top             =   2460
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74760
         MaxLength       =   5
         TabIndex        =   44
         Tag             =   "Chofer|N|N|||sclien_chofer|codchofe|||"
         Text            =   "chofe"
         Top             =   2460
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4120
         Left            =   -74760
         TabIndex        =   54
         Top             =   720
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   7276
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmGesSocios.frx":1D07
         Height          =   4120
         Left            =   -74760
         TabIndex        =   65
         Top             =   720
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   7276
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   23
         Left            =   -73320
         MaxLength       =   5
         TabIndex        =   77
         Tag             =   "Codigo Socio|N|N|||sclien|codtarif|||"
         Text            =   "Text"
         Top             =   780
         Width           =   870
      End
      Begin MSComctlLib.ListView lwCRM 
         Height          =   4305
         Left            =   -74160
         TabIndex        =   85
         Top             =   690
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7594
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   1710
         Left            =   -74940
         TabIndex        =   89
         Top             =   420
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3016
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Servicios"
               Object.Tag             =   "2"
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas Cliente"
               Object.Tag             =   "3"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas Proveedor"
               Object.Tag             =   "4"
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   4545
         Left            =   -74310
         TabIndex        =   91
         Top             =   390
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8017
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmGesSocios.frx":1D1C
         Height          =   3525
         Left            =   240
         TabIndex        =   99
         Top             =   720
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   6218
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3855
         Index           =   0
         Left            =   -74940
         TabIndex        =   101
         Top             =   360
         Width           =   615
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   4120
         Left            =   -74760
         TabIndex        =   102
         Top             =   720
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   7276
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.Label Label1 
         Caption         =   "Imp.a Cuenta"
         Height          =   255
         Index           =   21
         Left            =   -68280
         TabIndex        =   106
         Top             =   2010
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "IMPORTE TOTAL: "
         Height          =   195
         Left            =   6480
         TabIndex        =   100
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Left            =   -65430
         TabIndex        =   93
         Top             =   930
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   -64590
         Picture         =   "frmGesSocios.frx":1D31
         ToolTipText     =   "Buscar fecha"
         Top             =   930
         Width           =   240
      End
      Begin VB.Label LabelDoc 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   -65910
         TabIndex        =   92
         Top             =   450
         Width           =   1905
      End
      Begin VB.Label Label2 
         Caption         =   "Cálculo Cuotas"
         Height          =   255
         Left            =   -65550
         TabIndex        =   88
         Top             =   780
         Width           =   1170
      End
      Begin VB.Image imgDoc 
         Height          =   345
         Index           =   0
         Left            =   -64380
         ToolTipText     =   "Cálculo Importe Cuotas "
         Top             =   750
         Width           =   390
      End
      Begin VB.Image imgDoc 
         Height          =   345
         Index           =   1
         Left            =   -64380
         ToolTipText     =   "Impresión Documento Alta"
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label14 
         Caption         =   "Documento Alta"
         Height          =   255
         Left            =   -65550
         TabIndex        =   87
         Top             =   390
         Width           =   1230
      End
      Begin VB.Label LabelCRM 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   -74160
         TabIndex        =   86
         Top             =   390
         Width           =   5745
      End
      Begin VB.Label Label1 
         Caption         =   "Suplemen.Cuota"
         Height          =   255
         Index           =   20
         Left            =   -68280
         TabIndex        =   79
         Top             =   1620
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN Socios"
         Height          =   255
         Index           =   19
         Left            =   -68280
         TabIndex        =   78
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones:"
         Height          =   255
         Index           =   3
         Left            =   -68280
         TabIndex        =   76
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "CP:"
         Height          =   255
         Index           =   5
         Left            =   -74640
         TabIndex        =   75
         Top             =   780
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   -73710
         Picture         =   "frmGesSocios.frx":1DBC
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Población:"
         Height          =   255
         Index           =   6
         Left            =   -72330
         TabIndex        =   74
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia:"
         Height          =   255
         Index           =   7
         Left            =   -74640
         TabIndex        =   73
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "CIF"
         Height          =   255
         Index           =   8
         Left            =   -68280
         TabIndex        =   72
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono:"
         Height          =   255
         Index           =   9
         Left            =   -74640
         TabIndex        =   71
         Top             =   1650
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Movil:"
         Height          =   255
         Index           =   10
         Left            =   -71490
         TabIndex        =   70
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "E-mail"
         Height          =   255
         Index           =   11
         Left            =   -74610
         TabIndex        =   69
         Top             =   2070
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   -67050
         Picture         =   "frmGesSocios.frx":1EBE
         Tag             =   "-1"
         ToolTipText     =   "Ver observaciones"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   1
         Left            =   -73710
         Tag             =   "-1"
         ToolTipText     =   "Enviar e-mail"
         Top             =   2085
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio:"
         Height          =   255
         Index           =   4
         Left            =   -74640
         TabIndex        =   68
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Alta"
         Height          =   255
         Index           =   13
         Left            =   -68280
         TabIndex        =   67
         Top             =   810
         Width           =   795
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   -67410
         Picture         =   "frmGesSocios.frx":1FC0
         ToolTipText     =   "Buscar fecha"
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Baja"
         Height          =   255
         Index           =   14
         Left            =   -68280
         TabIndex        =   66
         Top             =   1200
         Width           =   825
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   -67410
         Picture         =   "frmGesSocios.frx":204B
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6360
      Top             =   6060
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5280
      Top             =   6060
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
      Left            =   8880
      TabIndex        =   31
      Top             =   6840
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10200
      TabIndex        =   32
      Top             =   6840
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10200
      TabIndex        =   33
      Top             =   6840
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   120
      TabIndex        =   35
      Top             =   510
      Width           =   11055
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   2670
         MaxLength       =   40
         TabIndex        =   2
         Tag             =   "Nombre|T|N|||sclien|nomclien|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   330
         Width           =   3825
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   1
         Left            =   8280
         MaxLength       =   5
         TabIndex        =   3
         Tag             =   "Num.Vehiculo|N|S|||sclien|numeruve|0000||"
         Text            =   "Text"
         Top             =   330
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   930
         MaxLength       =   5
         TabIndex        =   1
         Tag             =   "Codigo Socio|N|N|||sclien|codclien|0000|S|"
         Text            =   "Text"
         Top             =   330
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   0
         Left            =   2010
         TabIndex        =   40
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Número Vehículo"
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
         Index           =   2
         Left            =   6720
         TabIndex        =   37
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
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
            Object.ToolTipText     =   "Choferes"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Publicidad"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cuotas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Contadores"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   36
      Top             =   6600
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
         TabIndex        =   39
         Top             =   240
         Width           =   3615
      End
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   4230
      Top             =   6540
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnbuscar 
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
      Begin VB.Menu mnNuevo 
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
      Begin VB.Menu mnLineas 
         Caption         =   "&Choferes"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnLineas2 
         Caption         =   "&Publicidad"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnLineas3 
         Caption         =   "Cu&otas"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnLineas4 
         Caption         =   "Con&tadores"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnbarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmGesSocios"
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
Public WithEvents frmS As frmFacSituaciones
Attribute frmS.VB_VarHelpID = -1
Public WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Public WithEvents frmCond As frmGesConduc
Attribute frmCond.VB_VarHelpID = -1
Public WithEvents frmBanco As frmFacBancosPropios
Attribute frmBanco.VB_VarHelpID = -1
Public WithEvents frmCal As frmCal
Attribute frmCal.VB_VarHelpID = -1
Public WithEvents frmV As frmGesVehic
Attribute frmV.VB_VarHelpID = -1
Public WithEvents frmSerie As frmRepNumSerie2
Attribute frmSerie.VB_VarHelpID = -1
Public WithEvents frmArt As frmAlmArticulos
Attribute frmArt.VB_VarHelpID = -1

Private WithEvents frmLLam As frmGesHisLlam
Attribute frmLLam.VB_VarHelpID = -1
Private WithEvents frmDoc As frmDocAltaBaja
Attribute frmDoc.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim btnAnyadir As Byte
Dim btnPrimero As Byte
Dim NombreTabla As String
Dim Ordenacion As String
Dim CadenaConsulta As String
Dim HaDevueltoDatos As Boolean
Private Modo As Byte
Dim kCampo As Integer
Dim ModificaLineas As Byte
Dim Fecha As Date
Dim Situacion As Boolean

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

Dim cadB1 As String

Dim BuscaChekc As String


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


Private Sub cmdAccCRM_Click(Index As Integer)
    
    'Acciones parar el CRM
    Select Case Index
    Case 0
    
        Select Case CByte(RecuperaValor(lwCRM.Tag, 1))
        Case 0
            'NUEVA, modificar o insertar acciones comerciales
            Set frmSerie = New frmRepNumSerie2
            
            frmSerie.DatoAInsertar = Adodc1.Recordset!CodClien
            frmSerie.DatosADevolverBusqueda = ""   'NUEVA
            frmSerie.Show vbModal
        End Select
        
        Me.Refresh
        DoEvents
        CargaDatosLWCRM
        Screen.MousePointer = vbDefault
    
    Case 1
        ' impresion de equipamiento del socio
        
        If Modo <> 2 Then Exit Sub
        If Me.Adodc1.Recordset.EOF Then Exit Sub
        If Text1(0).Text = "" Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        
        frmListado.OpcionListado = 60
        frmListado.NumCod = Format(Adodc1.Recordset!CodClien, "000000")
        frmListado.Show vbModal
        Screen.MousePointer = vbDefault
        
    Case 2
        If CByte(RecuperaValor(lwCRM.Tag, 1)) = 4 Then
            If lwCRM.SelectedItem Is Nothing Then Exit Sub
            If MsgBox("¿Desea eliminar las observaciones del departamento " & Me.lwCRM.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
            BuscaChekc = "DELETE from scrmobsclien  WHERE codclien = " & Me.Adodc1.Recordset!CodClien & " AND dpto=" & lwCRM.SelectedItem.SubItems(3)
            If ejecutar(BuscaChekc, False) Then CargaDatosLWCRM
            BuscaChekc = ""
        End If
    End Select
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cadB As String
Dim cad As String
Dim aaa As String


    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOK Then
              If InsertaRegistro Then 'antes InsertadesdeForm(Me) Then
'                    CrearContadores
                    PosicionarData
              End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOK Then
                '[Monica]13/03/2012: Modificamos el registro
                If ModificarRegistro Then
                    PosicionarData
                End If
                If ModificaDesdeFormulario(Me, 1) Then
                    If vParamAplic.Cooperativa = 1 Then
                        If ModificarCuentaCble(vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000"), Text1(0), "sclien") Then
                            MsgBox "Se ha modificado la cuenta " & vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000") & " en el arimoney.", vbExclamation
                        Else
                            MsgBox "Error al modificar la cuenta " & vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000") & " en el arimoney.", vbExclamation
                        End If
                    '[Monica]13/03/2012: en teletaxi si me modifican miro si existen sus cuentas y las modifico
                    Else
                    
                    
                        
                    End If
                    TerminaBloquear
                    PosicionarData
                End If
            End If
         Case 5 'INSERTAR MODIFICAR LINEA
            
            If ModificaLineas = 1 Then 'INSERTAR lineas
                If SSTab1.Tab = 2 Then
                    'choferes
                    If InsertarLinea Then
                        CargaGrid DataGrid1, Adodc2
                        BotonAnyadirLinea
                    End If
                Else
                    If SSTab1.Tab = 3 Then
                        'publicidad
                        If InsertarLinea2 Then
                            CargaGrid DataGrid2, Adodc3
                            BotonAnyadirLinea2
                        End If
                    Else
                        If SSTab1.Tab = 6 Then
                            If InsertarLinea4 Then
                                CargaGrid DataGrid3, Adodc4
                                BotonAnyadirLinea4
                            End If
                        Else
                            'cuotas
                            If InsertarLinea3 Then
                                CargaGrid DataGrid4, Adodc5
                                BotonAnyadirLinea3
                            End If
                        End If
                    
                    End If
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If SSTab1.Tab = 2 Then 'chofer
                    If ModificarLinea Then
                        TerminaBloquear
                        CargaTxtAux False, False
                        CargaGrid DataGrid1, Adodc2
                        ModificaLineas = 0
                        PonerBotonCabecera True
                    End If
                Else
                    If SSTab1.Tab = 6 Then ' contadores
                        If ModificarLinea3 Then
                            TerminaBloquear
                            CargaTxtAux4 False, False
                            CargaGrid DataGrid3, Adodc4
                            PonerBotonCabecera True
                        End If
                    Else
                        '[Monica]05/11/2013: añdimos el poder modificar las cuotas
                        If SSTab1.Tab = 4 Then
                            ' cuotas
                            If ModificarLinea4 Then
                                TerminaBloquear
                                CargaTxtAux3 False, False
                                CargaGrid DataGrid4, Adodc5
                                PonerBotonCabecera True
                            End If
                        Else
                            'publicidad
                            If ModificarLinea2 Then
                                TerminaBloquear
                                CargaTxtAux2 False, False
                                CargaGrid DataGrid2, Adodc3
                                PonerBotonCabecera True
                            End If
                        End If
                    End If
                End If
                Me.DataGrid1.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault
Error1:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function InsertaRegistro() As Boolean
Dim C2 As String
Dim b As Boolean
    
    On Error GoTo EInsertaRegistro
    InsertaRegistro = False
    
    conn.BeginTrans
    ConnConta.BeginTrans
    
    b = InsertarDesdeForm(Me, 1)
    If b Then
         b = CrearContadores
         If b Then
            ' Insertamos en el hco
            C2 = "insert into shiuve (codsocio,numeruve,fechaalta,fechabaja) values ("
            C2 = C2 & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "N") & ","
            C2 = C2 & DBSet(Text1(12).Text, "F") & "," & DBSet(Text1(13).Text, "F", "S") & ")"
            conn.Execute C2
         End If
         If b And vParamAplic.Cooperativa = 1 Then
            If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000"), "T") = "" Then
                b = InsertarCuentaCble(vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000"), , , , Text1(0).Text)
            Else
                b = ModificarCuentaCble(vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
            End If
         End If
         
         '[Monica]13/03/2012: cuando inserten un socio que inserten en la contabilidad todas las cuentas contables
         If b And vParamAplic.Cooperativa = 0 Then
               ' Cuenta de retencion para liquidacion
               If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_Cta_Reten_Soc & Format(Text1(0).Text, "0000"), "T") = "" Then
                   b = InsertarCuentaCble(vParamAplic.Raiz_Cta_Reten_Soc & Format(Text1(0).Text, "0000"), , , , Text1(0).Text, True)
               Else
                   b = ModificarCuentaCble(vParamAplic.Raiz_Cta_Reten_Soc & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
               End If
         
               ' Cuenta de ventas de equipos
               If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_Cta_Soc_Equip & Format(Text1(0).Text, "0000"), "T") = "" Then
                   b = InsertarCuentaCble(vParamAplic.Raiz_Cta_Soc_Equip & Format(Text1(0).Text, "0000"), , , , Text1(0).Text)
               Else
                   b = ModificarCuentaCble(vParamAplic.Raiz_Cta_Soc_Equip & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
               End If
               
               ' Cuenta de liquidacion
               If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000"), "T") = "" Then
                   b = InsertarCuentaCble(vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000"), , , , Text1(0).Text)
               Else
                   b = ModificarCuentaCble(vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
               End If
               
               ' Cuenta de publicidad
               If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_Cta_Soc_publi & Format(Text1(0).Text, "0000"), "T") = "" Then
                   b = InsertarCuentaCble(vParamAplic.Raiz_Cta_Soc_publi & Format(Text1(0).Text, "0000"), , , , Text1(0).Text)
               Else
                   b = ModificarCuentaCble(vParamAplic.Raiz_Cta_Soc_publi & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
               End If
               
               ' Cuenta de alta/baja socios
               If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_CtaAltaSoc & Format(Text1(0).Text, "0000"), "T") = "" Then
                   b = InsertarCuentaCble(vParamAplic.Raiz_CtaAltaSoc & Format(Text1(0).Text, "0000"), , , , Text1(0).Text, True)
               Else
                   b = ModificarCuentaCble(vParamAplic.Raiz_CtaAltaSoc & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
               End If
               
               ' Cuenta de cuotas
               If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_CtaClien_Soc & Format(Text1(0).Text, "0000"), "T") = "" Then
                   b = InsertarCuentaCble(vParamAplic.Raiz_CtaClien_Soc & Format(Text1(0).Text, "0000"), , , , Text1(0).Text)
               Else
                   b = ModificarCuentaCble(vParamAplic.Raiz_CtaClien_Soc & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
               End If
         
         End If
    End If
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        InsertaRegistro = True
        Exit Function
    End If
EInsertaRegistro:
    conn.RollbackTrans
    ConnConta.RollbackTrans
    MuestraError Err.Number, "Inserta Registro"
End Function



Private Function CrearContadores() As Boolean
'creará los contadores del socio nuevo con contadores igual a 0 con movimientos
'que en stipom tengan tipodocu=2
Dim SQL As String

On Error GoTo EContadores


    CrearContadores = False

    Set miRsAux = New ADODB.Recordset
    SQL = "select * from stipom where tipodocu=2"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        SQL = "INSERT INTO sclien_contadores (codsocio,codtipom,contador) values ("
        SQL = SQL & Text1(0).Text & "," & DBSet(miRsAux!codtipom, "T") & ",0)"
        conn.Execute SQL
        miRsAux.MoveNext
    Wend
    
    miRsAux.Close
    Set miRsAux = Nothing


    CrearContadores = True
    Exit Function

EContadores:
If Err.Number <> 0 Then MsgBox "Error contadores: " & Err.Description

End Function

Private Function ModificarRegistro() As Boolean
Dim C2 As String
Dim b As Boolean
    
    On Error GoTo EModificarRegistro
    ModificarRegistro = False
    
    conn.BeginTrans
    ConnConta.BeginTrans
    
    b = ModificaDesdeFormulario(Me, 1)
    If b Then
        If vParamAplic.Cooperativa = 1 Then
            If ModificarCuentaCble(vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000"), Text1(0), "sclien") Then
                 MsgBox "Se ha modificado la cuenta " & vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000") & " en el arimoney.", vbExclamation
            Else
                 MsgBox "Error al modificar la cuenta " & vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000") & " en el arimoney.", vbExclamation
            End If
        End If
        
        '[Monica]13/03/2012: en teletaxi si me modifican miro si existen sus cuentas y las modifico
        If vParamAplic.Cooperativa = 0 Then
            ' Cuenta de retencion para liquidacion
            If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_Cta_Reten_Soc & Format(Text1(0).Text, "0000"), "T") = "" Then
                 b = InsertarCuentaCble(vParamAplic.Raiz_Cta_Reten_Soc & Format(Text1(0).Text, "0000"), , , , Text1(0).Text)
            Else
                 b = ModificarCuentaCble(vParamAplic.Raiz_Cta_Reten_Soc & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
            End If
        
            ' Cuenta de ventas de equipos
            If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_Cta_Soc_Equip & Format(Text1(0).Text, "0000"), "T") = "" Then
                 b = InsertarCuentaCble(vParamAplic.Raiz_Cta_Soc_Equip & Format(Text1(0).Text, "0000"), , , , Text1(0).Text)
            Else
                 b = ModificarCuentaCble(vParamAplic.Raiz_Cta_Soc_Equip & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
            End If
              
            ' Cuenta de liquidacion
            If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000"), "T") = "" Then
                 b = InsertarCuentaCble(vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000"), , , , Text1(0).Text)
            Else
                 b = ModificarCuentaCble(vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
            End If
              
            ' Cuenta de publicidad
            If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_Cta_Soc_publi & Format(Text1(0).Text, "0000"), "T") = "" Then
                 b = InsertarCuentaCble(vParamAplic.Raiz_Cta_Soc_publi & Format(Text1(0).Text, "0000"), , , , Text1(0).Text)
            Else
                 b = ModificarCuentaCble(vParamAplic.Raiz_Cta_Soc_publi & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
            End If
              
            ' Cuenta de alta/baja socios
            If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_CtaAltaSoc & Format(Text1(0).Text, "0000"), "T") = "" Then
                 b = InsertarCuentaCble(vParamAplic.Raiz_CtaAltaSoc & Format(Text1(0).Text, "0000"), , , , Text1(0).Text)
            Else
                 b = ModificarCuentaCble(vParamAplic.Raiz_CtaAltaSoc & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
            End If
              
            ' Cuenta de cuotas
            If DevuelveDesdeBDNew(conConta, "cuentas", "nommacta", "codmacta", vParamAplic.Raiz_CtaClien_Soc & Format(Text1(0).Text, "0000"), "T") = "" Then
                 b = InsertarCuentaCble(vParamAplic.Raiz_CtaClien_Soc & Format(Text1(0).Text, "0000"), , , , Text1(0).Text)
            Else
                 b = ModificarCuentaCble(vParamAplic.Raiz_CtaClien_Soc & Format(Text1(0).Text, "0000"), Text1(0).Text, "sclien")
            End If
         
        End If
    End If
                        
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        ModificarRegistro = True
        Exit Function
    End If
EModificarRegistro:
    conn.RollbackTrans
    ConnConta.RollbackTrans
    MuestraError Err.Number, "Modificar Registro"
End Function





Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a Mantenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    
    If b Then
        Me.lblIndicador(0).Caption = "Líneas "
        PonerFocoBtn Me.cmdRegresar
    End If
    
    'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim SQL As String
On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE sclien_chofer Set codchofe = " & txtAux1(0).Text & ", fechaalt='" & Format(txtAux1(2).Text, FormatoFecha) & "', "
        SQL = SQL & "fechabaj='" & Format(txtAux1(3).Text, FormatoFecha) & "', obsevac=" & DBSet(txtAux1(4).Text, "T")
        SQL = SQL & " where codsocio=" & Adodc2.Recordset!codSocio & " AND numlinea=" & Adodc2.Recordset!numlinea
        
        conn.Execute SQL
        
        ' modificamos en la tabla de hco de choferes
        SQL = "UPDATE schofe_historia Set fechaini = '" & Format(txtAux1(2).Text, FormatoFecha) & "', "
        SQL = SQL & "fechafin=" & DBSet(txtAux1(3).Text, "F", "S") & ", observac=" & DBSet(txtAux1(4).Text, "T")
        SQL = SQL & " where codchofe=" & DBSet(txtAux1(0).Text, "N") & " AND numeruve=" & DBSet(Adodc1.Recordset!NumerUve, "N")
        SQL = SQL & " and fechaini= " & DBSet(txtAux1(2).Text, "F")
        conn.Execute SQL
        
        
        ModificarLinea = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Chofer" & vbCrLf & Err.Description
End Function

Private Function ModificarLinea2() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim SQL As String
On Error GoTo EModificarLinea

    ModificarLinea2 = False
    SQL = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE sclien_publicidad Set codclien = " & txtAux2(0).Text & ", importes=" & TransformaComasPuntos(txtAux2(2).Text) & ", "
        SQL = SQL & "desdefec='" & Format(txtAux2(3).Text, FormatoFecha) & "',hastafec='" & Format(txtAux2(4).Text, FormatoFecha) & "', situacio=" & cmdAux1.ItemData(cmdAux1.ListIndex)
        SQL = SQL & " where codsocio=" & Adodc3.Recordset!codSocio & " AND numlinea=" & Adodc3.Recordset!numlinea
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea2 = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Publicidad" & vbCrLf & Err.Description
End Function

Private Function ModificarLinea3() As Boolean
'Modifica un registro en la tabla de lineas de contadores de socios: sclien_contadores
Dim SQL As String
On Error GoTo EModificarLinea

    ModificarLinea3 = False
    SQL = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE sclien_contadores set contador=" & txtAux4(2).Text
        SQL = SQL & " where codsocio=" & Adodc4.Recordset!codSocio & " AND codtipom='" & Adodc4.Recordset!codtipom & "'"
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea3 = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Contadores" & vbCrLf & Err.Description
End Function



Private Function ModificarLinea4() As Boolean
'Modifica un registro en la tabla de lineas de articulos de cuotas
Dim SQL As String
On Error GoTo EModificarLinea

    ModificarLinea4 = False
    SQL = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE sclien_cuotas set importes=" & DBSet(txtAux3(2).Text, "N")
        SQL = SQL & " where codsocio=" & Adodc5.Recordset!codSocio & " AND numlinea=" & Adodc5.Recordset!numlinea
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea4 = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Contadores" & vbCrLf & Err.Description
End Function






Private Function DatosOkLinea() As Boolean
Dim SQL As String

    DatosOkLinea = False
    If SSTab1.Tab = 2 Then
    'chofer
        If txtAux1(0).Text = "" Then
            MsgBox "Es necesario introducir el código de chofer.", vbExclamation
            DatosOkLinea = False
            Exit Function
        Else
            ' comprobamos que el chofer no se encuentre en ningún otro socio con fecha de alta
            If ModificaLineas = 1 Then ' solo si estamos insertando
                SQL = "select codsocio from sclien_chofer where codchofe = " & DBSet(txtAux1(0).Text, "N")
                SQL = SQL & " and (fechabaj is null or fechabaj = '0000-00-00')"
                If TotalRegistrosConsulta(SQL) > 0 Then
                    MsgBox "Este chofer está asignado al socio " & Format(DevuelveValor(SQL), "000000") & ". Revise.", vbExclamation
                    DatosOkLinea = False
                    Exit Function
                End If
            End If
        End If
    ElseIf SSTab1.Tab = 3 Then
        'publicidad
            If txtAux2(0).Text = "" Then
                MsgBox "Es necesario introducir el código de cliente.", vbExclamation
                DatosOkLinea = False
                Exit Function
            End If
            If txtAux2(2).Text = "" Then
                MsgBox "Es necesario introducir el importe.", vbExclamation
                DatosOkLinea = False
                Exit Function
            End If
            If txtAux2(3).Text = "" Then
                MsgBox "Es necesario introducir la fecha desde.", vbExclamation
                DatosOkLinea = False
                Exit Function
            End If
            If txtAux2(4).Text = "" Then
                MsgBox "Es necesario introducir la fecha hasta.", vbExclamation
                DatosOkLinea = False
                Exit Function
            End If
            If cmdAux1.Text = "" Then
                MsgBox "Es necesario introducir el código de situación.", vbExclamation
                DatosOkLinea = False
                Exit Function
            End If
    ElseIf SSTab1.Tab = 4 Then
        'cuotas
            If txtAux3(0).Text = "" Then
                MsgBox "Es necesario introducir el código de artículo.", vbExclamation
                DatosOkLinea = False
                Exit Function
            Else
                '[Monica]05/12/2013: añado la condicion para que solo lo compruebe si estoy insertando
                If ModificaLineas = 0 Then
                
                    ' comprobamos que el codigo de articulo no esté introducido ya
                    If TotalRegistros("select count(*) from sclien_cuotas where codsocio = " & DBSet(Text1(0).Text, "N") & " and codartic = " & DBSet(txtAux3(0).Text, "T")) Then
                        MsgBox "Este artículo ya está introducido en el socio. Revise.", vbExclamation
                        DatosOkLinea = False
                        Exit Function
                    End If
                    
                End If
            End If
    ElseIf SSTab1.Tab = 6 Then
        'contadores
            If txtAux4(0).Text = "" Then
                MsgBox "Es necesario introducir el tipo de movimiento.", vbExclamation
                DatosOkLinea = False
                Exit Function
            Else
                If ModificaLineas = 1 Then ' si estamos insertando
                    ' comprobamos que el codigo de articulo no esté introducido ya
                    If TotalRegistros("select count(*) from sclien_contadores where codsocio = " & DBSet(Text1(0).Text, "N") & " and codtipom = " & DBSet(txtAux4(0).Text, "T")) Then
                        MsgBox "Este tipo de movimiento ya está introducido en el socio. Revise.", vbExclamation
                        DatosOkLinea = False
                        Exit Function
                    End If
                End If
            End If
    End If
    DatosOkLinea = True

End Function


Private Function InsertarLinea() As Boolean
Dim SQL As String
Dim vWhere As String
Dim numF As String
On Error GoTo EInsertarLinea

    conn.BeginTrans

    InsertarLinea = False
    SQL = ""
    If DatosOkLinea Then
        vWhere = "codsocio=" & Val(Text1(0).Text)
        numF = SugerirCodigoSiguienteStr("sclien_chofer", "numlinea", vWhere)
        SQL = "INSERT INTO sclien_chofer "
        SQL = SQL & "(codsocio, numlinea, codchofe, fechaalt,fechabaj,obsevac) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
        SQL = SQL & DBSet(txtAux1(0).Text, "N") & "," & DBSet(txtAux1(2).Text, "F") & ","
        SQL = SQL & DBSet(txtAux1(3).Text, "F", "S") & "," & DBSet(txtAux1(4).Text, "T") & ")"
     
        conn.Execute SQL
     
        ' insertamos en el hco de schofe_historia
        vWhere = "codchofe=" & Val(txtAux1(0).Text)
        numF = SugerirCodigoSiguienteStr("schofe_historia", "numlinea", vWhere)
        SQL = "INSERT INTO schofe_historia "
        SQL = SQL & "(codchofe,numlinea,numeruve,fechaini,fechafin,observac) "
        SQL = SQL & "VALUES (" & Val(txtAux1(0).Text) & ", " & numF & ","
        SQL = SQL & DBSet(Text1(1), "N") & "," & DBSet(txtAux1(2).Text, "F") & ","
        SQL = SQL & DBSet(txtAux1(3).Text, "F", "S") & "," & DBSet(txtAux1(4).Text, "T") & ")"
        
        conn.Execute SQL
        
        InsertarLinea = True
    End If
    conn.CommitTrans
    Exit Function
EInsertarLinea:
    conn.RollbackTrans
    MuestraError Err.Number, "Insertar Lineas chofer" & vbCrLf & Err.Description
End Function

Private Function InsertarLinea2() As Boolean
Dim SQL As String
Dim vWhere As String
Dim numF As String
Dim Importe As Currency
On Error GoTo EInsertarLinea2

    InsertarLinea2 = False
    SQL = ""
    If DatosOkLinea Then
        Importe = ImporteFormateado(txtAux2(2).Text)
        vWhere = "codsocio=" & Val(Text1(0).Text)
        numF = SugerirCodigoSiguienteStr("sclien_publicidad", "numlinea", vWhere)
        SQL = "INSERT INTO sclien_publicidad "
        SQL = SQL & "(codsocio, numlinea, codclien, importes,desdefec,hastafec,situacio) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
        SQL = SQL & DBSet(txtAux2(0).Text, "N") & "," & TransformaComasPuntos(CStr(Importe)) & ",'"
        SQL = SQL & Format(txtAux2(3).Text, FormatoFecha) & "','" & Format(txtAux2(4).Text, FormatoFecha) & "'," & cmdAux1.ItemData(cmdAux1.ListIndex) & ")"
     End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea2 = True
    End If
    Exit Function
EInsertarLinea2:
    MuestraError Err.Number, "Insertar Lineas publicidad" & vbCrLf & Err.Description
End Function


Private Function InsertarLinea3() As Boolean
Dim SQL As String
Dim vWhere As String
Dim numF As String
Dim Importe As Currency
On Error GoTo EInsertarLinea3

    InsertarLinea3 = False
    SQL = ""
    If DatosOkLinea Then
        vWhere = "codsocio=" & Val(Text1(0).Text)
        numF = SugerirCodigoSiguienteStr("sclien_cuotas", "numlinea", vWhere)
        SQL = "INSERT INTO sclien_cuotas "
        SQL = SQL & "(codsocio, numlinea, codartic, importes) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
        SQL = SQL & DBSet(txtAux3(0).Text, "T") & "," & DBSet(txtAux3(2).Text, "N") & ")"
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea3 = True
    End If
    Exit Function
EInsertarLinea3:
    MuestraError Err.Number, "Insertar Lineas Cuotas" & vbCrLf & Err.Description
End Function


Private Function InsertarLinea4() As Boolean
Dim SQL As String
Dim vWhere As String
Dim numF As String
Dim Importe As Currency
On Error GoTo EInsertarLinea4

    InsertarLinea4 = False
    SQL = ""
    If DatosOkLinea Then
        vWhere = "codsocio=" & Val(Text1(0).Text)
        SQL = "INSERT INTO sclien_contadores "
        SQL = SQL & "(codsocio, codtipom, contador) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & DBSet(txtAux4(0).Text, "T") & ","
        SQL = SQL & DBSet(txtAux4(2).Text, "N") & ")"
     End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea4 = True
    End If
    Exit Function
EInsertarLinea4:
    MuestraError Err.Number, "Insertar Lineas Contadores" & vbCrLf & Err.Description
End Function


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    cadB1 = ObtenerBusqueda(Me, True)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select distinct sclien.* from ((" & NombreTabla & " LEFT JOIN sclien_cuotas ON sclien.codclien = sclien_cuotas.codsocio) LEFT JOIN sclien_chofer ON sclien.codclien = sclien_chofer.codsocio) LEFT JOIN sclien_publicidad ON sclien.codclien = sclien_publicidad.codsocio WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        PonerFoco Text1(0)
    End If
End Sub

Private Function DatosOK() As Boolean
Dim b As Boolean
Dim Cta As String
Dim cadMen As String
Dim SQL As String

    DatosOK = False
    
    If Text1(0).Text = "" Then
        MsgBox "Es necesario introducir el codigo de socio.", vbExclamation
        DatosOK = False
        Exit Function
    End If
    If ComprobarCero(Text1(1).Text) = "0" Then
        MsgBox "Es necesario introducir la V de socio.", vbExclamation
        DatosOK = False
        Exit Function
    End If
    If Text1(2).Text = "" Then
        MsgBox "Es necesario introducir el nombre del socio.", vbExclamation
        DatosOK = False
        Exit Function
    End If
    If Text1(3).Text = "" Then
        MsgBox "Es necesario introducir el domicilio del socio.", vbExclamation
        DatosOK = False
        Exit Function
    End If
    If Text1(4).Text = "" Then
        MsgBox "Es necesario introducir el código de población del socio.", vbExclamation
        DatosOK = False
        Exit Function
    End If
    If Text1(5).Text = "" Then
        MsgBox "Es necesario introducir la población del socio.", vbExclamation
        DatosOK = False
        Exit Function
    End If
    If Text1(6).Text = "" Then
        MsgBox "Es necesario introducir la provincia del socio.", vbExclamation
        DatosOK = False
        Exit Function
    End If
    If Text1(7).Text = "" Then
        MsgBox "Es necesario introducir el CIF del socio.", vbExclamation
        DatosOK = False
        Exit Function
    End If
    If Text1(12).Text = "" Then
        MsgBox "Es necesario introducir la fecha de alta del socio.", vbExclamation
        DatosOK = False
        Exit Function
    End If
    If Text1(14).Text = "" Then
        MsgBox "Es necesario introducir el còdigo de coche.", vbExclamation
        SSTab1.Tab = 0
        PonerFoco Text1(14)
        DatosOK = False
        Exit Function
    End If
    If Text1(11).Text = "" Then
        MsgBox "Es necesario introducir el código de situación.", vbExclamation
        SSTab1.Tab = 0
        PonerFoco Text1(11)
        DatosOK = False
        Exit Function
    End If
    If Text1(19).Text = "" Then
        MsgBox "Es necesario introducir el código del banco.", vbExclamation
        SSTab1.Tab = 0
        PonerFoco Text1(19)
        DatosOK = False
        Exit Function
    End If
    If Text1(20).Text = "" Then
        MsgBox "Es necesario introducir el código de sucursal.", vbExclamation
        SSTab1.Tab = 0
        PonerFoco Text1(19)
        DatosOK = False
        Exit Function
    End If
    If Text1(21).Text = "" Then
        MsgBox "Es necesario introducir el dígito de control.", vbExclamation
        SSTab1.Tab = 0
        PonerFoco Text1(21)
        DatosOK = False
        Exit Function
    End If
    If Text1(22).Text = "" Then
        MsgBox "Es necesario introducir la cuenta bancaria.", vbExclamation
        SSTab1.Tab = 0
        PonerFoco Text1(22)
        DatosOK = False
        Exit Function
    End If
    
    '[Monica]01/04/2014: al insertar modificar no comprobabamos que la uve estuviera asignada a otro socio
    If Modo = 3 Or Modo = 4 Then
        If Text1(1).Text <> "" Then
            SQL = "select count(*) from sclien where numeruve = " & DBSet(Text1(1).Text, "N") & " and codclien <> " & DBSet(Text1(0).Text, "N")
            If TotalRegistros(SQL) <> 0 Then
                If MsgBox("Esta V está asignada a otro socio. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                    DatosOK = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    '- Validar que la cuenta bancaria es correcta
'    If Not Comprueba_CuentaBan(Text1(19).Text & Text1(20).Text & Text1(21).Text & Text1(22).Text) Then
'        MsgBox "La cuenta bancaria no es correcta.", vbExclamation
'        DatosOk = False
'        Exit Function
'    End If
    b = True
    If (Modo = 3 Or Modo = 4) Then
        If Text1(19).Text = "" Or Text1(20).Text = "" Or Text1(21).Text = "" Or Text1(22).Text = "" Then
            Text1(28).Text = ""
            Text1(19).Text = ""
            Text1(20).Text = ""
            Text1(21).Text = ""
            Text1(22).Text = ""
        Else
            Cta = Format(Text1(19).Text, "0000") & Format(Text1(20).Text, "0000") & Format(Text1(21).Text, "00") & Format(Text1(22).Text, "0000000000")
            If Val(ComprobarCero(Cta)) = 0 Then
                cadMen = "El socio no tiene asignada cuenta bancaria."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(Cta) Then
                cadMen = "La cuenta bancaria del socio no es correcta. ¿ Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(19)
                    b = False
                End If
            Else
                BuscaChekc = ""
                If Me.Text1(28).Text <> "" Then BuscaChekc = Mid(Text1(28).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, Cta, Cta) Then
                    If Me.Text1(28).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(28).Text = BuscaChekc & Cta
                    Else
                        If Mid(Text1(28).Text, 3) <> Cta Then
                            Cta = "Calculado : " & BuscaChekc & Cta
                            Cta = "Introducido: " & Me.Text1(28).Text & vbCrLf & Cta & vbCrLf
                            Cta = "Error en codigo IBAN" & vbCrLf & Cta & "Continuar?"
                            If MsgBox(Cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco Text1(28)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    DatosOK = b
        
End Function

Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 1
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0"
            frmCli.Show vbModal
            Set frmCli = Nothing
        Case 0
            Set frmCond = New frmGesConduc
            frmCond.DatosADevolverBusqueda = "0"
            frmCond.Show vbModal
            Set frmCond = Nothing
        Case 2
            Set frmArt = New frmAlmArticulos
            frmArt.DatosADevolverBusqueda2 = "@1@"
            frmArt.Show vbModal
            Set frmArt = Nothing
    End Select
End Sub


Private Sub cmdAux1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAux1_LostFocus()
    cmdAceptar.SetFocus
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        Me.lblIndicador(0).Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
            DeseleccionaGrid DataGrid2
        End If
        cmdRegresar.Caption = "Regresar"
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Adodc1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        
        cad = Adodc1.Recordset.Fields(0) & "|"
        cad = cad & Adodc1.Recordset.Fields(2) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(2)
End Sub

Private Sub Form_Load()
    'Icono del form
    Me.Icon = frmPpal.Icon
    

    'ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 16
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(9).Image = 45 'Lineas chofer
        .Buttons(10).Image = 10 'Lineas publicidad
        .Buttons(11).Image = 11 'lineas de cuotas
        .Buttons(12).Image = 27 'lineas contadores
        .Buttons(13).Image = 16 'imprimir
        .Buttons(14).Image = 15 'salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    'La nevegacion para albaranes, facturas....
    ImagenesNavegacion
      
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    PrimeraVez = True
    
    ImgMail(1).Picture = frmPpal.imgListComun.ListImages(20).Picture
    Me.imgDoc(1).Picture = frmPpal.ImageListTPV.ListImages(8).Picture
    Me.imgDoc(0).Picture = frmPpal.ImageListTPV.ListImages(10).Picture

    '## A mano
    NombreTabla = "sclien"
    Ordenacion = " ORDER BY codclien"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Adodc1.ConnectionString = conn
    Adodc1.RecordSource = "Select * from " & NombreTabla & " where codclien=-1"
    Adodc1.Refresh
    
    LimpiarDataGrids
    
    Me.SSTab1.Tab = 0
    
    'Ponemos los datos del listview
    imgFecha(3).Tag = vEmpresa.FechaIni
    CargaColumnas 3
    
    CargaColumnasCRM 0

    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
    End If
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo


    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    
    Modo = Kmodo
    PonerIndicador lblIndicador(0), Modo
    
    
    'Modo 2. Hay datos y estamos visualizandolos
    '===========================================
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
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    Me.imgBuscar(2).Enabled = (Modo <= 4 And Modo > 1)
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    For i = 0 To Check1.Count - 1
        Me.Check1(i).Enabled = Not (Modo = 0 Or Modo = 2)
    Next i
    
    'CRM
    cmdAccCRM(0).visible = Modo = 2
    cmdAccCRM(1).visible = Modo = 2
    
    ' bloquear el numero de uve, no se puede modificar (solo por el hco de uves)
    BloquearTxt Text1(1), Not (Modo = 1 Or Modo = 3)
    
    
    '-----------------------------
    PonerLongCampos
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
                        
    'El listview
    If Modo <> 2 Then
        lw1.ListItems.Clear
    End If
                        
     ' solo si tenemos registro cargado podemos imprimir documentos
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    Me.imgDoc(1).visible = b
    Me.imgDoc(1).Enabled = b
    Me.Label14.visible = b
    
    Me.imgDoc(0).visible = False
    Me.imgDoc(0).Enabled = False
    Me.Label2.visible = False
    
    Me.Refresh
                       
    ' Para poder buscar por el codigo de articulo
    CargaTxtAux3 (Modo = 1 Or (Modo = 5 And ModificaLineas = 1)), True
    For i = 1 To 2
        txtAux3(i).visible = False
        txtAux3(i).Locked = False
    Next i
                       
    '[Monica]04/02/2015
    ' Para poder buscar por el codigo de conductor
     CargaTxtAux (Modo = 1 Or (Modo = 5 And ModificaLineas = 1)), True
    For i = 1 To 4
        txtAux1(i).visible = False
        txtAux1(i).Locked = False
    Next i
    
    '[Monica]04/02/2015
    ' Para poder buscar por el codigo de cliente y por importe
     CargaTxtAux2 (Modo = 1 Or (Modo = 5 And ModificaLineas = 1)), True
    For i = 1 To 4
        If i <> 2 Then
            txtAux2(i).visible = False
            txtAux2(i).Locked = False
        End If
    Next i
    If Modo = 1 Then
        cmdAux1.Locked = False
        cmdAux1.visible = False
    End If
                       
                       
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
    Me.mnnuevo.Enabled = b
    
    b = (Modo = 2 Or Modo = 5)
    'Modificar                                  '[Monica]05/12/2013: en cuotas dejamos modificar el importe
    Toolbar1.Buttons(6).Enabled = (Modo = 2) Or (Modo = 5) 'And SSTab1.Tab <> 4)
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    b = (Modo = 2)
    'lineas de chofer
    Toolbar1.Buttons(9).Enabled = b
    'lineas de publicidad
    Toolbar1.Buttons(10).Enabled = b
    'lineas de cuotas
    Toolbar1.Buttons(11).Enabled = b
    'lineas de contadores
    Toolbar1.Buttons(12).Enabled = b And vUsu.Nivel = 0
    
    b = (Modo = 0 Or Modo = 2)
    'imprimir
    Toolbar1.Buttons(13).Enabled = b
    
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
    lblIndicador(0).Caption = ""
    For i = 0 To Check1.Count - 1
        Check1(i).Value = 0
    Next i
    
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    
    txtAux3(0).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod articulo
    txtAux3(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre articulo
    
    txtAux3(2).Text = DevuelveValor("select preciove from sartic where codartic = " & DBSet(txtAux3(0).Text, "T"))
    txtAux3(2).Text = Format(txtAux3(2).Text, "###,##0.0000")
    
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        'Estamos en Cabecera
        'Recupera todo el registro de Tarifas de Precios
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmCal_Selec(vFecha As Date)
    Fecha = vFecha
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCond_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux1(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim indice As Byte
Dim devuelve As String
    
    indice = 4
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'Poblacion
    Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve)
    'provincia
    Text1(indice + 2).Text = devuelve

End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
'Situaciones
    If Situacion Then
        txtAux2(4).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        Text1(11).Text = RecuperaValor(CadenaSeleccion, 1)
        FormateaCampo Text1(10)
        Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmV_DatoSeleccionado(CadenaSeleccion As String)
    Text1(14).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

Select Case Index
    Case 2
            If Modo = 5 Or Modo = 0 Then
            Else
                
                If Modo = 3 Or Modo = 4 Then
                    CadenaDesdeOtroForm = Text1(17).Text
                Else
                    CadenaDesdeOtroForm = ""
                    If Not Adodc1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Adodc1.Recordset!observac, "T")
                End If
                frmFacClienteObser.Modificar = Modo >= 3
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 3 Or Modo = 4 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(17).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
            End If
            
    Case 0 'codigo postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            If Index = 9 Then
                indice = 4
            Else
                PonerFoco Text1(4)
            End If
            
    Case 1  'situaciones
            Situacion = False
            indice = 11
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
             If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
            frmS.Show vbModal
            Set frmS = Nothing
            
    Case 3 'coches
            Set frmV = New frmGesVehic
            frmV.DatosADevolverBusqueda = "0|1|"
            frmV.Show vbModal
            Set frmV = Nothing
    End Select
End Sub

Private Sub imgDoc_Click(Index As Integer)
    TerminaBloquear
    
    If Text1(1).Text = "" Then Exit Sub
    
    Select Case Index
        Case 1 'documentos de alta socio
            Set frmDoc = New frmDocAltaBaja
            frmDoc.NumCod = Text1(1).Text
            frmDoc.Show vbModal
            Set frmDoc = Nothing
            
       Case 0 ' calculo de cuotas
            CalculoCuotas
            
    End Select
End Sub

Private Sub CalculoCuotas()
Dim RS As ADODB.Recordset
Dim SQL As String
Dim cad As String
Dim base0 As Single
Dim base1 As Single
Dim BaseImp As Single
Dim iva As String
Dim porciva As Currency
        
        cad = "CALCULO DE IMPORTE DE CUOTAS SOCIO: " & vbCrLf & vbCrLf
        
        'busco el iva del articulo
        iva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArtCuotaSinChofer, "T")
        If iva = "" Then
            porciva = 0
        End If
        
        base0 = vParamAplic.PrecioCuotaSinChofe
        
            cad = cad & "Cuota sin Chofer:      " & Format(base0, "###,##0.00") & vbCrLf
        
        
        '[Monica]10/02/2011: si el socio tiene suplemento se le añade
        base0 = base0 + DBLet(Adodc1.Recordset!suplecuota, "N")
        '[Monica]10/02/2011:end
        
        If DBLet(Adodc1.Recordset!suplecuota, "N") <> 0 Then
            cad = cad & "Suplemento:               " & Format(Adodc1.Recordset!suplecuota, "##,###,##0.00") & vbCrLf
        End If
        
        BaseImp = base0
        
        base1 = 0
        If TieneChofer(CStr(Adodc1.Recordset!CodClien)) Then
            base1 = vParamAplic.PrecioCuotaConChofe
            BaseImp = BaseImp + base1
            cad = cad & "Cuota Chofer:             " & Format(base1, "#,###,##0.00") & vbCrLf
        End If
        
        If Adodc1.Recordset!essocio = 0 Then
            BaseImp = BaseImp + vParamAplic.PrecioPorAlquiler
            cad = cad & "Cuota Alquiler:           " & Format(vParamAplic.PrecioPorAlquiler, "#,###,##0.00") & vbCrLf
        End If
    
            cad = cad & vbCrLf
            cad = cad & "--------------------------------" & vbCrLf
            cad = cad & "TOTAL BRUTO:         " & Format(BaseImp, "###,##0.00")

        MsgBox cad, vbInformation, "Cálculo de Cuotas"


End Sub





Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then
        If Index <> 3 Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
   
    Set frmCal = New frmCal
    frmCal.Fecha = Now
       
    Select Case Index
        Case 0
            indice = 12
            PonerFormatoFecha Text1(indice)
            If Text1(indice).Text <> "" Then frmCal.Fecha = CDate(Text1(indice).Text)
        Case 1
            indice = 13
            PonerFormatoFecha Text1(indice)
            If Text1(indice).Text <> "" Then frmCal.Fecha = CDate(Text1(indice).Text)
        Case 2
            indice = 18
            PonerFormatoFecha Text1(indice)
            If Text1(indice).Text <> "" Then frmCal.Fecha = CDate(Text1(indice).Text)
        Case 3
            indice = 26
            PonerFormatoFecha Text1(indice)
            If Text1(indice).Text <> "" Then frmCal.Fecha = CDate(Text1(indice).Text)
    End Select
    Screen.MousePointer = vbDefault
    frmCal.Show vbModal
    If IsDate(Fecha) Then
        If Fecha <> "0:00:00" Then Text1(indice) = Fecha
    End If
    Set frmCal = Nothing
    
    'Para la fecha de la navegacion
    If Index = 3 And Text1(26).Text <> "" Then
        imgFecha(3).Tag = Text1(26).Text
        CargaDatosLWDoc
    End If
    
    PonerFoco Text1(indice)
    
    
End Sub


Private Sub lw1_DblClick()
Dim Seleccionado As Long
    
    If Modo <> 2 Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub


    If Me.DatosADevolverBusqueda <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un socio. No puede ver los documentos.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 2
        '[Monica] no son albaranes son llamadas
        'LLAMADAS
        Set frmLLam = New frmGesHisLlam
        frmLLam.HoraServ = lw1.SelectedItem.SubItems(1)
        frmLLam.FechaServ = lw1.SelectedItem.Text
        frmLLam.NumerUve = lw1.SelectedItem.SubItems(2)
        frmLLam.Show vbModal
        Set frmLLam = Nothing

    Case 3, 4
        'FACTURAS del cliente scafaccli (facturas de publicidad FPC y de llamadas FAC)
        'Este no necesitamos crear instancias
        
        'Lo que ocurre que esta preparado para abrir la factura a partir de un albaran, con lo cual
        'En la funcion abrir factura, buscare un albaran de la factura para abrirlo
        AbrirFacturaLW
        
        
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    lw1.SetFocus
    Seleccionado = lw1.SelectedItem.Index
    CargaDatosLWDoc
    lw1.SelectedItem.Selected = False
    Set lw1.SelectedItem = Nothing
    If lw1.ListItems.Count >= Seleccionado Then
            lw1.ListItems(Seleccionado).Selected = True
            lw1.ListItems(Seleccionado).EnsureVisible
    End If
    Screen.MousePointer = vbDefault
End Sub





Private Sub lwCRM_DblClick()
    If Modo <> 2 Then Exit Sub
    If lwCRM.ListItems.Count = 0 Then Exit Sub
    If lwCRM.SelectedItem Is Nothing Then Exit Sub


    If Me.DatosADevolverBusqueda <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un socio. No puede ver los documentos.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case CByte(RecuperaValor(lwCRM.Tag, 1))
    Case 0 ' Mantenimiento de Series
        Set frmSerie = New frmRepNumSerie2
        frmSerie.numSerie = lwCRM.SelectedItem.SubItems(3)
        frmSerie.codArtic = lwCRM.SelectedItem.SubItems(1)
        frmSerie.Show vbModal
        Set frmSerie = Nothing
    
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    lwCRM.SetFocus
'    Seleccionado = lwCRM.SelectedItem.Index
    CargaDatosLWCRM
    lwCRM.SelectedItem.Selected = False
    Set lwCRM.SelectedItem = Nothing
'    If lwCRM.ListItems.Count >= Seleccionado Then
'            lwCRM.ListItems(Seleccionado).Selected = True
'            lwCRM.ListItems(Seleccionado).EnsureVisible
'    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnImprimir_Click()
    printNou
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
Dim devuelve As String
Dim SQL As String

If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
'Text1(Index) = UCase(Text1(Index).Text)

Select Case Index
    Case 11 'cod situacion
        If Modo = 1 Then Exit Sub
        
        If Text1(Index).Text <> "" Then
            If IsNumeric(Text1(Index).Text) Then
               Text1(Index).Text = Format(Text1(Index).Text, "00")
                encontrado = DevuelveDesdeBD(conAri, "nomsitua", "ssitua", "codsitua", Text1(Index).Text, "T")
                If encontrado <> "" Then
                    Text2(0).Text = encontrado
                Else
                    MsgBox "El código de situación introducido no existe.", vbExclamation
                    PonerFoco Text1(Index)
                End If
            Else
                MsgBox "El código de situación debe ser numérico.", vbExclamation
                PonerFoco Text1(Index)
            End If
        End If
    Case 19 'banco
'        If Text1(Index).Text <> "" Then
'            If IsNumeric(Text1(Index).Text) Then
'                encontrado = DevuelveDesdeBD(conAri, "nombanpr", "sbanpr", "codbanpr", Text1(Index).Text, "T")
'                If encontrado <> "" Then
'                    'Text2(1).Text = encontrado
'                Else
'                    MsgBox "El código de banco introducido no existe.", vbExclamation
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                MsgBox "El código de banco debe ser numérico.", vbExclamation
'                PonerFoco Text1(Index)
'            End If
'        End If
    Case 0 'codsocio
        If Modo = 1 Then Exit Sub
        
        If Text1(Index).Text <> "" Then
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "El código de socio debe ser numérico.", vbExclamation
                PonerFoco Text1(Index)
            Else
                '[Monica]25/01/2012: comprobamos que el socio no exista
                If Modo = 3 Then
                    SQL = DevuelveDesdeBDNew(conAri, "sclien", "codclien", "codclien", Text1(0).Text, "N")
                    If SQL <> "" Then
                        MsgBox "El código de socio ya existe. Reintroduzca.", vbExclamation
                        PonerFoco Text1(Index)
                    End If
                End If
            End If
        End If
    Case 1 'numeruve
        If Modo = 1 Then Exit Sub
        
        If Text1(Index).Text <> "" Then
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "El código de vehiculo debe ser numérico.", vbExclamation
                PonerFoco Text1(Index)
            ElseIf Text1(Index).Text <= 0 Then
                MsgBox "El código de vehiculo tiene que tener un valor mayor que 0.", vbExclamation
                PonerFoco Text1(Index)
            Else
                VerificarVehiculo
            End If
        End If
    
    Case 12, 13, 18 'fecha alta,baja y situación
        If Text1(Index).Text <> "" Then
            PonerFormatoFecha Text1(Index)
        End If
    
    Case 4
        If Text1(Index).Text <> "" Then
            'Poblacion
            Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
            'provincia
            Text1(Index + 2).Text = devuelve
        End If
    
    Case 14
        If Modo = 1 Then Exit Sub
        
        If Text1(Index).Text <> "" Then
            encontrado = DevuelveDesdeBD(conAri, "nomchofe", "scoche", "codcoche", Text1(Index).Text, "T")
            If encontrado = "" Then
                MsgBox "El código de coche introducido no existe.", vbExclamation
                PonerFoco Text1(Index)
            Else
                Text1(Index).Text = Format(Text1(Index).Text, "0000")
                Text2(1).Text = encontrado
            End If
        End If
        
    Case 16
        Text1(Index).Text = Format(Text1(Index).Text, "00000000")
        
    Case 7 'nif
        If Text1(Index).Text <> "" Then
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
        End If
        
    Case 25, 27
        PonerFormatoDecimal Text1(Index), 6
    Case 42 ' codigo de iban
        Text1(Index).Text = UCase(Text1(Index).Text)
        
End Select
    
'[Monica]: calculo del iban si no lo ponen
If Index = 19 Or Index = 20 Or Index = 21 Or Index = 22 Then
    Dim Cta As String
    Dim CC As String
    If Text1(19).Text <> "" And Text1(20).Text <> "" And Text1(21).Text <> "" And Text1(22).Text <> "" Then
        
        Cta = Format(Text1(19).Text, "0000") & Format(Text1(20).Text, "0000") & Format(Text1(21).Text, "00") & Format(Text1(22).Text, "0000000000")
        If Len(Cta) = 20 Then
'        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)

            If Text1(28).Text = "" Then
                'NO ha puesto IBAN
                If DevuelveIBAN2("ES", Cta, Cta) Then Text1(28).Text = "ES" & Cta
            Else
                CC = CStr(Mid(Text1(28).Text, 1, 2))
                If DevuelveIBAN2(CStr(CC), Cta, Cta) Then
                    If Mid(Text1(28).Text, 3) <> Cta Then
                        
                        MsgBox "Codigo IBAN distinto del calculado [" & CC & Cta & "]", vbExclamation
                    End If
                End If
            End If
        End If
    End If
End If
        
End Sub

Private Sub VerificarVehiculo()
Dim encontrado As String
Dim Cliente As String

If Text1(13).Text <> "" Then 'si esta dado de baja no hace ninguna comprobación
    Cliente = "codclien"
    encontrado = DevuelveDesdeBD(conAri, "numeruve", "sclien", "numeruve", Text1(1).Text, "T", Cliente)
    Cliente = Format(Cliente, "0000")
    If encontrado <> "" Then
        If Not Cliente = Text1(0).Text Then
            MsgBox "El código de vehiculo ingresado esta asociado a otro Socio.", vbExclamation
        End If
    End If
End If

End Sub

Private Sub LimpiarDataGrids()
Dim SQL As String
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    'SQL = "select * from sclien_chofer where codsocio=-1"
    SQL = "select sclien_chofer.codsocio,sclien_chofer.numlinea,sclien_chofer.codchofe,schofe.nomchofe,sclien_chofer.fechaalt,sclien_chofer.fechabaj,sclien_chofer.obsevac from sclien_chofer inner join schofe on sclien_chofer.codsocio= -1 and sclien_chofer.codchofe=schofe.codchofe"
    CargaGridGnral DataGrid1, Adodc2, SQL, PrimeraVez
    CargaGrid DataGrid1, Adodc2

    SQL = "select sclien_publicidad.codsocio,sclien_publicidad.numlinea,sclien_publicidad.codclien,scliente.nomclien,sclien_publicidad.importes,sclien_publicidad.desdefec,sclien_publicidad.hastafec,if (sclien_publicidad.situacio=0, ""Activo"",""No Activo"") from sclien_publicidad inner join scliente on sclien_publicidad.codsocio= -1 and sclien_publicidad.codclien=scliente.codclien"
    CargaGridGnral DataGrid2, Adodc3, SQL, PrimeraVez
    CargaGrid DataGrid2, Adodc3

    SQL = "select sclien_contadores.codsocio,sclien_contadores.codtipom,stipom.nomtipom,sclien_contadores.contador from sclien_contadores inner join stipom on sclien_contadores.codsocio=-1 and sclien_contadores.codtipom=stipom.codtipom"
    CargaGridGnral DataGrid3, Adodc4, SQL, PrimeraVez
    CargaGrid DataGrid3, Adodc4
    '[Monica]05/12/2013: antes era sartic.preciove, ahora sclien_cuotas.importes
    SQL = "select sclien_cuotas.codsocio,sclien_cuotas.numlinea,sclien_cuotas.codartic,sartic.nomartic, sclien_cuotas.importes from sclien_cuotas inner join sartic on sclien_cuotas.codsocio=-1 and sclien_cuotas.codartic=sartic.codartic"
    CargaGridGnral DataGrid4, Adodc5, SQL, PrimeraVez
    CargaGrid DataGrid4, Adodc5
    
'    CargaGrid DataGrid1, Adodc2, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            
            TerminaBloquear
            CargaTxtAux False, False
            CargaTxtAux2 False, False
            CargaTxtAux3 False, False
            CargaTxtAux4 False, False
            
            If ModificaLineas = 1 Then 'INSERTAR
                If SSTab1.Tab = 2 Then
                    ModificaLineas = 0
                    DataGrid1.AllowAddNew = False
                    If Not Adodc2.Recordset.EOF Then Adodc2.Recordset.MoveFirst
                ElseIf SSTab1.Tab = 3 Then
                    ModificaLineas = 0
                    DataGrid2.AllowAddNew = False
                    If Not Adodc3.Recordset.EOF Then Adodc3.Recordset.MoveFirst
                ElseIf SSTab1.Tab = 6 Then
                    ModificaLineas = 0
                    DataGrid3.AllowAddNew = False
                    If Not Adodc4.Recordset.EOF Then Adodc4.Recordset.MoveFirst
                Else
                    ModificaLineas = 0
                    DataGrid4.AllowAddNew = False
                    If Not Adodc5.Recordset.EOF Then Adodc5.Recordset.MoveFirst
                End If
            Else
                ModificaLineas = 0
            End If
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            Me.DataGrid2.Enabled = True
            Me.DataGrid4.Enabled = True
            Me.DataGrid3.Enabled = True
            
            
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
        Case 9 'chofer
            mnLineas_Click
        Case 10 'publicidad
            mnLineas2_Click
        Case 11 'cuotas
            mnLineas3_Click
        Case 12 'contadores
            mnLineas4_Click
        Case 13  'imprimir
            printNou
        Case 14  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub
Private Sub mnLineas_Click()
    BotonMtoLineas "Choferes"
End Sub

Private Sub mnLineas2_Click()
    BotonMtoLineas "Publicidad"
End Sub

Private Sub mnLineas3_Click()
    BotonMtoLineas "Cuotas"
End Sub

Private Sub mnLineas4_Click()
    BotonMtoLineas "Contadores"
End Sub



Private Sub BotonMtoLineas(cad As String)
        Select Case cad
            Case "Choferes"
                SSTab1.Tab = 2
            Case "Publicidad"
                SSTab1.Tab = 3
            Case "Cuotas"
                SSTab1.Tab = 4
            Case "Contadores"
                SSTab1.Tab = 6
            
        End Select
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
        
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
        Select Case SSTab1.Tab
            Case 2
                BotonAnyadirLinea
            Case 3
                BotonAnyadirLinea2
            Case 4
                BotonAnyadirLinea3
            Case 6
                BotonAnyadirLinea4
        End Select
    Else 'Añadir Cabecera de Pedidos
         BotonAnyadir
    End If
End Sub
Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador(0).Caption = "INSERTAR Chofer"
    
    AnyadirLinea DataGrid1, Adodc2
    CargaTxtAux True, True
   
    PonerFoco txtAux1(0)
    Me.DataGrid1.Enabled = False
End Sub

Private Sub BotonAnyadirLinea2()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador(0).Caption = "INSERTAR Publicidad"
    
    AnyadirLinea DataGrid2, Adodc3
    CargaTxtAux2 True, True
   
    PonerFoco txtAux2(0)
    Me.DataGrid2.Enabled = False
End Sub

Private Sub BotonAnyadirLinea3()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador(0).Caption = "INSERTAR Cuotas"
    
    AnyadirLinea DataGrid4, Adodc5
    CargaTxtAux3 True, True
   
    PonerFoco txtAux3(0)
    Me.DataGrid2.Enabled = False
End Sub

Private Sub BotonAnyadirLinea4()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador(0).Caption = "INSERTAR Contadores"
    
    AnyadirLinea DataGrid3, Adodc4
    CargaTxtAux4 True, True
   
    PonerFoco txtAux4(0)
    Me.DataGrid3.Enabled = False
End Sub


Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux1.Count - 1 'TextBox
            txtAux1(i).Top = 290
            txtAux1(i).visible = visible
        Next i
        Me.cmdAux(0).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux1.Count - 1
                txtAux1(i).Text = ""
                BloquearTxt txtAux1(i), False
            Next i
        Else 'Vamos a modificar
            For i = 0 To txtAux1.Count - 1
                txtAux1(i).Text = DataGrid1.Columns(i + 2).Text
                If i >= 2 Then
                    txtAux1(i).Locked = False
                    txtAux1(i).BackColor = &H80000005
                Else
                    txtAux1(i).Locked = True
                End If
            Next i
            cmdAux(0).Enabled = False
        End If
        
        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 20)
        
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).Top = alto
            txtAux1(i).Height = DataGrid1.RowHeight
        Next i
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'chofer
        txtAux1(0).Left = DataGrid1.Left + 330
        txtAux1(0).Width = DataGrid1.Columns(2).Width - 160
        cmdAux(0).Left = txtAux1(0).Left + txtAux1(0).Width - 50
'        txtAux1(0).Left = DataGrid1.Left + 330
'        txtAux1(0).Width = DataGrid1.Columns(2).Width - 100
        
        'nombre
        txtAux1(1).Left = cmdAux(0).Left + cmdAux(0).Width + 10
        txtAux1(1).Width = DataGrid1.Columns(3).Width - 50
'        txtAux1(1).Width = DataGrid1.Columns(3).Width - 100
'        txtAux1(1).Left = txtAux1(0).Left + (txtAux1(0).Width + 100)
        
        'fecha alta
        txtAux1(2).Left = txtAux1(1).Left + txtAux1(1).Width + 55
        txtAux1(2).Width = DataGrid1.Columns(4).Width - 30
'        txtAux1(2).Width = DataGrid1.Columns(4).Width - 100
'        txtAux1(2).Left = txtAux1(1).Left + (txtAux1(1).Width + 100)
        
        'fecha baja
        txtAux1(3).Left = txtAux1(2).Left + txtAux1(2).Width + 35
        txtAux1(3).Width = DataGrid1.Columns(5).Width - 30
'        txtAux1(3).Width = DataGrid1.Columns(5).Width - 100
'        txtAux1(3).Left = txtAux1(2).Left + (txtAux1(2).Width + 100)
        'observaciones
        txtAux1(4).Left = txtAux1(3).Left + txtAux1(3).Width + 30
        txtAux1(4).Width = DataGrid1.Columns(6).Width - 25
        
'        txtAux1(4).Width = DataGrid1.Columns(5).Width - 100
'        txtAux1(4).Left = txtAux1(3).Left + (txtAux1(3).Width + 100)
        
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).visible = visible
        Next i
        Me.cmdAux(0).Height = Me.DataGrid1.RowHeight
        Me.cmdAux(0).Top = alto
        Me.cmdAux(0).visible = visible
        cmdAux1.Top = alto
        cmdAux1.visible = visible
    End If
End Sub

Private Sub CargaTxtAux2(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte


    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux2.Count - 1 'TextBox
            txtAux2(i).Top = 290
            txtAux2(i).visible = visible
        Next i
'        cmdAux1.Top = 290
'        cmdAux1.visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid2
            For i = 0 To txtAux2.Count - 1
                txtAux2(i).Text = ""
                BloquearTxt txtAux2(i), False
            Next i
            cmdAux1.ListIndex = 0
        Else 'Vamos a modificar
            For i = 0 To txtAux2.Count - 1
                txtAux2(i).Text = DataGrid2.Columns(i + 2).Text
                txtAux2(i).Locked = False
            Next i
        End If
            If Not Adodc3.Recordset.EOF Then
                If Adodc3.Recordset.Fields(7) = "Activo" Then
                    cmdAux1.ListIndex = 0
                Else
                    cmdAux1.ListIndex = 1
                End If
            End If

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid2, 20)
        
        For i = 0 To txtAux2.Count - 1
            txtAux2(i).Top = alto
            txtAux2(i).Height = DataGrid2.RowHeight
        Next i
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'cliente
        txtAux2(0).Left = DataGrid2.Left + 330
        txtAux2(0).Width = DataGrid2.Columns(2).Width - 100
        'nombre
        txtAux2(1).Width = DataGrid2.Columns(3).Width - 100
        txtAux2(1).Left = txtAux2(0).Left + (txtAux2(0).Width + 100)
        'importe
        txtAux2(2).Width = DataGrid2.Columns(4).Width - 100
        txtAux2(2).Left = txtAux2(1).Left + (txtAux2(1).Width + 100)
        'desde
        txtAux2(3).Width = DataGrid2.Columns(5).Width - 100
        txtAux2(3).Left = txtAux2(2).Left + (txtAux2(2).Width + 100)
        'hasta
        txtAux2(4).Width = DataGrid2.Columns(6).Width - 100
        txtAux2(4).Left = txtAux2(3).Left + (txtAux2(3).Width + 100)
        
        cmdAux1.Width = DataGrid2.Columns(7).Width - 100
        cmdAux1.Left = txtAux2(4).Left + (txtAux2(4).Width + 100)
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux2.Count - 1
            txtAux2(i).visible = visible
        Next i
    End If
    Me.cmdAux(1).Height = Me.DataGrid2.RowHeight
    Me.cmdAux(1).Top = alto
    Me.cmdAux(1).visible = visible
    cmdAux1.Top = alto
    cmdAux1.visible = visible
End Sub


Private Sub CargaTxtAux3(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte


    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux3.Count - 1 'TextBox
            txtAux3(i).Top = 290
            txtAux3(i).visible = visible
        Next i
        cmdAux(2).Top = 290
        cmdAux(2).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid4
            For i = 0 To txtAux3.Count - 1
                txtAux3(i).Text = ""
                BloquearTxt txtAux3(i), False
            Next i
        Else 'Vamos a modificar
            '[Monica]05/12/2013: solo dejamos modificar el importe
            For i = 0 To txtAux3.Count - 1
                txtAux3(i).Text = DataGrid4.Columns(i + 2).Text
                txtAux3(2).Locked = False
            Next i
            '[Monica]05/12/2013: dejamos modificar el importe
            BloquearTxt txtAux3(2), False
        End If
        

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid4, 20)
        
        For i = 0 To txtAux3.Count - 1
            txtAux3(i).Top = alto
            txtAux3(i).Height = DataGrid4.RowHeight
        Next i
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'articulo
        txtAux3(0).Left = DataGrid4.Left + 330
        txtAux3(0).Width = DataGrid4.Columns(2).Width - 100
        
        cmdAux(2).Left = txtAux3(0).Left + txtAux3(0).Width - 50
        
        'nombre
        txtAux3(1).Width = DataGrid4.Columns(3).Width - 100
        txtAux3(1).Left = cmdAux(2).Left + cmdAux(2).Width + 10
        
        'importe
        txtAux3(2).Width = DataGrid4.Columns(4).Width - 100
        txtAux3(2).Left = txtAux3(1).Left + (txtAux3(1).Width + 100)
        
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux3.Count - 1
            txtAux3(i).visible = visible
        Next i
    End If
    cmdAux(2).Top = alto
    cmdAux(2).visible = visible
End Sub


Private Sub CargaTxtAux4(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte


    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux4.Count - 1 'TextBox
            txtAux4(i).Top = 290
            txtAux4(i).visible = visible
        Next i
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid3
            For i = 0 To txtAux4.Count - 1
                txtAux4(i).Text = ""
                BloquearTxt txtAux4(i), False
            Next i
        Else 'Vamos a modificar
            For i = 0 To txtAux4.Count - 1
                txtAux4(i).Text = DataGrid3.Columns(i + 1).Text
                If i = 2 Then
                    txtAux4(i).Locked = False
                    txtAux4(i).BackColor = &H80000005
                Else
                    txtAux4(i).Locked = True
                End If
            Next i
        End If
        

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid3, 20)
        
        For i = 0 To txtAux4.Count - 1
            txtAux4(i).Top = alto
            txtAux4(i).Height = DataGrid3.RowHeight
        Next i
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'articulo
        txtAux4(0).Left = DataGrid3.Left + 330
        txtAux4(0).Width = DataGrid3.Columns(1).Width - 10
        
'        cmdAux(3).Left = TxtAux4(0).Left + TxtAux4(0).Width - 50
        
        'nombre
        txtAux4(1).Width = DataGrid3.Columns(2).Width - 10
        txtAux4(1).Left = txtAux4(0).Left + txtAux4(0).Width + 10
        
        'importe
        txtAux4(2).Width = DataGrid3.Columns(3).Width - 10
        txtAux4(2).Left = txtAux4(1).Left + (txtAux4(1).Width + 10)
        
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux4.Count - 1
            txtAux4(i).visible = visible
        Next i
    End If
End Sub




Private Sub BotonAnyadir()
Dim SQL As String
Dim Codigo As String
'Añadir registro en tabla de trabajadores: straba (Cabecera)

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
'    Text1(1).BackColor = &HFFFFC0
    PonerFoco Text1(0)
    'busco el codtarif correspondiente al menor codlista que tenga como valor
    'en bonifica=0
    SQL = "select min(codlista) from starif where bonifica=0"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'no puede ser eof
    Text1(23).Text = miRsAux.Fields(0)
    miRsAux.Close
    Set miRsAux = Nothing
    Text1(12).Text = Date
    '[Monica]25/01/2012: añado la condicion de codclien < 1998
    Codigo = SugerirCodigoSiguienteStr("sclien", "codclien", "codclien < 1998")
    Text1(0).Text = Codigo
End Sub
Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Pedido
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub
Private Sub BotonModificarLinea()
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String
On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If SSTab1.Tab = 2 Then
        'chofer
        If Adodc2.Recordset.EOF Then Exit Sub
        vWhere = "codsocio=" & Adodc2.Recordset!codSocio & " and numlinea=" & Adodc2.Recordset!numlinea
    
        If Not BloqueaRegistro("sclien_chofer", vWhere) Then Exit Sub
    
        CargaTxtAux True, False
        Me.lblIndicador(0).Caption = "MODIFICAR CHOFERES"
        PonerFoco txtAux1(2)
        Me.DataGrid1.Enabled = False
    Else
        If SSTab1.Tab = 3 Then
            'publicidad
            If Adodc3.Recordset.EOF Then Exit Sub
            vWhere = "codsocio=" & Adodc3.Recordset!codSocio & " and numlinea=" & Adodc3.Recordset!numlinea
        
            If Not BloqueaRegistro("sclien_publicidad", vWhere) Then Exit Sub
        
            CargaTxtAux2 True, False
            Me.lblIndicador(0).Caption = "MODIFICAR PUBLICIDAD"
            PonerFoco txtAux2(0)
            Me.DataGrid2.Enabled = False
        Else
            '[Monica]05/12/2013: Dejo modificar el articulo de cuota
            If SSTab1.Tab = 4 Then
                'cuotas
                If Adodc5.Recordset.EOF Then Exit Sub
                vWhere = "codsocio=" & Adodc5.Recordset!codSocio & " and numlinea=" & Adodc5.Recordset!numlinea
            
                If Not BloqueaRegistro("sclien_cuotas", vWhere) Then Exit Sub
            
                CargaTxtAux3 True, False
                Me.lblIndicador(0).Caption = "MODIFICAR CUOTAS"
                PonerFoco txtAux3(2)
                '[Monica]05/12/2013: cambio false por true en el datagrid.enabled
                Me.DataGrid4.Enabled = True
            Else
                If SSTab1.Tab = 6 Then
                    'publicidad
                    If Adodc4.Recordset.EOF Then Exit Sub
                    vWhere = "codsocio=" & Adodc4.Recordset!codSocio & " and codtipom='" & Adodc4.Recordset!codtipom & "'"
                
                    If Not BloqueaRegistro("sclien_contadores", vWhere) Then Exit Sub
                
                    CargaTxtAux4 True, False
                    Me.lblIndicador(0).Caption = "MODIFICAR CONTADORES"
                    PonerFoco txtAux4(2)
                    Me.DataGrid3.Enabled = False
                End If
            End If
        End If
    End If
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
        
    
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub
Private Sub BotonModificar()
'Prepara el Form para Modificar la cabecera de Pedidos (tabla: scaped)
Dim DeVarios As Boolean
Dim SQL As String
On Error GoTo EModificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(1)
    Text1(1).BackColor = &HFFFFC0
'    Text1(2).Enabled = False
'    Text1(2).BackColor = &H80000018
    
   
EModificar:
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub mnEliminar_Click()
If Modo = 5 Then
    BotonEliminarFila
Else
    BotonEliminar
End If

End Sub
Private Sub BotonEliminar()
Dim msg As String
Dim SQL As String
Dim encontrado As String

On Error GoTo EEliminar

msg = "Esta seguro que desea eliminar el socio:" & Me.Adodc1.Recordset!CodClien  '  Text1(0).Text & "?"
If MsgBox(msg, vbQuestion + vbYesNo) = vbYes Then
    NumRegElim = Adodc1.Recordset.AbsolutePosition
    encontrado = DevuelveDesdeBD(conAri, "codclien", "scafac", "codclien", Me.Adodc1.Recordset!CodClien, "T")
    If encontrado <> "" Then
        MsgBox "No es posible eliminar este socio, ya que tiene facturas asociadas.", vbExclamation
        Exit Sub
    End If
    encontrado = DevuelveDesdeBD(conAri, "codsocio", "shiuve", "codsocio", Me.Adodc1.Recordset!CodClien, "T")
    If encontrado <> "" Then
        MsgBox "No es posible eliminar este socio, se encuentra en el histórico de uves.", vbExclamation
        Exit Sub
    End If
    conn.BeginTrans
    'Primero borramos las lineas
    SQL = "delete from schofe_historia where (codchofe, numeruve) in (select codchofe, " & DBSet(Me.Adodc1.Recordset!NumerUve, "N") & "  from sclien_chofer where codsocio = " & DBSet(Me.Adodc1.Recordset!CodClien, "N") & ")"
    conn.Execute SQL
    
    SQL = "Delete from sclien_chofer where codsocio=" & Me.Adodc1.Recordset!CodClien
    conn.Execute SQL
    
    SQL = "Delete from sclien_publicidad where codsocio=" & Me.Adodc1.Recordset!CodClien
    conn.Execute SQL
    
    SQL = "Delete from sclien_cuotas where codsocio = " & Me.Adodc1.Recordset!CodClien
    conn.Execute SQL
    
    SQL = "Delete from sclien_contadores where codsocio = " & Me.Adodc1.Recordset!CodClien
    conn.Execute SQL
    
    
    'Ahora cabecera
    SQL = "Delete from sclien where codclien=" & Me.Adodc1.Recordset!CodClien
    conn.Execute SQL
    conn.CommitTrans
End If

If SituarDataTrasEliminar(Adodc1, NumRegElim) Then
    PonerCampos
End If

EEliminar:
If Err.Number <> 0 Then
    conn.RollbackTrans
    MsgBox "Error al eliminar Socio." & Err.Description
End If
End Sub

Private Sub BotonEliminarFila()
Dim msg As String
Dim SQL As String

On Error GoTo EEliminarLineas

msg = "Esta seguro que desea eliminar la linea?"
If MsgBox(msg, vbQuestion + vbYesNo) = vbYes Then
    Select Case SSTab1.Tab
        Case 2
            'Ahora borramos las lineas
            SQL = "delete from schofe_historia where codchofe = " & Adodc2.Recordset!codchofe & " and numeruve = " & DBSet(Text1(1).Text, "N")
            SQL = SQL & " and fechaini = " & DBSet(Adodc2.Recordset!Fechaalt, "F")
            conn.Execute SQL
            
            SQL = "Delete from sclien_chofer where codsocio=" & Text1(0).Text & " and numlinea = " & Adodc2.Recordset!numlinea
            conn.Execute SQL
            
            CargaGrid DataGrid1, Me.Adodc2
        Case 3
            SQL = "Delete from sclien_publicidad where codsocio=" & Text1(0).Text & " and numlinea = " & Adodc3.Recordset!numlinea
            conn.Execute SQL
            
            CargaGrid DataGrid2, Me.Adodc3
        Case 4
            SQL = "Delete from sclien_cuotas where codsocio=" & Text1(0).Text & " and numlinea = " & Adodc5.Recordset!numlinea
            conn.Execute SQL
            
            CargaGrid DataGrid4, Me.Adodc5
        Case 6
            SQL = "Delete from sclien_contadores where codsocio=" & Text1(0).Text & " and codtipom = " & DBSet(Adodc4.Recordset!codtipom, "T")
            conn.Execute SQL
            
            CargaGrid DataGrid3, Me.Adodc4
    End Select
End If

EEliminarLineas:
If Err.Number <> 0 Then
    MsgBox "Error al eliminar Lineas." & Err.Description
End If

End Sub

Private Sub mnSalir_Click()
    Unload Me
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
        BuscaChekc = ""
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
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
        PonerModo 2
        PonerCampos
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
    
    
    'data2 para el grid de las lineas chofer
    Adodc2.ConnectionString = conn
    Adodc2.RecordSource = "select sclien_chofer.codsocio,sclien_chofer.numlinea,sclien_chofer.codchofe,schofe.nomchofe,sclien_chofer.fechaalt,sclien_chofer.fechabaj,sclien_chofer.obsevac from sclien_chofer inner join schofe on sclien_chofer.codsocio=" & Text1(0).Text & " and sclien_chofer.codchofe=schofe.codchofe"
    Adodc2.Refresh
    
    CargaGrid DataGrid1, Adodc2
    
    'data3 para el grid de las lineas publicidad
    Adodc3.ConnectionString = conn
    Adodc3.RecordSource = "select sclien_publicidad.codsocio,sclien_publicidad.numlinea,sclien_publicidad.codclien,scliente.nomclien,sclien_publicidad.importes,sclien_publicidad.desdefec,sclien_publicidad.hastafec,if (sclien_publicidad.situacio=0, ""Activo"",""No Activo"") from sclien_publicidad inner join scliente on sclien_publicidad.codsocio= " & Text1(0).Text & " and sclien_publicidad.codclien=scliente.codclien"
    Adodc3.Refresh
    
    CargaGrid DataGrid2, Adodc3
    
    'data4 para el grid contadores
    Adodc4.ConnectionString = conn
    Adodc4.RecordSource = "select sclien_contadores.codsocio,sclien_contadores.codtipom,stipom.nomtipom,sclien_contadores.contador from sclien_contadores inner join stipom on sclien_contadores.codsocio=" & Text1(0).Text & " and sclien_contadores.codtipom=stipom.codtipom"
    Adodc4.Refresh
    CargaGrid DataGrid3, Adodc4
    
    'data5 para el grid cuotas
    Adodc5.ConnectionString = conn
    '[Monica]05/12/2013: Modificado enseñamos el importe almacenado no el preciove del articulo : sartic.preciove
    Adodc5.RecordSource = "select sclien_cuotas.codsocio, sclien_cuotas.numlinea, sclien_cuotas.codartic, sartic.nomartic, sclien_cuotas.importes from sclien_cuotas inner join sartic on sclien_cuotas.codsocio=" & Text1(0).Text & " and sclien_cuotas.codartic=sartic.codartic "
    Adodc5.Refresh
    CargaGrid DataGrid4, Adodc5
    
    
    
    If Text1(11).Text <> "" Then
        encontrado = DevuelveDesdeBD(conAri, "nomsitua", "ssitua", "codsitua", Text1(11).Text, "T")
        Text2(0).Text = encontrado
    End If
'    If Text1(19).Text <> "" Then
'        encontrado = DevuelveDesdeBD(conAri, "nombanpr", "sbanpr", "codbanpr", Text1(19).Text, "T")
'        Text2(1).Text = encontrado
'    End If
    If Text1(14).Text <> "" Then
        encontrado = DevuelveDesdeBD(conAri, "nomchofe", "scoche", "codcoche", Text1(14).Text, "T")
        Text2(1).Text = encontrado
    End If
    
    '-- Esto permanece para saber donde estamos
    lblIndicador(0).Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    
    CargaDatosLWDoc
    CargaDatosLWCRM
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)   ', enlaza As Boolean)
Dim i As Integer
Dim SQL As String

On Error GoTo ECargaGrid

    vData.Refresh
    Set vDataGrid.DataSource = vData
    vDataGrid.Columns(0).visible = False 'codcoche

    If vDataGrid.Name = "DataGrid1" Then
        vDataGrid.Columns(1).visible = False 'numlinea
        vDataGrid.Columns(2).Caption = "Chofer"
        vDataGrid.Columns(2).Width = 1000
        vDataGrid.Columns(2).NumberFormat = "0000"
        vDataGrid.Columns(3).Caption = "Nombre"
        vDataGrid.Columns(3).Width = 2800
        vDataGrid.Columns(4).Caption = "Fecha Alta"
        vDataGrid.Columns(4).Width = 1100
        vDataGrid.Columns(5).Caption = "Fecha Baja"
        vDataGrid.Columns(5).Width = 1100
        vDataGrid.Columns(6).Caption = "Observaciones"
        vDataGrid.Columns(6).Width = 3600
    ElseIf vDataGrid.Name = "DataGrid2" Then
        vDataGrid.Columns(1).visible = False 'numlinea
        vDataGrid.Columns(2).Caption = "Cliente"
        vDataGrid.Columns(2).Width = 1100
        vDataGrid.Columns(2).NumberFormat = "000000"
        vDataGrid.Columns(3).Caption = "Nombre"
        vDataGrid.Columns(3).Width = 3300
        vDataGrid.Columns(4).Caption = "Importe"
        vDataGrid.Columns(4).Width = 1500
        vDataGrid.Columns(4).NumberFormat = "#,###,###,##0.00"
        vDataGrid.Columns(4).Alignment = dbgRight
        vDataGrid.Columns(5).Caption = "Desde"
        vDataGrid.Columns(5).Width = 1200
        vDataGrid.Columns(6).Caption = "Hasta"
        vDataGrid.Columns(6).Width = 1200
        vDataGrid.Columns(7).Caption = "Situación"
        vDataGrid.Columns(7).Width = 1300
    ElseIf vDataGrid.Name = "DataGrid3" Then
        vDataGrid.Columns(1).Caption = "Tipo Movimiento"
        vDataGrid.Columns(1).Width = 1700
        vDataGrid.Columns(2).Caption = "Nombre"
        vDataGrid.Columns(2).Width = 5800
        vDataGrid.Columns(3).Caption = "Contador"
        vDataGrid.Columns(3).Width = 2100
        vDataGrid.Columns(3).Alignment = dbgRight
        vDataGrid.Columns(3).NumberFormat = "0000000"
    ElseIf vDataGrid.Name = "DataGrid4" Then
        vDataGrid.Columns(1).visible = False 'numlinea
        vDataGrid.Columns(2).Caption = "Artículo"
        vDataGrid.Columns(2).Width = 1700
        vDataGrid.Columns(3).Caption = "Nombre"
        vDataGrid.Columns(3).Width = 5800
        vDataGrid.Columns(4).Caption = "Importe"
        vDataGrid.Columns(4).Width = 2100
        vDataGrid.Columns(4).Alignment = dbgRight
        vDataGrid.Columns(4).NumberFormat = "###,##0.0000"
        
        If Not Me.Adodc1.Recordset.EOF Then
            '[Monica]05/12/2013: tengo que sumar lo que hay almacenado
            'antes
            'SQL = "select sum(preciove) from sartic inner join sclien_cuotas on sartic.codartic = sclien_cuotas.codartic where codsocio = " & DBSet(Me.Adodc1.Recordset!CodClien, "N")
            'ahora
            SQL = "select sum(importes) from sclien_cuotas where codsocio = " & DBSet(Me.Adodc1.Recordset!CodClien, "N")
            
            CalcularTotales SQL
        Else
            Text3.Text = ""
        End If
        
        
    End If


    vDataGrid.Enabled = (Modo = 0) Or (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    vDataGrid.ScrollBars = dbgAutomatic

    Exit Sub

ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description

End Sub





Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    '##A mano
    cad = ""
    cad = cad & ParaGrid(Text1(0), 14, "Código")
    cad = cad & ParaGrid(Text1(2), 55, "Nombre")
    cad = cad & ParaGrid(Text1(1), 14, "Uve")
    cad = cad & ParaGrid(Text1(8), 14, "Teléfono")

    Tabla = "sclien LEFT JOIN sclien_cuotas ON sclien.codclien = sclien_cuotas.codsocio "
    Titulo = "Socios"
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|2|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
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
Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 1: dirMail = Text1(10).Text
    End Select

    If LanzaMailGnral(dirMail) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(codclien=" & Text1(0).Text & ")"
    If SituarData(Adodc1, cad, Indicador) Then
       PonerModo 2
       lblIndicador(0).Caption = Indicador
        'data4 para el grid contadores
        Adodc4.ConnectionString = conn
        Adodc4.RecordSource = "select sclien_contadores.codsocio,sclien_contadores.codtipom,stipom.nomtipom,sclien_contadores.contador from sclien_contadores left join stipom on sclien_contadores.codsocio=" & Text1(0).Text & " and sclien_contadores.codtipom=stipom.codtipom"
        Adodc4.Refresh
        CargaGrid DataGrid3, Adodc4
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux1(Index), cadkey
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux1_LostFocus(Index As Integer)
Dim encontrado As String

txtAux1(Index).Text = UCase(txtAux1(Index).Text)

    Select Case Index
        Case 2, 3
            PonerFormatoFecha txtAux1(Index)
        Case 0
            If txtAux1(Index).Text <> "" Then
                txtAux1(Index).Text = Format(txtAux1(Index).Text, "0000")
                encontrado = DevuelveDesdeBD(conAri, "nomchofe", "schofe", "codchofe", txtAux1(Index).Text, "T")
                If encontrado <> "" Then
                    txtAux1(1).Text = encontrado
                Else
                    MsgBox "No existe el código de chofer introducido.", vbExclamation
                    PonerFoco txtAux1(Index)
                End If
            End If
            
        Case 4
            cmdAceptar.SetFocus
    End Select
End Sub

Private Sub txtAux2_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux2(Index), cadkey
End Sub

Private Sub txtAux2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub txtAux2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux2_LostFocus(Index As Integer)
Dim encontrado As String

txtAux2(Index).Text = UCase(txtAux2(Index).Text)

    Select Case Index
        Case 0
            If txtAux2(Index).Text <> "" Then
                txtAux2(Index).Text = Format(txtAux2(Index).Text, "000000")
                encontrado = DevuelveDesdeBD(conAri, "nomclien", "scliente", "codclien", txtAux2(Index).Text, "T")
                If encontrado <> "" Then
                    txtAux2(1).Text = encontrado
                Else
                    MsgBox "No existe el código de cliente introducido.", vbExclamation
                    PonerFoco txtAux2(Index)
                End If
            End If
        Case 3, 4
            PonerFormatoFecha txtAux2(Index)
        Case 2
            PonerFormatoDecimal txtAux2(Index), 1
    End Select
End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux3(Index), cadkey
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
Dim encontrado As String
Dim campo1 As String

    txtAux3(Index).Text = Trim(txtAux3(Index).Text)
    If txtAux3(Index).Text = "" Then Exit Sub

    Select Case Index
        Case 0
            If txtAux3(Index).Text <> "" Then
                campo1 = "preciove"
                encontrado = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtAux3(Index).Text, "T", campo1)
                If encontrado <> "" Then
                    txtAux3(1).Text = encontrado
                    txtAux3(2).Text = Format(campo1, "###,##0.0000")
                    
                    Me.cmdAceptar.SetFocus
                Else
                    MsgBox "No existe el código de artículo introducido.", vbExclamation
                    PonerFoco txtAux3(Index)
                End If
            Else
                Me.cmdCancelar.SetFocus
            End If
            
        '[Monica]05/12/2013: dejo modificar el importe
        Case 2 ' importe
            PonerFormatoDecimal txtAux3(Index), 2
    
    End Select
End Sub

Private Sub TxtAux4_GotFocus(Index As Integer)
Dim cadkey As Integer
    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux4(Index), cadkey
End Sub

Private Sub TxtAux4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub TxtAux4_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux4_LostFocus(Index As Integer)
Dim encontrado As String
Dim campo1 As String

    Select Case Index
        Case 0, 1
            UCase (txtAux4(Index))
'            If txtAux4(Index).Text <> "" Then
'                encontrado = DevuelveDesdeBD(conAri, "nomtipom", "stipom", "codtipomc", txtAux4(Index).Text, "T")
'                If encontrado <> "" Then
'                    txtAux4(1).Text = encontrado
'                    txtAux4(2).Text = Format(campo1, "000000")
'
'                    Me.cmdAceptar.SetFocus
'
'                Else
'                    MsgBox "No existe el tipom de movimiento introducido.", vbExclamation
'                    PonerFoco txtAux4(Index)
'                End If
'            Else
'
'                Me.cmdCancelar.SetFocus
'
'            End If
       Case 2
            If PonerFormatoEntero(txtAux4(Index)) Then
                cmdAceptar.SetFocus
            Else
                cmdCancelar.SetFocus
            End If
    
    End Select
End Sub


Private Sub printNou()
Dim Nombre As String
'    If MsgBox("¿Desea imprimir los datos para envio a socio?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'        Nombre = "rGesSociosEtiq.rpt"
'    Else
'        Nombre = "rGesSocios.rpt"
'    End If
    With frmImprimir2
        .cadTabla2 = "((sclien left join sclien_cuotas) left join sclien_chofer) left join sclien_publicidad"
        .Informe2 = "rGesSocios.rpt"
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


'[Monica]10/02/2011
Private Sub ImagenesNavegacion()
    With Me.Toolbar2
        .ImageList = frmPpal.ImgListPpal
        .Buttons(1).Image = 7
        .Buttons(3).Image = 8
        .Buttons(5).Image = 36
        
    End With
    
    Set lw1.SmallIcons = frmPpal.ImgListPpal

    With Me.Toolbar3
        .ImageList = frmPpal.ImgListPpal
        .Buttons(1).Image = 16
        .Buttons(3).Image = 30
        .Buttons(5).Image = 25
        .Buttons(7).Image = 13
        .Buttons(9).Image = 31
'            .Buttons(11).Image = 12
    End With
    
    Set lwCRM.SmallIcons = frmPpal.ImgListPpal
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Tag = "" Then Exit Sub
    LabelDoc.Caption = ""
    'Levantamos todos los botones y dejamos pulsado el de ahora
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Tag <> "" Then
            If Toolbar2.Buttons(NumRegElim).Index <> Button.Index Then Toolbar2.Buttons(NumRegElim).Value = tbrUnpressed
        End If
    Next NumRegElim
    If CByte(Button.Tag) = 0 Then
        Button.Tag = "1"
    End If
    CargaColumnas CByte(Button.Tag)
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLWDoc
End Sub

Private Sub CargaColumnas(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader

    Select Case OpcionList
    Case 2
        'Servicios del socios
        LabelDoc.Caption = "Servicios"
        Columnas = "Fecha|Hora|Nº V|Tipo|Domicilio|Importe|"
        Ancho = "1400|600|600|500|3400|1300|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|1|0|0|1|"
        'Formatos
        Formato = "dd/mm/yyyy|hh:mm|0000|0||" & FormatoImporte & "|"
        Ncol = 6
    
    Case 3, 4 ' tanto de venta como de compra
        'FACTURAS
        LabelDoc.Caption = "Facturas"
        Columnas = "Tipo|Numero|Fecha|Importe|"
        Ancho = "1000|2000|1200|2500|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|"
        'Formatos
        Formato = "|00000000|dd/mm/yyyy|" & FormatoImporte & "|"
        Ncol = 4
    End Select
    
    
    'Fecha incio busquedas
    Text1(26).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")
    'Guardo la opcion en el tag
    lw1.Tag = OpcionList & "|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
End Sub

Private Sub CargaDatosLWDoc()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador(0).Caption
    lblIndicador(0).Caption = "Leyendo " & LabelDoc.Caption
    lblIndicador(0).Refresh
    CargaDatosLWDoc2
    Me.lblIndicador(0).Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLWDoc2()
Dim cad As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim EsDTOFam As Boolean

    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Value = tbrPressed Then
            ElIcono = Toolbar2.Buttons(NumRegElim).Image
            Exit For
        End If
    Next
    
    
    'Fecha incio busquedas
    Text1(26).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")
    EsDTOFam = False
    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 2
        'LLAMADAS
        cad = "select fecha,hora,numeruve,tipservi,dirllama,impventa from shilla WHERE "
        cad = cad & " codsocio=" & Adodc1.Recordset!CodClien
        GroupBy = "1,2,3"
        BuscaChekc = "fecha"
        
    Case 3
        'FACTURAS como cliente
        cad = "select codtipom,numfactu,fecfactu,totalfac from scafac WHERE "
        cad = cad & " codclien=" & Adodc1.Recordset!CodClien
        BuscaChekc = "fecfactu"
        GroupBy = "1,2,3"
        
    Case 4
        'FACTURAS COMO PROVEEDOR
        cad = "select codtipom,numfactu,fecfactu,totalfac from sfactusoc WHERE "
        cad = cad & " codsocio=" & Adodc1.Recordset!CodClien
        BuscaChekc = "fecfactu"
        GroupBy = "1,2,3"
        
    End Select
    
    
    'La fecha
    If BuscaChekc <> "" Then cad = cad & " and " & BuscaChekc & " >='" & Format(imgFecha(3).Tag, FormatoFecha) & "'"
    
    'El group by
    If GroupBy <> "" Then cad = cad & " GROUP BY " & GroupBy
    
    'El ORDER BY
    cad = cad & " ORDER BY " & BuscaChekc & " DESC"
    
    BuscaChekc = ""
    
    lw1.ListItems.Clear
    If cad <> "" Then
        Set RS = New ADODB.Recordset
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            Set It = lw1.ListItems.Add()
            If lw1.ColumnHeaders(1).Tag <> "" Then
            It.Text = Format(RS.Fields(0), lw1.ColumnHeaders(1).Tag)
            Else
                It.Text = RS.Fields(0)
            End If
            'El resto de cmpos
            For NumRegElim = 2 To CInt(RecuperaValor(lw1.Tag, 2))
                If IsNull(RS.Fields(NumRegElim - 1)) Then
                    It.SubItems(NumRegElim - 1) = " "
                Else
                    If lw1.ColumnHeaders(NumRegElim).Tag <> "" Then
                        It.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lw1.ColumnHeaders(NumRegElim).Tag)
                    Else
                        It.SubItems(NumRegElim - 1) = RS.Fields(NumRegElim - 1)
                    End If
                End If
            Next
            It.SmallIcon = ElIcono
        
'            'Para familia /dto
'            If EsDTOFam Then
'                'Si codclien es >0 then
'                If DBLet(RS!CodClien, "N") > 0 Then It.Bold = True
'            End If
        
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End If
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set RS = Nothing
    
End Sub



Private Sub AbrirFacturaLW()
Dim s As String
'    Set miRsAux = New ADODB.Recordset
    
'
'    If lw1.SelectedItem.Text = "FAM" Then
        'Van directas
        s = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & lw1.SelectedItem.SubItems(2) & "|"
'    Else
'        s = "select codtipoa,numalbar,fechaalb from scafac1 where codtipom='"
'        s = s & lw1.SelectedItem.Text & "' and numfactu=" & lw1.SelectedItem.SubItems(1)
'        s = s & " and fecfactu='" & Format(lw1.SelectedItem.SubItems(2), FormatoFecha) & "' ORDER BY codtipoa desc"
'        miRsAux.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        s = ""
'        If Not miRsAux.EOF Then
'            s = miRsAux.Fields(0) & "|" & miRsAux.Fields(1) & "|" & miRsAux.Fields(2) & "|"
'        End If
'        miRsAux.Close
'        Set miRsAux = Nothing
'    End If
    
    If s <> "" Then
        Select Case RecuperaValor(s, 1)
            Case "FPS"
                With frmPubliHcoFacSoc
                        .DesdeFichaSocio = True
                        .hcoCodMovim = RecuperaValor(s, 2)
                        .hcoCodTipoM = RecuperaValor(s, 1)
                        .hcoFechaMov = RecuperaValor(s, 3)
                        .hcoCodSocio = Adodc1.Recordset!CodClien
                        .Show vbModal
                End With
            Case "FLI", "FRL"
                With frmLiqHcoFacSoc
                        .DesdeFichaSocio = True
                        .hcoCodMovim = RecuperaValor(s, 2)
                        .hcoCodTipoM = RecuperaValor(s, 1)
                        .hcoFechaMov = RecuperaValor(s, 3)
                        .hcoCodSocio = Adodc1.Recordset!CodClien
                        .Show vbModal
                End With
            Case "FCN", "FCE"
                With frmCuotasHcoFacturas
                        .DesdeFichaSocio = True
                        .hcoCodMovim = RecuperaValor(s, 2)
                        .hcoCodTipoM = RecuperaValor(s, 1)
                        .hcoFechaMov = RecuperaValor(s, 3)
                        .Show vbModal
                End With
            Case "FAV"
                With frmFacHcoFacturas2
                        .DesdeFichaCliente = True
                        .hcoCodMovim = RecuperaValor(s, 2)
                        .hcoCodTipoM = RecuperaValor(s, 1)
                        .hcoFechaMov = RecuperaValor(s, 3)
                        .Show vbModal
                End With
                
            
        End Select
            
    
    End If
End Sub




'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'
'  CRM
'
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Tag = "" Then Exit Sub
    LabelCRM.Caption = ""
    'Levantamos todos los botones y dejamos pulsado el de ahora
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar3.Buttons(NumRegElim).Tag <> "" Then
            If Toolbar3.Buttons(NumRegElim).Index <> Button.Index Then Toolbar3.Buttons(NumRegElim).Value = tbrUnpressed
        End If
    Next NumRegElim
    CargaColumnasCRM CByte(Button.Tag)
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLWCRM
End Sub





Private Sub CargaColumnasCRM(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader
Dim Ordena As Integer
    'Las llamadas cogera las llamadas recibidas desde sllama y las efectuadas desde acciones comerciales con tipoaccion=1
    'para poder ordenarlas tendremos una columna viiblefalse con yyymmddhhmmss
    Ordena = -1
    Select Case OpcionList
    Case 0
        'Acciones comerciales
        LabelCRM.Caption = "Equipamiento"
        
        Columnas = "Tipo de Artículo|Código|Artículo|Nro.Serie|"   'nro serie, articulo, tipo de articulo
        Ancho = "2100|2000|3000|2200|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|"
        'Formatos
        Formato = "|||"
        Ncol = 4
               
    
    End Select
    
    cmdAccCRM(2).visible = OpcionList = 4
    lwCRM.ColumnHeaders.Clear
    
    'Guardo la opcion en el tag
    lwCRM.Tag = OpcionList & "|" & Ncol & "|"
    
    For NumRegElim = 1 To Ncol
         Set C = lwCRM.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
    
    If Ordena < 0 Then
        lwCRM.Sorted = False
    Else
        lwCRM.Sorted = True
        lwCRM.SortKey = 4
        lwCRM.SortOrder = lvwDescending
    End If
    
End Sub



Private Sub CargaDatosLWCRM()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador(0).Caption
    lblIndicador(0).Caption = "Leyendo " & LabelCRM.Caption
    lblIndicador(0).Refresh
    CargaDatosLWcrm2
    Me.lblIndicador(0).Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLWcrm2()
Dim cad As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Kopc As Byte
Dim MeteIT As Boolean
    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    For NumRegElim = 1 To Toolbar3.Buttons.Count
        If Toolbar3.Buttons(NumRegElim).Value = tbrPressed Then
            ElIcono = Toolbar3.Buttons(NumRegElim).Image
            Exit For
        End If
    Next
    
    
    'Fecha incio busquedas
    Text1(26).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")

    'EL where del codclien     lo lleva cada sql
    Kopc = CByte(RecuperaValor(lwCRM.Tag, 1))
    
    Select Case Kopc
    Case 0
        'Nros de serie
        cad = "select nomtipar, sserie.codartic, nomartic, numserie from sartic, sserie, stipar WHERE sserie.codartic= sartic.codartic "
        cad = cad & " and sserie.codtipar = stipar.codtipar  "
        cad = cad & " and sserie.codclien=" & Adodc1.Recordset!CodClien
        GroupBy = ""
        BuscaChekc = "nomtipar, sserie.codartic, numserie "
    End Select
    
    'El group by
    If GroupBy <> "" Then cad = cad & " GROUP BY " & GroupBy
    
    'El ORDER BY
    cad = cad & " ORDER BY " & BuscaChekc
'     If Kopc <> 4 Then cad = cad & " DESC"

    BuscaChekc = ""
    
    lwCRM.ListItems.Clear
   
    Set RS = New ADODB.Recordset
    If Kopc <> 3 Then
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        'Va contra la contabilidad
        RS.Open cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    While Not RS.EOF
        If Kopc <> 3 Then
            MeteIT = True
        Else
            If RS!tot <> 0 Then
                MeteIT = True
            Else
                MeteIT = False
            End If
        End If
        
        If MeteIT Then
                Set It = lwCRM.ListItems.Add()
                 
                If lwCRM.ColumnHeaders(1).Tag <> "" Then
                    It.Text = Format(RS.Fields(0), lwCRM.ColumnHeaders(1).Tag)
                Else
                    It.Text = RS.Fields(0)
                End If
                'El resto de cmpos
                For NumRegElim = 2 To CInt(RecuperaValor(lwCRM.Tag, 2))
                    If IsNull(RS.Fields(NumRegElim - 1)) Then
                        It.SubItems(NumRegElim - 1) = " "
                    Else
                    
                        If lwCRM.ColumnHeaders(NumRegElim).Tag <> "" Then
                            It.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lwCRM.ColumnHeaders(NumRegElim).Tag)
                        Else
                        
                            
                            cad = RS.Fields(NumRegElim - 1)
                            'no TIENE FORMATO. sEGUN LO QUE SEA puedo hacer unas cosas u otras
                            If NumRegElim = 4 And Kopc = 1 Then cad = Replace(cad, vbCrLf, " ")
                            'Medio
                            If NumRegElim = 3 And Kopc = 0 Then 'DevuelveMedio cad
                            End If
                            If NumRegElim = 3 And Kopc = 4 Then cad = Replace(cad, vbCrLf, " ")
                            
                            It.SubItems(NumRegElim - 1) = cad
                        
                            
                            
                        End If
                    End If
                Next
                'El icono
                If Kopc = 1 Then
                    It.SmallIcon = 27
                ElseIf Kopc = 2 Then

                    If RS.Fields(1) = "Enviado" Then
                        It.SmallIcon = 28
                    Else
                        It.SmallIcon = 29
                    End If
                Else
                    'el resto ponemos el del toolbar
                    It.SmallIcon = ElIcono
                End If
        End If
        
        
    
        RS.MoveNext
    Wend
    RS.Close
    
    Set RS = Nothing
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set RS = Nothing
End Sub

Private Sub DevuelveMedio(ByRef cad As String)
    'pendiente,en curso finalizada
    If cad = "0" Then
        cad = "Pendiente"
    ElseIf cad = "1" Then
        cad = "En curso"
    Else
        cad = "Finalizada"
    End If
End Sub


Private Sub CalcularTotales(CADENA As String)
Dim Importe  As Currency
Dim Compleme As Currency
Dim Penaliza As Currency

Dim RS As ADODB.Recordset
Dim SQL As String

    On Error Resume Next
    
    SQL = CADENA ' "select sum(preciove) importe  from (" & CADENA & ") aaaaa"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Importe = 0
    Text3.Text = ""
    
    If TotalRegistrosConsulta(CADENA) = 0 Then Exit Sub
    
    If Not RS.EOF Then
        If RS.Fields(0).Value <> 0 Then Importe = DBLet(RS.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
    
        Text3.Text = Format(Importe, "###,###,##0.0000")
    End If
    RS.Close
    Set RS = Nothing

    
    DoEvents
    
End Sub

