VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfParamAplic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Par�metros de la Aplicaci�n"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   Icon            =   "frmConfParamAplic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   283
      Top             =   30
      Width           =   945
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   284
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   150
         TabIndex        =   285
         Top             =   150
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
         EndProperty
      End
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
      Left            =   10350
      TabIndex        =   87
      Top             =   7995
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   150
      TabIndex        =   89
      Top             =   7920
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Left            =   120
         TabIndex        =   90
         Top             =   180
         Width           =   2280
      End
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
      Left            =   9150
      TabIndex        =   86
      Top             =   7995
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Left            =   10350
      TabIndex        =   88
      Top             =   8010
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3600
      Top             =   8010
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Height          =   7095
      Left            =   150
      TabIndex        =   91
      Top             =   810
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   9907723
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Varios"
      TabPicture(0)   =   "frmConfParamAplic.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(6)"
      Tab(0).Control(1)=   "Label1(5)"
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(5)=   "Label1(14)"
      Tab(0).Control(6)=   "imgBuscar(1)"
      Tab(0).Control(7)=   "Label1(59)"
      Tab(0).Control(8)=   "Label1(76)"
      Tab(0).Control(9)=   "Label1(77)"
      Tab(0).Control(10)=   "FrameOpciones"
      Tab(0).Control(11)=   "Frame16"
      Tab(0).Control(12)=   "Text1(0)"
      Tab(0).Control(13)=   "FrameDiasMante"
      Tab(0).Control(14)=   "Text1(5)"
      Tab(0).Control(15)=   "Text1(4)"
      Tab(0).Control(16)=   "FramePrecioKm"
      Tab(0).Control(17)=   "Text1(1)"
      Tab(0).Control(18)=   "cboTipodtos"
      Tab(0).Control(19)=   "cboOrdenDtos"
      Tab(0).Control(20)=   "Text2(1)"
      Tab(0).Control(21)=   "cboCreaTarifa"
      Tab(0).Control(22)=   "Frame13"
      Tab(0).Control(23)=   "chkHayrepar"
      Tab(0).Control(24)=   "Text1(78)"
      Tab(0).Control(25)=   "chkHaynserie"
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Datos Facturaci�n"
      TabPicture(1)   =   "frmConfParamAplic.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame15"
      Tab(1).Control(1)=   "Frame10"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(4)=   "Frame9"
      Tab(1).Control(5)=   "Frame4"
      Tab(1).Control(6)=   "Frame12"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Internet"
      TabPicture(2)   =   "frmConfParamAplic.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1(81)"
      Tab(2).Control(1)=   "FrameSoporte"
      Tab(2).Control(2)=   "FrameEMail"
      Tab(2).Control(3)=   "Label1(80)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Datos Contabilidad "
      TabPicture(3)   =   "frmConfParamAplic.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CboModAnalitica"
      Tab(3).Control(1)=   "Text1(50)"
      Tab(3).Control(2)=   "Text1(52)"
      Tab(3).Control(3)=   "Text2(52)"
      Tab(3).Control(4)=   "cboObsFactura"
      Tab(3).Control(5)=   "Text2(48)"
      Tab(3).Control(6)=   "Text1(48)"
      Tab(3).Control(7)=   "Frame8"
      Tab(3).Control(8)=   "Text1(23)"
      Tab(3).Control(9)=   "Text1(22)"
      Tab(3).Control(10)=   "Text1(21)"
      Tab(3).Control(11)=   "Text1(20)"
      Tab(3).Control(12)=   "Label1(58)"
      Tab(3).Control(13)=   "Label1(51)"
      Tab(3).Control(14)=   "Label1(47)"
      Tab(3).Control(15)=   "imgBuscar(45)"
      Tab(3).Control(16)=   "Label1(53)"
      Tab(3).Control(17)=   "imgBuscar(41)"
      Tab(3).Control(18)=   "Label1(50)"
      Tab(3).Control(19)=   "Label1(49)"
      Tab(3).Control(20)=   "Label1(19)"
      Tab(3).Control(21)=   "Label1(18)"
      Tab(3).Control(22)=   "Label1(17)"
      Tab(3).Control(23)=   "Label1(15)"
      Tab(3).ControlCount=   24
      TabCaption(4)   =   "Publicidad / Cuotas"
      TabPicture(4)   =   "frmConfParamAplic.frx":007C
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "imgBuscar(30)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "imgBuscar(25)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "imgBuscar(27)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "imgBuscar(26)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label1(30)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label1(27)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label1(26)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label1(25)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "imgBuscar(31)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "imgBuscar(29)"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Label1(29)"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Label1(31)"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Text1(68)"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "Text1(64)"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "Text1(63)"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "Text1(62)"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "Frame7"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "Frame5"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "Text1(31)"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "Text2(31)"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "Text1(30)"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "Text2(30)"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "Text1(29)"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "Text2(29)"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "Text1(27)"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "Text2(27)"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "Text1(26)"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "Text2(26)"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "Text1(25)"
      Tab(4).Control(28).Enabled=   0   'False
      Tab(4).Control(29)=   "Text2(25)"
      Tab(4).Control(29).Enabled=   0   'False
      Tab(4).Control(30)=   "Frame11"
      Tab(4).Control(30).Enabled=   0   'False
      Tab(4).Control(31)=   "Frame6"
      Tab(4).Control(31).Enabled=   0   'False
      Tab(4).ControlCount=   32
      TabCaption(5)   =   "Varios"
      TabPicture(5)   =   "frmConfParamAplic.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ChkMarcarValidados"
      Tab(5).Control(1)=   "Frame17"
      Tab(5).Control(2)=   "Text1(82)"
      Tab(5).Control(3)=   "Frame14"
      Tab(5).Control(4)=   "Label1(81)"
      Tab(5).ControlCount=   5
      Begin VB.CheckBox ChkMarcarValidados 
         Caption         =   "Marcar en traspaso como validados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74685
         TabIndex        =   296
         Tag             =   "Marcar validados|N|S|||spara1|marcarvalidados|||"
         Top             =   6615
         Width           =   5010
      End
      Begin VB.Frame Frame17 
         Caption         =   "Datos de Intercambio entre empresas"
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
         Height          =   2175
         Left            =   -74820
         TabIndex        =   290
         Top             =   4365
         Width           =   10725
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
            Index           =   91
            Left            =   3870
            MaxLength       =   6
            TabIndex        =   253
            Tag             =   "Codigo empresa Taxitronic|N|S|||spara1|empresataxitronic|||"
            Text            =   "3"
            Top             =   1440
            Width           =   1260
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
            Index           =   90
            Left            =   3870
            MaxLength       =   6
            TabIndex        =   252
            Tag             =   "Cliente otra empresa|N|S|||spara1|codclicooperativa|||"
            Text            =   "3"
            Top             =   945
            Width           =   1260
         End
         Begin VB.TextBox Text2 
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
            Index           =   90
            Left            =   5175
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   294
            Text            =   "Text2"
            Top             =   945
            Width           =   5025
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
            Index           =   89
            Left            =   3855
            MaxLength       =   6
            TabIndex        =   251
            Tag             =   "Socio otra empresa|N|S|||spara1|codsoccooperativa|||"
            Text            =   "3"
            Top             =   495
            Width           =   1260
         End
         Begin VB.TextBox Text2 
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
            Index           =   89
            Left            =   5175
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   291
            Text            =   "Text2"
            Top             =   495
            Width           =   5025
         End
         Begin VB.Label Label2 
            Caption         =   "C�digo empresa Taxitronic"
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
            Left            =   180
            TabIndex        =   295
            Top             =   1485
            Width           =   3435
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   16
            Left            =   3555
            Tag             =   "-1"
            ToolTipText     =   "Buscar cliente"
            Top             =   975
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente a Facturar Servicios"
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
            Index           =   17
            Left            =   180
            TabIndex        =   293
            Top             =   990
            Width           =   3435
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   14
            Left            =   3555
            Tag             =   "-1"
            ToolTipText     =   "Buscar socio"
            Top             =   525
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Socio Global otra Empresa"
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
            Left            =   180
            TabIndex        =   292
            Top             =   495
            Width           =   2625
         End
      End
      Begin VB.CheckBox chkHaynserie 
         Caption         =   "Hay N� Serie en Compras"
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
         Left            =   -69780
         TabIndex        =   282
         Tag             =   "Hay N� Serie en Compras|N|N|||spara1|haynserie|||"
         Top             =   2160
         Width           =   4785
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72000
         MaxLength       =   100
         TabIndex        =   250
         Tag             =   "Impresora Tarjetas|T|S|||spara1|impretarjeta|||"
         Top             =   3870
         Visible         =   0   'False
         Width           =   7380
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
         Index           =   81
         Left            =   -72870
         MaxLength       =   255
         TabIndex        =   52
         Tag             =   "Path FacturaE|T|S|||spara1|pathfacturae|||"
         Top             =   5700
         Width           =   6030
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
         Index           =   78
         Left            =   -64890
         MaxLength       =   3
         TabIndex        =   12
         Tag             =   "Cooperativa|N|N|||spara1|cooperativa|000||"
         Text            =   "Text1"
         Top             =   1890
         Width           =   615
      End
      Begin VB.Frame Frame15 
         Caption         =   "Facturaci�n Equipamiento"
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
         Height          =   825
         Left            =   -74865
         TabIndex        =   262
         Top             =   6045
         Width           =   10635
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
            Index           =   77
            Left            =   2550
            MaxLength       =   10
            TabIndex        =   263
            Tag             =   "Raiz Cta Socio Publi|T|S|||spara1|raiz_cta_soc_equip|||"
            Text            =   "3"
            Top             =   330
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Ra�z Cuenta Socios"
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
            Left            =   150
            TabIndex        =   264
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.CheckBox chkHayrepar 
         Caption         =   "Realiza Reparaciones"
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
         Left            =   -69780
         TabIndex        =   11
         Tag             =   "Reparaciones|N|N|||spara1|hayrepar|||"
         Top             =   1770
         Width           =   2475
      End
      Begin VB.Frame Frame14 
         Caption         =   "Alta Socios"
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
         Height          =   3255
         Left            =   -74850
         TabIndex        =   260
         Top             =   555
         Width           =   10695
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
            Index           =   95
            Left            =   3915
            MaxLength       =   255
            TabIndex        =   249
            Tag             =   "Secretario Aportaci�n|T|S|||spara1|aporsecretario|||"
            Text            =   "3"
            Top             =   2745
            Width           =   6330
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
            Index           =   94
            Left            =   3915
            MaxLength       =   255
            TabIndex        =   248
            Tag             =   "Presidente Aportaci�n|T|S|||spara1|aporpresidente|||"
            Text            =   "3"
            Top             =   2340
            Width           =   6330
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
            Index           =   93
            Left            =   8955
            MaxLength       =   10
            TabIndex        =   247
            Tag             =   "Nro Titulo Aportacion|N|S|||spara1|apornumero|||"
            Text            =   "3"
            Top             =   1935
            Width           =   1290
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
            Index           =   92
            Left            =   9495
            MaxLength       =   10
            TabIndex        =   246
            Tag             =   "Serie de aportaciones|T|S|||spara1|aporserie|||"
            Text            =   "3"
            Top             =   1530
            Width           =   750
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
            Index           =   76
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   245
            Tag             =   "Raiz Cta Socio Altasi|T|S|||spara1|raiz_ctaaltasoc|||"
            Text            =   "3"
            Top             =   1935
            Width           =   1290
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
            Index           =   75
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   244
            Tag             =   "Importe Gastos Alta|N|S|||spara1|impgastoalta|###,###,##0.00||"
            Text            =   "3"
            Top             =   1530
            Width           =   1260
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
            Index           =   74
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   243
            Tag             =   "Importe Titulo Alta|N|S|||spara1|imptituloalta|###,###,##0.00||"
            Text            =   "3"
            Top             =   1140
            Width           =   1260
         End
         Begin VB.TextBox Text2 
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
            Left            =   5220
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   256
            Text            =   "Text2"
            Top             =   750
            Width           =   5025
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
            Index           =   73
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   242
            Tag             =   "Cta Gastos|N|S|||spara1|ctagastoalta|||"
            Text            =   "3"
            Top             =   750
            Width           =   1260
         End
         Begin VB.TextBox Text2 
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
            Index           =   72
            Left            =   5220
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   254
            Text            =   "Text2"
            Top             =   360
            Width           =   5025
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
            Index           =   72
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   241
            Tag             =   "Cta T�tulo|N|S|||spara1|ctatituloalta|||"
            Text            =   "3"
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "Secretario T�tulo Aportaciones"
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
            Index           =   88
            Left            =   180
            TabIndex        =   300
            Top             =   2790
            Width           =   3705
         End
         Begin VB.Label Label1 
            Caption         =   "Presidente T�tulo Aportaciones"
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
            Index           =   87
            Left            =   180
            TabIndex        =   299
            Top             =   2385
            Width           =   3705
         End
         Begin VB.Label Label1 
            Caption         =   "N�mero T�tulo"
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
            Index           =   86
            Left            =   7470
            TabIndex        =   298
            Top             =   1935
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Serie T�tulo"
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
            Index           =   85
            Left            =   7470
            TabIndex        =   297
            Top             =   1575
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz Cuenta Socios"
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
            Left            =   180
            TabIndex        =   261
            Top             =   1980
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Importe Reserva Legal Obligatoria"
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
            Index           =   66
            Left            =   180
            TabIndex        =   259
            Top             =   1590
            Width           =   3945
         End
         Begin VB.Label Label1 
            Caption         =   "Importe Aportaci�n Capital Social"
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
            Index           =   65
            Left            =   180
            TabIndex        =   258
            Top             =   1180
            Width           =   4035
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta Reserva Legal Obligatoria"
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
            Left            =   180
            TabIndex        =   257
            Top             =   770
            Width           =   3375
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   3600
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   780
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta Aportaci�n Capital Social"
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
            Left            =   180
            TabIndex        =   255
            Top             =   360
            Width           =   2625
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   3600
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   390
            Width           =   240
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Liquidaci�n Socios"
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
         Height          =   2760
         Left            =   -74850
         TabIndex        =   215
         Top             =   3210
         Width           =   10665
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
            Index           =   96
            Left            =   8670
            MaxLength       =   5
            TabIndex        =   35
            Tag             =   "%Dto Centralizacion|N|S|||spara1|pordtocentra|##0.00||"
            Text            =   "3"
            Top             =   1080
            Width           =   1350
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
            Index           =   88
            Left            =   2535
            MaxLength       =   10
            TabIndex        =   38
            Tag             =   "Cta liquidacion Suplidos|N|S|||spara1|ctaliqsuplidos|||"
            Text            =   "3"
            Top             =   2250
            Width           =   1350
         End
         Begin VB.TextBox Text2 
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
            Index           =   88
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   288
            Text            =   "Text2"
            Top             =   2250
            Width           =   6120
         End
         Begin VB.TextBox Text2 
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
            Index           =   71
            Left            =   3945
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   239
            Text            =   "Text2"
            Top             =   1875
            Width           =   6120
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
            Index           =   71
            Left            =   2550
            MaxLength       =   10
            TabIndex        =   37
            Tag             =   "Cta liquidacion|N|S|||spara1|ctaliquidacion|||"
            Text            =   "3"
            Top             =   1875
            Width           =   1350
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
            Index           =   65
            Left            =   2550
            MaxLength       =   10
            TabIndex        =   32
            Tag             =   "Raiz Cta Socio Publi|T|S|||spara1|raiz_cta_soc_liqui|||"
            Text            =   "3"
            Top             =   270
            Width           =   1350
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
            Index           =   49
            Left            =   2550
            MaxLength       =   5
            TabIndex        =   34
            Tag             =   "N� Contabilidad|N|S|||spara1|porreten|||"
            Text            =   "3"
            Top             =   1110
            Width           =   1350
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
            Index           =   47
            Left            =   2550
            MaxLength       =   10
            TabIndex        =   33
            Tag             =   "Ra�z Cta retencion|N|S|||spara1|raiz_ctareten|||"
            Text            =   "3"
            Top             =   690
            Width           =   1350
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
            Index           =   46
            Left            =   2550
            MaxLength       =   2
            TabIndex        =   36
            Tag             =   "REA|N|S|0||spara1|iva_rea|||"
            Text            =   "Text1"
            Top             =   1500
            Width           =   1350
         End
         Begin VB.TextBox Text2 
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
            Index           =   46
            Left            =   3945
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   216
            Text            =   "Text2"
            Top             =   1500
            Width           =   6120
         End
         Begin VB.Label Label2 
            Caption         =   "% Dto Centralizacion"
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
            Left            =   6255
            TabIndex        =   301
            Top             =   1110
            Width           =   2655
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   13
            Left            =   2280
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   2250
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta Base Suplidos"
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
            Left            =   135
            TabIndex        =   289
            Top             =   2250
            Width           =   2130
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta Base"
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
            Left            =   135
            TabIndex        =   240
            Top             =   1875
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   2295
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   1875
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Ra�z Cuenta Socios"
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
            Left            =   135
            TabIndex        =   220
            Top             =   300
            Width           =   2145
         End
         Begin VB.Label Label2 
            Caption         =   "Porcentaje Retenci�n"
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
            Left            =   135
            TabIndex        =   219
            Top             =   1140
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Ra�z Cuenta Retenci�n"
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
            Index           =   48
            Left            =   135
            TabIndex        =   218
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "IVA Liquidaci�n"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   52
            Left            =   135
            TabIndex        =   217
            Top             =   1530
            Width           =   1605
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   39
            Left            =   2295
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1530
            Width           =   240
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Publicidad"
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
         Height          =   2475
         Left            =   120
         TabIndex        =   210
         Top             =   4350
         Width           =   10755
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
            Index           =   60
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   78
            Tag             =   "Raiz Cta Socio Publi|T|S|||spara1|raiz_cta_soc_publi|||"
            Text            =   "3"
            Top             =   1290
            Width           =   1260
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
            Index           =   59
            Left            =   2400
            TabIndex        =   77
            Tag             =   "Concepto Facturaci�n Publicidad |T|S|||spara1|confactupubli|||"
            Text            =   "Text1 "
            Top             =   870
            Width           =   6405
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
            Index           =   32
            Left            =   2400
            MaxLength       =   16
            TabIndex        =   76
            Tag             =   "Recar |T|S|||spara1|codartictel|||"
            Text            =   "Text1 "
            Top             =   450
            Width           =   1605
         End
         Begin VB.TextBox Text2 
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
            Index           =   32
            Left            =   4020
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   211
            Text            =   "Text2"
            Top             =   450
            Width           =   4845
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz Cuenta Socios"
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
            Left            =   240
            TabIndex        =   214
            Top             =   1320
            Width           =   2115
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   2130
            Tag             =   "-1"
            ToolTipText     =   "Ver observaciones"
            Top             =   900
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto Factura"
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
            Left            =   240
            TabIndex        =   213
            Top             =   900
            Width           =   2100
         End
         Begin VB.Label Label1 
            Caption         =   "Art�culo a facturar"
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
            Index           =   32
            Left            =   240
            TabIndex        =   212
            Top             =   510
            Width           =   1830
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   32
            Left            =   2130
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   480
            Width           =   240
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Concepto Facturaci�n Publicidad"
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
         Height          =   855
         Left            =   5130
         TabIndex        =   208
         Top             =   5880
         Width           =   3375
         Begin VB.CheckBox chkTicketsAgrupads 
            Caption         =   "Contabilizar ticket TPV agrupados"
            Height          =   375
            Left            =   360
            TabIndex        =   209
            Tag             =   "Tickets agrupadsos|N|N|||spara1|conttickagrupado|||"
            Top             =   360
            Visible         =   0   'False
            Width           =   2775
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   25
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   201
         Text            =   "Text2"
         Top             =   4710
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   25
         Left            =   180
         MaxLength       =   3
         TabIndex        =   200
         Tag             =   "Actividad|N|S|0||spara1|defactividad|000||"
         Text            =   "Tex"
         Top             =   4710
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   26
         Left            =   5340
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   199
         Text            =   "Text2"
         Top             =   4710
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   26
         Left            =   4500
         MaxLength       =   3
         TabIndex        =   198
         Tag             =   "Envio|N|S|0|999|spara1|defenvio|000||"
         Text            =   "Tex"
         Top             =   4710
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   27
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   197
         Text            =   "Text2"
         Top             =   5310
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   27
         Left            =   180
         TabIndex        =   196
         Tag             =   "Zona|N|S|0|999|spara1|defzona|000||"
         Text            =   "Tex"
         Top             =   5310
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   29
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   195
         Text            =   "Text2"
         Top             =   6030
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   29
         Left            =   180
         MaxLength       =   3
         TabIndex        =   194
         Tag             =   "Situacion|N|S|0|999|spara1|defstituacion|000||"
         Text            =   "Tex"
         Top             =   6030
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   30
         Left            =   5340
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   193
         Text            =   "Text2"
         Top             =   6030
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   30
         Left            =   4500
         MaxLength       =   3
         TabIndex        =   192
         Tag             =   "Tarifa|N|S|0|999|spara1|deftarifa|||"
         Text            =   "Tex"
         Top             =   6030
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   31
         Left            =   5340
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   191
         Text            =   "Text2"
         Top             =   5310
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   31
         Left            =   4500
         MaxLength       =   3
         TabIndex        =   190
         Tag             =   "Agente|N|S|0|999|spara1|defagente|000||"
         Text            =   "Tex"
         Top             =   5310
         Width           =   735
      End
      Begin VB.Frame Frame13 
         Caption         =   "Garantia de reparaci�n"
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
         Height          =   735
         Left            =   -74640
         TabIndex        =   184
         Top             =   5760
         Width           =   4605
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
            Index           =   58
            Left            =   2670
            MaxLength       =   4
            TabIndex        =   10
            Tag             =   "Dias de garantia de Reparacion|N|S|0|9999|spara1|diasgaranrepa|||"
            Text            =   "Text1"
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Dias"
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
            Index           =   62
            Left            =   150
            TabIndex        =   185
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.ComboBox cboCreaTarifa 
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
         ItemData        =   "frmConfParamAplic.frx":00B4
         Left            =   -72510
         List            =   "frmConfParamAplic.frx":00C1
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo Descuentos|N|N|||spara1|creatarifart|||"
         Top             =   2790
         Width           =   2505
      End
      Begin VB.ComboBox CboModAnalitica 
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
         Left            =   -72090
         Style           =   2  'Dropdown List
         TabIndex        =   180
         Tag             =   "Modo anal�tica|N|N|0|9|spara1|modanalitica|||"
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   50
         Left            =   -72090
         MaxLength       =   2
         TabIndex        =   171
         Tag             =   "N�Conta|N|S|1|99|spara1|conta_B|||"
         Text            =   "Text1"
         Top             =   1125
         Visible         =   0   'False
         Width           =   675
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
         Index           =   52
         Left            =   -72120
         MaxLength       =   2
         TabIndex        =   59
         Tag             =   "IVAexento|N|S|0||spara1|IvaIntracom|||"
         Text            =   "Text1"
         Top             =   2520
         Width           =   675
      End
      Begin VB.TextBox Text2 
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
         Left            =   -71400
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   169
         Text            =   "Text2"
         Top             =   2520
         Width           =   3105
      End
      Begin VB.ComboBox cboObsFactura 
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
         Left            =   -67800
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Tag             =   "Orden Descuentos|N|S|||spara1|obsfactura|||"
         Top             =   1110
         Width           =   3135
      End
      Begin VB.TextBox Text2 
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
         Left            =   -71400
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   163
         Text            =   "Text2"
         Top             =   2085
         Width           =   2985
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
         Index           =   48
         Left            =   -72120
         MaxLength       =   2
         TabIndex        =   58
         Tag             =   "IVAexento|N|S|0||spara1|ivaexento|||"
         Text            =   "Text1"
         Top             =   2085
         Width           =   675
      End
      Begin VB.Frame Frame8 
         Caption         =   "IVA 's"
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
         Height          =   3735
         Left            =   -74940
         TabIndex        =   146
         Top             =   2970
         Width           =   10905
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
            Index           =   86
            Left            =   6690
            MaxLength       =   2
            TabIndex        =   69
            Tag             =   "IVRE Ant 2|N|S|0|99|spara1|ivaant2eq|||"
            Text            =   "Text1"
            Top             =   2130
            Width           =   615
         End
         Begin VB.TextBox Text2 
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
            Index           =   86
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   280
            Text            =   "Text2"
            Top             =   2130
            Width           =   3195
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
            Index           =   85
            Left            =   6690
            MaxLength       =   2
            TabIndex        =   68
            Tag             =   "IVA Ant 2|N|S|0|99|spara1|ivaant2|||"
            Text            =   "Text1"
            Top             =   1710
            Width           =   615
         End
         Begin VB.TextBox Text2 
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
            Index           =   85
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   277
            Text            =   "Text2"
            Top             =   1710
            Width           =   3195
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
            Index           =   84
            Left            =   6690
            MaxLength       =   2
            TabIndex        =   67
            Tag             =   "IVRE Ant 1|N|S|0|99|spara1|ivaant1eq|||"
            Text            =   "Text1"
            Top             =   1050
            Width           =   615
         End
         Begin VB.TextBox Text2 
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
            Index           =   84
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   275
            Text            =   "Text2"
            Top             =   1050
            Width           =   3195
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
            Index           =   83
            Left            =   6690
            MaxLength       =   2
            TabIndex        =   66
            Tag             =   "IVA Ant 1|N|S|0|99|spara1|ivaant1|||"
            Text            =   "Text1"
            Top             =   630
            Width           =   615
         End
         Begin VB.TextBox Text2 
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
            Index           =   83
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   272
            Text            =   "Text2"
            Top             =   630
            Width           =   3195
         End
         Begin VB.TextBox Text2 
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
            Index           =   42
            Left            =   2070
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   155
            Text            =   "Text2"
            Top             =   3210
            Width           =   2925
         End
         Begin VB.TextBox Text2 
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
            Index           =   45
            Left            =   2070
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   154
            Text            =   "Text2"
            Top             =   2790
            Width           =   2925
         End
         Begin VB.TextBox Text2 
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
            Index           =   41
            Left            =   2070
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   153
            Text            =   "Text2"
            Top             =   2100
            Width           =   2925
         End
         Begin VB.TextBox Text2 
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
            Index           =   44
            Left            =   2070
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   152
            Text            =   "Text2"
            Top             =   1680
            Width           =   2925
         End
         Begin VB.TextBox Text2 
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
            Index           =   40
            Left            =   2070
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   151
            Text            =   "Text2"
            Top             =   1050
            Width           =   2925
         End
         Begin VB.TextBox Text2 
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
            Index           =   43
            Left            =   2070
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   150
            Text            =   "Text2"
            Top             =   630
            Width           =   2925
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
            Index           =   42
            Left            =   1410
            MaxLength       =   2
            TabIndex        =   65
            Tag             =   "IVRE3|N|S|0|99|spara1|ivare3eq|||"
            Text            =   "Text1"
            Top             =   3210
            Width           =   615
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
            Index           =   41
            Left            =   1410
            MaxLength       =   2
            TabIndex        =   63
            Tag             =   "IVRE2|N|S|0|99|spara1|ivare2eq|||"
            Text            =   "Text1"
            Top             =   2100
            Width           =   615
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
            Index           =   40
            Left            =   1410
            MaxLength       =   2
            TabIndex        =   61
            Tag             =   "IVRE1|N|S|0|99|spara1|ivare1eq|||"
            Text            =   "Text1"
            Top             =   1050
            Width           =   615
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
            Index           =   43
            Left            =   1410
            MaxLength       =   2
            TabIndex        =   60
            Tag             =   "IVA1|N|S|0|99|spara1|ivare1|||"
            Text            =   "Text1"
            Top             =   630
            Width           =   615
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
            Index           =   44
            Left            =   1410
            MaxLength       =   2
            TabIndex        =   62
            Tag             =   "IVA2|N|S|0|99|spara1|ivare2|||"
            Text            =   "Text1"
            Top             =   1680
            Width           =   615
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
            Index           =   45
            Left            =   1410
            MaxLength       =   2
            TabIndex        =   64
            Tag             =   "IVA3|N|S|0|99|spara1|ivare3|||"
            Text            =   "Text1"
            Top             =   2790
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   6390
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   2160
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
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
            Left            =   5580
            TabIndex        =   281
            Top             =   2130
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   6390
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1740
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Reducido Antiguo"
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
            Height          =   195
            Index           =   83
            Left            =   5430
            TabIndex        =   279
            Top             =   1410
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
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
            Left            =   5580
            TabIndex        =   278
            Top             =   1740
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   6420
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
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
            Left            =   5610
            TabIndex        =   276
            Top             =   1050
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   6420
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   660
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "General Antiguo"
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
            Height          =   285
            Index           =   82
            Left            =   5430
            TabIndex        =   274
            Top             =   330
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
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
            Left            =   5610
            TabIndex        =   273
            Top             =   660
            Width           =   705
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
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
            Left            =   300
            TabIndex        =   161
            Top             =   3210
            Width           =   435
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
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
            Left            =   270
            TabIndex        =   160
            Top             =   2790
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
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
            Left            =   300
            TabIndex        =   159
            Top             =   2100
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
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
            Left            =   300
            TabIndex        =   158
            Top             =   1710
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
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
            Left            =   300
            TabIndex        =   157
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
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
            Left            =   270
            TabIndex        =   156
            Top             =   660
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "General"
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
            Height          =   195
            Index           =   45
            Left            =   150
            TabIndex        =   149
            Top             =   330
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Reducido"
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
            Height          =   195
            Index           =   44
            Left            =   150
            TabIndex        =   148
            Top             =   1380
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Super-Reducido"
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
            Height          =   195
            Index           =   46
            Left            =   90
            TabIndex        =   147
            Top             =   2460
            Width           =   2160
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   36
            Left            =   1140
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   660
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   33
            Left            =   1140
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   37
            Left            =   1110
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1710
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   34
            Left            =   1110
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   2130
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   38
            Left            =   1110
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   2790
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   35
            Left            =   1110
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   3240
            Width           =   240
         End
      End
      Begin VB.TextBox Text2 
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
         Left            =   -71640
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   133
         Text            =   "Text2"
         Top             =   1320
         Width           =   6135
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cuotas"
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
         Height          =   3645
         Left            =   120
         TabIndex        =   130
         Top             =   650
         Width           =   10755
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
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
            Index           =   69
            Left            =   9030
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   234
            Text            =   "Text3"
            Top             =   2280
            Width           =   1305
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
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
            Index           =   67
            Left            =   9030
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   232
            Text            =   "Text3"
            Top             =   1875
            Width           =   1305
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
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
            Index           =   70
            Left            =   9030
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   231
            Text            =   "Text3"
            Top             =   1065
            Width           =   1305
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
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
            Index           =   61
            Left            =   9030
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   230
            Text            =   "Text3"
            Top             =   1470
            Width           =   1305
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
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
            Index           =   28
            Left            =   9030
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   229
            Text            =   "Text3"
            Top             =   660
            Width           =   1305
         End
         Begin VB.TextBox Text2 
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
            Index           =   70
            Left            =   4440
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   226
            Text            =   "Text2"
            Top             =   1065
            Width           =   4545
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
            Index           =   70
            Left            =   2400
            TabIndex        =   71
            Tag             =   "Con Chofer|T|S|||spara1|artcuotaconchof|||"
            Text            =   "Tex"
            Top             =   1065
            Width           =   1995
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
            Index           =   69
            Left            =   2400
            TabIndex        =   74
            Tag             =   "Servicios|T|S|||spara1|artservcuotas|||"
            Text            =   "Tex"
            Top             =   2280
            Width           =   1995
         End
         Begin VB.TextBox Text2 
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
            Index           =   69
            Left            =   4440
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   224
            Text            =   "Text2"
            Top             =   2280
            Width           =   4545
         End
         Begin VB.TextBox Text2 
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
            Index           =   67
            Left            =   4440
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   222
            Text            =   "Text2"
            Top             =   1875
            Width           =   4545
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
            Index           =   67
            Left            =   2400
            TabIndex        =   73
            Tag             =   "Alquiler|T|S|||spara1|artalquiler|||"
            Text            =   "Tex"
            Top             =   1875
            Width           =   1995
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
            Index           =   66
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   75
            Tag             =   "Raiz Cta Socio Publi|T|S|||spara1|raiz_ctaclien_soc|||"
            Text            =   "3"
            Top             =   2850
            Width           =   1320
         End
         Begin VB.TextBox Text2 
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
            Index           =   61
            Left            =   4440
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   187
            Text            =   "Text2"
            Top             =   1470
            Width           =   4545
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
            Index           =   61
            Left            =   2400
            TabIndex        =   72
            Tag             =   "Cuota Ext|T|S|||spara1|artcuotaext|||"
            Text            =   "Tex"
            Top             =   1470
            Width           =   1995
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
            Index           =   28
            Left            =   2400
            TabIndex        =   70
            Tag             =   "Sin Chofer|T|S|||spara1|artcuotasinchof|||"
            Text            =   "Tex"
            Top             =   660
            Width           =   1995
         End
         Begin VB.TextBox Text2 
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
            Index           =   28
            Left            =   4440
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   131
            Text            =   "Text2"
            Top             =   660
            Width           =   4545
         End
         Begin VB.Label Label1 
            Caption         =   "Precio"
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
            Index           =   75
            Left            =   9030
            TabIndex        =   233
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Art�culo"
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
            Index           =   64
            Left            =   2430
            TabIndex        =   228
            Top             =   330
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Cuota Normal Con Chofer"
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
            Left            =   180
            TabIndex        =   227
            Top             =   1117
            Width           =   1875
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   2100
            ToolTipText     =   "Buscar articulo"
            Top             =   1095
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   2100
            ToolTipText     =   "Buscar articulo"
            Top             =   2310
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Servicios"
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
            Left            =   180
            TabIndex        =   225
            Top             =   2310
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Alquiler Equipos"
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
            Left            =   180
            TabIndex        =   223
            Top             =   1911
            Width           =   1515
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   2100
            ToolTipText     =   "Buscar articulo"
            Top             =   1905
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz Cuenta Socios"
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
            Left            =   180
            TabIndex        =   221
            Top             =   2850
            Width           =   1665
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   2100
            ToolTipText     =   "Buscar articulo"
            Top             =   1500
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuota Extraordinaria"
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
            Left            =   180
            TabIndex        =   186
            Top             =   1514
            Width           =   1515
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   28
            Left            =   2100
            ToolTipText     =   "Buscar articulo"
            Top             =   690
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuota Normal Sin Chofer"
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
            Index           =   28
            Left            =   180
            TabIndex        =   132
            Top             =   720
            Width           =   2025
         End
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
         Left            =   -72090
         MaxLength       =   30
         TabIndex        =   53
         Tag             =   "Servidor Contabilidad|T|S|||spara1|serconta|||"
         Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
         Top             =   555
         Width           =   1620
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
         Index           =   22
         Left            =   -65070
         MaxLength       =   2
         TabIndex        =   56
         Tag             =   "N� Contabilidad|N|S|||spara1|numconta|||"
         Text            =   "3"
         Top             =   555
         Width           =   390
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
         Left            =   -69570
         MaxLength       =   20
         TabIndex        =   54
         Tag             =   "Usuario Contabilidad|T|S|||spara1|usuconta|||"
         Text            =   "3wwwwwwwwwwwwwwwwwww"
         Top             =   555
         Width           =   1020
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
         IMEMode         =   3  'DISABLE
         Index           =   20
         Left            =   -67800
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   55
         Tag             =   "Password Contabilidad|T|S|||spara1|pasconta|||"
         Text            =   "3"
         Top             =   555
         Width           =   1140
      End
      Begin VB.ComboBox cboOrdenDtos 
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
         Left            =   -72510
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Orden Descuentos|N|N|||spara1|ordendto|||"
         Top             =   1860
         Width           =   2505
      End
      Begin VB.Frame Frame3 
         Caption         =   "Compras"
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
         Height          =   855
         Left            =   -74850
         TabIndex        =   119
         Top             =   630
         Width           =   10665
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
            Index           =   19
            Left            =   6540
            MaxLength       =   2
            TabIndex        =   28
            Tag             =   "Mes a no girar|N|S|0|12|spara1|mesnogir|||"
            Text            =   "Text1"
            Top             =   330
            Width           =   615
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
            Index           =   18
            Left            =   3990
            MaxLength       =   2
            TabIndex        =   27
            Tag             =   "Dia 3 de pago compras|N|S|0|31|spara1|diapago3|||"
            Text            =   "Text1"
            Top             =   330
            Width           =   615
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
            Index           =   17
            Left            =   3270
            MaxLength       =   2
            TabIndex        =   26
            Tag             =   "Dia 2 de pago compras|N|S|0|31|spara1|diapago2|||"
            Text            =   "Text1"
            Top             =   330
            Width           =   615
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
            Index           =   16
            Left            =   2550
            MaxLength       =   2
            TabIndex        =   25
            Tag             =   "Dia 1 de pago compras|N|S|0|31|spara1|diapago1|||"
            Text            =   "Text1"
            Top             =   330
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Mes a no girar"
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
            Left            =   4950
            TabIndex        =   121
            Top             =   360
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "D�as de pago"
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
            Left            =   180
            TabIndex        =   120
            Top             =   360
            Width           =   1755
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Facturaci�n Clientes"
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
         Height          =   1605
         Left            =   -74850
         TabIndex        =   116
         Top             =   1500
         Width           =   10665
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
            Index           =   87
            Left            =   2505
            MaxLength       =   16
            TabIndex        =   31
            Tag             =   "Art�culo Suplidos |T|S|||spara1|artsuplidos|||"
            Text            =   "Text1 "
            Top             =   1170
            Width           =   1815
         End
         Begin VB.TextBox Text2 
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
            Index           =   87
            Left            =   4365
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   286
            Text            =   "Text2"
            Top             =   1170
            Width           =   5715
         End
         Begin VB.TextBox Text2 
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
            Index           =   53
            Left            =   4380
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   188
            Text            =   "Text2"
            Top             =   750
            Width           =   5715
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
            Index           =   53
            Left            =   2520
            MaxLength       =   16
            TabIndex        =   30
            Tag             =   "Art�culo Gastos |T|S|||spara1|ArtReciclado|||"
            Text            =   "Text1 "
            Top             =   750
            Width           =   1815
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
            Index           =   15
            Left            =   2520
            MaxLength       =   16
            TabIndex        =   29
            Tag             =   "Art�culo Servicios |T|S|||spara1|codartid|||"
            Text            =   "Text1 "
            Top             =   327
            Width           =   1815
         End
         Begin VB.TextBox Text2 
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
            Left            =   4380
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   117
            Text            =   "Text2"
            Top             =   330
            Width           =   5715
         End
         Begin VB.Label Label1 
            Caption         =   "Art�culo Suplidos"
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
            Index           =   84
            Left            =   150
            TabIndex        =   287
            Top             =   1230
            Width           =   2010
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   87
            Left            =   2205
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   1215
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   53
            Left            =   2220
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   795
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Art�culo Gtos Admon"
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
            Left            =   150
            TabIndex        =   189
            Top             =   810
            Width           =   2010
         End
         Begin VB.Label Label1 
            Caption         =   "Art�culo Servicios"
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
            Left            =   150
            TabIndex        =   118
            Top             =   360
            Width           =   1905
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   2220
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame FrameSoporte 
         ForeColor       =   &H00972E0B&
         Height          =   1635
         Left            =   -74760
         TabIndex        =   111
         Top             =   3840
         Width           =   8355
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
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   49
            Tag             =   "Web de Soporte|T|S|||spara1|websoporte|||"
            Text            =   "3"
            Top             =   300
            Width           =   6060
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
            Index           =   13
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   50
            Tag             =   "Mail de Soporte|T|S|||spara1|mailsoporte|||"
            Text            =   "3"
            Top             =   720
            Width           =   6060
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
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   51
            Tag             =   "Version Web|T|S|||spara1|webversion|||"
            Text            =   "3"
            Top             =   1140
            Width           =   6060
         End
         Begin VB.Label Label1 
            Caption         =   "Web de soporte"
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
            Index           =   9
            Left            =   300
            TabIndex        =   115
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "Mail soporte"
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
            Left            =   300
            TabIndex        =   114
            Top             =   780
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Web check version"
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
            Left            =   300
            TabIndex        =   113
            Top             =   1200
            Width           =   1500
         End
         Begin VB.Label Label8 
            Caption         =   "Soporte"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   112
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.Frame FrameEMail 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   105
         Top             =   720
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
            IMEMode         =   3  'DISABLE
            Index           =   57
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   48
            Tag             =   "LanzaMailOutlook|T|S|||spara1|arigesmail|||"
            Text            =   "3"
            Top             =   2400
            Width           =   1620
         End
         Begin VB.CheckBox chkOutlook 
            Caption         =   "Enviar desde Outlook"
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
            Left            =   7740
            TabIndex        =   181
            Tag             =   "Outlook|N|N|||spara1|EnvioDesdeOutlook|||"
            Top             =   1560
            Width           =   2685
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
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   1440
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   47
            Tag             =   "Password SMTP|T|S|||spara1|smtppass|||"
            Text            =   "3"
            Top             =   1650
            Width           =   4260
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   46
            Tag             =   "Usuario SMTP|T|S|||spara1|smtpuser|||"
            Text            =   "3"
            Top             =   1245
            Width           =   4260
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
            Index           =   9
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   45
            Tag             =   "Servidor SMTP|T|S|||spara1|smtphost|||"
            Text            =   "3"
            Top             =   825
            Width           =   5700
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   44
            Tag             =   "Direccion e-mail|T|S|||spara1|diremail|||"
            Text            =   "3"
            Top             =   420
            Width           =   5700
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   8040
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label Label1 
            Caption         =   "Lanza pantalla mail outlook"
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
            Left            =   240
            TabIndex        =   183
            Top             =   2430
            Width           =   2280
         End
         Begin VB.Label Label8 
            Caption         =   "Envio E-Mail"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   110
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Left            =   300
            TabIndex        =   109
            Top             =   1710
            Width           =   1020
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
            Height          =   195
            Index           =   22
            Left            =   300
            TabIndex        =   108
            Top             =   1300
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
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
            Index           =   21
            Left            =   300
            TabIndex        =   107
            Top             =   890
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
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
            Left            =   300
            TabIndex        =   106
            Top             =   480
            Width           =   900
         End
      End
      Begin VB.ComboBox cboTipodtos 
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
         Left            =   -72510
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Tipo Descuentos|N|N|||spara1|tipodtos|||"
         Top             =   2325
         Width           =   2505
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
         Index           =   1
         Left            =   -72480
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "C�digo Tarifa PVP|N|N|||spara1|codtarif|000||"
         Text            =   "Text1"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Frame FramePrecioKm 
         Caption         =   "Precio Km"
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
         Height          =   1070
         Left            =   -74640
         TabIndex        =   96
         Top             =   3360
         Width           =   4605
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
            Index           =   2
            Left            =   2670
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Precio Km desplaz. Clientes|N|S|0|9999.0000|spara1|preukmcl|#,##0.0000||"
            Text            =   "Text1"
            Top             =   240
            Width           =   1755
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
            Left            =   2670
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Precio Km desplaz. T�cnicos|N|S|0|9999.0000|spara1|preukmtc|#,##0.0000||"
            Text            =   "Text1"
            Top             =   630
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Desplazamiento Clientes"
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
            Left            =   120
            TabIndex        =   98
            Top             =   255
            Width           =   2505
         End
         Begin VB.Label Label1 
            Caption         =   "Desplazamiento T�cnicos"
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
            Left            =   120
            TabIndex        =   97
            Top             =   660
            Width           =   2655
         End
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
         Left            =   -71640
         MaxLength       =   35
         TabIndex        =   0
         Tag             =   "Nombre Director Gerente|T|S|||spara1|nomgeren|||"
         Text            =   "Text1"
         Top             =   480
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
         Index           =   5
         Left            =   -71640
         MaxLength       =   35
         TabIndex        =   1
         Tag             =   "Nombre responsable Admon|T|S|||spara1|nomadmin|||"
         Text            =   "Text1"
         Top             =   900
         Width           =   4095
      End
      Begin VB.Frame FrameDiasMante 
         Caption         =   "D�as Reparaci�n"
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
         Height          =   1095
         Left            =   -74640
         TabIndex        =   92
         Top             =   4560
         Width           =   4605
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
            Index           =   6
            Left            =   2670
            MaxLength       =   4
            TabIndex        =   9
            Tag             =   "Dias Repar. sin Mantenimiento|N|N|0|9999|spara1|diasnoman|||"
            Text            =   "Text"
            Top             =   615
            Width           =   1755
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
            Index           =   7
            Left            =   2670
            MaxLength       =   4
            TabIndex        =   8
            Tag             =   "Dias Repar. con Mantenimiento|N|N|0|9999|spara1|diassiman|||"
            Text            =   "Text1"
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Sin Mantenimiento"
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
            Left            =   120
            TabIndex        =   94
            Top             =   675
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Con Mantenimiento"
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
            Left            =   120
            TabIndex        =   93
            Top             =   300
            Width           =   2085
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   -70740
         MaxLength       =   15
         TabIndex        =   103
         Tag             =   "C�digo Par�metros Aplic|N|N|||spara1|codigo||S|"
         Text            =   "Text1"
         Top             =   540
         Width           =   645
      End
      Begin VB.Frame Frame7 
         Caption         =   "Avisos"
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
         Height          =   2415
         Left            =   120
         TabIndex        =   134
         Top             =   1080
         Width           =   8535
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   79
            Tag             =   "ped. cli|N|S|0||spara1|avipedcli|||"
            Text            =   "3"
            Top             =   315
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   6720
            MaxLength       =   2
            TabIndex        =   80
            Tag             =   "ped.pro.|N|S|0||spara1|avipedpro|||"
            Text            =   "3"
            Top             =   315
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   81
            Tag             =   "alb.cli.|N|S|0||spara1|avialbcli|||"
            Text            =   "3"
            Top             =   720
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   36
            Left            =   6720
            MaxLength       =   2
            TabIndex        =   82
            Tag             =   "alb.pro.|N|S|0||spara1|avialbpro|||"
            Text            =   "3"
            Top             =   720
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   83
            Tag             =   "avi.mante|N|S|0||spara1|avimanteni|||"
            Text            =   "3"
            Top             =   1275
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   85
            Tag             =   "avi.avisos|N|S|0||spara1|aviavios|||"
            Text            =   "3"
            Top             =   1995
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   84
            Tag             =   "avi.repa.|N|S|0||spara1|avirepara|||"
            Text            =   "3"
            Top             =   1635
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pedidos clientes"
            Height          =   195
            Index           =   33
            Left            =   2040
            TabIndex        =   145
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Pedidos proveedores"
            Height          =   195
            Index           =   34
            Left            =   4680
            TabIndex        =   144
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Albaranes clientes"
            Height          =   195
            Index           =   35
            Left            =   2040
            TabIndex        =   143
            Top             =   765
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Albaranes proveedores"
            Height          =   195
            Index           =   36
            Left            =   4680
            TabIndex        =   142
            Top             =   765
            Width           =   1635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Mantenimientos"
            Height          =   195
            Index           =   37
            Left            =   2040
            TabIndex        =   141
            Top             =   1320
            Width           =   1110
         End
         Begin VB.Label Label1 
            Caption         =   "Reparaciones"
            Height          =   195
            Index           =   38
            Left            =   2040
            TabIndex        =   140
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Avisos "
            Height          =   195
            Index           =   39
            Left            =   2040
            TabIndex        =   139
            Top             =   2040
            Width           =   510
         End
         Begin VB.Label Label1 
            Caption         =   "Dias desde la fecha"
            Height          =   195
            Index           =   40
            Left            =   120
            TabIndex        =   138
            Top             =   360
            Width           =   7275
         End
         Begin VB.Label Label1 
            Caption         =   "No facturados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   41
            Left            =   4680
            TabIndex        =   137
            Top             =   1320
            Width           =   2955
         End
         Begin VB.Label Label1 
            Caption         =   "Pendientes de reparar sin motivo de reparaci�n"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   42
            Left            =   4680
            TabIndex        =   136
            Top             =   1680
            Width           =   3555
         End
         Begin VB.Label Label1 
            Caption         =   "Abiertos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   43
            Left            =   4680
            TabIndex        =   135
            Top             =   2040
            Width           =   2955
         End
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   62
         Left            =   1110
         TabIndex        =   235
         Text            =   "Tex"
         Top             =   1050
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   63
         Left            =   1170
         TabIndex        =   236
         Text            =   "Tex"
         Top             =   1020
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   64
         Left            =   1500
         TabIndex        =   237
         Text            =   "Tex"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   68
         Left            =   1230
         TabIndex        =   238
         Text            =   "Tex"
         Top             =   810
         Width           =   1335
      End
      Begin VB.Frame Frame9 
         Caption         =   "Aportaci�n en facturas"
         Enabled         =   0   'False
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
         Height          =   735
         Left            =   -74880
         TabIndex        =   165
         Top             =   6000
         Visible         =   0   'False
         Width           =   8655
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
            Index           =   51
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   43
            Tag             =   "Cta aportacion|N|S|||spara1|ctaaportacion|||"
            Text            =   "3"
            Top             =   240
            Width           =   1260
         End
         Begin VB.TextBox Text2 
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
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   166
            Text            =   "Text2"
            Top             =   240
            Width           =   4185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   42
            Left            =   2220
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta"
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
            Left            =   270
            TabIndex        =   167
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cheques  regalo"
         Enabled         =   0   'False
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
         Height          =   735
         Left            =   -74880
         TabIndex        =   127
         Top             =   6000
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   24
            Left            =   3960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   128
            Text            =   "Text2"
            Top             =   240
            Width           =   4550
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   24
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   42
            Tag             =   "Forma de pago para cheque regalo |N|S|0|999|spara1|codforpa|000||"
            Text            =   "Tex"
            Top             =   237
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   24
            Left            =   2640
            Tag             =   "-1"
            ToolTipText     =   "Buscar forma pago"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma de pago "
            Height          =   255
            Index           =   24
            Left            =   1320
            TabIndex        =   129
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "C�lculo Importes Llamada"
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
         Height          =   2025
         Left            =   -69810
         TabIndex        =   267
         Top             =   2550
         Width           =   5835
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
            Height          =   285
            Index           =   80
            Left            =   3540
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Precio por tpo de espera|N|S|||spara1|precioxtpoespera|###,##0.0000||"
            Text            =   "3"
            Top             =   840
            Width           =   1260
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
            Height          =   285
            Index           =   79
            Left            =   3540
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Precio por distancia|N|S|||spara1|precioxdistancia|###,##0.0000||"
            Text            =   "3"
            Top             =   390
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "Precio por tiempo de espera"
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
            Index           =   79
            Left            =   240
            TabIndex        =   269
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "Precio por distancia"
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
            Index           =   78
            Left            =   240
            TabIndex        =   268
            Top             =   420
            Width           =   2175
         End
      End
      Begin VB.Frame FrameOpciones 
         Caption         =   "Opciones"
         Enabled         =   0   'False
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
         Height          =   4365
         Left            =   -69810
         TabIndex        =   95
         Top             =   2520
         Visible         =   0   'False
         Width           =   5835
         Begin VB.CheckBox ChkDtoxCantidad 
            Caption         =   "Hay Dtos por cantidad"
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
            Left            =   3000
            TabIndex        =   178
            Tag             =   "Hay Dtos por cantidad|N|N|||spara1|dtoxcanti|||"
            Top             =   240
            Width           =   2715
         End
         Begin VB.CheckBox chkMataPrimaPorcen 
            Caption         =   "Materia prima como porcentaje"
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
            Left            =   240
            TabIndex        =   24
            Tag             =   "Descriptores|N|N|||spara1|compoporcen|||"
            Top             =   3810
            Width           =   4605
         End
         Begin VB.CheckBox chkDescriptores 
            Caption         =   "Usa descriptores especiales"
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
            Left            =   240
            TabIndex        =   23
            Tag             =   "Descriptores|N|N|||spara1|descriptores|||"
            Top             =   3405
            Width           =   5085
         End
         Begin VB.CheckBox chkProduccion 
            Caption         =   "Tiene produccion"
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
            Left            =   240
            TabIndex        =   22
            Tag             =   "Tiene produccion|N|N|||spara1|produccion|||"
            Top             =   3015
            Width           =   4275
         End
         Begin VB.CheckBox chkHayServicio 
            Caption         =   "Hay Servicios"
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
            Left            =   240
            TabIndex        =   18
            Tag             =   "Hay Servicios|N|N|||spara1|hayservicio|||"
            Top             =   1425
            Width           =   2175
         End
         Begin VB.CheckBox chkCajacomp 
            Caption         =   "Cajas completas precios"
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
            Left            =   240
            TabIndex        =   15
            Tag             =   "Cajas Completas Precios|N|N|||spara1|cajacomp|||"
            Top             =   240
            Width           =   2775
         End
         Begin VB.CheckBox chkHaymante 
            Caption         =   "Realiza Mantenimientos"
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
            Left            =   240
            TabIndex        =   16
            Tag             =   "Mantenimientos|N|N|||spara1|haymante|||"
            Top             =   630
            Width           =   2835
         End
         Begin VB.CheckBox chkHayfrecu 
            Caption         =   "Hay Frecuencias"
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
            Left            =   240
            TabIndex        =   17
            Tag             =   "Hay Frecuencias|N|N|||spara1|hayfrecu|||"
            Top             =   1035
            Width           =   2175
         End
         Begin VB.CheckBox chkHaydepar 
            Caption         =   "Tiene Departamentos (o Direcci�n)"
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
            Left            =   240
            TabIndex        =   19
            Tag             =   "Departamento/Direc.|N|N|||spara1|haydepar|||"
            Top             =   1830
            Width           =   4125
         End
         Begin VB.CheckBox chkctrstock 
            Caption         =   "Control de Stock estricto"
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
            Left            =   240
            TabIndex        =   20
            Tag             =   "Control de Stock|N|N|||spara1|ctrstock|||"
            Top             =   2220
            Width           =   3885
         End
         Begin VB.CheckBox chkInventar 
            Caption         =   "Realiza Inventario por Proveedor"
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
            Left            =   240
            TabIndex        =   21
            Tag             =   "Inventarios por Proveedor|N|N|||spara1|inventar|||"
            Top             =   2610
            Width           =   4335
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Portes"
         Enabled         =   0   'False
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
         Height          =   975
         Left            =   -74790
         TabIndex        =   173
         Top             =   4230
         Visible         =   0   'False
         Width           =   8655
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
            Height          =   320
            Index           =   56
            Left            =   6720
            MaxLength       =   16
            TabIndex        =   41
            Tag             =   "R|N|S||10000|spara1|impminped|#,##0.00||"
            Text            =   "Text1 "
            Top             =   600
            Width           =   1455
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
            Height          =   320
            Index           =   55
            Left            =   6720
            MaxLength       =   16
            TabIndex        =   40
            Tag             =   "i |N|S|||spara1|abonokilos|#,##0.0000||"
            Text            =   "Text1 "
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox Text2 
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
            Height          =   315
            Index           =   54
            Left            =   840
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   175
            Text            =   "Text2"
            Top             =   600
            Width           =   3705
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
            Height          =   320
            Index           =   54
            Left            =   1200
            MaxLength       =   16
            TabIndex        =   39
            Tag             =   "Reci. |T|S|||spara1|ArticuloPortes|||"
            Text            =   "Text1 "
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Importe minimo pedido"
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
            Left            =   4800
            TabIndex        =   177
            Top             =   600
            Width           =   1620
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Abono kilos"
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
            Left            =   4800
            TabIndex        =   176
            Top             =   300
            Width           =   1620
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   54
            Left            =   840
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   285
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Articulo"
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
            Left            =   120
            TabIndex        =   174
            Top             =   300
            Width           =   780
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Impresora Tarjetas"
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
         Index           =   81
         Left            =   -74670
         TabIndex        =   271
         Top             =   3915
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Label Label1 
         Caption         =   "Path FacturaE"
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
         Index           =   80
         Left            =   -74460
         TabIndex        =   270
         Top             =   5745
         Width           =   2070
      End
      Begin VB.Label Label1 
         Caption         =   "Cooperativa"
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
         Index           =   77
         Left            =   -66540
         TabIndex        =   266
         Top             =   1890
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo Tarifa de PVP"
         Height          =   255
         Index           =   76
         Left            =   -69450
         TabIndex        =   265
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Agente"
         Height          =   255
         Index           =   31
         Left            =   4500
         TabIndex        =   207
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Situaci�n"
         Height          =   255
         Index           =   29
         Left            =   180
         TabIndex        =   206
         Top             =   5760
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   29
         Left            =   900
         ToolTipText     =   "Buscar situacion"
         Top             =   5760
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   31
         Left            =   5100
         ToolTipText     =   "Buscar agente"
         Top             =   5040
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Actividad"
         Height          =   255
         Index           =   25
         Left            =   180
         TabIndex        =   205
         Top             =   4470
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Envio"
         Height          =   195
         Index           =   26
         Left            =   4500
         TabIndex        =   204
         Top             =   4470
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Forma de Pago"
         Height          =   255
         Index           =   27
         Left            =   180
         TabIndex        =   203
         Top             =   5070
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Tarifa"
         Height          =   255
         Index           =   30
         Left            =   4500
         TabIndex        =   202
         Top             =   5790
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   4980
         ToolTipText     =   "Buscar forma de envio"
         Top             =   4470
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1380
         ToolTipText     =   "Buscar forma de pago"
         Top             =   5070
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   900
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   4470
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   30
         Left            =   4980
         ToolTipText     =   "Buscar tarifa"
         Top             =   5790
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Crear tarifas"
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
         Index           =   59
         Left            =   -74640
         TabIndex        =   182
         Top             =   2790
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Modo anal�tica"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   58
         Left            =   -74760
         TabIndex        =   179
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "N� Conta presupuestos *"
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   51
         Left            =   -74760
         TabIndex        =   172
         Top             =   1125
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "IVA intracomunitario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   47
         Left            =   -74760
         TabIndex        =   170
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   45
         Left            =   -72480
         Tag             =   "-1"
         ToolTipText     =   "Buscar I.V.A."
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Integ.Fras.Observaciones "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   53
         Left            =   -70410
         TabIndex        =   168
         Top             =   1125
         Width           =   2610
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   41
         Left            =   -72480
         Tag             =   "-1"
         ToolTipText     =   "Buscar I.V.A."
         Top             =   2085
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "IVA exento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   50
         Left            =   -74760
         TabIndex        =   164
         Top             =   2085
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Contabilidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   49
         Left            =   -74760
         TabIndex        =   162
         Top             =   600
         Width           =   1185
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   -73035
         Tag             =   "-1"
         ToolTipText     =   "Buscar tarifa"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Servidor"
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
         Left            =   -73020
         TabIndex        =   126
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "N� conta"
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
         Left            =   -66000
         TabIndex        =   125
         Top             =   600
         Width           =   1290
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
         Height          =   195
         Index           =   17
         Left            =   -70410
         TabIndex        =   124
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Pass."
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
         Left            =   -68280
         TabIndex        =   123
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Orden Descuentos"
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
         Left            =   -74640
         TabIndex        =   122
         Top             =   1860
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Descuentos"
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
         Left            =   -74640
         TabIndex        =   102
         Top             =   2325
         Width           =   2205
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo Tarifa de PVP"
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
         Left            =   -74640
         TabIndex        =   101
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del Director Gerente"
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
         Left            =   -74640
         TabIndex        =   100
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre responsable Administraci�n"
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
         Left            =   -74640
         TabIndex        =   99
         Top             =   900
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo"
         Height          =   255
         Index           =   6
         Left            =   -71280
         TabIndex        =   104
         Top             =   1020
         Width           =   495
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 102

Private WithEvents frmMtoArt As frmAlmArticulos
Attribute frmMtoArt.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1

Private WithEvents frmSoc As frmGesSocios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmA As frmFacActividades
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmE As frmFacFormasEnvio
Attribute frmE.VB_VarHelpID = -1
Private WithEvents frmAC As frmFacAgentesCom
Attribute frmAC.VB_VarHelpID = -1
Private WithEvents frmT As frmFacTarifas
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmS As frmFacSituaciones
Attribute frmS.VB_VarHelpID = -1


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1


Private NombreTabla As String  'Nombre de la tabla o de la
Private CadenaConsulta As String
Dim indice As Byte
Dim indCodigo As Integer

Dim PrimeraVez As Boolean
Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: A�adir
'4: Modificar






Private Sub cboCreaTarifa_KeyPress(KeyAscii As Integer)
 KEYpress KeyAscii
End Sub

Private Sub cboObsFactura_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboOrdenDtos_KeyPress(KeyAscii As Integer)
      KEYpress KeyAscii
End Sub

Private Sub cboTipodtos_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub





Private Sub chkCajacomp_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkCajacomp_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkctrstock_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkctrstock_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub






Private Sub ChkDtoxCantidad_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub ChkDtoxCantidad_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkHaydepar_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHaydepar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkHayfrecu_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHayfrecu_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkHaymante_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHaymante_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkHaynserie_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHaynserie_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkHayrepar_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHayrepar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkHayServicio_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHayServicio_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkInventar_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkInventar_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkInventar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub chkMataPrimaPorcen_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkMataPrimaPorcen_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkOutlook_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkOutlook_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkProduccion_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkProduccion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDescriptores_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkDescriptores_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub ChkTarifArt_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub ChkTarifArt_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkTicketsAgrupads_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkTicketsAgrupads_KeyPress(KeyAscii As Integer)
  KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim actualiza As Boolean
Dim kms As Currency


    If Modo = 4 Then 'MODIFICAR
        If DatosOk Then
            vParamAplic.TipoDtos = Me.cboTipodtos.ListIndex
            vParamAplic.OrdenDtos = Me.cboOrdenDtos.ListIndex
            vParamAplic.ObsFactura = Me.cboObsFactura.ListIndex
            vParamAplic.CodTarifa = Text1(1).Text
            vParamAplic.NomGerente = Text1(4).Text
            vParamAplic.NomRespAdmin = Text1(5).Text
            kms = ImporteFormateado(ComprobarCero(Text1(2).Text))
            vParamAplic.PrecioKmClientes = CSng(CStr(kms))
            kms = ImporteFormateado(ComprobarCero(Text1(3).Text))
            vParamAplic.PrecioKmTecnicos = CSng(CStr(kms))
            vParamAplic.CajasCompletas = Me.chkCajacomp.Value
            vParamAplic.Mantenimientos = Me.chkHaymante.Value
            vParamAplic.Reparaciones = Me.chkHayrepar.Value
            vParamAplic.Frecuencias = Me.chkHayfrecu.Value
            vParamAplic.Servicios = Me.chkHayServicio.Value
            vParamAplic.Departamento = Me.chkHaydepar.Value
            vParamAplic.ControlStock = Me.chkCtrStock.Value
            vParamAplic.InventarioxProv = Me.chkInventar.Value
            vParamAplic.NumSeries = Me.chkHaynserie.Value  'Hay N� Serie en Compras?
            vParamAplic.DiasSiMante = Me.Text1(7).Text 'Dias Rep. con Mantenimiento
            vParamAplic.DiasNoMante = Me.Text1(6).Text 'Dias Rep. sin Mantenimiento
            'datos publicidad
            vParamAplic.ConFactuPubli = Text1(59).Text 'concepto factura publicidad
            vParamAplic.Raiz_Cta_Soc_publi = Text1(60).Text 'raiz cuenta socios publicidad
            ' datos de liquidacion socio
            vParamAplic.Raiz_Cta_Soc_Liqui = Text1(65).Text 'raiz cuenta socios liquidacion
            ' raiz de la cta cliente del socio (430)
            vParamAplic.Raiz_CtaClien_Soc = Text1(66).Text 'raiz cuenta socios cuotas
            vParamAplic.CtaLiquidacion = Text1(71).Text ' cta base de liquidacion de socios
            
            '[Monica]16/11/2017: cta de liquidacion de suplidos
            vParamAplic.CtaLiqSuplidos = Text1(88).Text ' cta base de liq. suplidos de socios
            
            
            ' raiz de la cta cliente del socio para facturas FAV (equipamiento)
            vParamAplic.Raiz_Cta_Soc_Equip = Text1(77).Text 'raiz cuenta socios equipamiento
            
            'alta Socio
            vParamAplic.CtaTituloAlta = Text1(72).Text ' cta titulo 1 (alta socio)
            vParamAplic.CtaGastoAlta = Text1(73).Text ' cta gastos (alta socio)
            vParamAplic.ImpTituloAlta = ComprobarCero(Text1(74).Text) ' importe titulo
            vParamAplic.ImpGastoAlta = ComprobarCero(Text1(75).Text) ' importe gastos
            vParamAplic.Raiz_CtaAltaSoc = Text1(76).Text 'raiz cuenta alta socios
            
            
            'Articulo para facturar servicios
            vParamAplic.ArticServ = Me.Text1(15).Text
            'dias de pago para compras
            vParamAplic.DiaPago1 = CByte(DBLet(ComprobarCero(Text1(16).Text), "N"))
            vParamAplic.DiaPago2 = CByte(DBSet(Text1(17).Text, "N"))
            vParamAplic.DiaPago3 = CByte(DBSet(Text1(18).Text, "N"))
            vParamAplic.MesNoGirar = CByte(DBSet(Text1(19).Text, "N"))
            vParamAplic.ForPagoChequeRegalo = Me.Text1(24).Text
            
            vParamAplic.DireMail = Text1(8).Text 'Direccion email
            vParamAplic.SMTPhost = Text1(9).Text
            vParamAplic.SMTPuser = Text1(10).Text
            vParamAplic.SMTPpass = Text1(11).Text
            vParamAplic.WebSoporte = Text1(12).Text
            vParamAplic.MailSoporte = Text1(13).Text
            vParamAplic.WebVersion = Text1(14).Text
            
            'Datos contabilidad
            vParamAplic.ServidorConta = Text1(23).Text
            vParamAplic.UsuarioConta = Text1(21).Text
            vParamAplic.PasswordConta = Text1(20).Text
            vParamAplic.NumeroConta = ComprobarCero(Text1(22).Text)
            
            'Valores por defecto
            vParamAplic.PorDefecto_Activ = ComprobarCero(Text1(25).Text)
            vParamAplic.PorDefecto_Envio = ComprobarCero(Text1(26).Text)
            vParamAplic.PorDefecto_Situ = ComprobarCero(Text1(29).Text)
            vParamAplic.PorDefecto_Tarifa = ComprobarCero(Text1(30).Text)
            vParamAplic.PorDefecto_Agente = ComprobarCero(Text1(31).Text)
            
            vParamAplic.ArtCuotaExtraor = Text1(61).Text
            vParamAplic.ArtCuotaSinChofer = Text1(28).Text
            vParamAplic.ArtCuotaConChofer = Text1(70).Text
            vParamAplic.ArtAlquiler = Me.Text1(67).Text
            vParamAplic.ArtServCuotas = Me.Text1(69).Text
            kms = 0
            If Text3(28).Text <> "" Then kms = ImporteFormateado(ComprobarCero(Text3(28).Text))
            vParamAplic.PrecioCuotaSinChofe = CSng(CStr(kms))
            kms = 0
            If Text3(70).Text <> "" Then kms = ImporteFormateado(ComprobarCero(Text3(70).Text))
            vParamAplic.PrecioCuotaConChofe = CSng(CStr(kms))
            kms = 0
            If Text3(69).Text <> "" Then kms = ImporteFormateado(ComprobarCero(Text3(69).Text))
            vParamAplic.PrecioPorServicio = CSng(CStr(kms))
            kms = 0
            If Text3(67).Text <> "" Then kms = ImporteFormateado(ComprobarCero(Text3(67).Text))
            vParamAplic.PrecioPorAlquiler = CSng(CStr(kms))
            
            vParamAplic.PorDefecto_Zona = ComprobarCero(Text1(27).Text) '[Monica]08/02/2011 lo pido en el frame de cuotas 'lo usamos como forma de pago

            
            'Telefonia
            vParamAplic.CodarticTfnia = Me.Text1(32).Text
           
            'Los avisos
            vParamAplic.avipedcli = ComprobarCero(Text1(33).Text)
            vParamAplic.avipedpro = ComprobarCero(Text1(34).Text)
            vParamAplic.avialbcli = ComprobarCero(Text1(35).Text)
            vParamAplic.avialbpro = ComprobarCero(Text1(36).Text)
            vParamAplic.avimanteni = ComprobarCero(Text1(37).Text)
            vParamAplic.aviavisos = ComprobarCero(Text1(38).Text)
            vParamAplic.avirepara = ComprobarCero(Text1(39).Text)
            
            
            'Los tipos de IVA
            vParamAplic.TipoIVAre1 = ComprobarCero(Text1(40).Text)
            vParamAplic.TipoIVAre2 = ComprobarCero(Text1(41).Text)
            vParamAplic.TipoIVAre3 = ComprobarCero(Text1(42).Text)
             
            vParamAplic.TipoIVA1 = ComprobarCero(Text1(43).Text)
            vParamAplic.TipoIVA2 = ComprobarCero(Text1(44).Text)
            vParamAplic.TipoIVA3 = ComprobarCero(Text1(45).Text)
            
            
            'Los tipos de IVA antiguos
            vParamAplic.TipoIVAAntre1 = ComprobarCero(Text1(84).Text)
            vParamAplic.TipoIVAAntre2 = ComprobarCero(Text1(86).Text)
             
            vParamAplic.TipoIVAAnt1 = ComprobarCero(Text1(83).Text)
            vParamAplic.TipoIVAAnt2 = ComprobarCero(Text1(85).Text)
            
            'REtencion y REA
            vParamAplic.IVA_REA = ComprobarCero(Text1(46).Text)
            vParamAplic.Raiz_Cta_Reten_Soc = ComprobarCero(Text1(47).Text)
            vParamAplic.PorReten = ComprobarCero(Text1(49).Text)
            
            vParamAplic.PorDtoCentra = ComprobarCero(Text1(96).Text)
            
            'IVA exento
            vParamAplic.IVA_Exento2 = ComprobarCero(Text1(48).Text)
            vParamAplic.IVA_Intracomunitario = ComprobarCero(Text1(52).Text)

            
            'Tickets acgrupados
            vParamAplic.ContabilizarTicketAgrupados = Me.chkTicketsAgrupads.Value
            
            vParamAplic.ContabilidadB = ComprobarCero(Text1(50).Text)
            vParamAplic.ctaAportacion = Text1(51).Text
            
            vParamAplic.Produccion = Me.chkProduccion.Value
            vParamAplic.Descriptores = Me.chkDescriptores.Value
            
            vParamAplic.ArtGastosAdmon = Text1(53).Text
            vParamAplic.ArtSuplidos = Text1(87).Text
            
            'Portes(FOntenas)
            vParamAplic.ArtPortes = Text1(54).Text
            vParamAplic.AbonoKilos = ComprobarCero(Text1(55).Text)
            vParamAplic.ImporteMinimo = ComprobarCero(Text1(56).Text)
            
            vParamAplic.ComponentePorcentaje = Me.chkMataPrimaPorcen.Value
            
            ' ---- [14/09/2009] (LAURA)
            vParamAplic.DtoxCantidad = Me.ChkDtoxCantidad.Value
            vParamAplic.CreaTarifasArticulo = Me.cboCreaTarifa.ItemData(cboCreaTarifa.ListIndex)
            
            ' ----
            
            ' ---- [19/10/2009] [LAURA]: a�adir campo modo analitica
            If Me.CboModAnalitica.ListIndex >= 0 Then
                vParamAplic.ModoAnalitica = Me.CboModAnalitica.ListIndex
            End If
            
            vParamAplic.EnvioDesdeOutlook = Me.chkOutlook.Value
            
            
            vParamAplic.ExeEnvioMail = Trim(Text1(57).Text)
            vParamAplic.DiasGarantia = ComprobarCero(Text1(58).Text)
            
            If Text1(78).visible Then
                vParamAplic.Cooperativa = ComprobarCero(Text1(78).Text)
            End If
            vParamAplic.PrecioxDistancia = ComprobarCero(Text1(79).Text)
            vParamAplic.PrecioxTpoEspera = ComprobarCero(Text1(80).Text)
            
            '[Monica]05/09/2012: el Path de destino de las facturas lo indicamos en paramentros
            vParamAplic.PathFacturaE = Text1(81).Text 'Replace(Text1(81).Text, "\", "\\")
            
            '[Monica]28/09/2012: Impresora de tarjetas
            vParamAplic.ImpresoraTarjetas = Text1(82).Text
            
            '[Monica]05/12/2017: socio de la otra cooperativa y cliente de la otra cooperativa
            vParamAplic.SocioCooperativa = Text1(89).Text
            vParamAplic.ClienteCooperativa = Text1(90).Text
            vParamAplic.EmpresaTaxitronic = Text1(91).Text
            
            
            '[Monica]28/02/2018: se marcan como validados
            vParamAplic.MarcarValidados = Me.ChkMarcarValidados.Value
            
            
            
            '[Monica]14/03/2019: impresion de certificados de aportaciones
            vParamAplic.AporSerie = Text1(92).Text
            vParamAplic.AporNumero = Text1(93).Text
            vParamAplic.AporPresidente = Text1(94).Text
            vParamAplic.AporSecretario = Text1(95).Text
            
            
            actualiza = vParamAplic.Modificar(Text1(0).Text)
            TerminaBloquear

            vParamAplic.ComprobarProgramaEnvioMail
            
            
            

            If actualiza Then  'Inserta o Modifica
                'Abrir la conexion a la conta q hemos modificado
                CerrarConexionConta
                If vParamAplic.NumeroConta <> 0 Then
                    If Not AbrirConexionConta(False) Then End
                End If
                PonerModo 2
                PonerFocoBtn Me.cmdSalir
            End If
        End If
    End If
End Sub


Private Sub cmdCancelar_Click()
    TerminaBloquear
    If Data1.Recordset.EOF Then
        PonerModo 0
        LimpiarCampos
    Else
        PonerCampos
        PonerModo 2
    End If
End Sub


Private Sub cmdSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerModo 0
    Else
        If Modo <> 4 Then PonerCadenaBusqueda
        PonerFoco Text1(0)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim Im
    'Icono del formulario
    Me.Icon = frmppal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .Buttons(1).Image = 4   'Modificar
    End With
    
'    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With
    
    
    
    'cargar iconos de busqueda
    For Each Im In Me.imgBuscar
        Im.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    'imgBuscar(1).Picture = frmPpal.imgListComun.ListImages(19).Picture
    'imgBuscar(15).Picture = frmPpal.imgListComun.ListImages(19).Picture
   '
   ' For NumRegElim = 24 To 42
   '     Me.imgBuscar(NumRegElim).Picture = frmPpal.imgListComun.ListImages(19).Picture
   ' Next NumRegElim
    
    '[Monica]19/02/2018: Entra Cordoba
        '[Monica]19/11/2018: Entra Sevilla
    Frame16.visible = (vParamAplic.Cooperativa <> 0 And vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 3)
    Frame16.Enabled = (vParamAplic.Cooperativa <> 0 And vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 3)


    LimpiarCampos   'Limpia los campos TextBox
    Me.SSTab1.Tab = 0
    
    CargarComboTipoDtos
    CargarComboOrdenDtos
    CargaComoboObsFactura
    CargarComboModoAnalitica
    
    
    ' ---- [21/10/2009] [LAURA]
    '-- modo analitica si contabilidad lleva analitica
    If vEmpresa.LeerNiveles Then
        Label1(58).visible = vEmpresa.TieneAnalitica
        Me.CboModAnalitica.visible = vEmpresa.TieneAnalitica
    End If
    
    Me.Label1(77).visible = (vUsu.Login = "root")
    Me.Label1(77).Enabled = (vUsu.Login = "root")
    Me.Text1(78).visible = (vUsu.Login = "root")
    Me.Text1(78).Enabled = (vUsu.Login = "root")
    
    '[Monica]01/02/2019: cambiamos el label de path facturae por lo de URL para sevilla
    If vParamAplic.Cooperativa = 3 Then
        Label1(80).Caption = "Direccion URL"
    End If
    
    
    NombreTabla = "spara1"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    PonerModo 0
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Me.Toolbar1.Buttons(1).Enabled = False 'Modificar
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub






Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    'ENVIO
    Text1(25).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(25).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmAC_DatoSeleccionado(CadenaSeleccion As String)
    'agentes
    Text1(31).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(31).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadenaDesdeOtroForm = CadenaDevuelta
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmE_DatoSeleccionado(CadenaSeleccion As String)
    'ENVIO
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago


    
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub

Private Sub frmMtoArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    
    Text1(CInt(imgBuscar(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod articulo
    Text2(CInt(imgBuscar(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre articulo
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    'SITUACION
    Text1(29).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(29).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    'TARIFA
    If Not IsNumeric(Me.imgBuscar(1).Tag) Then Exit Sub
    
    If CInt(Me.imgBuscar(1).Tag) = 1 Then
        Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        Text1(30).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(30).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim i As Integer
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'obervaciones de concepto de publicidad
            If Modo = 5 Or Modo = 0 Then
            Else
                
                If Modo = 3 Or Modo = 4 Then
                    CadenaDesdeOtroForm = Text1(59).Text
                Else
                    CadenaDesdeOtroForm = ""
                    If Not Data1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Data1.Recordset!ConFactuPubli, "T")
                End If
                frmFacClienteObser.Modificar = Modo >= 3
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 3 Or Modo = 4 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(59).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
            End If
            PonerFoco Text1(59)
        Case 15, 32, 53, 54, 28, 70, 87 'cod. articulo
            Me.imgBuscar(1).Tag = Index
            Set frmMtoArt = New frmAlmArticulos
            frmMtoArt.DatosADevolverBusqueda2 = "@1@"
            frmMtoArt.Show vbModal
            Set frmMtoArt = Nothing
            
        Case 2 'codarticulo
            Me.imgBuscar(1).Tag = 61
            Set frmMtoArt = New frmAlmArticulos
            frmMtoArt.DatosADevolverBusqueda2 = "@1@"
            frmMtoArt.Show vbModal
            Set frmMtoArt = Nothing
        
        Case 4 'codarticulo
            Me.imgBuscar(1).Tag = 69
            Set frmMtoArt = New frmAlmArticulos
            frmMtoArt.DatosADevolverBusqueda2 = "@1@"
            frmMtoArt.Show vbModal
            Set frmMtoArt = Nothing
        
        
        Case 24 'forma de pago
            If Modo = 4 Then TerminaBloquear
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            indice = 24
            Set frmFP = Nothing
            If Modo = 4 Then
                If Not BLOQUEADesdeFormulario(Me) Then cmdCancelar_Click
            End If
    
        Case 25 'Codigo Actividad
            Set frmA = New frmFacActividades
            frmA.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 26  'Cod. Envio
            Set frmE = New frmFacFormasEnvio
            frmE.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmE.Show vbModal
            Set frmE = Nothing
            
            
        Case 27 'forma de pago
            indice = 27
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
'        Case 4  'Cod. Forma de Pago
'            Set frmFP = New frmFacFormasPago
'            frmFP.DatosADevolverBusqueda = "0"
'            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
'            frmFP.Show vbModal
'            Set frmFP = Nothing
            
            
        Case 31 'C�digo de Agente
            Set frmAC = New frmFacAgentesCom
            frmAC.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmAC.Show vbModal
            Set frmAC = Nothing
            
        Case 1, 30 'C�digo de Tarifa
            Me.imgBuscar(1).Tag = Index
            Set frmT = New frmFacTarifas
            frmT.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 29 'C�digo de Situaci�n
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
             If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmS.Show vbModal
            Set frmS = Nothing
            
            
        Case 33 To 42, 45, 6, 7, 8, 9 To 13 'Todos los ivas y la Cta de retencion, y cuenta aportacion TERMINAL
            CadenaDesdeOtroForm = ""
                        
            BuscaBuscaGRid2 (Index <> 40 And Index <> 42 And Index <> 6 And Index <> 7 And Index <> 8 And Index <> 13)
            If CadenaDesdeOtroForm <> "" Then
                Select Case Index
                    Case 42
                        i = 9 'Para la cta aportacion
                    Case 6
                        i = 65 ' cta de liquidacion
                    Case 7
                        i = 65 ' cta de titulo d alta de socio
                    Case 8
                        i = 65 ' cta de gastos de alta de socio
                    Case 9, 10, 11, 12
                        i = 74
                    Case 13
                        i = 75
                    Case Else
                        i = 7
                End Select
                Text1(Index + i).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                Text2(Index + i).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            End If
        
        Case 3 ' articulo de alquiler de equipos
            Me.imgBuscar(1).Tag = 67
            Set frmMtoArt = New frmAlmArticulos
            frmMtoArt.DatosADevolverBusqueda2 = "@1@"
            frmMtoArt.Show vbModal
            Set frmMtoArt = Nothing
            
        Case 4 ' articulo de servicios de cuotas normales
            Me.imgBuscar(1).Tag = 69
            Set frmMtoArt = New frmAlmArticulos
            frmMtoArt.DatosADevolverBusqueda2 = "@1@"
            frmMtoArt.Show vbModal
            Set frmMtoArt = Nothing
        
        Case 5 ' articulo de  cuotas normales sin con
            Me.imgBuscar(1).Tag = 70
            Set frmMtoArt = New frmAlmArticulos
            frmMtoArt.DatosADevolverBusqueda2 = "@1@"
            frmMtoArt.Show vbModal
            Set frmMtoArt = Nothing
        
        Case 14 ' codigo de socio que es la otra cooperativa
            indCodigo = 89
            Set frmSoc = New frmGesSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            If CadenaDesdeOtroForm <> "" Then
                Text1(indCodigo).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                Text2(indCodigo).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            End If
            
        Case 16 ' codigo de cliente que es la otra cooperativa
            indCodigo = 90
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0|1|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            PonerFoco Text1(indCodigo)
            
    End Select
    PonerFoco Text1(Index)
End Sub


Private Sub BuscaBuscaGRid2(EsIVa As Boolean)


    Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        If EsIVa Then
            'Busco IVAS
            frmB.vCampos = "C�digo|tiposiva|codigiva|N||20�Denominacion|tiposiva|nombriva|T||70�"
            frmB.vTabla = "tiposiva"
            frmB.vTitulo = "Tipos de IVA"
        Else
                
            frmB.vCampos = "C�digo|cuentas|codmacta|T||20�Denominacion|cuentas|nommacta|T||70�"
            frmB.vTabla = "cuentas"
            frmB.vTitulo = "Cta contable"
            frmB.vSQL = "apudirec = 'S'"
        
        End If
        frmB.vDevuelve = "0|1|"
        frmB.vselElem = 1
        frmB.vConexionGrid = conConta

        frmB.vCargaFrame = False
      
        frmB.Show vbModal
        Set frmB = Nothing


    Screen.MousePointer = vbDefault

End Sub











Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

'*********************
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 1: KEYBusqueda KeyAscii, 1 'tarifa
                
                Case 15: KEYBusqueda KeyAscii, 15 'articulo servicios
                Case 53: KEYBusqueda KeyAscii, 53 'articulo gastos adm
                Case 87: KEYBusqueda KeyAscii, 87 'articulo suplidos
                Case 46: KEYBusqueda KeyAscii, 39 'iva liquidacion
                
                Case 71: KEYBusqueda KeyAscii, 6 'cta base liquidacion
                Case 88: KEYBusqueda KeyAscii, 13 'cta base de suplidos
                Case 48: KEYBusqueda KeyAscii, 41 'iva exento
                Case 52: KEYBusqueda KeyAscii, 45 'iva intracomunitario
                
                'ivas
                Case 43: KEYBusqueda KeyAscii, 36 'ivas
                Case 40: KEYBusqueda KeyAscii, 33 'ivas
                Case 44: KEYBusqueda KeyAscii, 37 'ivas
                Case 41: KEYBusqueda KeyAscii, 34 'ivas
                Case 45: KEYBusqueda KeyAscii, 38 'ivas
                Case 42: KEYBusqueda KeyAscii, 35 'ivas
                Case 83: KEYBusqueda KeyAscii, 9 'ivas
                Case 84: KEYBusqueda KeyAscii, 10 'ivas
                Case 85: KEYBusqueda KeyAscii, 11 'ivas
                Case 86: KEYBusqueda KeyAscii, 12 'ivas
                
                Case 28: KEYBusqueda KeyAscii, 28 'cuota normal con
                Case 70: KEYBusqueda KeyAscii, 5 'cuota sin
                Case 61: KEYBusqueda KeyAscii, 2 'cuota
                Case 67: KEYBusqueda KeyAscii, 3 'alquiler
                Case 69: KEYBusqueda KeyAscii, 4 'servicios
                Case 32: KEYBusqueda KeyAscii, 32 'publicidad
                Case 59: KEYBusqueda KeyAscii, 0 'concepto
                Case 84: KEYBusqueda KeyAscii, 10 'ivas
                
                Case 72: KEYBusqueda KeyAscii, 7 'cta aportacion
                Case 73: KEYBusqueda KeyAscii, 8 'cta reserva
                Case 89: KEYBusqueda KeyAscii, 14 'socio cooperativa
                Case 90: KEYBusqueda KeyAscii, 16 'cliente a facturar servicios
                
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub


'*********************
Private Sub Text1_LostFocus(Index As Integer)
'    If Text1(Index).Text = "" Then Exit Sub

    'Quitar espacios en blanco
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    Select Case Index
        Case 1 'tarifa de PVP
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "starif", "nomlista", "codlista", , "N")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 2 'Km desplaz clientes
            PonerFormatoDecimal Text1(Index), 5 'Tipo 4: Decimal(8,4)
        Case 3 'Km desplaz tecnicos
            PonerFormatoDecimal Text1(Index), 5 'Tipo 4: Decimal(8,4)
            
'        Case 6, 7 'Dias Reparacion con/sin mantenimiento
'            If Not EsNumerico(Text1(Index).Text) Then
'                Text1(Index).Text = ""
'                PonerFoco Text1(Index)
'            End If
        Case 14
            'PonerFocoBtn Me.cmdAceptar
            
        Case 15, 32, 53, 54, 87 'cod. artic
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic", "codartic", "Articulo")
            If Text2(Index).Text = "" Then Text1(Index).Text = ""
            
        Case 28, 61, 67, 69, 70 'cod. artic
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic", "codartic", "Articulo")
            If Text1(Index).Text <> "" Then Text3(Index).Text = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", Text1(Index).Text, "T")
            If Text2(Index).Text = "" Then Text1(Index).Text = ""
            
        
        Case 22 'n� conta
            'PonerFocoBtn Me.cmdAceptar
            
        Case 24 'FORMA DE PAGO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            'PonerFocoBtn Me.cmdAceptar
            
            
        Case 25 To 31
            'Campos por defecto
            Debug.Print Index & "-" & Text1(Index).Tag & ": " & Text1(Index).Text; ""
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
            Else
                Select Case Index
                Case 25
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sactiv", "nomactiv", "codactiv")
                Case 26
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "senvio", "nomenvio", "codenvio")
                Case 29
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "ssitua", "nomsitua", "codsitua")
                Case 30
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "starif", "nomlista", "codlista")
                Case 31
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent", "codagent")
                End Select
            End If
            
        Case 40 To 46, 48, 52
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conConta, "tiposiva", "nombriva", "codigiva")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 83 To 86
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conConta, "tiposiva", "nombriva", "codigiva")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 51
            'Cta retencion y Cta aportacion al terminal
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conConta, "cuentas", "nommacta", "codmacta")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 71, 72, 73, 88
            '71=Cta liquidacion
            '72=cta de titulo de alta de socio
            '73=cta de gastos de alta de socio
            '88=cta de liquidacion de suplidos
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conConta, "cuentas", "nommacta", "codmacta")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 49, 96
            'pORCE RETENCION, %dto centralizacion
            PonerFormatoDecimal Text1(48), 4
            
        Case 50, 58, 78, 91
            PonerFormatoEntero Text1(Index)
            
        Case 55
            PonerFormatoDecimal Text1(Index), 5   'cuatro decimales
            
        Case 74, 75
            PonerFormatoDecimal Text1(Index), 3
            
        Case 56
            PonerFormatoDecimal Text1(Index), 3
        Case 60 'raiz cuenta socio publicidad
            If Len(Text1(Index).Text) <> vEmpresa.DigitosNivelAnterior Then
                MsgBox "La raiz de la cuenta de socios publicidad debe tener " & vEmpresa.DigitosNivelAnterior & " digitos.", vbExclamation
'                PonerFoco Text1(60)
            End If
        Case 65 'raiz cuenta socio liquidacion
            If Len(Text1(Index).Text) <> vEmpresa.DigitosNivelAnterior Then
                MsgBox "La raiz de la cuenta de socios liquidaci�n debe tener " & vEmpresa.DigitosNivelAnterior & " digitos.", vbExclamation
'                PonerFoco Text1(65)
            End If
            
            
        Case 47 'raiz cuenta retencion liquidacion
            If Len(Text1(Index).Text) <> vEmpresa.DigitosNivelAnterior Then
                MsgBox "La raiz de la cuenta de Retenci�n socios debe tener " & vEmpresa.DigitosNivelAnterior & " digitos.", vbExclamation
'                PonerFoco Text1(47)
            End If
        Case 66 'raiz cuenta socio como cliente para cuotas
            If Len(Text1(Index).Text) <> vEmpresa.DigitosNivelAnterior Then
                MsgBox "La raiz de la cuenta cliente de socios debe tener " & vEmpresa.DigitosNivelAnterior & " digitos.", vbExclamation
'                PonerFoco Text1(66)
            End If
            
        Case 76 'raiz alta cuenta socio
            If Len(Text1(Index).Text) <> vEmpresa.DigitosNivelAnterior Then
                MsgBox "La raiz de la cuenta alta de socios debe tener " & vEmpresa.DigitosNivelAnterior & " digitos.", vbExclamation
'                PonerFoco Text1(76)
            End If
    
        Case 77 'raiz cuenta socio equipamiento
            If Len(Text1(Index).Text) <> vEmpresa.DigitosNivelAnterior Then
                MsgBox "La raiz de la cuenta de socios facturas venta debe tener " & vEmpresa.DigitosNivelAnterior & " digitos.", vbExclamation
'                PonerFoco Text1(65)
            End If
    
        Case 79, 80 ' precio por distancia y precio por tpo de espera
            PonerFormatoDecimal Text1(Index), 5   'cuatro decimales
    
        Case 89 ' codigo de socio que engloba a los de la otra empresa
            PonerFormatoEntero Text1(Index)
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "N")

        Case 90 ' codigo de cliente que es la otra empresa
            PonerFormatoEntero Text1(Index)
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "scliente", "nomclien", "codclien", "N")
    
        '[Monica]14/03/2019: tema de impresion de aportaciones
        Case 92 ' serie de aportaciones
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index))
    End Select
End Sub


Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 6, 7, 16, 17, 18
            If Text1(Index).Text <> "" Then
                If Not EsNumerico(Text1(Index).Text) Then
                    Cancel = True
                    ConseguirFoco Text1(Index), Modo
                End If
            End If
    End Select
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
'        Case 1  'Anyadir
'            BotonAnyadir
        Case 1  'Modificar
            mnModificar_Click
        Case 4 'Salir
            mnSalir_Click
    End Select
End Sub


'Private Sub BotonAnyadir()
'    LimpiarCampos
'    PonerModo 3
'
'    Text1(0).Text = SugerirCodigoSiguienteStr("scryst", "codcryst")
'    PonerFoco Text1(0)
'End Sub


Private Sub BotonModificar()
    PonerModo 4
    
    Select Case Me.SSTab1.Tab
        Case 0:    PonerFoco Text1(4)
        Case 1: PonerFoco Text1(32)
        Case 2: PonerFoco Text1(8)
        Case 3: PonerFoco Text1(23)
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean

    On Error GoTo ErrOK

    DatosOk = False
    
    'Para que no de errores insesperados
    If Text1(6).Text = "" Then Text1(6).Text = "0"
    If Text1(7).Text = "" Then Text1(7).Text = "0"
    
    
    
    b = CompForm(Me, 1)
    
    '--- forma de pago de CHEQUE regalo
    'comprobar q el tipo de la forma de pago es EFECTIVO
    If b And Text1(24).Text <> "" Then
            If DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", Text1(24).Text, "N") <> "0" Then
                MsgBox "La forma de pago del cheque debe ser del tipo EFECTIVO", vbExclamation
                b = False
            End If
    End If
    
    If Text1(47).Text = "" Xor Text1(49).Text = "" Then
        MsgBox "Ra�z Cta retenci�n o % retenci�n vacios", vbExclamation
        Exit Function
    End If
    
    
    If cboCreaTarifa.ListIndex < 0 Then
        MsgBox "Seleccion valor para crear tarifa", vbExclamation
        Exit Function
    End If
    
    DatosOk = b
    Exit Function
    
ErrOK:
    MuestraError Err.Number, "Comprobar datos", Err.Description
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdSalir.visible = b
    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    
    'poner descripcion del articulo
    Text2(15).Text = PonerNombreDeCod(Text1(15), conAri, "sartic", "nomartic", "codartic", "Art�culos")
    Text2(32).Text = PonerNombreDeCod(Text1(32), conAri, "sartic", "nomartic", "codartic", "Art�culos")
    
    Text2(87).Text = PonerNombreDeCod(Text1(87), conAri, "sartic", "nomartic", "codartic", "Art�culos")
    
    Text2(89).Text = PonerNombreDeCod(Text1(89), conAri, "sclien", "nomclien", "codclien", "Socios")
    Text2(90).Text = PonerNombreDeCod(Text1(90), conAri, "scliente", "nomclien", "codclien", "Clientes")
    
        
    ' Art�culos para cuotas
    Text2(28).Text = PonerNombreDeCod(Text1(28), conAri, "sartic", "nomartic", "codartic", "Art�culos")
    If Text1(28).Text <> "" Then
        Text3(28).Text = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", Text1(28).Text, "T")
        PonerFormatoDecimal Text3(28), 2
    End If
    Text2(70).Text = PonerNombreDeCod(Text1(28), conAri, "sartic", "nomartic", "codartic", "Art�culos")
    If Text1(70).Text <> "" Then
        Text3(70).Text = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", Text1(70).Text, "T")
        PonerFormatoDecimal Text3(70), 2
    End If
    Text2(61).Text = PonerNombreDeCod(Text1(61), conAri, "sartic", "nomartic", "codartic", "Art�culos")
    If Text1(61).Text <> "" Then
        Text3(61).Text = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", Text1(61).Text, "T")
        PonerFormatoDecimal Text3(61), 2
    End If
    Text2(67).Text = PonerNombreDeCod(Text1(67), conAri, "sartic", "nomartic", "codartic", "Art�culos")
    If Text1(67).Text <> "" Then
        Text3(67).Text = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", Text1(67).Text, "T")
        PonerFormatoDecimal Text3(67), 2
    End If
    Text2(69).Text = PonerNombreDeCod(Text1(69), conAri, "sartic", "nomartic", "codartic", "Art�culos")
    If Text1(69).Text <> "" Then
        Text3(69).Text = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", Text1(69).Text, "T")
        PonerFormatoDecimal Text3(69), 2
    End If
    
    'poner descripcion de la forma de pago
    Text2(24).Text = PonerNombreDeCod(Text1(24), conAri, "sforpa", "nomforpa", "codforpa")
    Text2(27).Text = PonerNombreDeCod(Text1(27), conAri, "sforpa", "nomforpa", "codforpa")
    
    'poner descripcion de la tarifa de PVP
    Text2(1).Text = PonerNombreDeCod(Text1(1), conAri, "starif", "nomlista", "codlista", , "N")
    
    
    For NumRegElim = 25 To 54
        If NumRegElim < 49 Or NumRegElim > 50 Or NumRegElim <> 47 Then
            If Text1(NumRegElim).Text <> "" Then Text1_LostFocus CInt(NumRegElim)
        End If
    Next NumRegElim
    NumRegElim = 0
    
    Text1_LostFocus (71) ' cta base de liquidacion de socios
    Text1_LostFocus (72) ' cta titulo de alta de socio
    Text1_LostFocus (73) ' cta gasto de alta de socio
    '[Monica]16/11/2017: liquidacion de suplidos
    Text1_LostFocus (88) ' cta liquidacion de suplidos
    
    ' tipos de iva
    For NumRegElim = 83 To 86
        If Text1(NumRegElim).Text <> "" Then Text1_LostFocus CInt(NumRegElim)
    Next NumRegElim
    
    
    
    BloquearChecks Me, Modo
    
    Exit Sub
    
EPonerCampos:
    MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim i As Integer

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
      
    For i = 0 To Text1.Count - 1
        Text1(i).BackColor = vbWhite
    Next i
      
      
    '------------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    PonerBotonCabecera Not b
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1 y bloquea la clave primaria
    BloquearText1 Me, Modo
    
    'Bloquear los checkbox
    BloquearChecks Me, Modo

    'Bloquear el combobox
    b = Modo = 4
    Me.cboTipodtos.Enabled = b
    Me.cboOrdenDtos.Enabled = b
    Me.cboObsFactura.Enabled = b
    Me.cboCreaTarifa.Enabled = b
    BloquearCmb Me.CboModAnalitica, Not b
    
    
    'Bloquear imagen de Busqueda
    Dim img As Image
    For Each img In Me.imgBuscar
        BloquearImg img, Not b
    Next
'    BloquearImg Me.imgBuscar(1), (Modo <> 4)
'    BloquearImg Me.imgBuscar(15), (Modo <> 4)
'    For NumRegElim = 24 To 42
'        BloquearImg Me.imgBuscar(NumRegElim), (Modo <> 4)
'    Next NumRegElim
'    NumRegElim = 0
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n el Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
    PonerModoUsuarioGnral Modo, "aritaxi"
End Sub

 

Private Sub PonerModoUsuarioGnral(Modo As Byte, Aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(Aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = Toolbar1.Buttons(1).Enabled And DBLet(Rs!Modificar, "N")
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not b 'Modificar
    Me.mnModificar.Enabled = Not b
'    Me.Toolbar1.Buttons(2).Enabled = (Not b) 'Modificar
End Sub


Private Sub CargarComboTipoDtos()
'### Combo Tipo Descuento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Aditivo, 1-sobre Resto

    cboTipodtos.Clear
    cboTipodtos.AddItem "Aditivo"
    cboTipodtos.ItemData(cboTipodtos.NewIndex) = 0
    
    cboTipodtos.AddItem "sobre Resto"
    cboTipodtos.ItemData(cboTipodtos.NewIndex) = 1
End Sub


Private Sub CargarComboOrdenDtos()
'### Combo Tipo Descuento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Aditivo, 1-sobre Resto

    Me.cboOrdenDtos.Clear
    Me.cboOrdenDtos.AddItem "Familia/Marca"
    cboOrdenDtos.ItemData(cboOrdenDtos.NewIndex) = 0
    
    cboOrdenDtos.AddItem "Marca/Familia"
    cboOrdenDtos.ItemData(cboOrdenDtos.NewIndex) = 1
End Sub

Private Sub CargaComoboObsFactura()
'## Cuando contabilice, que valor pondra en el campo observaciones del
'   la factura, tanto cliente como de proveedores

    Me.cboObsFactura.Clear
    Me.cboObsFactura.AddItem "Sin observaciones"
    cboObsFactura.ItemData(cboObsFactura.NewIndex) = 0
    
    cboObsFactura.AddItem "N�mero factura"
    cboObsFactura.ItemData(cboObsFactura.NewIndex) = 1

    cboObsFactura.AddItem "Fecha integraci�n"
    cboObsFactura.ItemData(cboObsFactura.NewIndex) = 2

End Sub



' ---- [19/10/2009] [LAURA]: a�adir campo modo analitica
Private Sub CargarComboModoAnalitica()
'### Combo modo analitica
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Trabajador, 1-Familia, 2-Proyecto

    Me.CboModAnalitica.Clear
    Me.CboModAnalitica.AddItem "Trabajador"
    CboModAnalitica.ItemData(CboModAnalitica.NewIndex) = 0
    
    CboModAnalitica.AddItem "Familia"
    CboModAnalitica.ItemData(CboModAnalitica.NewIndex) = 1
    
    CboModAnalitica.AddItem "Proyecto"
    CboModAnalitica.ItemData(CboModAnalitica.NewIndex) = 2
End Sub
