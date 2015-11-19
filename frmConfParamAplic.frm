VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfParamAplic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de la Aplicación"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9705
   Icon            =   "frmConfParamAplic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8400
      TabIndex        =   84
      Top             =   7875
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   240
      TabIndex        =   87
      Top             =   7800
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   210
         Width           =   2280
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7200
      TabIndex        =   83
      Top             =   7875
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8400
      TabIndex        =   85
      Top             =   7875
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   86
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3600
      Top             =   7920
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
      TabIndex        =   89
      Top             =   630
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Datos Varios"
      TabPicture(0)   =   "frmConfParamAplic.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(14)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(59)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(76)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(77)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FrameDiasMante"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "FrameOpciones"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "FramePrecioKm"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboTipodtos"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cboOrdenDtos"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text2(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboCreaTarifa"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Frame13"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chkHayrepar"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(78)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Frame16"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkHaynserie"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Datos Facturación"
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
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).Control(1)=   "Frame11"
      Tab(4).Control(2)=   "Text2(25)"
      Tab(4).Control(3)=   "Text1(25)"
      Tab(4).Control(4)=   "Text2(26)"
      Tab(4).Control(5)=   "Text1(26)"
      Tab(4).Control(6)=   "Text2(27)"
      Tab(4).Control(7)=   "Text1(27)"
      Tab(4).Control(8)=   "Text2(29)"
      Tab(4).Control(9)=   "Text1(29)"
      Tab(4).Control(10)=   "Text2(30)"
      Tab(4).Control(11)=   "Text1(30)"
      Tab(4).Control(12)=   "Text2(31)"
      Tab(4).Control(13)=   "Text1(31)"
      Tab(4).Control(14)=   "Frame5"
      Tab(4).Control(15)=   "Frame7"
      Tab(4).Control(16)=   "Text1(62)"
      Tab(4).Control(17)=   "Text1(63)"
      Tab(4).Control(18)=   "Text1(64)"
      Tab(4).Control(19)=   "Text1(68)"
      Tab(4).Control(20)=   "Label1(31)"
      Tab(4).Control(21)=   "Label1(29)"
      Tab(4).Control(22)=   "imgBuscar(29)"
      Tab(4).Control(23)=   "imgBuscar(31)"
      Tab(4).Control(24)=   "Label1(25)"
      Tab(4).Control(25)=   "Label1(26)"
      Tab(4).Control(26)=   "Label1(27)"
      Tab(4).Control(27)=   "Label1(30)"
      Tab(4).Control(28)=   "imgBuscar(26)"
      Tab(4).Control(29)=   "imgBuscar(27)"
      Tab(4).Control(30)=   "imgBuscar(25)"
      Tab(4).Control(31)=   "imgBuscar(30)"
      Tab(4).ControlCount=   32
      TabCaption(5)   =   "Varios"
      TabPicture(5)   =   "frmConfParamAplic.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text1(82)"
      Tab(5).Control(1)=   "Frame14"
      Tab(5).Control(2)=   "Label1(81)"
      Tab(5).ControlCount=   3
      Begin VB.CheckBox chkHaynserie 
         Caption         =   "Hay Nº Serie en Compras"
         Height          =   375
         Left            =   4320
         TabIndex        =   273
         Tag             =   "Hay Nº Serie en Compras|N|N|||spara1|haynserie|||"
         Top             =   2130
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   82
         Left            =   -72000
         MaxLength       =   100
         TabIndex        =   244
         Tag             =   "Impresora Tarjetas|T|S|||spara1|impretarjeta|||"
         Top             =   3690
         Visible         =   0   'False
         Width           =   4590
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   81
         Left            =   -72870
         MaxLength       =   255
         TabIndex        =   49
         Tag             =   "Path FacturaE|T|S|||spara1|pathfacturae|||"
         Top             =   5700
         Width           =   6030
      End
      Begin VB.Frame Frame16 
         Caption         =   "Cálculo Importes Llamada"
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
         Height          =   2025
         Left            =   4020
         TabIndex        =   258
         Top             =   2880
         Width           =   5055
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   80
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Precio por tpo de espera|N|S|||spara1|precioxtpoespera|###,##0.0000||"
            Text            =   "3"
            Top             =   810
            Width           =   1260
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   79
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Precio por distancia|N|S|||spara1|precioxdistancia|###,##0.0000||"
            Text            =   "3"
            Top             =   390
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "Precio por tiempo de espera"
            Height          =   195
            Index           =   79
            Left            =   240
            TabIndex        =   260
            Top             =   840
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "Precio por distancia"
            Height          =   195
            Index           =   78
            Left            =   240
            TabIndex        =   259
            Top             =   420
            Width           =   1515
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   78
         Left            =   7830
         MaxLength       =   3
         TabIndex        =   12
         Tag             =   "Cooperativa|N|N|||spara1|cooperativa|000||"
         Text            =   "Text1"
         Top             =   1890
         Width           =   615
      End
      Begin VB.Frame Frame15 
         Caption         =   "Facturación Equipamiento"
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
         Height          =   825
         Left            =   -74850
         TabIndex        =   253
         Top             =   5910
         Width           =   8655
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   77
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   254
            Tag             =   "Raiz Cta Socio Publi|T|S|||spara1|raiz_cta_soc_equip|||"
            Text            =   "3"
            Top             =   330
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Raíz Cuenta Socios"
            Height          =   195
            Index           =   72
            Left            =   150
            TabIndex        =   255
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.CheckBox chkHayrepar 
         Caption         =   "Realiza Reparaciones"
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Tag             =   "Reparaciones|N|N|||spara1|hayrepar|||"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Frame Frame14 
         Caption         =   "Alta Socios"
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
         Height          =   2715
         Left            =   -74850
         TabIndex        =   251
         Top             =   650
         Width           =   8655
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   76
            Left            =   2850
            MaxLength       =   10
            TabIndex        =   243
            Tag             =   "Raiz Cta Socio Altasi|T|S|||spara1|raiz_ctaaltasoc|||"
            Text            =   "3"
            Top             =   2160
            Width           =   1320
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   75
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   242
            Tag             =   "Importe Gastos Alta|N|S|||spara1|impgastoalta|###,###,##0.00||"
            Text            =   "3"
            Top             =   1530
            Width           =   1260
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   74
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   241
            Tag             =   "Importe Titulo Alta|N|S|||spara1|imptituloalta|###,###,##0.00||"
            Text            =   "3"
            Top             =   1140
            Width           =   1260
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   73
            Left            =   4200
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   247
            Text            =   "Text2"
            Top             =   750
            Width           =   4185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   73
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   240
            Tag             =   "Cta Gastos|N|S|||spara1|ctagastoalta|||"
            Text            =   "3"
            Top             =   750
            Width           =   1260
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   345
            Index           =   72
            Left            =   4200
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   245
            Text            =   "Text2"
            Top             =   360
            Width           =   4185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   72
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   239
            Tag             =   "Cta Título|N|S|||spara1|ctatituloalta|||"
            Text            =   "3"
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz Cuenta Socios"
            Height          =   195
            Index           =   67
            Left            =   180
            TabIndex        =   252
            Top             =   2130
            Width           =   1665
         End
         Begin VB.Label Label1 
            Caption         =   "Importe Reserva Legal Obligatoria"
            Height          =   195
            Index           =   66
            Left            =   180
            TabIndex        =   250
            Top             =   1530
            Width           =   2715
         End
         Begin VB.Label Label1 
            Caption         =   "Importe Aportación Capital Social"
            Height          =   195
            Index           =   65
            Left            =   180
            TabIndex        =   249
            Top             =   1170
            Width           =   2745
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta Reserva Legal Obligatoria"
            Height          =   255
            Index           =   10
            Left            =   180
            TabIndex        =   248
            Top             =   750
            Width           =   2415
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   2580
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   780
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta Aportación Capital Social"
            Height          =   255
            Index           =   9
            Left            =   180
            TabIndex        =   246
            Top             =   390
            Width           =   2385
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   2580
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   390
            Width           =   240
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Liquidación Socios"
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
         Height          =   2475
         Left            =   -74850
         TabIndex        =   213
         Top             =   3210
         Width           =   8655
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   71
            Left            =   3450
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   237
            Text            =   "Text2"
            Top             =   1980
            Width           =   4185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   71
            Left            =   2130
            MaxLength       =   10
            TabIndex        =   35
            Tag             =   "Cta liquidacion|N|S|||spara1|ctaliquidacion|||"
            Text            =   "3"
            Top             =   1980
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   65
            Left            =   2130
            MaxLength       =   10
            TabIndex        =   31
            Tag             =   "Raiz Cta Socio Publi|T|S|||spara1|raiz_cta_soc_liqui|||"
            Text            =   "3"
            Top             =   450
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   49
            Left            =   2130
            MaxLength       =   5
            TabIndex        =   33
            Tag             =   "Nº Contabilidad|N|S|||spara1|porreten|||"
            Text            =   "3"
            Top             =   1230
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   47
            Left            =   2130
            MaxLength       =   10
            TabIndex        =   32
            Tag             =   "Raíz Cta retencion|N|S|||spara1|raiz_ctareten|||"
            Text            =   "3"
            Top             =   840
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   46
            Left            =   2130
            MaxLength       =   2
            TabIndex        =   34
            Tag             =   "REA|N|S|0||spara1|iva_rea|||"
            Text            =   "Text1"
            Top             =   1590
            Width           =   1185
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   46
            Left            =   3450
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   214
            Text            =   "Text2"
            Top             =   1590
            Width           =   4185
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta Base"
            Height          =   255
            Index           =   7
            Left            =   150
            TabIndex        =   238
            Top             =   1980
            Width           =   1425
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1800
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   2010
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Raíz Cuenta Socios"
            Height          =   195
            Index           =   68
            Left            =   120
            TabIndex        =   218
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "Porcentaje de Retención"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   217
            Top             =   1260
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Raíz Cuenta Retención"
            Height          =   195
            Index           =   48
            Left            =   120
            TabIndex        =   216
            Top             =   870
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "IVA Liquidación"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   52
            Left            =   150
            TabIndex        =   215
            Top             =   1620
            Width           =   1365
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   39
            Left            =   1800
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1590
            Width           =   240
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Publicidad"
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
         Height          =   2385
         Left            =   -74880
         TabIndex        =   208
         Top             =   4350
         Width           =   8955
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   60
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   75
            Tag             =   "Raiz Cta Socio Publi|T|S|||spara1|raiz_cta_soc_publi|||"
            Text            =   "3"
            Top             =   1140
            Width           =   1260
         End
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   59
            Left            =   2040
            TabIndex        =   74
            Tag             =   "Concepto Facturación Publicidad |T|S|||spara1|confactupubli|||"
            Text            =   "Text1 "
            Top             =   720
            Width           =   6405
         End
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   32
            Left            =   2040
            MaxLength       =   16
            TabIndex        =   73
            Tag             =   "Recar |T|S|||spara1|codartictel|||"
            Text            =   "Text1 "
            Top             =   297
            Width           =   1545
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   32
            Left            =   3660
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   209
            Text            =   "Text2"
            Top             =   300
            Width           =   4785
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz Cuenta Socios"
            Height          =   195
            Index           =   61
            Left            =   240
            TabIndex        =   212
            Top             =   1170
            Width           =   1665
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1680
            Tag             =   "-1"
            ToolTipText     =   "Ver observaciones"
            Top             =   750
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto Factura"
            Height          =   195
            Index           =   69
            Left            =   240
            TabIndex        =   211
            Top             =   750
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Artículo a facturar"
            Height          =   195
            Index           =   32
            Left            =   240
            TabIndex        =   210
            Top             =   360
            Width           =   1380
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   32
            Left            =   1680
            Tag             =   "-1"
            ToolTipText     =   "Buscar artículo"
            Top             =   337
            Width           =   240
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Concepto Facturación Publicidad"
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
         Left            =   -69870
         TabIndex        =   206
         Top             =   5880
         Width           =   3375
         Begin VB.CheckBox chkTicketsAgrupads 
            Caption         =   "Contabilizar ticket TPV agrupados"
            Height          =   375
            Left            =   360
            TabIndex        =   207
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
         Left            =   -73980
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
         Index           =   25
         Left            =   -74820
         MaxLength       =   3
         TabIndex        =   198
         Tag             =   "Actividad|N|S|0||spara1|defactividad|000||"
         Text            =   "Tex"
         Top             =   4710
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   26
         Left            =   -69660
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   197
         Text            =   "Text2"
         Top             =   4710
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   26
         Left            =   -70500
         MaxLength       =   3
         TabIndex        =   196
         Tag             =   "Envio|N|S|0|999|spara1|defenvio|000||"
         Text            =   "Tex"
         Top             =   4710
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   27
         Left            =   -73980
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   195
         Text            =   "Text2"
         Top             =   5310
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   27
         Left            =   -74820
         TabIndex        =   194
         Tag             =   "Zona|N|S|0|999|spara1|defzona|000||"
         Text            =   "Tex"
         Top             =   5310
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   29
         Left            =   -73980
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
         Index           =   29
         Left            =   -74820
         MaxLength       =   3
         TabIndex        =   192
         Tag             =   "Situacion|N|S|0|999|spara1|defstituacion|000||"
         Text            =   "Tex"
         Top             =   6030
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   30
         Left            =   -69660
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   191
         Text            =   "Text2"
         Top             =   6030
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   30
         Left            =   -70500
         MaxLength       =   3
         TabIndex        =   190
         Tag             =   "Tarifa|N|S|0|999|spara1|deftarifa|||"
         Text            =   "Tex"
         Top             =   6030
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   31
         Left            =   -69660
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   189
         Text            =   "Text2"
         Top             =   5310
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   31
         Left            =   -70500
         MaxLength       =   3
         TabIndex        =   188
         Tag             =   "Agente|N|S|0|999|spara1|defagente|000||"
         Text            =   "Tex"
         Top             =   5310
         Width           =   735
      End
      Begin VB.Frame Frame13 
         Caption         =   "Garantia de reparación"
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
         Left            =   360
         TabIndex        =   182
         Top             =   5760
         Width           =   3375
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   58
            Left            =   960
            MaxLength       =   4
            TabIndex        =   10
            Tag             =   "Dias de garantia de Reparacion|N|S|0|9999|spara1|diasgaranrepa|||"
            Text            =   "Text1"
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dias"
            Height          =   255
            Index           =   62
            Left            =   120
            TabIndex        =   183
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.ComboBox cboCreaTarifa 
         Height          =   315
         ItemData        =   "frmConfParamAplic.frx":00B4
         Left            =   1920
         List            =   "frmConfParamAplic.frx":00C1
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo Descuentos|N|N|||spara1|creatarifart|||"
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ComboBox CboModAnalitica 
         Height          =   315
         Left            =   -73200
         Style           =   2  'Dropdown List
         TabIndex        =   178
         Tag             =   "Modo analítica|N|N|0|9|spara1|modanalitica|||"
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   320
         Index           =   50
         Left            =   -72480
         MaxLength       =   2
         TabIndex        =   169
         Tag             =   "NºConta|N|S|1|99|spara1|conta_B|||"
         Text            =   "Text1"
         Top             =   1000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   52
         Left            =   -71760
         MaxLength       =   2
         TabIndex        =   56
         Tag             =   "IVAexento|N|S|0||spara1|IvaIntracom|||"
         Text            =   "Text1"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   52
         Left            =   -71040
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   167
         Text            =   "Text2"
         Top             =   2400
         Width           =   3105
      End
      Begin VB.ComboBox cboObsFactura 
         Height          =   315
         Left            =   -69480
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Tag             =   "Orden Descuentos|N|S|||spara1|obsfactura|||"
         Top             =   1000
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   48
         Left            =   -71040
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   161
         Text            =   "Text2"
         Top             =   1960
         Width           =   2985
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   48
         Left            =   -71760
         MaxLength       =   2
         TabIndex        =   55
         Tag             =   "IVAexento|N|S|0||spara1|ivaexento|||"
         Text            =   "Text1"
         Top             =   1960
         Width           =   615
      End
      Begin VB.Frame Frame8 
         Caption         =   "IVA 's"
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
         Height          =   3075
         Left            =   -74940
         TabIndex        =   144
         Top             =   2970
         Width           =   9105
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   86
            Left            =   5670
            MaxLength       =   2
            TabIndex        =   66
            Tag             =   "IVRE Ant 2|N|S|0|99|spara1|ivaant2eq|||"
            Text            =   "Text1"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   86
            Left            =   6330
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   271
            Text            =   "Text2"
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   85
            Left            =   5670
            MaxLength       =   2
            TabIndex        =   65
            Tag             =   "IVA Ant 2|N|S|0|99|spara1|ivaant2|||"
            Text            =   "Text1"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   85
            Left            =   6330
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   268
            Text            =   "Text2"
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   84
            Left            =   5670
            MaxLength       =   2
            TabIndex        =   64
            Tag             =   "IVRE Ant 1|N|S|0|99|spara1|ivaant1eq|||"
            Text            =   "Text1"
            Top             =   900
            Width           =   615
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   84
            Left            =   6330
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   266
            Text            =   "Text2"
            Top             =   900
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   83
            Left            =   5670
            MaxLength       =   2
            TabIndex        =   63
            Tag             =   "IVA Ant 1|N|S|0|99|spara1|ivaant1|||"
            Text            =   "Text1"
            Top             =   540
            Width           =   615
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   83
            Left            =   6330
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   263
            Text            =   "Text2"
            Top             =   540
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   42
            Left            =   1830
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   153
            Text            =   "Text2"
            Top             =   2670
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   45
            Left            =   1830
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   152
            Text            =   "Text2"
            Top             =   2310
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   41
            Left            =   1830
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   151
            Text            =   "Text2"
            Top             =   1770
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   44
            Left            =   1830
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   150
            Text            =   "Text2"
            Top             =   1410
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   40
            Left            =   1830
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   149
            Text            =   "Text2"
            Top             =   900
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   43
            Left            =   1830
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   148
            Text            =   "Text2"
            Top             =   540
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   42
            Left            =   1170
            MaxLength       =   2
            TabIndex        =   62
            Tag             =   "IVRE3|N|S|0|99|spara1|ivare3eq|||"
            Text            =   "Text1"
            Top             =   2670
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   41
            Left            =   1170
            MaxLength       =   2
            TabIndex        =   60
            Tag             =   "IVRE2|N|S|0|99|spara1|ivare2eq|||"
            Text            =   "Text1"
            Top             =   1770
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   40
            Left            =   1170
            MaxLength       =   2
            TabIndex        =   58
            Tag             =   "IVRE1|N|S|0|99|spara1|ivare1eq|||"
            Text            =   "Text1"
            Top             =   900
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   43
            Left            =   1170
            MaxLength       =   2
            TabIndex        =   57
            Tag             =   "IVA1|N|S|0|99|spara1|ivare1|||"
            Text            =   "Text1"
            Top             =   540
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   44
            Left            =   1170
            MaxLength       =   2
            TabIndex        =   59
            Tag             =   "IVA2|N|S|0|99|spara1|ivare2|||"
            Text            =   "Text1"
            Top             =   1410
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   45
            Left            =   1170
            MaxLength       =   2
            TabIndex        =   61
            Tag             =   "IVA3|N|S|0|99|spara1|ivare3|||"
            Text            =   "Text1"
            Top             =   2310
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   5370
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1830
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
            Height          =   255
            Index           =   14
            Left            =   4800
            TabIndex        =   272
            Top             =   1800
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   5370
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1470
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Reducido Antiguo"
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
            Height          =   195
            Index           =   83
            Left            =   4650
            TabIndex        =   270
            Top             =   1230
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
            Height          =   255
            Index           =   13
            Left            =   4800
            TabIndex        =   269
            Top             =   1470
            Width           =   555
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   5400
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   930
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
            Height          =   255
            Index           =   12
            Left            =   4800
            TabIndex        =   267
            Top             =   900
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   5400
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   570
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "General Antiguo"
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
            Height          =   195
            Index           =   82
            Left            =   4650
            TabIndex        =   265
            Top             =   330
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
            Height          =   255
            Index           =   11
            Left            =   4770
            TabIndex        =   264
            Top             =   570
            Width           =   555
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
            Height          =   255
            Index           =   5
            Left            =   330
            TabIndex        =   159
            Top             =   2670
            Width           =   435
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
            Height          =   255
            Index           =   4
            Left            =   300
            TabIndex        =   158
            Top             =   2310
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
            Height          =   255
            Index           =   3
            Left            =   300
            TabIndex        =   157
            Top             =   1770
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
            Height          =   255
            Index           =   2
            Left            =   300
            TabIndex        =   156
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
            Height          =   255
            Index           =   1
            Left            =   300
            TabIndex        =   155
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   270
            TabIndex        =   154
            Top             =   570
            Width           =   555
         End
         Begin VB.Label Label1 
            Caption         =   "General"
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
            Height          =   195
            Index           =   45
            Left            =   150
            TabIndex        =   147
            Top             =   330
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Reducido"
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
            Height          =   195
            Index           =   44
            Left            =   150
            TabIndex        =   146
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Super-Reducido"
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
            Height          =   195
            Index           =   46
            Left            =   120
            TabIndex        =   145
            Top             =   2070
            Width           =   1380
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   36
            Left            =   900
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   570
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   33
            Left            =   900
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   930
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   37
            Left            =   870
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1440
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   34
            Left            =   870
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1800
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   38
            Left            =   870
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   2310
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   35
            Left            =   870
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   2700
            Width           =   240
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   3360
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   131
         Text            =   "Text2"
         Top             =   1320
         Width           =   4065
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cuotas"
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
         Height          =   3645
         Left            =   -74880
         TabIndex        =   128
         Top             =   650
         Width           =   8955
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   69
            Left            =   7800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   232
            Text            =   "Text3"
            Top             =   2280
            Width           =   915
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   67
            Left            =   7800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   230
            Text            =   "Text3"
            Top             =   1875
            Width           =   915
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   70
            Left            =   7800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   229
            Text            =   "Text3"
            Top             =   1065
            Width           =   915
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   61
            Left            =   7800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   228
            Text            =   "Text3"
            Top             =   1470
            Width           =   915
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   28
            Left            =   7800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   227
            Text            =   "Text3"
            Top             =   660
            Width           =   915
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   70
            Left            =   3810
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   224
            Text            =   "Text2"
            Top             =   1065
            Width           =   3945
         End
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   70
            Left            =   2400
            TabIndex        =   68
            Tag             =   "Con Chofer|T|S|||spara1|artcuotaconchof|||"
            Text            =   "Tex"
            Top             =   1065
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   69
            Left            =   2400
            TabIndex        =   71
            Tag             =   "Servicios|T|S|||spara1|artservcuotas|||"
            Text            =   "Tex"
            Top             =   2280
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   69
            Left            =   3810
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   222
            Text            =   "Text2"
            Top             =   2280
            Width           =   3945
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   67
            Left            =   3810
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   220
            Text            =   "Text2"
            Top             =   1875
            Width           =   3945
         End
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   67
            Left            =   2400
            TabIndex        =   70
            Tag             =   "Alquiler|T|S|||spara1|artalquiler|||"
            Text            =   "Tex"
            Top             =   1875
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   66
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   72
            Tag             =   "Raiz Cta Socio Publi|T|S|||spara1|raiz_ctaclien_soc|||"
            Text            =   "3"
            Top             =   2850
            Width           =   1320
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   61
            Left            =   3810
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   185
            Text            =   "Text2"
            Top             =   1470
            Width           =   3945
         End
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   61
            Left            =   2400
            TabIndex        =   69
            Tag             =   "Cuota Ext|T|S|||spara1|artcuotaext|||"
            Text            =   "Tex"
            Top             =   1470
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   28
            Left            =   2400
            TabIndex        =   67
            Tag             =   "Sin Chofer|T|S|||spara1|artcuotasinchof|||"
            Text            =   "Tex"
            Top             =   660
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   28
            Left            =   3810
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   129
            Text            =   "Text2"
            Top             =   660
            Width           =   3945
         End
         Begin VB.Label Label1 
            Caption         =   "Precio"
            Height          =   255
            Index           =   75
            Left            =   7830
            TabIndex        =   231
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Artículo"
            Height          =   255
            Index           =   64
            Left            =   2430
            TabIndex        =   226
            Top             =   330
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Cuota Normal Con Chofer"
            Height          =   195
            Index           =   74
            Left            =   180
            TabIndex        =   225
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
            Height          =   195
            Index           =   73
            Left            =   180
            TabIndex        =   223
            Top             =   2310
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Alquiler Equipos"
            Height          =   195
            Index           =   71
            Left            =   180
            TabIndex        =   221
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
            Height          =   195
            Index           =   70
            Left            =   180
            TabIndex        =   219
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
            Height          =   195
            Index           =   63
            Left            =   180
            TabIndex        =   184
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
            Height          =   195
            Index           =   28
            Left            =   180
            TabIndex        =   130
            Top             =   720
            Width           =   2025
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   -72720
         MaxLength       =   30
         TabIndex        =   50
         Tag             =   "Servidor Contabilidad|T|S|||spara1|serconta|||"
         Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
         Top             =   555
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   -66600
         MaxLength       =   2
         TabIndex        =   53
         Tag             =   "Nº Contabilidad|N|S|||spara1|numconta|||"
         Text            =   "3"
         Top             =   555
         Width           =   300
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   -70440
         MaxLength       =   20
         TabIndex        =   51
         Tag             =   "Usuario Contabilidad|T|S|||spara1|usuconta|||"
         Text            =   "3wwwwwwwwwwwwwwwwwww"
         Top             =   555
         Width           =   900
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   20
         Left            =   -68880
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   52
         Tag             =   "Password Contabilidad|T|S|||spara1|pasconta|||"
         Text            =   "3"
         Top             =   555
         Width           =   1140
      End
      Begin VB.ComboBox cboOrdenDtos 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Orden Descuentos|N|N|||spara1|ordendto|||"
         Top             =   1860
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Compras"
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
         Left            =   -74850
         TabIndex        =   117
         Top             =   630
         Width           =   5475
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   19
            Left            =   4440
            MaxLength       =   2
            TabIndex        =   28
            Tag             =   "Mes a no girar|N|S|0|12|spara1|mesnogir|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   18
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   27
            Tag             =   "Dia 3 de pago compras|N|S|0|31|spara1|diapago3|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   17
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   26
            Tag             =   "Dia 2 de pago compras|N|S|0|31|spara1|diapago2|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   16
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   25
            Tag             =   "Dia 1 de pago compras|N|S|0|31|spara1|diapago1|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Mes a no girar"
            Height          =   255
            Index           =   13
            Left            =   3360
            TabIndex        =   119
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Días de pago"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   118
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Facturación Clientes"
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
         Height          =   1425
         Left            =   -74850
         TabIndex        =   114
         Top             =   1680
         Width           =   8655
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   53
            Left            =   3870
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   186
            Text            =   "Text2"
            Top             =   750
            Width           =   4665
         End
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   53
            Left            =   2070
            MaxLength       =   16
            TabIndex        =   30
            Tag             =   "Artículo Gastos |T|S|||spara1|ArtReciclado|||"
            Text            =   "Text1 "
            Top             =   750
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   15
            Left            =   2070
            MaxLength       =   16
            TabIndex        =   29
            Tag             =   "Artículo Servicios |T|S|||spara1|codartid|||"
            Text            =   "Text1 "
            Top             =   327
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   15
            Left            =   3870
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   115
            Text            =   "Text2"
            Top             =   330
            Width           =   4665
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   53
            Left            =   1710
            Tag             =   "-1"
            ToolTipText     =   "Buscar artículo"
            Top             =   795
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Artículo Gtos Admon"
            Height          =   195
            Index           =   54
            Left            =   150
            TabIndex        =   187
            Top             =   810
            Width           =   1500
         End
         Begin VB.Label Label1 
            Caption         =   "Artículo Servicios"
            Height          =   255
            Index           =   10
            Left            =   150
            TabIndex        =   116
            Top             =   360
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   1710
            Tag             =   "-1"
            ToolTipText     =   "Buscar artículo"
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame FrameSoporte 
         ForeColor       =   &H00972E0B&
         Height          =   1635
         Left            =   -74760
         TabIndex        =   109
         Top             =   3840
         Width           =   8355
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   12
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   46
            Tag             =   "Web de Soporte|T|S|||spara1|websoporte|||"
            Text            =   "3"
            Top             =   300
            Width           =   6060
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   47
            Tag             =   "Mail de Soporte|T|S|||spara1|mailsoporte|||"
            Text            =   "3"
            Top             =   690
            Width           =   6060
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   48
            Tag             =   "Version Web|T|S|||spara1|webversion|||"
            Text            =   "3"
            Top             =   1080
            Width           =   6060
         End
         Begin VB.Label Label1 
            Caption         =   "Web de soporte"
            Height          =   195
            Index           =   9
            Left            =   300
            TabIndex        =   113
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "Mail soporte"
            Height          =   195
            Index           =   12
            Left            =   300
            TabIndex        =   112
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Web check version"
            Height          =   195
            Index           =   16
            Left            =   300
            TabIndex        =   111
            Top             =   1140
            Width           =   1500
         End
         Begin VB.Label Label8 
            Caption         =   "Soporte"
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
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   110
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.Frame FrameEMail 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   103
         Top             =   720
         Width           =   8355
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   57
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   45
            Tag             =   "LanzaMailOutlook|T|S|||spara1|arigesmail|||"
            Text            =   "3"
            Top             =   2400
            Width           =   1620
         End
         Begin VB.CheckBox chkOutlook 
            Caption         =   "Enviar desde Outlook"
            Height          =   375
            Left            =   5880
            TabIndex        =   179
            Tag             =   "Outlook|N|N|||spara1|EnvioDesdeOutlook|||"
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   1440
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   44
            Tag             =   "Password SMTP|T|S|||spara1|smtppass|||"
            Text            =   "3"
            Top             =   1560
            Width           =   4260
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   43
            Tag             =   "Usuario SMTP|T|S|||spara1|smtpuser|||"
            Text            =   "3"
            Top             =   1180
            Width           =   4260
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   42
            Tag             =   "Servidor SMTP|T|S|||spara1|smtphost|||"
            Text            =   "3"
            Top             =   800
            Width           =   5700
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   8
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   41
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
            Height          =   195
            Index           =   60
            Left            =   240
            TabIndex        =   181
            Top             =   2460
            Width           =   2280
         End
         Begin VB.Label Label8 
            Caption         =   "Envio E-Mail"
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
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   108
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   23
            Left            =   300
            TabIndex        =   107
            Top             =   1620
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   22
            Left            =   300
            TabIndex        =   106
            Top             =   1260
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
            Height          =   195
            Index           =   21
            Left            =   300
            TabIndex        =   105
            Top             =   840
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
            Height          =   195
            Index           =   20
            Left            =   300
            TabIndex        =   104
            Top             =   480
            Width           =   900
         End
      End
      Begin VB.ComboBox cboTipodtos 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Tipo Descuentos|N|N|||spara1|tipodtos|||"
         Top             =   2340
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   1
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "Código Tarifa PVP|N|N|||spara1|codtarif|000||"
         Text            =   "Text1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.Frame FramePrecioKm 
         Caption         =   "Precio Km"
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
         Height          =   1070
         Left            =   360
         TabIndex        =   94
         Top             =   3360
         Width           =   3375
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   2
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Precio Km desplaz. Clientes|N|S|0|9999.0000|spara1|preukmcl|#,##0.0000||"
            Text            =   "Text1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   3
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Precio Km desplaz. Técnicos|N|S|0|9999.0000|spara1|preukmtc|#,##0.0000||"
            Text            =   "Text1"
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Desplazamiento Clientes"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   96
            Top             =   255
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Desplazamiento Técnicos"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   95
            Top             =   660
            Width           =   1935
         End
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   4
         Left            =   3360
         MaxLength       =   35
         TabIndex        =   0
         Tag             =   "Nombre Director Gerente|T|S|||spara1|nomgeren|||"
         Text            =   "Text1"
         Top             =   540
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   5
         Left            =   3360
         MaxLength       =   35
         TabIndex        =   1
         Tag             =   "Nombre responsable Admon|T|S|||spara1|nomadmin|||"
         Text            =   "Text1"
         Top             =   900
         Width           =   4095
      End
      Begin VB.Frame FrameOpciones 
         Caption         =   "Opciones"
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
         Height          =   4575
         Left            =   4080
         TabIndex        =   93
         Top             =   2670
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CheckBox ChkDtoxCantidad 
            Caption         =   "Hay Dtos por cantidad"
            Height          =   375
            Left            =   3000
            TabIndex        =   176
            Tag             =   "Hay Dtos por cantidad|N|N|||spara1|dtoxcanti|||"
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox chkMataPrimaPorcen 
            Caption         =   "Materia prima como porcentaje"
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Tag             =   "Descriptores|N|N|||spara1|compoporcen|||"
            Top             =   4080
            Width           =   2775
         End
         Begin VB.CheckBox chkDescriptores 
            Caption         =   "Usa descriptores especiales"
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Tag             =   "Descriptores|N|N|||spara1|descriptores|||"
            Top             =   3720
            Width           =   2775
         End
         Begin VB.CheckBox chkProduccion 
            Caption         =   "Tiene produccion"
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Tag             =   "Tiene produccion|N|N|||spara1|produccion|||"
            Top             =   3360
            Width           =   2775
         End
         Begin VB.CheckBox chkHayServicio 
            Caption         =   "Hay Servicios"
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Tag             =   "Hay Servicios|N|N|||spara1|hayservicio|||"
            Top             =   1590
            Width           =   2175
         End
         Begin VB.CheckBox chkCajacomp 
            Caption         =   "Cajas completas precios"
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Tag             =   "Cajas Completas Precios|N|N|||spara1|cajacomp|||"
            Top             =   240
            Width           =   2175
         End
         Begin VB.CheckBox chkHaymante 
            Caption         =   "Realiza Mantenimientos"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Tag             =   "Mantenimientos|N|N|||spara1|haymante|||"
            Top             =   585
            Width           =   2175
         End
         Begin VB.CheckBox chkHayfrecu 
            Caption         =   "Hay Frecuencias"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Tag             =   "Hay Frecuencias|N|N|||spara1|hayfrecu|||"
            Top             =   1275
            Width           =   2175
         End
         Begin VB.CheckBox chkHaydepar 
            Caption         =   "Tiene Departamentos (o Dirección)"
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Tag             =   "Departamento/Direc.|N|N|||spara1|haydepar|||"
            Top             =   2280
            Width           =   2775
         End
         Begin VB.CheckBox chkctrstock 
            Caption         =   "Control de Stock estricto"
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Tag             =   "Control de Stock|N|N|||spara1|ctrstock|||"
            Top             =   2640
            Width           =   2775
         End
         Begin VB.CheckBox chkInventar 
            Caption         =   "Realiza Inventario por Proveedor"
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Tag             =   "Inventarios por Proveedor|N|N|||spara1|inventar|||"
            Top             =   3000
            Width           =   2775
         End
      End
      Begin VB.Frame FrameDiasMante 
         Caption         =   "Días Reparación"
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
         Height          =   1095
         Left            =   360
         TabIndex        =   90
         Top             =   4560
         Width           =   3375
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   6
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   9
            Tag             =   "Dias Repar. sin Mantenimiento|N|N|0|9999|spara1|diasnoman|||"
            Text            =   "Text"
            Top             =   680
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   7
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   8
            Tag             =   "Dias Repar. con Mantenimiento|N|N|0|9999|spara1|diassiman|||"
            Text            =   "Text1"
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Sin Mantenimiento"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   92
            Top             =   675
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Con Mantenimiento"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   91
            Top             =   300
            Width           =   1455
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   4260
         MaxLength       =   15
         TabIndex        =   101
         Tag             =   "Código Parámetros Aplic|N|N|||spara1|codigo||S|"
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
         Left            =   -74880
         TabIndex        =   132
         Top             =   1080
         Width           =   8535
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   76
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
            TabIndex        =   77
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
            TabIndex        =   78
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
            TabIndex        =   79
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
            TabIndex        =   80
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   143
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Pedidos proveedores"
            Height          =   195
            Index           =   34
            Left            =   4680
            TabIndex        =   142
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Albaranes clientes"
            Height          =   195
            Index           =   35
            Left            =   2040
            TabIndex        =   141
            Top             =   765
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Albaranes proveedores"
            Height          =   195
            Index           =   36
            Left            =   4680
            TabIndex        =   140
            Top             =   765
            Width           =   1635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Mantenimientos"
            Height          =   195
            Index           =   37
            Left            =   2040
            TabIndex        =   139
            Top             =   1320
            Width           =   1110
         End
         Begin VB.Label Label1 
            Caption         =   "Reparaciones"
            Height          =   195
            Index           =   38
            Left            =   2040
            TabIndex        =   138
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Avisos "
            Height          =   195
            Index           =   39
            Left            =   2040
            TabIndex        =   137
            Top             =   2040
            Width           =   510
         End
         Begin VB.Label Label1 
            Caption         =   "Dias desde la fecha"
            Height          =   195
            Index           =   40
            Left            =   120
            TabIndex        =   136
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
            TabIndex        =   135
            Top             =   1320
            Width           =   2955
         End
         Begin VB.Label Label1 
            Caption         =   "Pendientes de reparar sin motivo de reparación"
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
            TabIndex        =   134
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
            TabIndex        =   133
            Top             =   2040
            Width           =   2955
         End
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   62
         Left            =   -73890
         TabIndex        =   233
         Text            =   "Tex"
         Top             =   1050
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   63
         Left            =   -73830
         TabIndex        =   234
         Text            =   "Tex"
         Top             =   1020
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   64
         Left            =   -73500
         TabIndex        =   235
         Text            =   "Tex"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   68
         Left            =   -73770
         TabIndex        =   236
         Text            =   "Tex"
         Top             =   810
         Width           =   1335
      End
      Begin VB.Frame Frame9 
         Caption         =   "Aportación en facturas"
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
         TabIndex        =   163
         Top             =   6000
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   40
            Tag             =   "Cta aportacion|N|S|||spara1|ctaaportacion|||"
            Text            =   "3"
            Top             =   240
            Width           =   1260
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   51
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   164
            Text            =   "Text2"
            Top             =   240
            Width           =   4185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   42
            Left            =   1800
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta"
            Height          =   255
            Index           =   8
            Left            =   1080
            TabIndex        =   165
            Top             =   240
            Width           =   615
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
         TabIndex        =   125
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
            TabIndex        =   126
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
            TabIndex        =   39
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
            TabIndex        =   127
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Portes"
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
         Height          =   975
         Left            =   -74880
         TabIndex        =   171
         Top             =   6030
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   56
            Left            =   6720
            MaxLength       =   16
            TabIndex        =   38
            Tag             =   "R|N|S||10000|spara1|impminped|#,##0.00||"
            Text            =   "Text1 "
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   55
            Left            =   6720
            MaxLength       =   16
            TabIndex        =   37
            Tag             =   "i |N|S|||spara1|abonokilos|#,##0.0000||"
            Text            =   "Text1 "
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   54
            Left            =   840
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   173
            Text            =   "Text2"
            Top             =   600
            Width           =   3705
         End
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   54
            Left            =   1200
            MaxLength       =   16
            TabIndex        =   36
            Tag             =   "Reci. |T|S|||spara1|ArticuloPortes|||"
            Text            =   "Text1 "
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Importe minimo pedido"
            Height          =   195
            Index           =   57
            Left            =   4800
            TabIndex        =   175
            Top             =   600
            Width           =   1620
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Abono kilos"
            Height          =   195
            Index           =   56
            Left            =   4800
            TabIndex        =   174
            Top             =   300
            Width           =   1620
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   54
            Left            =   840
            Tag             =   "-1"
            ToolTipText     =   "Buscar artículo"
            Top             =   285
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Articulo"
            Height          =   195
            Index           =   55
            Left            =   120
            TabIndex        =   172
            Top             =   300
            Width           =   780
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Impresora Tarjetas"
         Enabled         =   0   'False
         Height          =   195
         Index           =   81
         Left            =   -74670
         TabIndex        =   262
         Top             =   3735
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Label Label1 
         Caption         =   "Path FacturaE"
         Height          =   195
         Index           =   80
         Left            =   -74460
         TabIndex        =   261
         Top             =   5745
         Width           =   2070
      End
      Begin VB.Label Label1 
         Caption         =   "Cooperativa"
         Height          =   255
         Index           =   77
         Left            =   6870
         TabIndex        =   257
         Top             =   1890
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Código Tarifa de PVP"
         Height          =   255
         Index           =   76
         Left            =   4320
         TabIndex        =   256
         Top             =   1860
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Agente"
         Height          =   255
         Index           =   31
         Left            =   -70500
         TabIndex        =   205
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Situación"
         Height          =   255
         Index           =   29
         Left            =   -74820
         TabIndex        =   204
         Top             =   5760
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   29
         Left            =   -74100
         ToolTipText     =   "Buscar situacion"
         Top             =   5760
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   31
         Left            =   -69900
         ToolTipText     =   "Buscar agente"
         Top             =   5040
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Actividad"
         Height          =   255
         Index           =   25
         Left            =   -74820
         TabIndex        =   203
         Top             =   4470
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Envio"
         Height          =   195
         Index           =   26
         Left            =   -70500
         TabIndex        =   202
         Top             =   4470
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Forma de Pago"
         Height          =   255
         Index           =   27
         Left            =   -74820
         TabIndex        =   201
         Top             =   5070
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Tarifa"
         Height          =   255
         Index           =   30
         Left            =   -70500
         TabIndex        =   200
         Top             =   5790
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   -70020
         ToolTipText     =   "Buscar forma de envio"
         Top             =   4470
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   -73620
         ToolTipText     =   "Buscar forma de pago"
         Top             =   5070
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   -74100
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   4470
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   30
         Left            =   -70020
         ToolTipText     =   "Buscar tarifa"
         Top             =   5790
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Crear tarifas"
         Height          =   255
         Index           =   59
         Left            =   360
         TabIndex        =   180
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Modo analítica"
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
         Height          =   195
         Index           =   58
         Left            =   -74760
         TabIndex        =   177
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Conta presupuestos *"
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   51
         Left            =   -74760
         TabIndex        =   170
         Top             =   1000
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "IVA intracomunitario"
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
         Height          =   195
         Index           =   47
         Left            =   -74760
         TabIndex        =   168
         Top             =   2400
         Width           =   1725
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   45
         Left            =   -72120
         Tag             =   "-1"
         ToolTipText     =   "Buscar I.V.A."
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Integ.  fras. Observaciones "
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
         Height          =   195
         Index           =   53
         Left            =   -71880
         TabIndex        =   166
         Top             =   1000
         Width           =   2385
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   41
         Left            =   -72120
         Tag             =   "-1"
         ToolTipText     =   "Buscar I.V.A."
         Top             =   1960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "IVA exento"
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
         Height          =   195
         Index           =   50
         Left            =   -74760
         TabIndex        =   162
         Top             =   1960
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Contabilidad"
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
         Height          =   195
         Index           =   49
         Left            =   -74760
         TabIndex        =   160
         Top             =   600
         Width           =   1185
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1965
         Tag             =   "-1"
         ToolTipText     =   "Buscar tarifa"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Servidor"
         Height          =   195
         Index           =   19
         Left            =   -73440
         TabIndex        =   124
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Nº conta"
         Height          =   195
         Index           =   18
         Left            =   -67560
         TabIndex        =   123
         Top             =   600
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   195
         Index           =   17
         Left            =   -71040
         TabIndex        =   122
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Pass."
         Height          =   195
         Index           =   15
         Left            =   -69360
         TabIndex        =   121
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Orden Descuentos"
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   120
         Top             =   1860
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Descuentos"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   100
         Top             =   2340
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Código Tarifa de PVP"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   99
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del Director Gerente"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   98
         Top             =   540
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre responsable Administración"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   97
         Top             =   900
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Index           =   6
         Left            =   3720
         TabIndex        =   102
         Top             =   1020
         Width           =   495
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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

Private WithEvents frmMtoArt As frmAlmArticulos
Attribute frmMtoArt.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1


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
Dim Indice As Byte


Dim PrimeraVez As Boolean
Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: Añadir
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
            vParamAplic.ControlStock = Me.chkctrstock.Value
            vParamAplic.InventarioxProv = Me.chkInventar.Value
            vParamAplic.NumSeries = Me.chkHaynserie.Value  'Hay Nº Serie en Compras?
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
            
            'Portes(FOntenas)
            vParamAplic.ArtPortes = Text1(54).Text
            vParamAplic.AbonoKilos = ComprobarCero(Text1(55).Text)
            vParamAplic.ImporteMinimo = ComprobarCero(Text1(56).Text)
            
            vParamAplic.ComponentePorcentaje = Me.chkMataPrimaPorcen.Value
            
            ' ---- [14/09/2009] (LAURA)
            vParamAplic.DtoxCantidad = Me.ChkDtoxCantidad.Value
            vParamAplic.CreaTarifasArticulo = Me.cboCreaTarifa.ItemData(cboCreaTarifa.ListIndex)
            
            ' ----
            
            ' ---- [19/10/2009] [LAURA]: añadir campo modo analitica
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
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 3   'Anyadir
        .Buttons(1).Image = 4   'Modificar
        .Buttons(4).Image = 15  'Salir
    End With
    
    'cargar iconos de busqueda
    For Each Im In Me.imgBuscar
        Im.Picture = frmPpal.imgListComun.ListImages(19).Picture
    Next
    'imgBuscar(1).Picture = frmPpal.imgListComun.ListImages(19).Picture
    'imgBuscar(15).Picture = frmPpal.imgListComun.ListImages(19).Picture
   '
   ' For NumRegElim = 24 To 42
   '     Me.imgBuscar(NumRegElim).Picture = frmPpal.imgListComun.ListImages(19).Picture
   ' Next NumRegElim
    
    Frame16.visible = (vParamAplic.Cooperativa <> 0)
    Frame16.Enabled = (vParamAplic.Cooperativa <> 0)


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
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
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

Private Sub frmE_DatoSeleccionado(CadenaSeleccion As String)
    'ENVIO
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago


    
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
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
        Case 15, 32, 53, 54, 28, 70 'cod. articulo
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
        Case 24 'forma de pago
            If Modo = 4 Then TerminaBloquear
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Indice = 24
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
            Indice = 27
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
        Case 4  'Cod. Forma de Pago
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmFP.Show vbModal
            Set frmFP = Nothing
            
            
        Case 31 'Código de Agente
            Set frmAC = New frmFacAgentesCom
            frmAC.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmAC.Show vbModal
            Set frmAC = Nothing
            
        Case 1, 30 'Código de Tarifa
            Me.imgBuscar(1).Tag = Index
            Set frmT = New frmFacTarifas
            frmT.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 29 'Código de Situación
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
             If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmS.Show vbModal
            Set frmS = Nothing
            
            
        Case 33 To 42, 45, 6, 7, 8, 9 To 12 'Todos los ivas y la Cta de retencion, y cuenta aportacion TERMINAL
            CadenaDesdeOtroForm = ""
                        
            BuscaBuscaGRid2 (Index <> 40 And Index <> 42 And Index <> 6 And Index <> 7 And Index <> 8)
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
        
            
            
    End Select
    PonerFoco Text1(Index)
End Sub


Private Sub BuscaBuscaGRid2(EsIVa As Boolean)


    Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        If EsIVa Then
            'Busco IVAS
            frmB.vCampos = "Código|tiposiva|codigiva|N||20·Denominacion|tiposiva|nombriva|T||70·"
            frmB.vTabla = "tiposiva"
            frmB.vTitulo = "Tipos de IVA"
        Else
                
            frmB.vCampos = "Código|cuentas|codmacta|T||20·Denominacion|cuentas|nommacta|T||70·"
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

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub


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
            
        Case 15, 32, 53, 54 'cod. artic
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic", "codartic", "Articulo")
            If Text2(Index).Text = "" Then Text1(Index).Text = ""
            
        Case 28, 61, 67, 69, 70 'cod. artic
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic", "codartic", "Articulo")
            If Text1(Index).Text <> "" Then Text3(Index).Text = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", Text1(Index).Text, "T")
            If Text2(Index).Text = "" Then Text1(Index).Text = ""
            
        
        Case 22 'nº conta
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
        
        Case 71, 72, 73
            '71=Cta liquidacion
            '72=cta de titulo de alta de socio
            '73=cta de gastos de alta de socio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conConta, "cuentas", "nommacta", "codmacta")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 49
            'pORCE RETENCION
            PonerFormatoDecimal Text1(48), 4
        Case 50, 58, 78
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
                MsgBox "La raiz de la cuenta de socios liquidación debe tener " & vEmpresa.DigitosNivelAnterior & " digitos.", vbExclamation
'                PonerFoco Text1(65)
            End If
            
            
        Case 47 'raiz cuenta retencion liquidacion
            If Len(Text1(Index).Text) <> vEmpresa.DigitosNivelAnterior Then
                MsgBox "La raiz de la cuenta de Retención socios debe tener " & vEmpresa.DigitosNivelAnterior & " digitos.", vbExclamation
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
        MsgBox "Raíz Cta retención o % retención vacios", vbExclamation
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
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
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
    Text2(15).Text = PonerNombreDeCod(Text1(15), conAri, "sartic", "nomartic", "codartic", "Artículos")
    Text2(32).Text = PonerNombreDeCod(Text1(32), conAri, "sartic", "nomartic", "codartic", "Artículos")
    
    ' Artículos para cuotas
    Text2(28).Text = PonerNombreDeCod(Text1(28), conAri, "sartic", "nomartic", "codartic", "Artículos")
    If Text1(28).Text <> "" Then
        Text3(28).Text = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", Text1(28).Text, "T")
        PonerFormatoDecimal Text3(28), 2
    End If
    Text2(70).Text = PonerNombreDeCod(Text1(28), conAri, "sartic", "nomartic", "codartic", "Artículos")
    If Text1(70).Text <> "" Then
        Text3(70).Text = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", Text1(70).Text, "T")
        PonerFormatoDecimal Text3(70), 2
    End If
    Text2(61).Text = PonerNombreDeCod(Text1(61), conAri, "sartic", "nomartic", "codartic", "Artículos")
    If Text1(61).Text <> "" Then
        Text3(61).Text = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", Text1(61).Text, "T")
        PonerFormatoDecimal Text3(61), 2
    End If
    Text2(67).Text = PonerNombreDeCod(Text1(67), conAri, "sartic", "nomartic", "codartic", "Artículos")
    If Text1(67).Text <> "" Then
        Text3(67).Text = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", Text1(67).Text, "T")
        PonerFormatoDecimal Text3(67), 2
    End If
    Text2(69).Text = PonerNombreDeCod(Text1(69), conAri, "sartic", "nomartic", "codartic", "Artículos")
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
   
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
      
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
    'Si estamos en Insertar además limpia los campos Text1 y bloquea la clave primaria
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
    
    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
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
    
    cboObsFactura.AddItem "Número factura"
    cboObsFactura.ItemData(cboObsFactura.NewIndex) = 1

    cboObsFactura.AddItem "Fecha integración"
    cboObsFactura.ItemData(cboObsFactura.NewIndex) = 2

End Sub



' ---- [19/10/2009] [LAURA]: añadir campo modo analitica
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
