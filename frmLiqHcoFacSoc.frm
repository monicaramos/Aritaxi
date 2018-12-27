VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLiqHcoFacSoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas Liquidación Socios"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   13290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   11010
      TabIndex        =   84
      Top             =   240
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3645
      TabIndex        =   82
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   83
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
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   80
      Top             =   30
      Width           =   3495
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   81
         Top             =   180
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Rectificativa"
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
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   49
      Top             =   705
      Width           =   12975
      Begin VB.CheckBox Check1 
         Caption         =   "FacturaE"
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
         Index           =   1
         Left            =   3720
         TabIndex        =   79
         Tag             =   "FacturaE|N|N|0|1|sfactusoc|exportada||N|"
         Top             =   390
         Width           =   1515
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
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   360
         Width           =   1005
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
         Index           =   24
         Left            =   11610
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "V Socio|N|N|0|999999|sfactusoc|numeruve|000000||"
         Text            =   "Text1"
         Top             =   270
         Width           =   990
      End
      Begin VB.TextBox Text1 
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
         Index           =   5
         Left            =   7140
         MaxLength       =   40
         TabIndex        =   5
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   270
         Width           =   4440
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
         Index           =   4
         Left            =   6240
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Cod. Socio|N|N|0|999999|sfactusoc|codsocio|000000|S|"
         Text            =   "Text1"
         Top             =   270
         Width           =   870
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
         Left            =   1410
         TabIndex        =   1
         Tag             =   "Tipo Factura|T|N|||sfactusoc|codtipom||S|"
         Text            =   "Text3"
         Top             =   375
         Width           =   735
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
         Index           =   2
         Left            =   2430
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||sfactusoc|fecfactu|dd/mm/yyyy|S|"
         Top             =   375
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         BeginProperty Font 
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
         Left            =   240
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Factura|N|S|||sfactusoc|numfactu|0000000|S|"
         Text            =   "Text1 7"
         Top             =   375
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
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
         Left            =   3720
         TabIndex        =   3
         Tag             =   "Contabilizado|N|N|0|1|sfactusoc|intconta||N|"
         Top             =   150
         Width           =   1665
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
         Index           =   0
         Left            =   5400
         TabIndex        =   53
         Top             =   270
         Width           =   525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   5970
         ToolTipText     =   "Buscar socio"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fact."
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
         Left            =   2460
         TabIndex        =   52
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
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
         Index           =   28
         Left            =   240
         TabIndex        =   51
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Fact."
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
         Left            =   1410
         TabIndex        =   50
         Top             =   120
         Width           =   1095
      End
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
      Height          =   4500
      Left            =   120
      TabIndex        =   31
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1605
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   7938
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmLiqHcoFacSoc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameCliente"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameFactura"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Llamadas"
      TabPicture(1)   =   "frmLiqHcoFacSoc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(48)"
      Tab(1).Control(1)=   "Data2"
      Tab(1).Control(2)=   "DataGrid1"
      Tab(1).Control(3)=   "txtAux3(2)"
      Tab(1).Control(4)=   "txtAux3(1)"
      Tab(1).Control(5)=   "txtAux3(0)"
      Tab(1).Control(6)=   "txtAux3(3)"
      Tab(1).Control(7)=   "txtAux3(4)"
      Tab(1).Control(8)=   "txtAux3(5)"
      Tab(1).Control(9)=   "txtAux3(6)"
      Tab(1).Control(10)=   "txtAux3(7)"
      Tab(1).Control(11)=   "txtAux3(8)"
      Tab(1).Control(12)=   "txtAux3(9)"
      Tab(1).Control(13)=   "txtAux3(10)"
      Tab(1).Control(14)=   "txtAux3(11)"
      Tab(1).Control(15)=   "txtAux3(12)"
      Tab(1).Control(16)=   "txtAux2(0)"
      Tab(1).Control(17)=   "FrameToolAux"
      Tab(1).ControlCount=   18
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   -74700
         TabIndex        =   85
         Top             =   360
         Width           =   1500
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Top             =   150
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
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
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtAux2 
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
         Left            =   -73560
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   76
         Top             =   4050
         Visible         =   0   'False
         Width           =   6285
      End
      Begin VB.Frame Frame3 
         Height          =   675
         Left            =   180
         TabIndex        =   73
         Top             =   1860
         Width           =   12645
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
            Height          =   375
            Index           =   23
            Left            =   1320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Tag             =   "Concepto|T|S|||sfactusoc|concepto|||"
            Text            =   "frmLiqHcoFacSoc.frx":0038
            Top             =   210
            Width           =   4560
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
            Left            =   7560
            MaxLength       =   3
            TabIndex        =   15
            Tag             =   "Servicios|N|S|||sfactusoc|numserv|###,##0|N|"
            Text            =   "Text1"
            Top             =   210
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto"
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
            Left            =   180
            TabIndex        =   75
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "SERVICIOS"
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
            Left            =   6090
            TabIndex        =   74
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   12
         Left            =   -63840
         MaxLength       =   15
         TabIndex        =   68
         Tag             =   "Importe |N|N|||sfactusoc_serv|impventa|###,##0.00|S|"
         Text            =   "importe"
         Top             =   2250
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   11
         Left            =   -64590
         MaxLength       =   10
         TabIndex        =   67
         Tag             =   "Telfono |T|N|||sfactusoc_serv|telefono|||"
         Text            =   "tfno"
         Top             =   2250
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   10
         Left            =   -65310
         MaxLength       =   30
         TabIndex        =   66
         Tag             =   "Ciudad|T|N|||sfactusoc_serv|ciudadre|||"
         Text            =   "ciudad"
         Top             =   2250
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Left            =   -66240
         MaxLength       =   10
         TabIndex        =   65
         Tag             =   "Puerllama|T|N|||sfactusoc_serv|puerllama|||"
         Text            =   "puerllama"
         Top             =   2250
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   8
         Left            =   -67140
         MaxLength       =   10
         TabIndex        =   64
         Tag             =   "Numllama|T|N|||sfactusoc_serv|numllama|||"
         Text            =   "numllama"
         Top             =   2250
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   7
         Left            =   -68100
         MaxLength       =   30
         TabIndex        =   63
         Tag             =   "Direccion|T|N|||sfactusoc_serv|dirllama|||"
         Text            =   "direccion"
         Top             =   2250
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   6
         Left            =   -69030
         MaxLength       =   30
         TabIndex        =   62
         Tag             =   "Nombre|T|N|||sfactusoc_serv|nomclien||S|"
         Text            =   "nombre"
         Top             =   2250
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   5
         Left            =   -69990
         MaxLength       =   6
         TabIndex        =   61
         Tag             =   "Cliente |N|N|||sfactusoc_serv|codclien|000000||"
         Text            =   "cliente"
         Top             =   2250
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   4
         Left            =   -70950
         MaxLength       =   8
         TabIndex        =   60
         Tag             =   "Hora |H|N|||sfactusoc_serv|hora|hh:mm:ss||"
         Text            =   "hora"
         Top             =   2250
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   3
         Left            =   -72180
         MaxLength       =   10
         TabIndex        =   59
         Tag             =   "Fecha|F|N|||sfactusoc_serv|fecha|dd/mm/yyyy|N|"
         Text            =   "fecha"
         Top             =   2250
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   0
         Left            =   -74520
         MaxLength       =   7
         TabIndex        =   69
         Tag             =   "Tipo Movimiento|T|N|||sfactusoc_serv|codtipom||S|"
         Text            =   "codtipoa"
         Top             =   2250
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   1
         Left            =   -73800
         MaxLength       =   15
         TabIndex        =   57
         Tag             =   "NºFactura |N|N|||sfactusoc_serv|numfactu|0000000|S|"
         Text            =   "numfactu"
         Top             =   2250
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Index           =   2
         Left            =   -72870
         MaxLength       =   30
         TabIndex        =   56
         Tag             =   "Fecha Factura|F|N|||sfactusoc_serv|fecfactu|dd/mm/yyyy|S|"
         Text            =   "fecfactu"
         Top             =   2250
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame FrameFactura 
         Height          =   1560
         Left            =   180
         TabIndex        =   40
         Top             =   2520
         Width           =   12645
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
            Index           =   25
            Left            =   120
            MaxLength       =   15
            TabIndex        =   21
            Tag             =   "Suplidos|N|S|||sfactusoc|suplidos|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   1635
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
            Index           =   21
            Left            =   6090
            MaxLength       =   5
            TabIndex        =   23
            Tag             =   "% Reten|N|S|0|99.90|sfactusoc|porcreten|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   675
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
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
            Left            =   6840
            MaxLength       =   15
            TabIndex        =   24
            Tag             =   "Importe Retencion|N|N|||sfactusoc|impreten|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   1695
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
            Left            =   4440
            MaxLength       =   15
            TabIndex        =   22
            Tag             =   "Base Retencion|N|N|||sfactusoc|basereten|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   1515
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
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
            Left            =   9210
            MaxLength       =   15
            TabIndex        =   25
            Tag             =   "Total Factura|N|N|||sfactusoc|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
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
            Left            =   6840
            MaxLength       =   15
            TabIndex        =   20
            Tag             =   "Importe IVA 1|N|N|||sfactusoc|impoiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   450
            Width           =   1695
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
            Left            =   6090
            MaxLength       =   5
            TabIndex        =   19
            Tag             =   "% IVA 1|N|S|0|99.90|sfactusoc|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   450
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
            Index           =   16
            Left            =   4440
            MaxLength       =   15
            TabIndex        =   18
            Tag             =   "Base Imponible 1|N|N|||sfactusoc|baseiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   450
            Width           =   1515
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
            Index           =   15
            Left            =   3780
            MaxLength       =   4
            TabIndex        =   17
            Tag             =   "Cod. IVA 1|N|S|0|9999|sfactusoc|codiiva1|0000|N|"
            Text            =   "Text1 7"
            Top             =   450
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
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
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   55
            Text            =   "Text1 7"
            Top             =   450
            Width           =   1545
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
            Left            =   120
            MaxLength       =   15
            TabIndex        =   16
            Tag             =   "Imp.Bruto|N|N|||sfactusoc|importel|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   450
            Width           =   1635
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
            Index           =   6
            Left            =   120
            TabIndex        =   87
            Top             =   810
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Base Retención"
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
            Left            =   4440
            TabIndex        =   72
            Top             =   810
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "% Ret"
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
            Left            =   6120
            TabIndex        =   71
            Top             =   810
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Importe Retención"
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
            Left            =   6870
            TabIndex        =   70
            Top             =   810
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
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
            Left            =   6840
            TabIndex        =   54
            Top             =   180
            Width           =   1425
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
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
            Index           =   41
            Left            =   6090
            TabIndex        =   48
            Top             =   180
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL FACTURA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   39
            Left            =   9210
            TabIndex        =   47
            Top             =   810
            Width           =   1830
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
            Left            =   8670
            TabIndex        =   46
            Top             =   1080
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
            TabIndex        =   45
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base impo. IVA"
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
            Left            =   4440
            TabIndex        =   44
            Top             =   180
            Width           =   1395
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
            Left            =   1920
            TabIndex        =   43
            Top             =   450
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Suma Importes"
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
            Left            =   2160
            TabIndex        =   42
            Top             =   180
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL Importes"
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
            TabIndex        =   41
            Top             =   180
            Width           =   1635
         End
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Socio"
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
         Height          =   1455
         Left            =   180
         TabIndex        =   32
         Top             =   360
         Width           =   12645
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
            Index           =   12
            Left            =   7560
            MaxLength       =   3
            TabIndex        =   12
            Tag             =   "Forma de Pago|N|N|0|999|sfactusoc|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   645
            Width           =   600
         End
         Begin VB.TextBox Text1 
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
            Index           =   11
            Left            =   7560
            MaxLength       =   30
            TabIndex        =   13
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1020
            Width           =   2445
         End
         Begin VB.TextBox Text1 
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
            Index           =   9
            Left            =   1305
            MaxLength       =   6
            TabIndex        =   10
            Text            =   "Text15"
            Top             =   1020
            Width           =   780
         End
         Begin VB.TextBox Text1 
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
            Index           =   10
            Left            =   2115
            MaxLength       =   30
            TabIndex        =   11
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1020
            Width           =   3765
         End
         Begin VB.TextBox Text1 
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
            Index           =   7
            Left            =   3555
            MaxLength       =   20
            TabIndex        =   8
            Text            =   "12345678911234567899"
            Top             =   285
            Width           =   2325
         End
         Begin VB.TextBox Text1 
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
            Left            =   1305
            MaxLength       =   15
            TabIndex        =   7
            Text            =   "123456789"
            Top             =   285
            Width           =   1290
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
            Left            =   8190
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   34
            Text            =   "Text2"
            Top             =   645
            Width           =   4245
         End
         Begin VB.TextBox Text1 
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
            Index           =   8
            Left            =   1305
            MaxLength       =   35
            TabIndex        =   9
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   645
            Width           =   4575
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   1065
            ToolTipText     =   "Buscar población"
            Top             =   1035
            Width           =   240
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
            Index           =   17
            Left            =   6090
            TabIndex        =   39
            Top             =   1020
            Width           =   975
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
            Index           =   16
            Left            =   120
            TabIndex        =   38
            Top             =   1020
            Width           =   915
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
            Index           =   19
            Left            =   2625
            TabIndex        =   37
            Top             =   285
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
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
            Left            =   120
            TabIndex        =   36
            Top             =   285
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
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
            Left            =   6090
            TabIndex        =   35
            Top             =   675
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   7290
            ToolTipText     =   "Buscar forma de pago"
            Top             =   660
            Width           =   240
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
            Index           =   7
            Left            =   120
            TabIndex        =   33
            Top             =   645
            Width           =   885
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmLiqHcoFacSoc.frx":003E
         Height          =   3030
         Left            =   -74670
         TabIndex        =   58
         Top             =   960
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   5345
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
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
      Begin MSAdodcLib.Adodc Data2 
         Height          =   330
         Left            =   -65880
         Top             =   780
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
         Caption         =   "Concepto"
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
         Index           =   48
         Left            =   -74670
         TabIndex        =   77
         Top             =   4050
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   6180
      Width           =   2175
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
         Height          =   240
         Left            =   240
         TabIndex        =   30
         Top             =   120
         Width           =   1755
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
      Left            =   12000
      TabIndex        =   27
      Top             =   6300
      Width           =   1135
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
      Left            =   10710
      TabIndex        =   26
      Top             =   6300
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
      Left            =   12000
      TabIndex        =   28
      Top             =   6300
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
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
      Begin VB.Menu mnRectifica 
         Caption         =   "&Rectificativa"
         Shortcut        =   ^R
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
Attribute VB_Name = "frmLiqHcoFacSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 420


Public DesdeFichaSocio As Boolean

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento
Public hcoCodSocio As String 'codigo de socio



Dim PrimeraVez As Boolean
Dim NombreTabla As String
Dim Ordenacion As String
Dim CadenaConsulta As String
Private kCampo As Integer
Private btnPrimero As Byte
Private HaDevueltoDatos As Boolean
Private Modo As Byte
Private BuscaChekc As String
Private CodTipoMov As String

Private cadFormula As String
Private cadParam As String
Private numParam As Byte

Private WithEvents frmC As frmGesSocios
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmHcoFacSocPre As frmLiqHcoFacSocPrev
Attribute frmHcoFacSocPre.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmBan As frmFacBancosPropios
Attribute frmBan.VB_VarHelpID = -1

Private NomTablaLineas As String 'Nombre de la Tabla de lineas


Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas


Dim cadB1 As String

Dim UnaVez As Boolean

Dim FPago As Currency
Dim Serv As Currency
Dim TotImp As Currency
Dim ImpIVA As Currency
Dim ImpRet As Currency
Dim TotFac As Currency

Dim cadban As String



Private Sub ComprobarDatosTotales()
Dim I As Byte

    For I = 13 To 14
        Text1(I).Text = ComprobarCero(Text1(I).Text)
    Next I
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean

    On Error GoTo EDatosOK

    DatosOk = False
    
    ComprobarDatosTotales
    
    'comprobamos datos OK de la tabla scafac
    b = CompForm(Me, 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function DatosOkLin() As Boolean
Dim b As Boolean

    On Error GoTo EDatosOK

    DatosOkLin = False

    RecalcularImportes txtAux3(12), False
    
    DatosOkLin = True

EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function ModificarFactura(Optional sqlLineas As String) As Boolean
'si se ha modificado la linea de slifac, añadir a la transaccion la modificación de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim Sql As String
Dim vFactu As CFacturaSoc
On Error GoTo EModFact

    bol = False
    conn.BeginTrans
    ConnConta.BeginTrans
    
    If sqlLineas <> "" Then
        'actualizar el importe de la linea modificada
        MenError = "Modificando lineas de Factura."
        conn.Execute sqlLineas
    End If
    
    'recalcular las bases imponibles x IVA
    MenError = "Recalcular importes IVA"
    If vParamAplic.Cooperativa = 1 Then
        bol = ActualizarDatosFactura
    Else
        bol = True
    End If
    If bol Then
        'modificamos la scafpc
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario(Me, 1)
        
        If bol Then
'            MenError = "Modificando albaranes de factura"
'            'modificar la tabla: scafpa
'            bol = ModificaAlbxFac

            If bol Then 'si se ha modificado la factura
                MenError = "Actualizando en Tesoreria"
                'y eliminar de tesoreria conta.spagop los registros de la factura

                'antes de Eliminar en las tablas de la Contabilidad
                Set vFactu = New CFacturaSoc
                bol = vFactu.LeerDatosLiq(Text1(4).Text, Text1(1).Text, Text1(0).Text, Text1(2).Text)

                If bol Then
                    'Modificamos la retencion en la sreten
                    MenError = "Modificando datos de retenciones"
                    bol = ActualizarRetencion(Text1(24).Text, vFactu, False)
                    
                    'Eliminar de la spagop
'[Monica]10/07/2012: tanto liquidaciones como rectificativas han de estar en la spagop
'                    If Text1(1).Text = "FLI" Then
                        If vParamAplic.ContabilidadNueva Then
                            Sql = " numserie = " & DBSet(SerieFraPro, "T")
                            Sql = Sql & " AND codmacta='" & vFactu.CtaSocio & "' AND numfactu='" & ObtenerLetraSerie(vFactu.tipoMov) & Format(CLng(Data1.Recordset.Fields!NumFactu), "0000000") & "'"
                            Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                            ConnConta.Execute "Delete from pagos WHERE " & Sql
                        Else
                            Sql = " ctaprove='" & vFactu.CtaSocio & "' AND numfactu='" & ObtenerLetraSerie(vFactu.tipoMov) & Format(CLng(Data1.Recordset.Fields!NumFactu), "0000000") & "'"
                            Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                            ConnConta.Execute "Delete from spagop WHERE " & Sql
                        End If
'                    Else
'                        Sql = " codmacta='" & vFactu.CtaSocio & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu & " "
'                        Sql = Sql & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
'                        Sql = Sql & " AND numserie = '" & ObtenerLetraSerie(Text1(1).Text) & "'"
'                        ConnConta.Execute "Delete from scobro WHERE " & Sql
'                    End If

                    'Volvemos a grabar en TESORERIA. Tabla de Contabilidad: sconta.spagop
                    If bol Then
'[Monica]10/07/2012: tanto liquidaciones como rectificativas han de estar en la spagop
'                        If Text1(1).Text = "FLI" Then
                            bol = vFactu.InsertarEnTesoreria(MenError)
'                        Else
'                            vFactu.TotalFac = CCur(TransformaPuntosComas(ImporteSinFormato(Text1(19).Text))) * (-1)
'                            bol = vFactu.InsertarEnTesoreriaCobro("", MenError)
'                        End If
                    End If
                End If
                Set vFactu = Nothing
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
Dim I As Integer
Dim vFactu As CFacturaCom
Dim FacOK As Boolean

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 22 To 38
         Text1(I).Text = ""
    Next I
    
    
    Set vFactu = New CFacturaCom
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(16).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(17).Text))
    vFactu.Proveedor = Text1(4).Text
    
    If vFactu.CalcularDatosFactura(ObtenerWhereCP(False), NombreTabla, "") Then
        FacOK = True
        Text1(22).Text = vFactu.BrutoFac
        Text1(23).Text = vFactu.ImpPPago
        Text1(24).Text = vFactu.ImpGnral
        Text1(25).Text = vFactu.BaseImp
        Text1(26).Text = QuitarCero(vFactu.TipoIVA1)
        Text1(27).Text = QuitarCero(vFactu.TipoIVA2)
        Text1(28).Text = QuitarCero(vFactu.TipoIVA3)
        Text1(29).Text = vFactu.PorceIVA1
        Text1(30).Text = vFactu.PorceIVA2
        Text1(31).Text = vFactu.PorceIVA3
        Text1(32).Text = vFactu.BaseIVA1
        Text1(33).Text = vFactu.BaseIVA2
        Text1(34).Text = vFactu.BaseIVA3
        Text1(35).Text = vFactu.ImpIVA1
        Text1(36).Text = vFactu.ImpIVA2
        Text1(37).Text = vFactu.ImpIVA3
        Text1(38).Text = vFactu.TotalFac
        FormatoDatosTotales
    Else
        FacOK = False
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
    CalcularDatosFactura = FacOK
End Function


Private Function ActualizarDatosFactura() As Boolean
Dim vFactu As CFacturaSoc
Dim cadSel As String

    Set vFactu = New CFacturaSoc
    cadSel = Trim(ObtenerWhereCP(False))
'    cadSel = "sfactusoc_serv." & cadSel
    
'    'Si tiene RETENCION
'    If Me.FrmRetencionSocios.visible Then
'        vFactu.PorRet = ImporteFormateado(Text1(32).Text)
'        vFactu.ImpRet2 = ImporteFormateado(Text1(33).Text)
'    End If

    
    If vFactu.CalcularDatosFacturaLiq(cadSel, "sfactusoc", "sfactusoc_serv") Then
        Text1(13).Text = vFactu.BrutoFac
        Text1(14).Text = vFactu.BrutoFac
        Text1(15).Text = vFactu.TipoIVA1
        Text1(16).Text = vFactu.BaseIVA1
        Text1(17).Text = vFactu.PorceIVA1
        Text1(18).Text = vFactu.ImpIVA1
        Text1(3).Text = vFactu.BrutoFac
        Text1(20).Text = vFactu.ImpRet2
        Text1(19).Text = vFactu.TotalFac
'        If Me.FrmRetencionSocios.visible Then
'            Text1(32).Text = vFactu.PorRet
'            Text1(33).Text = vFactu.ImpRet2
'        End If
        
        FormatoDatosTotales
        
        ActualizarDatosFactura = True
    Else
        ActualizarDatosFactura = False
        MuestraError Err.Number, "Recalculando Factura", Err.Description
    End If
    Set vFactu = Nothing
End Function

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then CheckCadenaBusqueda Check1(Index), BuscaChekc
End Sub


Private Sub cmdAceptar_Click()
Dim I As Integer
Dim cad1 As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarCabecera Then
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    PonerModo 2
                End If
            End If


        Case 4  'MODIFICAR
            If DatosOk Then
               If ModificarFactura Then
                                        
                    '[Monica]05/03/2012: guardo en el slog los campos que me han cambiado
                    cad1 = ""
                    If Text1(12).Text <> FPago Then cad1 = cad1 & "FPago: " & FPago & " a " & Text1(12).Text & ";"
                    If ComprobarCero(Text1(22).Text) <> Serv Then cad1 = cad1 & "Serv.: " & Serv & " a " & Text1(22).Text & ";"
                    If Text1(13).Text <> TotImp Then cad1 = cad1 & "Tot.Imp.: " & TotImp & " a " & Text1(13).Text & ";"
                    If Text1(18).Text <> ImpIVA Then cad1 = cad1 & "Imp.Iva.: " & ImpIVA & " a " & Text1(18).Text & ";"
                    If Text1(20).Text <> ImpRet Then cad1 = cad1 & "Imp.Ret.: " & ImpRet & " a " & Text1(20).Text & ";"
                    If Text1(19).Text <> TotFac Then cad1 = cad1 & "Tot.Fac.: " & TotFac & " a " & Text1(19).Text & ";"
                                        
                                        
                    Set LOG = New cLOG
                    LOG.Insertar 8, vUsu, "Factura Liq.modificada: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & " " & Text1(4).Text & " " & cad1 & vbCrLf
                    Set LOG = Nothing
               
               
                    Espera 0.2
                    TerminaBloquear
                    PosicionarData
'                    FormatoDatosTotales
'                    i = Data1.Recordset.AbsolutePosition
'                    PonerCamposLineas
'                    SituarDataPosicion Data1, CLng(i), ""
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            If ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then

                    'INSERTA LOG
                    '-------------------------------------------------
                    Set LOG = New cLOG
                    BuscaChekc = "   Linea: " & Data2.Recordset!numlinea
                    BuscaChekc = "Modificar linea Fact.Liq: " & Text1(1).Text & "-" & Text1(0).Text & " " & Text1(2).Text & " " & Text1(4).Text & BuscaChekc
                    LOG.Insertar 8, vUsu, BuscaChekc
                    Set LOG = Nothing
                    BuscaChekc = ""

                    TerminaBloquear
                    CargaGrid DataGrid1, Data2, True
                    ModificaLineas = 0
'                    PonerBotonCabecera True
'                    BloquearTxt Text2(16), True

                    LLamaLineas Modo, 0, "DataGrid1"
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

Private Function InsertarCabecera() As Boolean
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String
Dim NumFactu As Long
Dim vSocio As CSocio
Dim bol As Boolean
Dim devuelve As Long
Dim Existe As Boolean
Dim MenError As String
Dim vFacSoc As CFacturaSoc
Dim CtaBanco As String


    On Error GoTo EInsertarCab
    
    CodTipoMov = "FRL"
    
    bol = False
    
    conn.BeginTrans
    ConnConta.BeginTrans
    
    Set vSocio = New CSocio
    If vSocio.LeerDatos(Text1(4).Text) Then
        NumFactu = vSocio.ConseguirContador(CodTipoMov)
        If NumFactu = -1 Then bol = False
        Do
            NumFactu = vSocio.ConseguirContador(CodTipoMov)
            Sql = "select numfactu from rfactusoc where codtipom = " & DBSet(CodTipoMov, "T") & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(Text1(2).Text, "F") & " and codsocio = " & DBSet(vSocio.Codigo, "N")
            devuelve = DevuelveValor(Sql) 'DevuelveDesdeBDNew(cAgro, "rfacttra", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> 0 Then
                'Ya existe el contador incrementarlo
                Existe = True
                vSocio.IncrementarContador (CodTipoMov)
                NumFactu = vSocio.ConseguirContador(CodTipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
        Text1(0).Text = NumFactu
        
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            
            MenError = "Error al insertar en la tabla Cabecera de Factura (" & NombreTabla & ")."
            conn.Execute Sql, , adCmdText
            
            Set vFacSoc = New CFacturaSoc
            '[Monica]22/11/2013: iban
            vFacSoc.CCC_Iban = vSocio.Iban
            vFacSoc.CCC_Entidad = vSocio.Banco
            vFacSoc.CCC_Oficina = vSocio.Sucursal
            vFacSoc.CCC_CC = vSocio.DigControl
            vFacSoc.CCC_CTa = vSocio.CuentaBan
            vFacSoc.ForPago = Text1(12).Text
            vFacSoc.tipoMov = CodTipoMov
            vFacSoc.NumFactu = Text1(0).Text
            vFacSoc.FecFactu = Text1(2).Text
            '[Monica]10/07/2012: Tiene que estar en negativo en la spagop
            vFacSoc.TotalFac = CCur(TransformaPuntosComas(ImporteSinFormato(Text1(19).Text))) '* (-1)
            vFacSoc.ImpRet2 = Text1(20).Text
            vFacSoc.Socio = Text1(4).Text
            
            vFacSoc.CtaSocio = vSocio.CtaSocioLiq
            
            cadban = ""
            
            Set frmBan = New frmFacBancosPropios
            frmBan.DatosADevolverBusqueda = "1|"
            frmBan.Show vbModal
            Set frmBan = Nothing
            
            CtaBanco = cadban ' InputBox("Introduzca el Banco de pago: ", "Tesoreria", , 5000, 4000)

            If CtaBanco = "" Then
                MsgBox "No ha seleccionado cuenta de banco.", vbExclamation
                bol = False
            Else
                bol = True
                vFacSoc.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", CtaBanco, "N")
            End If

            If bol Then bol = ActualizarRetencion(vSocio.UveSocio, vFacSoc, True)
            
            '[Monica]10/07/2012: Tiene que estar en negativo en la spagop
            If bol Then bol = vFacSoc.InsertarEnTesoreria(MenError)   'vFacSoc.InsertarEnTesoreriaCobro("", MenError)
            
            If bol Then bol = vSocio.IncrementarContador(CodTipoMov)
            
            Set vFacSoc = Nothing
            
        End If
        
        Text1(0).Text = Format(Text1(0).Text, "0000000")
    End If
    Set vSocio = Nothing
    
EInsertarCab:
    Screen.MousePointer = vbDefault

    If Err.Number <> 0 Then
        MenError = "Insertando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ConnConta.CommitTrans
        InsertarCabecera = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        InsertarCabecera = False
    End If
End Function


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
Dim vSocio As CSocio
Dim vFacSoc As CFacturaSoc
Dim CtaBanco As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
'    cambiaSQL = False
'    'Comprobar si mientras tanto se incremento el contador de Pedidos
'    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
'    Do
'        devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numfactu", "codtipom", Text1(1).Text, "T", , "numfactu", Text1(0).Text, "N")
'        If devuelve <> "" Then
'            'Ya existe el contador incrementarlo
'            Existe = True
'            vTipoMov.IncrementarContador (CodTipoMov)
'            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
'            cambiaSQL = True
'        Else
'            Existe = False
'        End If
'    Loop Until Not Existe
'    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    ConnConta.BeginTrans
    
    MenError = "Error al insertar en la tabla Cabecera de Factura (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    If bol Then
        MenError = "Error al insertar en Tesoreria."
        Set vSocio = New CSocio
        If vSocio.LeerDatos(Text1(4).Text) Then
            
            Set vFacSoc = New CFacturaSoc
            '[Monica]22/11/2013: iban
            vFacSoc.CCC_Iban = vSocio.Iban
            vFacSoc.CCC_Entidad = vSocio.Banco
            vFacSoc.CCC_Oficina = vSocio.Sucursal
            vFacSoc.CCC_CC = vSocio.DigControl
            vFacSoc.CCC_CTa = vSocio.CuentaBan
            vFacSoc.ForPago = Text1(12).Text
            vFacSoc.NumFactu = Text1(0).Text
            vFacSoc.FecFactu = Text1(2).Text
            vFacSoc.TotalFac = Text1(19).Text
            vFacSoc.ImpRet2 = Text1(20).Text
            vFacSoc.Socio = Text1(4).Text
            
            vFacSoc.CtaSocio = vSocio.CtaSocioLiq
            
            CtaBanco = InputBox("Introduzca el Banco de pago: ", "Tesoreria", , 5000, 4000)

            vFacSoc.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", CtaBanco, "N")
            
            If bol Then bol = ActualizarRetencion(vSocio.UveSocio, vFacSoc, True)
            
            If bol Then bol = vFacSoc.InsertarEnTesoreria(MenError)
            
            Set vFacSoc = Nothing
        End If
    
    End If
    
    If bol Then
        'Actualizar el campo fechamov (ult. movimiento) de la tabla de clientes (sclien)
        MenError = "Error al actualizar el contador de la Factura."
        vTipoMov.IncrementarContador (CodTipoMov)
    End If
    
EInsertarOferta:
    If Err.Number <> 0 Then
        MenError = "Insertando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ConnConta.CommitTrans
        InsertarOferta = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        InsertarOferta = False
    End If
End Function


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

Private Sub FormatoDatosTotales()
Dim I As Byte

    For I = 16 To 19
'        Text1(I).Text = QuitarCero(Text1(I).Text)
        FormateaCampo Text1(I)
    Next I
    
'    For i = 24 To 26
'        If Text1(i).Text <> "" Then
'            'Si la Base Imp. es 0
'            If CSng(Text1(i).Text) = 0 Then
'                Text1(i).Text = QuitarCero(Text1(i).Text)
'                Text1(i - 3).Text = QuitarCero(Text1(i - 3).Text)
'                Text1(i - 6).Text = QuitarCero(Text1(i - 6).Text)
'                Text1(i + 3).Text = QuitarCero(Text1(i + 3).Text)
'            Else
'                FormateaCampo Text1(i)
'                FormateaCampo Text1(i - 3)
'                FormateaCampo Text1(i - 6)
'                FormateaCampo Text1(i + 3)
'            End If
'        Else 'No hay Base Imponible
'            Text1(i - 3).Text = QuitarCero(Text1(i - 3).Text)
'            Text1(i - 6).Text = QuitarCero(Text1(i - 6).Text)
'            Text1(i + 3).Text = ""
'        End If
'    Next i
'
'    If Me.FrmRetencionSocios.visible Then
'        FormateaCampo Text1(32)
'        FormateaCampo Text1(33)
'    End If
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas 0, 0, "DataGrid1" 'modo,0,"DataGrid2"
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
            DataGrid1.Enabled = True
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid1"
'            PonerBotonCabecera True
            PonerModo 2
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

            Me.DataGrid1.Enabled = True
    End Select

End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid1.Enabled = True
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If

End Sub




Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus()
    Text1(1).Text = Mid(Combo1.List(Combo1.ListIndex), 1, 3)
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        txtAux2(0).Text = DBLet(Data2.Recordset.Fields!observac2, "T")
    Else
        txtAux2(0).Text = ""
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

Private Sub Form_Load()
    PrimeraVez = True
    UnaVez = True
    
    'Icono del formulario
    Me.Icon = frmppal.Icon
    
     'Icono de busqueda
    Me.imgBuscar(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgBuscar(2).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgBuscar(5).Picture = frmppal.imgIcoForms.ListImages(1).Picture

    ' ICONITOS DE LA BARRA
'    btnPrimero = 13
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(4).Image = 3   'Insertar
'        .Buttons(5).Image = 4   'Modificar
'        .Buttons(6).Image = 5   'Borrar
'        .Buttons(7).Image = 10   'lineas
'        .Buttons(8).Image = 16 'Imprimir
'        .Buttons(11).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'    End With

    With Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.imgListComun1
        'ASignamos botones
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2 'Ver Todos
        .Buttons(1).Image = 3 'Añadir
        .Buttons(2).Image = 4 'Modificar
        .Buttons(3).Image = 5 'Eliminar
        .Buttons(8).Image = 16 'Imprimir
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.imgListComun1
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With

    With Me.ToolAux(0)
        .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
    End With

    LimpiarCampos   'Limpia los campos TextBox
    
    CargaCombo
    
        
    '## A mano
    NombreTabla = "sfactusoc"
    NomTablaLineas = "sfactusoc_serv" 'Tabla lineas de llamadas
    Ordenacion = " ORDER BY sfactusoc.codtipom, sfactusoc.codsocio, sfactusoc.numfactu, sfactusoc.fecfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Dim T1 As Single
    T1 = Timer
    
    ' facturas de liquidacion o rectificativas de liquidacion de socio
    cadB1 = "(sfactusoc.codtipom = 'FLI' or sfactusoc.codtipom = 'FRL')"
    
    CadenaConsulta = "Select * from " & NombreTabla ' & " where codtipom is null and " & cadB1
    
  '**
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        'como tenemos aqui el nº de albaran, buscar a que factura corresponde
        'en la scafaccli1
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura & " and " & cadB1
'        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        'CadenaConsulta = CadenaConsulta & " where numfactu=-1"
        'Cambio sugerido por Msoler
        'mariela 01/07/2010
        CadenaConsulta = CadenaConsulta & " WHERE codtipom is null and numfactu is null and fecfactu is null and " & cadB1
    End If
  
  '**
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    
    Data1.Refresh
    
    Me.SSTab1.Tab = 0
   
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        BotonBuscar
'    End If
'    LimpiarDataGrids
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

            
End Sub

Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        BuscaChekc = ""

        Select Case Me.Combo1.ListIndex
            Case 0
                Text1(1).Text = "FLI"
            Case 1
                Text1(1).Text = "FRL"
        End Select


        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue 'vbYellow
        
        
        anc = DataGrid1.top
        If DataGrid1.Row < 0 Then
            anc = anc + 430
        Else
            anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If
        CargaGrid DataGrid1, Me.Data2, False
        LLamaLineas 1, anc, "DataGrid1"
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue 'vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub HacerBusqueda()
Dim CadB As String
    
    If Text1(4).Text <> "" Then Text1(5).Text = ""
    
    CadB = ObtenerBusqueda(Me, False, BuscaChekc)
    If CadB = "" Then
        CadB = cadB1
    Else
        CadB = CadB & " and " & cadB1
    End If
    
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
'        cadB = cadB & " and scafac.codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los FAV
        CadenaConsulta = "select sfactusoc.* from " & NombreTabla & "  "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY sfactusoc.codtipom,sfactusoc.codsocio,sfactusoc.numfactu,sfactusoc.fecfactu " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    Me.Check1(0).Value = 0
    Me.Check1(1).Value = 0
    Me.Combo1.ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    hcoCodMovim = ""
    hcoCodTipoM = "" 'Codigo detalle de Movimiento(ALC)
    hcoFechaMov = "" 'fecha del movimiento
    hcoCodSocio = "" ' codigo de socio
    DesdeFichaSocio = False
End Sub

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        cadban = RecuperaValor(CadenaSeleccion, 1)
    Else
        cadban = ""
    End If
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

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim indice As Byte
    indice = 12
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(indice + 3).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub

Private Sub frmHcoFacSocPre_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaSeleccion, 1)
        CadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaSeleccion, 3)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(4), CadenaSeleccion, 4)
        CadB = CadB & " and " & Aux
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        
        PonerCadenaBusqueda
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

'    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            PonerFoco Text1(4)
            Set frmC = New frmGesSocios
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
            PonerFoco Text1(4)
            PonerDatosSocio
            
        Case 5 'forma de pago
            indice = 12
            PonerFoco Text1(indice)
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 2 'codpobla
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            indice = 9
            PonerFoco Text1(indice)
            
        Case 3 ' observaciones
            If Modo = 5 Or Modo = 0 Then
            Else
                If Modo = 3 Or Modo = 4 Then
                    CadenaDesdeOtroForm = Text1(3).Text
                Else
                    CadenaDesdeOtroForm = ""
                    If Not Me.Data1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Data1.Recordset!Concepto, "T")
                End If
                frmFacClienteObser.Modificar = Modo >= 3
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 3 Or Modo = 4 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(3).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
            End If
        
    End Select

End Sub

Private Sub PonerDatosSocio()
Dim Cad As String

    If Text1(4).Text = "" Then Exit Sub

    Set miRsAux = New ADODB.Recordset
    
    Cad = "select * from sclien where codclien=" & Text1(4).Text
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Text1(5).Text = miRsAux!nomclien
        Text1(6).Text = miRsAux!nifClien
        Text1(7).Text = DBLet(miRsAux!telclie1, "T")
        Text1(8).Text = miRsAux!domclien
        Text1(9).Text = miRsAux!codpobla
        Text1(10).Text = miRsAux!pobclien
        Text1(11).Text = miRsAux!proclien
        Text1(24).Text = DBLet(miRsAux!NumerUve, "N")
    End If
    
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnLineas_Click()
    BotonMtoLineas 1, "Facturas"
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
                If BloqueaLineasFac Then BotonModificarLinea
        End If
    Else
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: sfactusoc
            BotonModificar
        End If
    End If

End Sub

Private Function BloqueaLineasFac() As Boolean
'bloquea todas las lineas de la factura
Dim Sql As String

    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM sfactusoc_serv "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaLineasFac = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasFac = False
End Function




Private Sub BotonModificar()
Dim DeVarios As Boolean
Dim EnTesoreria  As String

    '[Monica]16/03/2012: Si la factura es de liquidacion se inserto en spagop, si es rectificativa en la scobro
    'solo se puede modificar la factura si no esta contabilizada
    
    If FactContabilizada(EnTesoreria) Then
        TerminaBloquear
        Exit Sub
    End If
    
    '[Monica]05/03/2012: guardo el valor de los campos que pueden modificar para el slog
    FPago = Text1(12).Text
    Serv = ComprobarCero(Text1(22).Text)
    TotImp = Text1(13).Text
    ImpIVA = Text1(18).Text
    ImpRet = Text1(20).Text
    TotFac = Text1(19).Text
    
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    'PonerFocoChk Me.Check1
    PonerFoco Text1(12)
    
'    'Inserto en slog
'
'    Set LOG = New cLOG
'    If EnTesoreria <> "" Then EnTesoreria = "Tesoreria: " & vbCrLf & EnTesoreria
'    EnTesoreria = Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & vbCrLf & EnTesoreria
'    EnTesoreria = "Pulsa mod factura: " & EnTesoreria
'    LOG.Insertar 8, vUsu, EnTesoreria
'    Set LOG = Nothing
'    Espera 0.3
'    '
    
End Sub

Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte
Dim EstaEnTesoreria As String
    On Error GoTo EModificarLinea

   '[Monica]16/03/2012:
   '                    si la factura es de liquidacion se inserto en spagop
   '                    si es rectificativa se inserto en la scobro
   

     'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada(EstaEnTesoreria) Then
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
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If


    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        J = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, J
        DataGrid1.Refresh
    End If
    
'    anc = ObtenerAlto(Me.DataGrid1)
    anc = DataGrid1.top
    If DataGrid1.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 10
    End If


    txtAux3(12).Text = DataGrid1.Columns(15).Text
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
'    PonerBotonCabecera False
    PonerFoco txtAux3(12)
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
            b = (xModo = 1) Or (xModo = 2)
            For jj = 3 To 11
                txtAux3(jj).Height = DataGrid1.RowHeight
                txtAux3(jj).top = alto
                txtAux3(jj).visible = (xModo = 1)
            Next jj
            
            For jj = 12 To 12
                txtAux3(jj).Height = DataGrid1.RowHeight
                txtAux3(jj).top = alto
                txtAux3(jj).visible = b
            Next jj
    End Select
End Sub

Private Function FactContabilizada2(ByRef EstaEnTesoreria As String) As Boolean
Dim LEtra As String, numasien As String

    On Error GoTo EContab
    
    
    
    
    'NO deberia poder modificar fras anteriors a fecha inicio ejercicio
    
    
    'Cojo la letra de serie
    LEtra = ObtenerLetraSerie(Text1(1).Text)

    'Primero comprobaremos que esta el cobro en contabilidad
    EstaEnTesoreria = ""
    
    
    If Not ComprobarCobroArimoney(EstaEnTesoreria, LEtra, "", CLng(Text1(0).Text), CDate(Text1(2).Text)) Then
        FactContabilizada2 = True
        Exit Function
    End If


    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1(0).Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
      
        If LEtra <> "" Then
            If vParamAplic.ContabilidadNueva Then
                numasien = DevuelveDesdeBDNew(conConta, "factcli", "numasien", "numserie", LEtra, "T", , "numfactu", Text1(0).Text, "N", "anofactu", Year(Text1(2).Text), "N")
            Else
                numasien = DevuelveDesdeBDNew(conConta, "cabfact", "numasien", "numserie", LEtra, "T", , "codfaccl", Text1(0).Text, "N", "anofaccl", Year(Text1(2).Text), "N")
            End If
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
Dim cta As String, numasien As String
    
    On Error GoTo EContab
    
    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1(0).Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        
'        Cta = vParamAplic.Raiz_Cta_Soc_publi & Format(Text1(4).Text, "0000")
''        Cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
'        If Cta <> "" Then
'            numasien = DevuelveDesdeBDNew(conConta, "cabfactprov", "numasien", "codmacta", Cta, "T", , "numfacpr", Text1(0).Text, "T", "fecfacpr", Text1(1).Text, "F")
'            If numasien <> "" Then
                FactContabilizada3 = True
                MsgBox "La factura esta contabilizada y no se puede modificar ni eliminar.", vbInformation
                Exit Function
'            Else
'                FactContabilizada = False
'            End If
'        Else
'            FactContabilizada = True
'            Exit Function
'        End If
    Else
        FactContabilizada3 = False
    End If
    
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function




Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim Sql As String
Dim vWhere As String
Dim b As Boolean

    On Error GoTo EModificarLinea

    ModificarLinea = False
    If Data2.Recordset.EOF Then Exit Function
    
    
    If Not DatosOkLin Then Exit Function
    
    
    
    vWhere = ObtenerWhereCP(True)
    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
    
    Sql = "UPDATE sfactusoc_serv SET "
    Sql = Sql & " impventa=" & DBSet(txtAux3(12).Text, "N")
    Sql = Sql & vWhere
    
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
    DataGrid1.Enabled = Not b
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function FactContabilizada(ByRef EstaEnTesoreria As String) As Boolean
Dim cta As String, numasien As String
Dim vSocio As CSocio
Dim LEtra As String
Dim NumFac As String


    On Error GoTo EContab
    
    FactContabilizada = False
    
    Set vSocio = New CSocio
    If vSocio.LeerDatos(Text1(4).Text) Then
        cta = vSocio.CtaSocioLiq
    
    
        'Primero comprobaremos que esta el pago en contabilidad
        EstaEnTesoreria = ""

'[Monica]10/07/2012: tanto liquidaciones como rectificativas han de estar en la spago
'        If Text1(1).Text = "FLI" Then
            If Not ComprobarPagoArimoney(EstaEnTesoreria, vSocio.CtaSocioLiq, CLng(Text1(0).Text), CDate(Text1(2).Text)) Then
                FactContabilizada = True
                Exit Function
            End If
'        Else
'            If Not ComprobarCobroArimoney(EstaEnTesoreria, ObtenerLetraSerie(Text1(1).Text), vSocio.CtaSocioLiq, CLng(Text1(0).Text), CDate(Text1(2).Text)) Then
'                FactContabilizada = True
'                Exit Function
'            End If
'        End If
        'comprabar que se puede modificar/eliminar la factura
        If Me.Check1(0).Value = 1 Then 'si esta contabilizada
            'comprobar en la contabilidad si esta contabilizada
        
    '        Cta = vParamAplic.Raiz_Cta_Soc_Liqui & Format(Text1(4).Text, "0000")
    '        Cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
            If vSocio.CtaSocioLiq <> "" Then
                LEtra = DevuelveDesdeBDNew(conAri, "stipom", "letraser", "codtipom", Text1(1).Text, "T")
                NumFac = LEtra & Text1(0).Text
                If vParamAplic.ContabilidadNueva Then
                    numasien = DevuelveDesdeBDNew(conConta, "factpro", "numasien", "codmacta", vSocio.CtaSocioLiq, "T", , "numfactu", NumFac, "T", "fecfactu", Text1(2).Text, "F")
                Else
                    numasien = DevuelveDesdeBDNew(conConta, "cabfactprov", "numasien", "codmacta", vSocio.CtaSocioLiq, "T", , "numfacpr", NumFac, "T", "fecfacpr", Text1(2).Text, "F")
                End If
                If numasien <> "" Then
'                    FactContabilizada = True
'                    MsgBox "La factura esta contabilizada y no se puede modificar ni eliminar.", vbInformation
'                    Exit Function
                Else
'                    FactContabilizada = False
                End If
                
                LEtra = "La factura esta en la contabilidad"
                If numasien <> "" Then LEtra = LEtra & vbCrLf & "Nº asiento: " & numasien
                LEtra = LEtra & vbCrLf & vbCrLf & "¿Continuar?"
                
                numasien = String(50, "*") & vbCrLf
                numasien = numasien & numasien & vbCrLf & vbCrLf
                LEtra = numasien & LEtra & vbCrLf & vbCrLf & numasien
                If MsgBox(LEtra, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    FactContabilizada = False
                Else
                    FactContabilizada = True
                End If
                
            Else
                FactContabilizada = True
                Exit Function
            End If
        End If
    Else
        FactContabilizada = False
    End If
    Set vSocio = Nothing
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function

Private Sub mnRectifica_Click()
    If Modo = 5 Then 'Añadir lineas
'         BotonAnyadirLinea
    Else 'Añadir Cabecera
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub

Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String
Dim Cad As String
Dim Rs As ADODB.Recordset

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    NomTraba = ""

    Text1(2).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    Text1(1).Text = "FRL"
    Text1(15).Text = vParamAplic.IVA_REA
    Text1(17).Text = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", vParamAplic.IVA_REA, "N")
    Text1(21).Text = vParamAplic.PorReten
    PonerFoco Text1(2)
End Sub



Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Select Case Index
        Case 2 ' fecha de factura
            PonerFormatoFecha Text1(Index)
            
        Case 4 'socio
            If Modo = 1 Then 'Modo=1 Busqueda
                '-- Laura 12/01/2007
                'Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, NombreTabla, "nomclien")
                Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", Text1(Index).Text, "N")
                '--
            Else
                PonerDatosSocio
            End If
        
        Case 12 'forma de pago
            If Text1(Index).Text <> "" Then
                devuelve = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", Text1(Index).Text, "T")
                If devuelve = "" Then
                    MsgBox "El código de forma de pago introducido no existe.", vbExclamation
                    PonerFoco Text1(Index)
                    Exit Sub
                Else
                    Text2(Index + 3).Text = devuelve
                End If
            End If
            
       Case 3, 13, 16, 18, 19, 20
            PonerFormatoDecimal Text1(Index), 3
            
       Case 17, 21
            PonerFormatoDecimal Text1(Index), 7
       
            
    End Select
    
    If (Modo = 4 Or Modo = 3) And ((Index >= 13 And Index <= 21) Or Index = 3) Then
        Dim PorceIVA As Currency
        Dim BaseImpo As Currency
        Dim BaseReten As Currency
        Dim TotalFac As Currency
        Dim TotalFac1 As Currency
        Dim PorceReten As Currency
        Dim ImpoReten As Currency
        Dim ImpoIva As Currency
        Dim Suplidos As Currency
        
        If Index = 15 Or Index = 21 Then
            BaseImpo = CCur(ImporteSinFormato(Text1(16).Text))
            
            Suplidos = CCur(ComprobarCero(Text1(25).Text))
            
            Text1(17).Text = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", Text1(15).Text, "N")
            PorceIVA = 0
            If Text1(17).Text <> "" Then PorceIVA = CCur(Text1(17).Text)
            ImpoIva = Round2(BaseImpo * PorceIVA / 100, 2)
            PorceReten = CCur(ImporteSinFormato(Text1(21).Text))
            ImpoReten = Round2((BaseImpo + ImpoIva) * PorceReten / 100, 2)
            
            TotalFac1 = BaseImpo + ImpoIva - ImpoReten + Suplidos
        
            Text1(16).Text = Format(BaseImpo, "#,###,###,##0.00")
            Text1(17).Text = Format(PorceIVA, "#0.00")
            Text1(18).Text = Format(ImpoIva, "#,###,###,##0.00")
            
            Text1(20).Text = Format(ImpoReten, "#,###,###,##0.00")
            Text1(19).Text = Format(TotalFac1, "#,###,###,##0.00")
        End If
        
        If Index = 13 Then
            TotalFac = CCur(ImporteSinFormato(Text1(13).Text))
            
            Text1(17).Text = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", Text1(15).Text, "N")
            PorceIVA = 0
            If Text1(17).Text <> "" Then PorceIVA = CCur(Text1(17).Text)
            BaseImpo = Round2(TotalFac / (1 + (PorceIVA / 100)), 2)
            'BaseReten = TotalFac
            ImpoIva = TotalFac - BaseImpo
            If Text1(21).Text = "" Then Text1(21).Text = "0"
            PorceReten = CCur(ImporteSinFormato(Text1(21).Text))
            ImpoReten = Round2(TotalFac * PorceReten / 100, 2)
            TotalFac1 = TotalFac - ImpoReten
        
            Text1(13).Text = Format(TotalFac, "#,###,###,##0.00")
            Text1(14).Text = Text1(13).Text
            Text1(16).Text = Format(BaseImpo, "#,###,###,##0.00")
            Text1(17).Text = Format(PorceIVA, "#0.00")
            Text1(18).Text = Format(ImpoIva, "#,###,###,##0.00")
            Text1(3).Text = Text1(13).Text
            
            Text1(20).Text = Format(ImpoReten, "#,###,###,##0.00")
            Text1(19).Text = Format(TotalFac1, "#,###,###,##0.00")
            
            cmdAceptar.SetFocus
            
        End If
        
        If Index = 16 Or Index = 18 Or Index = 20 Then
            BaseImpo = CCur(ImporteSinFormato(Text1(16).Text))
            ImpoIva = CCur(ImporteSinFormato(Text1(18).Text))
            BaseReten = BaseImpo + ImpoIva
            
            Text1(3).Text = Format(BaseReten, "#,###,###,##0.00")
            PorceReten = CCur(ImporteSinFormato(Text1(21).Text))
            ImpoReten = Round2(BaseReten * PorceReten / 100, 2)
            TotalFac1 = BaseImpo + ImpoIva - ImpoReten
            
            Text1(20).Text = Format(ImpoReten, "#,###,###,##0.00")
            Text1(19).Text = Format(TotalFac1, "#,###,###,##0.00")
        End If
    End If
    
End Sub

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    If Me.DataGrid1.visible Then
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "La factura no tiene lineas.", vbInformation
            Exit Sub
        End If
    End If
    If vUsu.Nivel >= 1 Then
        MsgBox "No tiene permiso", vbExclamation
        Exit Sub
    End If
    
    ModificaLineas = 0
    PonerModo 5


    Select Case Button.Index
'        Case 1
'            BotonAnyadirLinea
        Case 2
        
            mnModificar_Click
'            BotonModificarLinea
'        Case 3
'            BotonEliminarLinea
    End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5: mnBuscar_Click  'Buscar
        Case 6: BotonVerTodos  'Todos

        Case 1: mnRectifica_Click  'Rectifica
        
        Case 2: mnModificar_Click  'Modificar
        Case 3: mnEliminar_Click  'Borrar
        
'        Case 7: mnLineas_Click ' Lineas de factura
        Case 8: mnImprimir_Click 'Imprimir factura
        
    End Select
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnImprimir_Click()
    HacerImpresionFacturas
End Sub

Private Sub HacerImpresionFacturas()
    cadFormula = "({sfactusoc.codtipom}=" & DBSet(Text1(1).Text, "T") & " and {sfactusoc.numfactu}=" & Text1(0).Text
    cadFormula = cadFormula & " and {sfactusoc.fecfactu} = Date(" & Year(Text1(2).Text) & "," & Month(Text1(2).Text) & "," & Day(Text1(2).Text) & ")"
    cadFormula = cadFormula & " and {sfactusoc.codsocio} = " & Text1(4).Text & ")"
    
    '[Monica]29/02/2012: en la impresion de factura de liquidacion de socio hemos metido el tmpinformes
    If InsertResumen(Text1(1).Text, Text1(0).Text, Text1(4).Text, Text1(2).Text) Then
        cadFormula = "{tmpinformes.codusu} =" & vUsu.Codigo
        LlamarImprimir True
    End If
End Sub

'Insertar Resumen
Private Function InsertResumen(Tipo As String, NumFactu As String, Socio As String, FecFac As String) As Boolean
Dim MensError As String
Dim Sql As String
    
    On Error GoTo eInsertResumen
    
    MensError = ""
    InsertResumen = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
                                        ' codtipom, numfactu, codsocio, fecfactu
    Sql = "insert into tmpinformes (codusu, nombre1, importe1, codigo1, fecha1) values ( " & vUsu.Codigo
    Sql = Sql & ",'" & Tipo & "'," & DBSet(NumFactu, "N") & "," & DBSet(Socio, "N") & "," & DBSet(FecFac, "F") & ")"
    
    conn.Execute Sql
    
    InsertResumen = True
    
    Exit Function

eInsertResumen:
    MensError = "Error en la inserción de la factura " & NumFactu & " en el Resumen "
    MuestraError Err.Number, MensError
End Function



Private Sub LlamarImprimir(duplicado As Boolean)
Dim devuelve As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal

    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    cadParam = cadParam & "pDuplicado= " & Abs(duplicado) & " |"
    numParam = 2
        
    '[Monica]31/03/2014
    '[Monica]19/02/2018: Entra Cordoba
        '[Monica]19/02/2018: Entra Sevilla
    If (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3) And Text1(1).Text = "FLI" Then
        'preguntamos si quiere imprimirlo o no con los servicios
        If MsgBox("¿ Desea imprimir el detalle de servicios ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            cadParam = cadParam & "pDetalle=0|"
        Else
            cadParam = cadParam & "pDetalle=1|"
        End If
        numParam = numParam + 1
    End If
        
        
    indRPT = 51 'Impresion de facturas de liquidacion a socios
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, False, "") Then Exit Sub
    'Nombre fichero .rpt a Imprimir
        
        
        
    devuelve = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
        
    With frmImprimir
        
        'Nuevo. Febrero 2010
        .outClaveNombreArchiv = devuelve & Format(Text1(0).Text, "000")
        .outCodigoCliProv = Text1(4).Text
        .outTipoDocumento = 100
        
        .Titulo = "Impresión de Facturas de Liquidación"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .NombreRPT = nomDocu
        .Opcion = 101
        .ConSubInforme = False
        .Show vbModal
    End With

End Sub

Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia "" & cadB1
    Else
        lblIndicador.Caption = "Preparando bus."
        lblIndicador.Refresh
        LimpiarCampos
        DoEvents
        
        CadenaConsulta = "Select sfactusoc.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB1

        lblIndicador.Caption = "Obteniendo reg."
        lblIndicador.Refresh
        PonerCadenaBusqueda
    End If
End Sub
Private Sub mnEliminar_Click()
    BotonEliminar

End Sub

Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (sfactusoc)
Dim Cad As String
Dim EstaEnTesoreria As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada(EstaEnTesoreria) Then Exit Sub
    
    Cad = "Cabecera de Facturas." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Factura:            "
    Cad = Cad & vbCrLf & "Tipo:  " & Text1(1).Text
    Cad = Cad & vbCrLf & "Nº Fact.:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Socio.:  " & Format(Text1(4).Text, "000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(2).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        CodTipoMov = Text1(1).Text
        
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            Set LOG = New cLOG
            LOG.Insertar 8, vUsu, "Factura eliminada: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & " " & Text1(4).Text & vbCrLf & EstaEnTesoreria
            Set LOG = Nothing
        
            If SituarDataTrasEliminar(Data1, NumRegElim) Then
                PonerCampos
            Else
                LimpiarCampos
                LimpiarDataGrids
                'Poner los grid sin apuntar a nada
                PonerModo 0
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Albaran", Err.Description
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
        End If
        lblIndicador.Caption = ""
        LimpiarDataGrids
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
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


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I

    BuscaChekc = ""
    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
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
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
          
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
    Me.Check1(0).Enabled = (Modo = 1)
    Me.Check1(1).Enabled = (Modo = 1)
    Me.Combo1.visible = (Modo = 1)
    
    b = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(3), b 'referencia
    BloquearTxt Text1(19), b, True
    BloquearTxt Text1(17), b 'referencia
    BloquearTxt Text1(1), b ' tipo de movimiento
    BloquearTxt Text1(24), b  ' Numero de V
    
    'Importes siempre bloqueados, excepto para busquedas. ivas y aportacion tb bloqueado
    '[Monica]19/02/2018: Entra Cordoba
    For I = 13 To 16                            '[Monica]02/03/2012: dejamos modificar totales de la factura si es Teletaxi
            '[Monica]19/11/2018: Entra Sevilla
        BloquearTxt Text1(I), (Modo <> 1) And Not (Modo = 4 And (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3)) And Modo <> 3
    Next I
    BloquearTxt Text1(18), (Modo <> 1) And Not (Modo = 4 And (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3)) And Modo <> 3
    For I = 20 To 21                            '[Monica]02/03/2012: dejamos modificar totales de la factura si es Teletaxi
        BloquearTxt Text1(I), (Modo <> 1) And Not (Modo = 4 And (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 3)) And Modo <> 3
    Next I
    'Campo B.Imp y Imp. IVA siempre en azul
    BloquearTxt Text1(14), True
    Text1(14).Enabled = False
    Text1(14).BackColor = &HFFFFC0
'    BloquearTxt Text1(18), True
    Text1(3).BackColor = &HFFFFC0
    
    
    '[Monica]15/11/2017: el campo de suplidos siempre debe de estar bloqueado
    BloquearTxt Text1(25), True
    Text1(25).Enabled = False
    Text1(25).BackColor = &HFFFFC0
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImg imgBuscar(2), True
    BloquearImg imgBuscar(0), (Modo <> 1) And (Modo <> 3)  '(Modo = 2 Or Modo = 0 Or Modo = 5 Or Modo = 4)
    BloquearImg imgBuscar(5), (Modo <> 1) And (Modo <> 4)  '(Modo = 2 Or Modo = 0 Or Modo = 5)
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 3 To 11
        BloquearTxt txtAux3(I), (Modo <> 1)
        txtAux3(I).visible = (Modo = 1)
    Next I
    
    For I = 12 To 12
        BloquearTxt txtAux3(I), (Modo <> 5) And (Modo <> 1)
        txtAux3(I).visible = (Modo = 1)
    Next I
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
    PonerModoUsuarioGnral Modo, "aritaxi"
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
        Toolbar1.Buttons(2).Enabled = Toolbar1.Buttons(2).Enabled And DBLet(Rs!Modificar, "N")
        Toolbar1.Buttons(3).Enabled = Toolbar1.Buttons(3).Enabled And DBLet(Rs!creareliminar, "N")
        
        Toolbar1.Buttons(5).Enabled = Toolbar1.Buttons(5).Enabled And DBLet(Rs!ver, "N")
        Toolbar1.Buttons(6).Enabled = Toolbar1.Buttons(6).Enabled And DBLet(Rs!ver, "N")
        
        Toolbar1.Buttons(8).Enabled = Toolbar1.Buttons(8).Enabled And DBLet(Rs!Imprimir, "N")
        
        'subclientes
        For I = 0 To ToolAux.Count - 1
            ToolAux(I).Buttons(1).Enabled = ToolAux(I).Buttons(1).Enabled And DBLet(Rs!creareliminar, "N")
            ToolAux(I).Buttons(2).Enabled = ToolAux(I).Buttons(2).Enabled And DBLet(Rs!Modificar, "N")
            ToolAux(I).Buttons(3).Enabled = ToolAux(I).Buttons(3).Enabled And DBLet(Rs!creareliminar, "N")
        Next I
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean
Dim I As Integer
Dim bAux As Boolean


    b = (Modo = 2)
'    'Insertar
    Toolbar1.Buttons(1).Enabled = True
    Me.mnRectifica.Enabled = True
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = (Modo = 2)
    Me.mnEliminar.Enabled = (Modo = 2)
        
    b = (Modo = 2)
    'Lineas
'    Toolbar1.Buttons(7).Enabled = b
'    Me.mnLineas.Enabled = b
    
    'Imprimir
    Toolbar1.Buttons(8).Enabled = b
    Me.mnImprimir.Enabled = b
    
    b = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
        
    b = (Modo = 2)
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = False
        If Not Data2.Recordset Is Nothing Then
            If b Then bAux = (b And Me.Data2.Recordset.RecordCount > 0)
        End If
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = False
    Next I
               
        
        
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerCampos()
Dim BrutoFac As Single
    
    On Error Resume Next
    
    If Data1.Recordset.EOF Then
        LimpiarDataGrids
        Exit Sub
    End If
    PonerCamposForma Me, Data1
    
    BrutoFac = CSng(Text1(13).Text)
    Text1(14).Text = Format(BrutoFac, FormatoImporte)
    
    
    'poner descripcion campos
    Modo = 4
    Text1_LostFocus (12) 'forma de pago
    Modo = 2
    
    'Datos del socio
    PonerDatosSocio
    
    CargaGrid DataGrid1, Data2, True
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
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
    If Index <> 23 Then KEYpress KeyAscii
End Sub
Private Sub MandaBusquedaPrevia(CadB As String)
''Carga el formulario frmBuscaGrid con los valores correspondientes
'Dim Cad As String
'Dim Tabla As String
'Dim Titulo As String
'Dim Desc As String, devuelve As String
'    'Llamamos a al form
'    '##A mano
'    Cad = ""
'    Cad = Cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
'    Cad = Cad & ParaGrid(Text1(0), 15, "Nº Factura")
'    Cad = Cad & ParaGrid(Text1(2), 15, "Fecha Fac.")
'    Cad = Cad & ParaGrid(Text1(4), 10, "Socio")
'    Cad = Cad & ParaGrid(Text1(5), 50, "Nombre Socio")
'    Tabla = NombreTabla
'
'    Titulo = "Facturas"
'    devuelve = "0|1|2|"
'
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vTabla = Tabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        frmB.vDevuelve = devuelve
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 0
'        frmB.vConexionGrid = conAri  'Conexión a BD: Aritaxi
'        frmB.Show vbModal
'        Set frmB = Nothing
'        PonerCadenaBusqueda
'        Text1(0).Text = Format(Text1(0).Text, "0000000")
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'    End If
'    Screen.MousePointer = vbDefault


    Set frmHcoFacSocPre = New frmLiqHcoFacSocPrev

    frmHcoFacSocPre.DatosADevolverBusqueda = "0|1|2|3|"
    frmHcoFacSocPre.cWhere = CadB
    frmHcoFacSocPre.Show vbModal

    Set frmHcoFacSocPre = Nothing


End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Socios
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Socio
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
        CadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
    End If
    Screen.MousePointer = vbDefault
End Sub


'En vTesoreria pondremos como estaba el recibo
'Es decir. El  msgbox que pondra al final lo guardo en esta variable
Private Function ComprobarCobroArimoney(vTesoreria As String, LEtra As String, codmacta As String, Codfaccl As Long, Fecha As Date) As Boolean
Dim vR As ADODB.Recordset
Dim Cad As String


On Error GoTo EComprobarCobroArimoney
    ComprobarCobroArimoney = False
    Set vR = New ADODB.Recordset
    
    If vParamAplic.ContabilidadNueva Then
        Cad = "Select * from cobros where numserie='" & LEtra & "'"
        Cad = Cad & " AND numfactu =" & Codfaccl
        Cad = Cad & " AND codmacta =" & DBSet(codmacta, "T")
        Cad = Cad & " AND fecfactu =" & DBSet(Fecha, "F")
    Else
        Cad = "Select * from scobro where numserie='" & LEtra & "'"
        Cad = Cad & " AND codfaccl =" & Codfaccl
        Cad = Cad & " AND codmacta =" & DBSet(codmacta, "T")
        Cad = Cad & " AND fecfaccl =" & DBSet(Fecha, "F")
    End If
    '
    vTesoreria = ""
    vR.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If vR.EOF Then
        vTesoreria = "NO se ha encotrado ningun vencimiento en la tesoreria"
    Else
        While Not vR.EOF
            Cad = ""
            If DBLet(vR!codrem, "T") <> "" Then
                Cad = "El cobro asociado a la factura esta remesado(" & vR!codrem & ")"
            Else
                If DBLet(vR!recedocu, "N") = 1 Then
                    Cad = "Documento recibido"
                Else
                
                    If vParamAplic.ContabilidadNueva Then
                            
                        If DBLet(vR!transfer, "N") = 1 Then
                            Cad = "Esta en una transferencia"
                        Else
                           If DBLet(vR!impcobro, "N") > 0 Then Cad = "Esta parcialmente cobrado: " & vR!impcobro
                        
                            
                                    'Si hubeira que poner mas coas iria aqui
                        End If 'transfer
                    
                    Else
                        If DBLet(vR!Estacaja, "N") = 1 Then
                            Cad = "Cobrado por caja"
                        Else
                            If DBLet(vR!transfer, "N") = 1 Then
                                Cad = "Esta en una transferencia"
                            Else
                               If DBLet(vR!impcobro, "N") > 0 Then Cad = "Esta parcialmente cobrado: " & vR!impcobro
                            
                                
                                        'Si hubeira que poner mas coas iria aqui
                            End If 'transfer
                        End If 'estacaja
                    End If
                End If 'recdedocu
            End If 'remesado
            If Cad <> "" Then vTesoreria = vTesoreria & "Vto: " & vR!numorden & "      " & Cad & vbCrLf
            vR.MoveNext
        Wend
    End If
    vR.Close
    
    
    
    If vTesoreria <> "" Then
        Cad = vTesoreria & vbCrLf & vbCrLf
        If vUsu.Nivel > 1 Then
            MsgBox Cad, vbExclamation
        Else
            Cad = Cad & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then ComprobarCobroArimoney = True
        End If
    Else
        ComprobarCobroArimoney = True
    End If
            
EComprobarCobroArimoney:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set vR = Nothing
End Function



'En vTesoreria pondremos como estaba el recibo
'Es decir. El  msgbox que pondra al final lo guardo en esta variable
Private Function ComprobarPagoArimoney(vTesoreria As String, vCta As String, Codfaccl As Long, Fecha As Date) As Boolean
Dim vR As ADODB.Recordset
Dim Cad As String

On Error GoTo EComprobarPagoArimoney
    
    ComprobarPagoArimoney = False
    
    Set vR = New ADODB.Recordset
    
    
    If Not vParamAplic.ContabilidadNueva Then
        Cad = "Select * from spagop where ctaprove='" & vCta & "'"
        Cad = Cad & " AND numfactu =" & DBSet(ObtenerLetraSerie(Text1(1).Text) & Format(CLng(Codfaccl), "0000000"), "T")
        Cad = Cad & " AND fecfactu =" & DBSet(Fecha, "F")
    Else
        Cad = "Select * from pagos where codmacta='" & vCta & "'"
        Cad = Cad & " AND numfactu =" & DBSet(ObtenerLetraSerie(Text1(1).Text) & Format(CLng(Codfaccl), "0000000"), "T")
        Cad = Cad & " AND fecfactu =" & DBSet(Fecha, "F")
    End If
    
    '
    vTesoreria = ""
    vR.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If vR.EOF Then
        vTesoreria = "NO se ha encotrado ningun pago en la tesoreria"
    Else
        While Not vR.EOF
            Cad = ""
            If Not vParamAplic.ContabilidadNueva Then
                If DBLet(vR!Estacaja, "N") = 1 Then
                    Cad = "Pagado por caja"
                Else
                    If DBLet(vR!transfer, "N") <> 0 Then
                        Cad = "Esta en una transferencia"
                    Else
                       If DBLet(vR!imppagad, "N") > 0 Then Cad = "Esta parcialmente pagado: " & vR!impcobro
                        
                                'Si hubeira que poner mas coas iria aqui
                    End If 'transfer
                End If 'estacaja
                If Cad <> "" Then vTesoreria = vTesoreria & "Pago: " & vR!numorden & "      " & Cad & vbCrLf
            
            Else
                If DBLet(vR!nrodocum, "N") <> 0 Then
                    Cad = "Esta en una transferencia"
                Else
                   If DBLet(vR!imppagad, "N") > 0 Then Cad = "Esta parcialmente pagado: " & vR!imppagad
                    
                            'Si hubeira que poner mas coas iria aqui
                End If 'transfer
                If Cad <> "" Then vTesoreria = vTesoreria & "Pago: " & vR!numorden & "      " & Cad & vbCrLf
            
            End If
            vR.MoveNext
        Wend
    End If
    vR.Close
    
    
    
    If vTesoreria <> "" Then
        Cad = vTesoreria & vbCrLf & vbCrLf
        If vUsu.Nivel > 1 Then
            MsgBox Cad, vbExclamation
        Else
            Cad = Cad & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then ComprobarPagoArimoney = True
        End If
    Else
        ComprobarPagoArimoney = True
    End If
            
EComprobarPagoArimoney:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set vR = Nothing
End Function





Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim vSocio As CSocio
Dim vFactu As CFacturaSoc
Dim bol As Boolean

    On Error GoTo FinEliminar

'    B = False
'    If Data1.Recordset.EOF Then Exit Function
'
'    conn.BeginTrans
'
'    'Eliminar en las tablas de la Contabilidad
'    '------------------------------------------
'    Letra = ObtenerLetraSerie(Data1.Recordset!codtipom)
'
'    If Letra <> "" Then
'        SQL = " numserie='" & Letra & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
'        SQL = SQL & " AND anofaccl=" & Year(Data1.Recordset.Fields!FecFactu)
'
'        'Lineas
'        ConnConta.Execute "Delete from linfact WHERE " & SQL
'
'        'cabecera
'        ConnConta.Execute "Delete from cabfact WHERE " & SQL
'
'        'cobros
'        SQL = " numserie='" & Letra & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
'        SQL = SQL & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
'        ConnConta.Execute "Delete from scobro WHERE " & SQL
'        B = True
'    Else
'        B = False
'    End If
'
'    'Eliminar en tablas de factura de Aritaxi
'    '------------------------------------------
'    If B Then
'        SQL = " " & ObtenerWhereCP(True)
'
'
'        'Eliminar los vencimientos
'        conn.Execute "Delete from svenci " & SQL
'
'        'Cabecera de facturas (sfactusoc)
'        conn.Execute "Delete from " & NombreTabla & SQL
'
'        'Decrementar contador si borramos la ult. factura
'        Set vTipoMov = New CTiposMov
'        vTipoMov.DevolverContador Data1.Recordset!codtipom, Val(Text1(0).Text)
'        Set vTipoMov = Nothing
'    End If
'
'    B = True

        b = False
        Eliminar = False
        If Data1.Recordset.EOF Then Exit Function
        
        conn.BeginTrans
        ConnConta.BeginTrans
        
        'Eliminar en la tabla pagos de la Contabilidad: spagop
        '------------------------------------------------
'        Cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
        
        Set vSocio = New CSocio
        If vSocio.LeerDatos(Text1(4).Text) Then
            
            'antes de Eliminar en las tablas de la Contabilidad
            Set vFactu = New CFacturaSoc
            bol = vFactu.LeerDatosLiq(Text1(4).Text, Text1(1).Text, Text1(0).Text, Text1(2).Text)
            If bol Then
'[Monica]10/07/2012: tanto liquidaciones como rectificativas han de estar en la spagop
'                If Text1(1).Text = "FLI" Then

                    If vParamAplic.ContabilidadNueva Then
                        Sql = " numserie = " & DBSet(SerieFraPro, "T")
                        Sql = Sql & " AND codmacta='" & vFactu.CtaSocio & "' AND numfactu='" & ObtenerLetraSerie(CodTipoMov) & Format(Data1.Recordset.Fields!NumFactu, "0000000") & "'"
                        Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                        ConnConta.Execute "Delete from pagos WHERE " & Sql
                    
                    Else
                        Sql = " ctaprove='" & vFactu.CtaSocio & "' AND numfactu='" & ObtenerLetraSerie(CodTipoMov) & Format(Data1.Recordset.Fields!NumFactu, "0000000") & "'"
                        Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                        ConnConta.Execute "Delete from spagop WHERE " & Sql
                    End If
'                Else
'                    Sql = " codmacta='" & vFactu.CtaSocio & "' AND numserie='" & ObtenerLetraSerie(CodTipoMov) & "' AND codfaccl = " & Data1.Recordset.Fields!NumFactu & " "
'                    Sql = Sql & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
'                    ConnConta.Execute "Delete from scobro WHERE " & Sql
'                End If
                b = True
                
                'Eliminar en tablas de factura de Aritaxi: scafpc, scafpa, slifpc
                '---------------------------------------------------------------
                If b Then
                    Sql = " " & ObtenerWhereCP(True)
                
                    If Text1(1).Text = "FLI" Then
                        b = vFactu.DesmarcarLLamadas(Text1(4).Text, Text1(0).Text, Text1(2).Text)
                    End If
                
                    If b Then
                        conn.Execute "Delete from " & NomTablaLineas & Sql
                
                        'Cabecera de facturas (sfactusoc)
                        conn.Execute "Delete from " & NombreTabla & Sql
                    
                        'retenciones de la sreten
                        If Text1(1).Text = "FLI" Then
                            conn.Execute "delete from sreten where numfactu='" & Data1.Recordset.Fields!NumFactu & "'" & _
                                         " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'" & _
                                         " AND codsocio= " & Data1.Recordset.Fields!codSocio & _
                                         " AND tiporeten=0 "
                        Else
                            ' rectificativa de liquidacion de socio
                            conn.Execute "delete from sreten where numfactu='" & Data1.Recordset.Fields!NumFactu & "'" & _
                                         " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'" & _
                                         " AND codsocio= " & Data1.Recordset.Fields!codSocio & _
                                         " AND tiporeten=2 "
                        End If
                    
                    End If
                    
                End If
            End If
        
            Set vFactu = Nothing
        Else
            b = False
        End If
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
        conn.CommitTrans
        ConnConta.CommitTrans
        Eliminar = True
        
        vSocio.DevolverContador Text1(4).Text, Text1(0).Text, Text1(1).Text  ' "FLI"
    
    End If
End Function


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = " codtipom= '" & Text1(1).Text & "' and numfactu= " & Val(Text1(0).Text) & " and fecfactu='" & Format(Text1(2).Text, FormatoFecha) & "' "
    Sql = Sql & " and codsocio = " & Val(Text1(4).Text)
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function

Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGrid

    Select Case vDataGrid.Name
'        Case "DataGrid1" 'Cod. Almacen
'            'SQL = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea,
'            'codalmac, codartic, nomartic, ampliaci, cantidad,numbultos, precioar, origpre, dtoline1, dtoline2, importel "
'            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
'            tots = tots & "S|txtAux(0)|T|Alm.|500|;S|txtAux(1)|T|Artículo|1600|;S|txtAux(2)|T|Nombre Art.|3300|;"
'            tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|900|;S|txtAux(11)|T|Bultos|700|;S|txtAux(4)|T|Precio|1200|;S|txtAux(5)|T|OP|400|;S|txtAux(6)|T|Dto 1|600|;S|txtAux(7)|T|Dto 2|600|;S|txtAux(8)|T|Importe|1240|;"
'            'TRAZA
''            tots = tots & "S|txtAux(9)|T|Prov.|750|;S|cmdaux|B|||;S|txtAux(10)|T|Nom. prov.|2000|;"
''            If vEmpresa.TieneAnalitica Then
''                'codprove,nomprove, codccost
''                tots = tots & "N||||0|;S|cmdaux|B|||;N||||0|;S|txtAux(9)|T|CCoste|750|;"
''
''            Else
''                tots = tots & "S|txtAux(9)|T|Prov.|750|;S|cmdaux|B|||;N||||0|;N||||0|;"
''            End If
'            'numlote
''            tots = tots & "S|txtAux(10)|T|Nº Lote|1300|;"
'
'
'            arregla tots, DataGrid1, Me
'            DataGrid1.Columns(9).Alignment = dbgRight
'            DataGrid1.Columns(10).Alignment = dbgRight
'            DataGrid1.Columns(12).Alignment = dbgCenter
'            DataGrid1.Columns(13).Alignment = dbgRight
'            DataGrid1.Columns(14).Alignment = dbgRight
'            DataGrid1.Columns(15).Alignment = dbgRight
'            'ocultar las q no usamos
'            DataGrid1.Columns(16).visible = False
'            DataGrid1.Columns(17).visible = False
'            DataGrid1.Columns(18).visible = False
'            DataGrid1.Columns(19).visible = False
            
         Case "DataGrid1" 'albaranes x articulo
'             SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb,"
             'numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa  "
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(3)|T|Fecha|1200|;S|txtAux3(4)|T|Hora|1000|;N||||0|;S|txtAux3(5)|T|Código|950|;"
            tots = tots & "S|txtAux3(6)|T|Nombre|2500|;S|txtAux3(7)|T|Direccion|1880|;S|txtAux3(8)|T|Nro|600|;"
            tots = tots & "S|txtAux3(9)|T|Puerta|1000|;N|txtAux3(10)|T|Ciudad|1000|;S|txtAux3(11)|T|Tfno.|1200|;"
            tots = tots & "S|txtAux3(12)|T|Importe|1300|;N||||0|;"
            
            
            arregla tots, DataGrid1, Me, 350
                     
'            DataGrid1_RowColChange 1, 1
    End Select
    
    vDataGrid.HoldFields
    PonerModoOpcionesMenu Modo
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
'     B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
     b = False
     vDataGrid.Enabled = Not b
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

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
    
    Sql = "SELECT codtipom,codsocio,numfactu,fecfactu,numlinea,fecha,hora,numeruve,codclien,nomclien,dirllama,numllama,puerllama,"
    Sql = Sql & " ciudadre,telefono,impventa, observac2 " ',idservic"
    Sql = Sql & " FROM sfactusoc_serv " 'lineas de factura
    
    If enlaza Then
        Sql = Sql & " " & ObtenerWhereCP(True)
    Else
        'aNTES
        'SQL = SQL & " WHERE numfactu = -1 "
        'AHORA     Cambio sugerido por mangel para acelerar la entrada
        Sql = Sql & " WHERE codtipom is null and numfactu is null and fecfactu is null and codsocio is null  "
        If Opcion = 1 Then Sql = Sql & " AND numlinea is null"
    End If
    Sql = Sql & " ORDER BY codtipom, numfactu, fecfactu, numlinea "
    MontaSQLCarga = Sql
End Function

Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid1, Data2, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
    
    Me.SSTab1.Tab = numTab
    If Me.DataGrid1.visible Then
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "La factura no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = Cad
    End If
    If vUsu.Nivel >= 1 Then
        MsgBox "No tiene permiso", vbExclamation
        Exit Sub
    End If
    
    
    ModificaLineas = 0
    PonerModo 5
'    PonerBotonCabecera True
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), 4
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 12 ' importe del servicio
'            If Modo = 1 Then Exit Sub
            If PonerFormatoDecimal(txtAux3(Index), 3) Then cmdAceptar.SetFocus
        
    End Select
End Sub


Private Function ObtenerSelFactura() As String
Dim Cad As String
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Cad = ""
    If Me.DesdeFichaSocio Then
        '
        Cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F") & " and codsocio = " & DBSet(hcoCodSocio, "N")
        
    Else
        'Tengo YA el codigo de la factura
                '******************************************************
                'laura: esto se puede comentar, ya no hay movimiento FTI en la smoval
                If hcoCodTipoM = "FTI" Then
                    'no hay albaran directamente va a factura de ticket
                    
                    'ver si lo encontramos como factura: codtipom, numfactu,fecfactu
                    Cad = "SELECT COUNT(*) FROM sfactusoc "
                    Cad = Cad & " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    If RegistrosAListar(Cad) > 0 Then
                        Cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    Else
                        Cad = ""
                    End If
                Else
                    If hcoCodTipoM = "FAM" Then
                        Cad = "  WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    End If
                End If
                '******************************************************
                    
                If Cad = "" Then
                    'En la smoval estaba e mov. de ALbaran
                    Cad = "SELECT codtipom,numfactu,fecfactu FROM sfactusoc "
                    Cad = Cad & " WHERE codtipoa=" & DBSet(hcoCodTipoM, "T") & " AND numalbar=" & hcoCodMovim & " AND fechaalb=" & DBSet(hcoFechaMov, "F")
                    
                    Set Rs = New ADODB.Recordset
                    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not Rs.EOF Then 'where para la factura
                        Cad = " WHERE codtipom='" & Rs!codtipom & "' AND numfactu= " & Rs!NumFactu & " AND fecfactu=" & DBSet(Rs!FecFactu, "F")
                    Else
                        Cad = " WHERE numfactu=-1"
                    End If
                    Rs.Close
                    Set Rs = Nothing
                End If
    
    End If
    ObtenerSelFactura = Cad
End Function

Private Sub LimpiarDatosSocio()
Dim I As Byte
    
    For I = 4 To 12
        Text1(I).Text = ""
    Next I
    If (Modo = 3 Or Modo = 4) Then PonerFoco Text1(4)
End Sub

Private Sub BloquearDatosSocio(bol As Boolean)
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol
        Me.imgBuscar(1).Enabled = bol
        Me.imgBuscar(2).Enabled = bol
        
        For I = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(I), Not bol
        Next I
    End If
End Sub


Private Function ActualizarRetencion(Uve As String, ByRef vFac As CFacturaSoc, Inserta As Boolean) As Boolean
Dim Precio As Currency
Dim Sql As String
Dim SQL2 As String
Dim SqlValues As String
Dim linea As Long
Dim MensError As String

    On Error GoTo eInsertLinea
    
    ActualizarRetencion = False
    
    MensError = ""
    
    If Not Inserta Then
        Sql = "update sreten set impreten = " & DBSet(Text1(20).Text, "N")
        Sql = Sql & " where codsocio = " & DBSet(vFac.Socio, "N")
        Sql = Sql & " and numfactu = " & DBSet(vFac.NumFactu, "N")
        Sql = Sql & " and fecfactu = " & DBSet(vFac.FecFactu, "F")
        If Text1(1).Text = "FRL" Then
            Sql = Sql & " and tiporeten = 2"
        Else
            Sql = Sql & " and tiporeten = 0"
        End If
    Else
        Sql = "insert into sreten (codsocio, numeruve, fecfactu, numfactu, impreten, tiporeten) values ("
        Sql = Sql & DBSet(vFac.Socio, "N") & "," & DBSet(Uve, "N") & "," & DBSet(vFac.FecFactu, "F") & ","
        Sql = Sql & DBSet(vFac.NumFactu, "N") & "," & DBSet(vFac.ImpRet2, "N") & ",2)"
    End If
    
    conn.Execute Sql
    
    ActualizarRetencion = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la modificación de retencion de la factura del socio NºV " & Uve
        MuestraError Err.Number, MensError, Err.descripc
    End If
End Function


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim I As Byte
    
    Combo1.Clear
    
    '[Monica]11/02/2011: todo tipo de facturas excepto las de liquidacion,publicidad y cuotas de socio
    '                    y las facturas de cliente FAC y FPC
    Sql = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom in ('FLI','FRL')"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Sql = Rs!nomtipom
        Sql = Replace(Sql, "Factura", "")
        Combo1.AddItem Rs!codtipom & "-" & Sql
        Combo1.ItemData(Combo1.NewIndex) = I
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub


