VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPubliHcoFacCli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hist�rico de Facturas Publicidad Clientes"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   14460
   Icon            =   "frmPubliHcoFacCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   14460
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
      TabIndex        =   156
      Top             =   240
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3240
      TabIndex        =   154
      Top             =   60
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   155
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
               Object.ToolTipText     =   "�ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   152
      Top             =   60
      Width           =   3045
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   153
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
      Index           =   9
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   143
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   7785
      Visible         =   0   'False
      Width           =   6045
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   123
      Top             =   795
      Width           =   14175
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
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1305
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
         Left            =   8010
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Nombre Cliente|T|N|||scafaccli|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   300
         Width           =   5820
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
         Left            =   7125
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Cod. Cliente|N|N|0|999999|scafaccli|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   300
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
         Left            =   1290
         TabIndex        =   1
         Tag             =   "Tipo Factura|T|N|||scafaccli|codtipom||S|"
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
         Left            =   2670
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Factura|F|N|||scafaccli|fecfactu|dd/mm/yyyy|S|"
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
         Tag             =   "N� Factura|N|N|||scafaccli|numfactu|0000000|S|"
         Text            =   "Text1 7"
         Top             =   375
         Width           =   980
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
         Height          =   375
         Left            =   4110
         TabIndex        =   4
         Tag             =   "Contabilizado|N|N|0|1|scafaccli|intconta||N|"
         Top             =   285
         Width           =   1665
      End
      Begin VB.Label Label1 
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
         Left            =   6090
         TabIndex        =   127
         Top             =   330
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6855
         ToolTipText     =   "Buscar cliente"
         Top             =   330
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
         Left            =   2670
         TabIndex        =   126
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "N� Factura"
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
         Left            =   180
         TabIndex        =   125
         Top             =   120
         Width           =   1095
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
         Left            =   1320
         TabIndex        =   124
         Top             =   120
         Width           =   1275
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
      Height          =   5370
      Left            =   150
      TabIndex        =   32
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1710
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   9472
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmPubliHcoFacCli.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameFactura"
      Tab(0).Control(1)=   "FrameCliente"
      Tab(0).Control(2)=   "Text1(16)"
      Tab(0).Control(3)=   "Text1(17)"
      Tab(0).Control(4)=   "Label1(26)"
      Tab(0).Control(5)=   "Label1(25)"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmPubliHcoFacCli.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "imgBuscar(7)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(9)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(23)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(24)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(21)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(18)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(22)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(40)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "imgBuscar(8)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "imgBuscar(9)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "imgBuscar(6)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text3(15)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Text3(14)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Text2(3)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text3(3)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Text3(4)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Text3(5)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Text3(7)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Text3(6)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Text3(8)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Text2(1)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Text3(1)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Text2(2)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Text3(2)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "FrameObserva"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "DataGrid2"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "DataGrid1"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "txtAux(8)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "txtAux(7)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txtAux(6)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "txtAux(4)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Text3(0)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Text2(0)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "cmdObserva"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "txtAux(0)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "txtAux(1)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "txtAux(2)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "txtAux(3)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "txtAux(5)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "txtAux3(0)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "txtAux3(1)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "txtAux3(2)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "txtAux(9)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "txtAux(10)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "cmdaux"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "txtAux(11)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "FrameToolAux"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).ControlCount=   49
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   240
         TabIndex        =   150
         Top             =   2730
         Width           =   1500
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   151
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
      Begin VB.TextBox txtAux 
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
         Height          =   330
         Index           =   11
         Left            =   4080
         MaxLength       =   9
         TabIndex        =   144
         Tag             =   "N� Bultos|N|N|0||slifac|numbultos|#,###,##0|N|"
         Text            =   "numbultos"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdaux 
         Caption         =   "+"
         Height          =   320
         Left            =   9480
         TabIndex        =   117
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox txtAux 
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
         Height          =   330
         Index           =   10
         Left            =   9720
         MaxLength       =   15
         TabIndex        =   139
         Tag             =   "N� Lote|T|S|||slifac|numlote||N|"
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
         Index           =   9
         Left            =   8880
         MaxLength       =   30
         TabIndex        =   116
         Tag             =   "Cod. Proveedor|N|N|||slifac|codprovex|0||"
         Text            =   "prove"
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
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
         Height          =   330
         Index           =   2
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   122
         Tag             =   "Fecha Albaran|F|N|||scafac1|fechaalb|dd/mm/yyyy|N|"
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
         Index           =   1
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   121
         Tag             =   "N� Albaran|N|N|||scafac1|numalbar|0000000|N|"
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
         Index           =   0
         Left            =   360
         MaxLength       =   7
         TabIndex        =   120
         Tag             =   "Tipo Albaran|T|N|||scafac1|codtipoa||N|"
         Text            =   "codtipoa"
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Height          =   330
         Index           =   5
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   111
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
         Index           =   3
         Left            =   3240
         MaxLength       =   12
         TabIndex        =   109
         Tag             =   "Cantidad|N|N|0||slifac|cantidad|#,###,###,##0.00|N|"
         Text            =   "cantidad"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
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
         Height          =   330
         Index           =   2
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   108
         Tag             =   "Nombre Art.|T|N|||slifac|nomartic||N|"
         Text            =   "nomartic"
         Top             =   4320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
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
         Height          =   330
         Index           =   1
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   107
         Tag             =   "Art.|T|N|||slifac|codartic||N|"
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
         Index           =   0
         Left            =   360
         MaxLength       =   12
         TabIndex        =   106
         Tag             =   "Almacen|N|N|0|999|slifac|codalmac|000|N|"
         Text            =   "almacen"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdObserva 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   930
         Visible         =   0   'False
         Width           =   375
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
         Index           =   0
         Left            =   7500
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   87
         Text            =   "Text2"
         Top             =   420
         Width           =   5265
      End
      Begin VB.TextBox Text3 
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
         Left            =   6780
         MaxLength       =   30
         TabIndex        =   25
         Tag             =   "Trabajador Albaran|N|N|0|9999|scafac1|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   420
         Width           =   660
      End
      Begin VB.Frame FrameFactura 
         Height          =   2160
         Left            =   -74820
         TabIndex        =   55
         Top             =   2790
         Width           =   13725
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
            Left            =   7350
            MaxLength       =   5
            TabIndex        =   135
            Tag             =   "% IVA 1|N|S|0|99.90|scafaccli|porciva3re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1695
            Width           =   735
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
            Index           =   43
            Left            =   8175
            MaxLength       =   15
            TabIndex        =   134
            Tag             =   "Importe IVA 1|N|S|||scafaccli|imporiv3re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1695
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
            Index           =   42
            Left            =   7350
            MaxLength       =   5
            TabIndex        =   133
            Tag             =   "% IVA 1|N|S|0|99.90|scafaccli|porciva2re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1335
            Width           =   735
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
            Index           =   41
            Left            =   8175
            MaxLength       =   15
            TabIndex        =   132
            Tag             =   "Importe IVA 1|N|S|||scafaccli|imporiv2re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1335
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
            Index           =   40
            Left            =   7350
            MaxLength       =   5
            TabIndex        =   131
            Tag             =   "% IVA 1|N|S|0|99.90|scafaccli|porciva1re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   735
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
            Index           =   39
            Left            =   8175
            MaxLength       =   15
            TabIndex        =   130
            Tag             =   "Importe IVA 1|N|S|||scafaccli|imporiv1re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
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
            Index           =   38
            Left            =   10215
            MaxLength       =   15
            TabIndex        =   80
            Tag             =   "Total Factura|N|N|||scafaccli|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1350
            Width           =   1755
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
            Index           =   37
            Left            =   5745
            MaxLength       =   15
            TabIndex        =   75
            Tag             =   "Importe IVA 3|N|S|||scafaccli|imporiv3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1695
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
            Index           =   31
            Left            =   4740
            MaxLength       =   5
            TabIndex        =   74
            Tag             =   "% IVA 3|N|S|0|99.90|scafaccli|porciva3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1695
            Width           =   780
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
            Index           =   28
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   73
            Tag             =   "Cod. IVA 3|N|S|0|9999|scafaccli|codigiv3|0000|N|"
            Text            =   "Text1 7"
            Top             =   1695
            Width           =   555
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
            Index           =   34
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   72
            Tag             =   "Base Imponible 3|N|S|||scafaccli|baseimp3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1695
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
            Index           =   36
            Left            =   5745
            MaxLength       =   15
            TabIndex        =   71
            Tag             =   "Importe IVA 2|N|S|||scafaccli|imporiv2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1335
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
            Index           =   30
            Left            =   4740
            MaxLength       =   5
            TabIndex        =   70
            Tag             =   "% IVA 2|N|S|0|99.90|scafaccli|porciva2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1335
            Width           =   780
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
            Index           =   27
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   69
            Tag             =   "Cod. IVA 2|N|S|0|9999|scafaccli|codigiv2|0000|N|"
            Text            =   "Text1 7"
            Top             =   1335
            Width           =   555
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
            Index           =   33
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   68
            Tag             =   "Base Imponible 2 |N|S|||scafaccli|baseimp2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1335
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
            Index           =   35
            Left            =   5745
            MaxLength       =   15
            TabIndex        =   67
            Tag             =   "Importe IVA 1|N|N|||scafaccli|imporiv1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
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
            Index           =   29
            Left            =   4740
            MaxLength       =   5
            TabIndex        =   66
            Tag             =   "% IVA 1|N|S|0|99.90|scafaccli|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   780
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
            Index           =   26
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   65
            Tag             =   "Cod. IVA 1|N|S|0|9999|scafaccli|codigiv1|0000|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   555
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
            Index           =   32
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   64
            Tag             =   "Base Imponible 1|N|N|||scafaccli|baseimp1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
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
            Index           =   25
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   59
            Text            =   "Text1 7"
            Top             =   375
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
            Index           =   24
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   58
            Tag             =   "Imp. Dto Gn|N|N|||scafaccli|impdtogr|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   375
            Width           =   1395
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
            Index           =   23
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   57
            Tag             =   "Imp. Dto PP|N|N|||scafaccli|impdtopp|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   375
            Width           =   1395
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
            Left            =   240
            MaxLength       =   15
            TabIndex        =   56
            Tag             =   "Imp.Bruto|N|N|||scafaccli|brutofac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Importe RE"
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
            Left            =   8190
            TabIndex        =   138
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
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
            Index           =   43
            Left            =   7350
            TabIndex        =   137
            Top             =   720
            Width           =   765
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
            Left            =   5775
            TabIndex        =   136
            Top             =   720
            Width           =   1215
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
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   42
            Left            =   600
            TabIndex        =   119
            Top             =   1320
            Width           =   1515
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
            Left            =   4740
            TabIndex        =   118
            Top             =   720
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
            Left            =   10215
            TabIndex        =   83
            Top             =   1110
            Width           =   1890
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
            Left            =   9900
            TabIndex        =   82
            Top             =   1350
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
            TabIndex        =   81
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base imponible"
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
            Left            =   3030
            TabIndex        =   79
            Top             =   720
            Width           =   1560
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
            Left            =   5520
            TabIndex        =   78
            Top             =   390
            Width           =   135
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
            Left            =   3720
            TabIndex        =   77
            Top             =   390
            Width           =   135
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
            Left            =   1920
            TabIndex        =   76
            Top             =   390
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
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
            Left            =   5760
            TabIndex        =   63
            Top             =   120
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
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
            Left            =   4080
            TabIndex        =   62
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
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
            Left            =   2280
            TabIndex        =   61
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
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
            Left            =   240
            TabIndex        =   60
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.TextBox txtAux 
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
         Height          =   330
         Index           =   4
         Left            =   4800
         MaxLength       =   12
         TabIndex        =   110
         Tag             =   "Precio|N|N|0|999999.0000|slifac|precioar|###,##0.0000|N|"
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
         Index           =   6
         Left            =   6360
         MaxLength       =   5
         TabIndex        =   112
         Tag             =   "Dto 1|N|N|0|99.90|slifac|dtoline1|#0.00|N|"
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
         Index           =   7
         Left            =   7080
         MaxLength       =   30
         TabIndex        =   113
         Tag             =   "Dto 2|N|N|0|99.90|slifac|dtolinea|#0.00|N|"
         Text            =   "Dto2"
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Height          =   330
         Index           =   8
         Left            =   7680
         MaxLength       =   12
         TabIndex        =   115
         Tag             =   "Importe|N|N|0||slifac|importel|#,###,###,##0.00|N|"
         Text            =   "Importe"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Cliente"
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
         Height          =   2295
         Left            =   -74820
         TabIndex        =   34
         Top             =   420
         Width           =   13725
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
            Left            =   4890
            MaxLength       =   4
            TabIndex        =   18
            Tag             =   "IBAN|T|S|||scafaccli|iban|||"
            Text            =   "Text1 7"
            Top             =   1770
            Width           =   705
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
            Left            =   9840
            MaxLength       =   10
            TabIndex        =   145
            Tag             =   "Aportacion|N|S|||scafaccli|portes|#,##0.00|N|"
            Text            =   "Portes"
            Top             =   1770
            Visible         =   0   'False
            Width           =   1065
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
            Left            =   11040
            MaxLength       =   10
            TabIndex        =   140
            Tag             =   "Aportacion|N|S|||scafaccli|aportacion|#,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1770
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
            Index           =   21
            Left            =   7980
            MaxLength       =   10
            TabIndex        =   22
            Tag             =   "Cuenta Bancaria|T|S|||scafaccli|cuentaba|0000000000|N|"
            Text            =   "Text1 7"
            Top             =   1755
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
            Index           =   20
            Left            =   7410
            MaxLength       =   2
            TabIndex        =   21
            Tag             =   "Digito Control|T|S|||scafaccli|digcontr|00|N|"
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   465
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
            Index           =   19
            Left            =   6540
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "Sucursal|N|S|0|9999|scafaccli|codsucur|0000|N|"
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   705
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
            Left            =   5700
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "Banco|N|S|0|9999|scafaccli|codbanco|0000|N|"
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   705
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
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
            Left            =   1425
            MaxLength       =   20
            TabIndex        =   13
            Tag             =   "Refere. Cliente|T|S|||scafac1|referenc|||"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   1830
            Width           =   1725
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
            Left            =   8340
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   15
            Tag             =   "Direccion/Dpto.|T|S|||scafaccli|nomdirec||N|"
            Text            =   "Text1"
            Top             =   285
            Width           =   4305
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
            Index           =   12
            Left            =   7755
            MaxLength       =   3
            TabIndex        =   14
            Tag             =   "Direccion/Dpto.|N|S|0|999|scafaccli|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   285
            Width           =   540
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
            Index           =   11
            Left            =   1425
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||scafaccli|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1410
            Width           =   2445
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
            Left            =   1425
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "CPostal|T|N|||scafaccli|codpobla||N|"
            Text            =   "Text15"
            Top             =   1020
            Width           =   840
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
            Left            =   2295
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Poblaci�n|T|N|||scafaccli|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1020
            Width           =   3405
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
            Index           =   7
            Left            =   3675
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "tel�fono Cliente|T|S|||scafaccli|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   285
            Width           =   2025
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
            Index           =   6
            Left            =   1425
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "NIF Cliente|T|N|||scafaccli|nifclien||N|"
            Text            =   "123456789"
            Top             =   285
            Width           =   1230
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
            Index           =   14
            Left            =   7755
            MaxLength       =   4
            TabIndex        =   16
            Tag             =   "Cod. Agente|N|N|0|9999|scafaccli|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   675
            Width           =   540
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
            Index           =   14
            Left            =   8340
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   38
            Text            =   "Text2"
            Top             =   675
            Width           =   4305
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
            Left            =   7755
            MaxLength       =   3
            TabIndex        =   17
            Tag             =   "Forma de Pago|N|N|0|999|scafaccli|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   1080
            Width           =   540
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
            Left            =   8340
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   36
            Text            =   "Text2"
            Top             =   1080
            Width           =   4305
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
            Left            =   1425
            MaxLength       =   35
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||scafaccli|domclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   645
            Width           =   4275
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
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
            Index           =   47
            Left            =   4890
            TabIndex        =   157
            Top             =   1530
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Aportaci�n"
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
            Index           =   45
            Left            =   11040
            TabIndex        =   141
            Top             =   1530
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   7470
            ToolTipText     =   "Buscar agente"
            Top             =   705
            Width           =   240
         End
         Begin VB.Label Label1 
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
            Left            =   7980
            TabIndex        =   54
            Top             =   1530
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "DC"
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
            Left            =   7500
            TabIndex        =   53
            Top             =   1530
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal"
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
            Left            =   6480
            TabIndex        =   52
            Top             =   1530
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Banco"
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
            Left            =   5700
            TabIndex        =   51
            Top             =   1530
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
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
            Left            =   120
            TabIndex        =   45
            Top             =   1830
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   1155
            ToolTipText     =   "Buscar poblaci�n"
            Top             =   1035
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc./Dpto"
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
            Left            =   6240
            TabIndex        =   44
            Top             =   285
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   3
            Left            =   7470
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   315
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
            Left            =   120
            TabIndex        =   43
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
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
            TabIndex        =   42
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
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
            Left            =   2745
            TabIndex        =   41
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
            TabIndex        =   40
            Top             =   285
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1155
            ToolTipText     =   "Buscar cliente varios"
            Top             =   300
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
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
            Index           =   34
            Left            =   6240
            TabIndex        =   39
            Top             =   675
            Width           =   765
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
            Left            =   6240
            TabIndex        =   37
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   7470
            ToolTipText     =   "Buscar forma de pago"
            Top             =   1110
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
            TabIndex        =   35
            Top             =   645
            Width           =   915
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmPubliHcoFacCli.frx":0044
         Height          =   1905
         Left            =   240
         TabIndex        =   50
         Top             =   3330
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   3360
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmPubliHcoFacCli.frx":0059
         Height          =   1950
         Left            =   240
         TabIndex        =   84
         Top             =   780
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   3440
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
      Begin VB.Frame FrameObserva 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   2145
         Left            =   4050
         TabIndex        =   99
         Tag             =   "Observaci�n 4|T|S|||scafac1|observa4||N|"
         Top             =   690
         Width           =   9945
         Begin VB.TextBox Text3 
            BeginProperty Font 
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
            Left            =   450
            MaxLength       =   80
            TabIndex        =   104
            Tag             =   "Observaci�n 5|T|S|||scafac1|observa5||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1680
            Width           =   9270
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
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
            Left            =   450
            MaxLength       =   80
            TabIndex        =   103
            Tag             =   "Observaci�n 4|T|S|||scafac1|observa4||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1320
            Width           =   9270
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
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
            Left            =   450
            MaxLength       =   80
            TabIndex        =   102
            Tag             =   "Observaci�n 3|T|S|||scafac1|observa3||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   960
            Width           =   9270
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
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
            Left            =   450
            MaxLength       =   80
            TabIndex        =   101
            Tag             =   "Observaci�n 2|T|S|||scafac1|observa2||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   600
            Width           =   9270
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
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
            Left            =   450
            MaxLength       =   80
            TabIndex        =   100
            Tag             =   "Observaci�n 1|T|S|||scafac1|observa1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   240
            Width           =   9270
         End
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   5010
         MaxLength       =   30
         TabIndex        =   27
         Tag             =   "Preparador materia|N|N|0|9999|scafac1|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   3675
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   5730
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   85
         Text            =   "Text2"
         Top             =   3675
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   5010
         MaxLength       =   30
         TabIndex        =   26
         Tag             =   "Trabajador pedido|N|S|0|9999|scafac1|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   3915
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   5730
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   86
         Text            =   "Text2"
         Top             =   3915
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   11970
         MaxLength       =   10
         TabIndex        =   88
         Tag             =   "Semana Entrega|N|S|||scafac1|sementre||N|"
         Top             =   3840
         Width           =   705
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   9690
         MaxLength       =   7
         TabIndex        =   89
         Tag             =   "N� Pedido|N|S|||scafac1|numpedcl|0000000|N|"
         Text            =   "Text1 7"
         Top             =   3840
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   10650
         MaxLength       =   10
         TabIndex        =   90
         Tag             =   "Fecha Pedido|F|S|||scafac1|fecpedcl|dd/mm/yyyy|N|"
         Top             =   3840
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   11370
         MaxLength       =   10
         TabIndex        =   91
         Tag             =   "Fecha Oferta|F|S|||scafac1|fecofert|dd/mm/yyyy|N|"
         Top             =   3390
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   10170
         MaxLength       =   7
         TabIndex        =   92
         Tag             =   "N� Oferta|N|S|||scafac1|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   3390
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   5010
         MaxLength       =   30
         TabIndex        =   28
         Tag             =   "Cod. Env�o|N|N|0|999|scafac1|codenvio|000|N|"
         Text            =   "Text1"
         Top             =   4230
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   5730
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   98
         Text            =   "Text2"
         Top             =   4230
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   14
         Left            =   11250
         MaxLength       =   7
         TabIndex        =   128
         Tag             =   "N� Venta|N|S|||scafac1|numventa|0000000|N|"
         Text            =   "Text1 7"
         Top             =   3870
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   15
         Left            =   10050
         MaxLength       =   7
         TabIndex        =   129
         Tag             =   "N� Terminal|N|S|||scafac1|numtermi||N|"
         Text            =   "Text1 7"
         Top             =   3870
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   16
         Left            =   -62835
         MaxLength       =   5
         TabIndex        =   146
         Tag             =   "Descuento P.Pago|N|N|0|99.90|scafaccli|dtoppago|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   900
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   17
         Left            =   -62835
         MaxLength       =   5
         TabIndex        =   147
         Tag             =   "Descuento General|N|N|0|99.90|scafaccli|dtognral|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   1260
         Width           =   525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   6450
         ToolTipText     =   "Buscar trabajador"
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   4770
         ToolTipText     =   "Buscar forma de envio"
         Top             =   4230
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   4770
         ToolTipText     =   "Buscar trabajador"
         Top             =   3750
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "N� Oferta"
         Height          =   255
         Index           =   40
         Left            =   10050
         TabIndex        =   97
         Top             =   3630
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
         Height          =   255
         Index           =   22
         Left            =   11490
         TabIndex        =   96
         Top             =   3630
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         Height          =   255
         Index           =   18
         Left            =   10650
         TabIndex        =   95
         Top             =   3675
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "N� Pedido"
         Height          =   255
         Index           =   6
         Left            =   9810
         TabIndex        =   94
         Top             =   3675
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Sem. Entrega"
         Height          =   255
         Index           =   2
         Left            =   11850
         TabIndex        =   93
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Albar�n"
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
         Index           =   21
         Left            =   4440
         TabIndex        =   49
         Top             =   435
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo  Env�o"
         Height          =   195
         Index           =   24
         Left            =   3330
         TabIndex        =   48
         Top             =   4230
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Prepar. Material"
         Height          =   255
         Index           =   23
         Left            =   3330
         TabIndex        =   47
         Top             =   3750
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Pedido"
         Height          =   255
         Index           =   9
         Left            =   3330
         TabIndex        =   46
         Top             =   3960
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   4770
         ToolTipText     =   "Buscar trabajador"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Dto. Gral"
         Height          =   255
         Index           =   26
         Left            =   -63525
         TabIndex        =   149
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Dto. P.P"
         Height          =   255
         Index           =   25
         Left            =   -63540
         TabIndex        =   148
         Top             =   900
         Width           =   615
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
      Index           =   16
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   114
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   7410
      Visible         =   0   'False
      Width           =   7125
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   150
      TabIndex        =   30
      Top             =   7245
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
         TabIndex        =   31
         Top             =   150
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
      Left            =   13170
      TabIndex        =   24
      Top             =   7320
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
      Left            =   11880
      TabIndex        =   23
      Top             =   7320
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
      Left            =   13170
      TabIndex        =   29
      Top             =   7320
      Visible         =   0   'False
      Width           =   1135
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
      Index           =   46
      Left            =   2400
      TabIndex        =   142
      Top             =   7815
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliaci�n L�nea"
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
      Index           =   35
      Left            =   2400
      TabIndex        =   33
      Top             =   7170
      Visible         =   0   'False
      Width           =   1335
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
         Caption         =   "&Rectifica"
         Enabled         =   0   'False
         Shortcut        =   ^R
         Visible         =   0   'False
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
         Caption         =   "Imprimir &albar�n"
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
Attribute VB_Name = "frmPubliHcoFacCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 603

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

Private WithEvents frmHcoFacCliPre As frmPubliHcoFacCliPrev
Attribute frmHcoFacCliPre.VB_VarHelpID = -1


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
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en funci�n del valor en la tabla de par�metros: stipom

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
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Private UnaVez As Boolean
Private BuscaChekc As String

Dim cadB1 As String



Private Sub Check1_Click()
    If Modo = 1 Then CheckCadenaBusqueda Check1, BuscaChekc
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 4  'MODIFICAR
            If DatosOk Then
               If ModificarFactura Then
               
                                        
                    Set LOG = New cLOG
                    LOG.Insertar 8, vUsu, "Factura modificada: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & vbCrLf
                    Set LOG = Nothing
               
               
               
               
                    Espera 0.2
                    TerminaBloquear
                    PosicionarData
                    FormatoDatosTotales
                    I = Data3.Recordset.AbsolutePosition
                    PonerCamposLineas
                    SituarDataPosicion Data3, CLng(I), ""
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            If ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                
                        'INSERTA LOG
                        '-------------------------------------------------
                        Set LOG = New cLOG
                        BuscaChekc = "     Alb : " & Data2.Recordset!NumAlbar & "   Linea: " & Data2.Recordset!numlinea
                        BuscaChekc = "Modificar linea: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & BuscaChekc
                        LOG.Insertar 8, vUsu, BuscaChekc
                        Set LOG = Nothing
                        BuscaChekc = ""
                
                    TerminaBloquear
                    CargaGrid DataGrid1, Data2, True
                    ModificaLineas = 0
'                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
            
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


Private Sub cmdAux_Click()
'    Set frmP = New frmComProveedores
'    frmP.DatosADevolverBusqueda = "0|1|"
'    frmP.Show vbModal
'    Set frmP = Nothing

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
'            PonerBotonCabecera True
            PonerModo 2
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

            Me.DataGrid1.Enabled = True
    End Select
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        BuscaChekc = ""
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
        'poner los txtaux para buscar por lineas de albaran
        anc = DataGrid2.top
        If DataGrid2.Row < 0 Then
            anc = anc + 440
        Else
            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid2"
        
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue 'vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue 'vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    
    cadB1 = "scafaccli.codtipom in ('FPC','FRP') "


    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia "" & cadB1
    Else
        lblIndicador.Caption = "Preparando bus."
        lblIndicador.Refresh
        LimpiarCampos
        LimpiarDataGrids
        DoEvents
        
        CadenaConsulta = "Select scafaccli.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE scafaccli.codtipom='" & CodTipoMov & "'"
        CadenaConsulta = CadenaConsulta & " where " & cadB1
        lblIndicador.Caption = "Obteniendo reg."
        lblIndicador.Refresh
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean
Dim EnTesoreria  As String
    'solo se puede modificar la factura si no esta contabilizada
    
'   [Monica]18/02/2011: no contabilizamos las facturas
    If FactContabilizada2(EnTesoreria) Then
        TerminaBloquear
        Exit Sub
    End If
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFocoChk Me.Check1
        
    'Inserto en slog
    
    Set LOG = New cLOG
    If EnTesoreria <> "" Then EnTesoreria = "Tesoreria: " & vbCrLf & EnTesoreria
    EnTesoreria = Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & vbCrLf & EnTesoreria
    EnTesoreria = "Pulsa mod factura: " & EnTesoreria
    LOG.Insertar 8, vUsu, EnTesoreria
    Set LOG = Nothing
    Espera 0.3
    '
    
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
'        [Monica]18/02/2011: no contabilizamos las facturas
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
    Set LOG = New cLOG
    If EstaEnTesoreria <> "" Then EstaEnTesoreria = "Tesoreria: " & EstaEnTesoreria
    EstaEnTesoreria = "     Alb : " & Data2.Recordset!NumAlbar & "   Linea: " & Data2.Recordset!numlinea & vbCrLf & EstaEnTesoreria
    EstaEnTesoreria = "Pulsa mod linea: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & EstaEnTesoreria
    LOG.Insertar 8, vUsu, EstaEnTesoreria
    Set LOG = Nothing




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

    For J = 0 To 2
        txtAux(J).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(16).Text = DataGrid1.Columns(J + 5).Text
    
    'cantidad
    J = 4
    txtAux(J - 1).Text = DataGrid1.Columns(J + 5).Text
    'num bultos
    J = 5
    txtAux(11).Text = DataGrid1.Columns(J + 5).Text
    
    J = 4
    For J = J + 1 To 9
        txtAux(J - 1).Text = DataGrid1.Columns(J + 6).Text
    Next J
    
    txtAux(9).Text = DataGrid1.Columns(16).Text
    txtAux2(9).Text = DataGrid1.Columns(17).Text
    'num lote
    txtAux(10).Text = DataGrid1.Columns(19).Text
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
'    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    PonerFoco txtAux(4)
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
                If jj = 4 Then 'Or (jj >= 6 And jj <= 10) Then
                    txtAux(jj).Height = DataGrid1.RowHeight
                    txtAux(jj).top = alto
                    txtAux(jj).visible = b
                End If
            Next jj
'            cmdaux.Top = alto
'            cmdaux.visible = b
            
        Case "DataGrid2"
            DeseleccionaGrid Me.DataGrid2
            b = (xModo = 1)
             For jj = 0 To txtAux3.Count - 1
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).top = alto
                txtAux3(jj).visible = b
            Next jj
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
'        [Monica]18/02/2011: no contabilizamos las facturas
    If FactContabilizada3(EstaEnTesoreria) Then Exit Sub
    
    cad = "Cabecera de Facturas." & vbCrLf
    cad = cad & "-------------------------------------      " & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar la Factura:            "
    cad = cad & vbCrLf & "Tipo:  " & Text1(1).Text
    cad = cad & vbCrLf & "N� Fact.:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(2).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " �Desea Eliminarla? "

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
    If Modo <> 2 And Modo <> 4 And Modo <> 1 Then Exit Sub
    If Me.FrameObserva.visible = False Then
        Me.DataGrid1.visible = False
        Me.FrameObserva.visible = True
        Me.cmdObserva.Picture = frmppal.imgListComun1.ListImages(18).Picture
'        CargarICO Me.cmdObserva, "volver.ico"
        Me.cmdObserva.ToolTipText = "volver lineas albaran"
        BloqueaText3
    Else
        Me.DataGrid1.visible = True
        Me.FrameObserva.visible = False
        Me.cmdObserva.Picture = frmppal.imgListComun1.ListImages(41).Picture
'        CargarICO Me.cmdObserva, "message.ico"
        Me.cmdObserva.ToolTipText = "ver observaciones albaran"
    End If
    SSTab1_Click 0
End Sub

Private Sub BloqueaText3()
Dim I As Byte
Dim b As Boolean

    'bloquear los Text3 que son las lineas de scafaccli1
    'b = Modo <> 4 And Modo <> 1
    b = Modo <> 1
    For I = 0 To 3
        BloquearTxt Text3(I), b
    Next I
    BloquearTxt Text3(16), b
    
    
'    If Me.FrameObserva.visible Then
        For I = 9 To 13
            BloquearTxt Text3(I), (Modo <> 4 And Modo <> 1)
        Next I
'    End If
    
    b = Modo <> 1
    For I = 4 To 8
        BloquearTxt Text3(I), b
    Next I
    'datos venta TPV
    BloquearTxt Text3(14), True
    BloquearTxt Text3(15), True
 
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

    Else 'Se llama desde alg�n Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
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
'Ayuda de Etiqueta de precio de salida de la Funci�n de Precios
On Error Resume Next

    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        Me.DataGrid1.ToolTipText = ""
        If X > 7790 And X < 8170 Then
            Select Case DataGrid1.Columns(11).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoci�n"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Art�culo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Art�culo"
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
    On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        If ModificaLineas <> 1 Then
            Text2(16).Text = DBLet(Data2.Recordset.Fields!Ampliaci, "T")
            If vEmpresa.TieneAnalitica Then
                '- centro de coste
                ' ---- [19/10/2009] [LAURA]: a�adir campo centro de coste familia
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
Dim I As Byte

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
        For I = 0 To Text3.Count - 1
            Text3(I).Text = ""
        Next I
        For I = 0 To 3
            Text2(I).Text = ""
        Next I
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, False
    End If
    
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
    Me.Icon = frmppal.Icon
    
     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next kCampo

'    ' ICONITOS DE LA BARRA
'    btnPrimero = 15
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Bot�n Buscar
'        .Buttons(2).Image = 2   'Bot�n Todos
'        .Buttons(5).Image = 4   'Modificar
'        .Buttons(6).Image = 5   'Borrar
'        .Buttons(9).Image = 10 'Mto Lineas
'        .Buttons(10).Image = 16 'Imprimir
'        .Buttons(11).Image = 40 'Imprimir albaran
'        .Buttons(13).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
'    End With
    
    With Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.imgListComun1
        'ASignamos botones
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2 'Ver Todos
        .Buttons(1).Image = 3 'A�adir
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
    
    
    Me.SSTab1.Tab = 0
    LimpiarCampos   'Limpia los campos TextBox
    
    
    CargaCombo
    
    Combo1.ListIndex = 0
    cadB1 = "scafaccli.codtipom in ('FPC','FRP') "
    
    
    'cargar icono de observaciones de los albaranes de factura
    Me.cmdObserva.Picture = frmppal.imgListComun1.ListImages(41).Picture
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
    
'    'Mariela 01/07/2010 consulto si estoy llamando desde publicidad
'    If publicidad Then
'        CadenaConsulta = "Select * from " & NombreTabla & " where codtipom='FPC'"
'    Else
        CadenaConsulta = "Select * from " & NombreTabla
'    End If
    
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        'como tenemos aqui el n� de albaran, buscar a que factura corresponde
        'en la scafaccli1
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura
'        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        'CadenaConsulta = CadenaConsulta & " where numfactu=-1"
        'Cambio sugerido por Msoler
        'mariela 01/07/2010
        'If Not publicidad Then
        CadenaConsulta = CadenaConsulta & " WHERE codtipom is null and numfactu is null and fecfactu is null " & " and " & cadB1
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

End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1.Value = 0
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
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
            CadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
            CadB = CadB & " and " & Aux
            Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
            CadB = CadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
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


Private Sub frmHcoFacCliPre_DatoSeleccionado(CadenaSeleccion As String)
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
Dim devuelve As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim ImprimeDirecto As Boolean

    cadFormula = "({scafaccli.codtipom}= '" & Data1.Recordset!codtipom & "' and {scafaccli.numfactu}= "
    cadFormula = cadFormula & Data1.Recordset!NumFactu & " and {scafaccli.fecfactu}= Date(" & Year(Data1.Recordset!FecFactu) & "," & Month(Data1.Recordset!FecFactu) & "," & Day(Data1.Recordset!FecFactu) & "))"
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    cadParam = cadParam & "pDuplicado= 1|"
    numParam = 2
    
    
    devuelve = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
    
    indRPT = 47 'Facturas Publicidad Clientes
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, ImprimeDirecto, pPdfRpt) Then Exit Sub
    
    
    If Not ImprimeDirecto Then frmImprimir.NombreRPT = nomDocu

    
    With frmImprimir
        '[Monica]18/01/2018: a�adido
        .outClaveNombreArchiv = devuelve & Format(Text1(0).Text, "000")
        .outCodigoCliProv = Text1(4).Text
        .outTipoDocumento = 2
        
        .Titulo = "Impresi�n de Facturas de publicidad"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .NombreRPT = nomDocu
        .NombrePDF = pPdfRpt

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
                If BloqueaLineasFac Then BotonModificarLinea
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


Private Sub mnRectifica_Click()
    If Modo = 5 Then 'A�adir lineas
'         BotonAnyadirLinea
    Else 'A�adir Cabecera
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub

Private Sub BotonAnyadir()
'A�adir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String
Dim cad As String
Dim Rs As ADODB.Recordset

    LimpiarCampos 'Vac�a los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    NomTraba = ""

    Text1(2).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    Text1(1).Text = "FRP"
    Text1(15).Text = vParamAplic.IVA_REA
    Text1(17).Text = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", vParamAplic.IVA_REA, "N")
    Text1(21).Text = vParamAplic.PorReten
    PonerFoco Text1(2)
End Sub



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
                Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
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
        
        Case 22, 23, 24, 32, 33, 34, 35, 36, 37, 38, 39, 41, 43
            PonerFormatoDecimal Text1(Index), 1
            
        Case 29, 30, 31, 40, 42, 44
            PonerFormatoDecimal Text1(Index), 7


    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim cadAux As String
    
    '--- Laura 12/01/2007
    cadAux = Text1(5).Text
    If Text1(4).Text <> "" Then Text1(5).Text = ""
    
    '---
    
    CadB = ObtenerBusqueda(Me, False, BuscaChekc)
    If CadB = "" Then
        CadB = CadB & " where " & cadB1
    Else
        CadB = CadB & " and " & cadB1
    End If
    
    
    '--- David.  No se pq referencia NO lleva tag. Si han puesto algo lo paso a la cadena de busqueda
    
    
    
    '--- Laura 12/01/2007
    Text1(5).Text = cadAux
    '---
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
'        cadB = cadB & " and scafaccli.codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los FAV
        CadenaConsulta = "select scafaccli.* from " & NombreTabla & " INNER JOIN scafaccli1 ON scafaccli.codtipom=scafaccli1.codtipom AND scafaccli.numfactu=scafaccli1.numfactu AND scafaccli.fecfactu=scafaccli1.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY scafaccli.codtipom,scafaccli.numfactu,scafaccli.fecfactu " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    If EsCabecera Then
'        Cad = Cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
'        Cad = Cad & ParaGrid(Text1(0), 15, "N� Factura")
'        Cad = Cad & ParaGrid(Text1(2), 15, "Fecha Fac.")
'        Cad = Cad & ParaGrid(Text1(4), 10, "Cliente")
'        Cad = Cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
''      If publicidad Then
''        Tabla = NombreTabla & " INNER JOIN scafaccli1 ON scafaccli.codtipom='FPC' and scafaccli.codtipom=scafaccli1.codtipom AND scafaccli.numfactu=scafaccli1.numfactu AND scafaccli.fecfactu=scafaccli1.fecfactu "
''      Else
'        Tabla = NombreTabla & " INNER JOIN scafaccli1 ON scafaccli.codtipom=scafaccli1.codtipom AND scafaccli.numfactu=scafaccli1.numfactu AND scafaccli.fecfactu=scafaccli1.fecfactu "
''      End If
'        'CadenaConsulta = "select scafaccli.* from " & NombreTabla & " INNER JOIN scafaccli1 ON scafaccli.codtipom=scafaccli1.codtipom AND scafaccli.numfactu=scafacli1.numfactu AND scafaccli.fecfactu=scafaccli1.fecfactu "
'        'CadenaConsulta = CadenaConsulta & " WHERE " & cadB & " GROUP BY scafaccli.codtipom,scafaccli.numfactu,scafaccli.fecfactu " & Ordenacion
'
'        Titulo = "Facturas"
'        devuelve = "0|1|2|"

        
        Set frmHcoFacCliPre = New frmPubliHcoFacCliPrev
    
        frmHcoFacCliPre.DatosADevolverBusqueda = "0|1|2|3|"
        frmHcoFacCliPre.cWhere = CadB
        frmHcoFacCliPre.Show vbModal
    
        Set frmHcoFacCliPre = Nothing


    Else
        If vParamAplic.Departamento Then
            Titulo = "Dptos Cliente: "
            Desc = "Dpto."
        Else
            Titulo = "Direc. Cliente: "
            Desc = "Direc."
        End If
        Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N||15�"
        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35�"
        Tabla = "sdirec"
        devuelve = "0|1|"
    End If
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri  'Conexi�n a BD: Aritaxi
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
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
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
    Label1(6).Caption = "N� Pedido"
    Label1(18).Caption = "Fecha Pedido"
    If b Then
        If b2 Then
            Label1(6).Caption = "N� Ticket"
            Label1(18).Caption = "Fecha Ticket"
        End If
        Label1(40).Caption = "N� Terminal"
        Label1(22).Caption = "N� Venta"
    Else
        Label1(40).Caption = "N� Oferta"
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
Dim I As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general

    For I = 0 To txtAux.Count - 1
        txtAux(I).BackColor = vbWhite
    Next I

'    ActualizarToolbar Modo, Kmodo
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
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
     DesplazamientoVisible b And Data1.Recordset.RecordCount > 1

        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
    Me.Check1.Enabled = (Modo = 1)
    
    b = (Modo <> 1)
    'Campos N� Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(3), b 'referencia
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(3), b 'referencia
    BloquearTxt Text1(4), b 'cliente
    BloquearTxt Text1(12), b 'direccion
    BloquearTxt Text1(14), b 'agente
    BloquearTxt Text1(13), b 'direccion / departamento
    
    For I = 18 To 21
        BloquearTxt Text1(I), b
    Next I
    
    'Importes siempre bloqueados, excepto para busquedas. ivas y aportacion tb bloqueado
    For I = 22 To 45
        BloquearTxt Text1(I), (Modo <> 1)
    Next I
    'Campo B.Imp y Imp. IVA siempre en azul
    BloquearTxt Text1(25), True
    Text1(25).BackColor = &HFFFFC0
    
    If Modo <> 1 Then
        Text1(35).BackColor = &HFFFFC0
        Text1(36).BackColor = &HFFFFC0
        Text1(37).BackColor = &HFFFFC0
'    Text1(38).BackColor = &HC0C0FF    'Total factura
        Text1(38).BackColor = &HC0FFC0
    End If
    
    'bloquear los Text3 que son las lineas de scafaccli1
    BloqueaText3
    If Modo = 1 Then
        'Busqueda. Habilitamos numero pedido y fecha pedido
'        BloquearTxt Text3(6), False
'        BloquearTxt Text3(7), False
'        BloquearTxt Text3(16), False
    End If
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
    BloquearTxt txtAux(8), True
    BloquearTxt txtAux(10), True
    BloquearTxt txtAux(11), True
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los albaranes de factura)
    For I = 0 To txtAux3.Count - 1
        BloquearTxt txtAux3(I), (Modo <> 1)
    Next I
    
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


    Me.Combo1.visible = (Modo = 1)

    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
'    For I = 0 To 5
'        Me.imgBuscar(I).Enabled = b
'    Next I
'    For I = 6 To 9
'        Me.imgBuscar(I).Enabled = b And (Modo <> 1)
'    Next I
'
'    Me.imgBuscar(1).visible = False
    For I = 0 To 5
        Me.imgBuscar(I).Enabled = (Modo = 1)
        If I = 5 Then Me.imgBuscar(I).Enabled = (Modo = 1 Or Modo = 4)

    Next I
    For I = 6 To 9
        Me.imgBuscar(I).Enabled = False 'B And (Modo <> 1)
    Next I
    
    Me.imgBuscar(1).visible = False
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario
    PonerModoUsuarioGnral Modo, "aritaxi"
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerModoUsuarioGnral(Modo As Byte, Aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(Aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(2).Enabled = Toolbar1.Buttons(2).Enabled And DBLet(Rs!Modificar, "N")
        Toolbar1.Buttons(3).Enabled = Toolbar1.Buttons(3).Enabled And DBLet(Rs!creareliminar, "N")
        
        Toolbar1.Buttons(5).Enabled = Toolbar1.Buttons(5).Enabled And DBLet(Rs!Ver, "N")
        Toolbar1.Buttons(6).Enabled = Toolbar1.Buttons(6).Enabled And DBLet(Rs!Ver, "N")
        
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

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
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
Dim I As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For I = 0 To txtAux.Count - 1
        If I = 4 Or I = 6 Or I = 7 Then
            If txtAux(I).Text = "" Then
                MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        End If
    Next I
            
    'PRoveedor
    If txtAux(9).Text <> "" And txtAux2(9).Text = "" Then
        MsgBox "Codigo proveedor incorrecto", vbExclamation
        PonerFoco txtAux(9)
        b = False
        Exit Function
    End If
            
'    'Febrero 2010   Si han apretado Alt+A NO recalculaba
'    '----------------------------------------------------------------------------------
'    'txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
'    Aux = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
'    Aux = Format(Aux, FormatoImporte)
'    If Aux <> txtAux(8).Text Then txtAux(8).Text = Aux
    RecalcularImportes txtAux(8), True, txtAux(3), txtAux(4), txtAux(6), txtAux(7)

            
            
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
        Case 2: mnModificar_Click  'Modificar
    End Select



End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5: mnBuscar_Click  'Buscar
        Case 6: BotonVerTodos  'Todos
            

        Case 1: mnRectifica_Click ' Insertar rectificativa
        Case 2: mnModificar_Click  'Modificar
        Case 3: mnEliminar_Click  'Borrar
        
'        Case 9: mnLineas_Click  'Lineas
        Case 8: mnImprimir_Click 'Imprimir Albaran
        
'        Case 11: mnImprimirAlbaran_Click
'
'        Case 13: mnSalir_Click    'Salir
'
'        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
'            Desplazamiento (Button.Index - btnPrimero)
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
    vWhere = vWhere & " AND codtipoa='" & Data3.Recordset.Fields!codtipoa & "' "
    vWhere = vWhere & " AND numalbar=" & Data3.Recordset.Fields!NumAlbar
    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
    
    If DatosOkLinea() Then
        Sql = "UPDATE slifaccli SET "
        Sql = Sql & " ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        Sql = Sql & "precioar = " & DBSet(txtAux(4).Text, "N") & ", "
        Sql = Sql & "dtoline1= " & DBSet(txtAux(6).Text, "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N") & ", "
        Sql = Sql & "importel = " & DBSet(txtAux(8).Text, "N") & ", "
        Sql = Sql & "origpre='" & txtAux(5) & "',"
        'TRAZA
        Sql = Sql & "codprovex= " & DBSet(txtAux(9).Text, "N", "S")
        Sql = Sql & vWhere
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




Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    If vDataGrid.Name = "DataGrid1" Then
        Opcion = 1
    Else
        Opcion = 2
    End If
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
' cambiado como en hco de facturas de clientes
'            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
'            tots = tots & "S|txtAux(0)|T|Alm.|500|;S|txtAux(1)|T|Art�culo|1600|;S|txtAux(2)|T|Nombre Art.|3300|;"
'            tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|900|;S|txtAux(11)|T|Bultos|700|;S|txtAux(4)|T|Precio|1200|;S|txtAux(5)|T|OP|400|;S|txtAux(6)|T|Dto 1|600|;S|txtAux(7)|T|Dto 2|600|;S|txtAux(8)|T|Importe|1240|;"
'
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(0)|T|Alm.|500|;S|txtAux(1)|T|Art�culo|1600|;S|txtAux(2)|T|Nombre Art.|5050|;"
            tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|900|;N||||0|;S|txtAux(4)|T|Precio|1200|;S|txtAux(5)|T|OP|400|;S|txtAux(6)|T|Dto 1|600|;S|txtAux(7)|T|Dto 2|600|;S|txtAux(8)|T|Importe|2240|;"
            
            
            'TRAZA
'            tots = tots & "S|txtAux(9)|T|Prov.|750|;S|cmdaux|B|||;S|txtAux(10)|T|Nom. prov.|2000|;"
            If vEmpresa.TieneAnalitica Then
                'codprove,nomprove, codccost
                tots = tots & "N||||0|;S|cmdaux|B|||;N||||0|;S|txtAux(9)|T|CCoste|750|;"

            Else
'                tots = tots & "S|txtAux(9)|T|Prov.|750|;S|cmdaux|B|||;N||||0|;N||||0|;"
                tots = tots & "N||||0|;S|cmdaux|B|||;N||||0|;N||||0|;"
            End If
            'numlote
'            tots = tots & "S|txtAux(10)|T|N� Lote|1300|;"
            tots = tots & "N||||0|;"

            
            arregla tots, DataGrid1, Me, 350
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
            
            
            arregla tots, DataGrid2, Me, 350
                     
            DataGrid2_RowColChange 1, 1
    End Select
    
    vDataGrid.HoldFields
    
    PonerModoOpcionesMenu Modo
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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
             If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 2
            
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
                        MsgBox "Campo proveedor debe ser num�rico", vbExclamation
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
        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(8), 1
    End If
End Sub




Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
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
'        [Monica]18/02/2011: no contabilizamos las facturas
'        SQL = " numserie='" & LEtra & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
'        SQL = SQL & " AND anofaccl=" & Year(Data1.Recordset.Fields!FecFactu)
'
'        'Lineas
'        ConnConta.Execute "Delete from linfact WHERE " & SQL
'
'        'cabecera
'        ConnConta.Execute "Delete from cabfact WHERE " & SQL
        
        'cobros
        If vParamAplic.ContabilidadNueva Then
            Sql = " numserie='" & LEtra & "' AND numfactu=" & Data1.Recordset.Fields!NumFactu
            Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
            ConnConta.Execute "Delete from cobros WHERE " & Sql
        
        Else
            Sql = " numserie='" & LEtra & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
            Sql = Sql & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
            ConnConta.Execute "Delete from scobro WHERE " & Sql
        End If
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
        
        'Eliminar los vencimientos
        conn.Execute "Delete from svencicli " & Sql
        
        'Cabecera de facturas (scafaccli)
        conn.Execute "Delete from " & NombreTabla & Sql
        
        'Decrementar contador si borramos la ult. factura
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador Data1.Recordset!codtipom, Val(Text1(0).Text)
        Set vTipoMov = Nothing
    End If
    
    b = True
    
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
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ning�n registro
On Error Resume Next

    CargaGrid DataGrid2, Data3, False
    CargaGrid DataGrid1, Data2, False
    
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
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    If Opcion = 1 Then
        Sql = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,numbultos, precioar, origpre, dtoline1, dtoline2, importel ,codprovex, nomprove,codccost,numlote"
        Sql = Sql & " FROM slifaccli left join sprove on codprovex=codprove " 'lineas de factura
    ElseIf Opcion = 2 Then
        Sql = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb, numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa  "
        Sql = Sql & " FROM scafaccli1 " 'cabeceras albaranes de la factura
    End If
    
    If enlaza Then
        Sql = Sql & " " & ObtenerWhereCP(True)
        If Opcion = 1 Then Sql = Sql & " AND numalbar=" & Data3.Recordset.Fields!NumAlbar
    Else
        'aNTES
        'SQL = SQL & " WHERE numfactu = -1 "
        'AHORA     Cambio sugerido por mangel para acelerar la entrada
        Sql = Sql & " WHERE codtipom is null and numfactu is null and fecfactu is null and codtipoa is null and numalbar is null "
        If Opcion = 1 Then Sql = Sql & " AND numlinea is null"
    End If
    Sql = Sql & " ORDER BY codtipom, numfactu, fecfactu,numalbar "
    If Opcion = 1 Then Sql = Sql & ", numlinea "
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean
Dim I As Integer
Dim bAux As Boolean

    b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = (Modo = 2)
    Me.mnEliminar.Enabled = (Modo = 2)
        
    b = (Modo = 2)
    'Mantenimiento lineas
'        Toolbar1.Buttons(9).Enabled = b
'        Me.mnLineas.Enabled = b
    'Imprimir
    Toolbar1.Buttons(10).Enabled = b
    Me.mnImprimir.Enabled = b
'        Toolbar1.Buttons(11).Enabled = b
'        mnImprimirAlbaran.Enabled = b
    
    b = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnvertodos.Enabled = Not b

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
Dim I As Byte
    
    For I = 4 To 13
        Text1(I).Text = ""
    Next I
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
    If (OpcionListado = 53) Then
        If Text1(1).Text = "FAZ" Then
            'Factura B
            indRPT = 30
        ElseIf Text1(1).Text = "FRP" Then
                indRPT = 47
        Else
            indRPT = 12 'Facturas Clientes
        End If
    Else
        '-----------------------------------------------
        indRPT = 18 'Facturas Clientes TPV
    End If
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
    'Cadena para seleccion N� de Factura
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Cod Tipo Movimiento
        devuelve = "{" & NombreTabla & ".codtipom}='" & Text1(1).Text & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        'N� Factura
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
                .Titulo = ""
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
'        MsgBox "No se ha podido establecer la impresora de ticket." & vbCrLf & "Debe configurar primero los par�metros del TPV.", vbExclamation
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
'si se ha modificado la linea de slifaccli, a�adir a la transaccion la modificaci�n de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim Sql As String, LEtra As String
Dim vFactura As CFactura
Dim recalcular As Boolean

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
    '    bol = RecalcularFactura
        bol = CalcularDatosFactura
        
    '    bol = True
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
                    If vParamAplic.ContabilidadNueva Then
                        'Eliminar de la scobro
                        Sql = " numserie='" & vFactura.LetraSerie & "' AND numfactu=" & Data1.Recordset.Fields!NumFactu
                        Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                        ConnConta.Execute "Delete from cobros WHERE " & Sql
                    
                    Else
                        'Eliminar de la scobro
                        Sql = " numserie='" & vFactura.LetraSerie & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
                        Sql = Sql & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                        ConnConta.Execute "Delete from scobro WHERE " & Sql
                    End If
                    bol = True

                    'Volvemos a Insertar los Vencimientos de la Factura. Tabla: svenci
                    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
                    If bol Then
                        vFactura.Agente = Text1(14).Text
                        bol = vFactura.InsertarEnTesoreriaFACcli("", MenError)
                    End If
                End If
                Set vFactura = Nothing
            End If
'            End If
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
Dim vFactu As CFactura
Dim FacOK As Boolean

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 22 To 38
         Text1(I).Text = ""
    Next I
    
    
    Set vFactu = New CFactura
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(16).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(17).Text))
    vFactu.Cliente = Text1(4).Text
    
    If vFactu.CalcularDatosFactura(True, ObtenerWhereCP(False), NombreTabla, NomTablaLineas) Then
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


Private Sub FormatoDatosTotales()
Dim I As Byte

    For I = 22 To 25
        Text1(I).Text = QuitarCero(Text1(I).Text)
        Text1(I).Text = Format(Text1(I).Text, FormatoImporte)
    Next I
    
    'Desglose B.Imponible por IVA
    For I = 32 To 34
        If Text1(I).Text <> "" Then
             If CSng(Text1(I).Text) = 0 And Text1(I - 6).Text = "" Then
                Text1(I).Text = QuitarCero(Text1(I).Text)
                Text1(I - 3).Text = QuitarCero(Text1(I - 3).Text)
                Text1(I - 6).Text = QuitarCero(Text1(I - 6).Text)
                Text1(I + 3).Text = QuitarCero(Text1(I).Text)
            Else
                Text1(I).Text = Format(Text1(I).Text, FormatoImporte)
                Text1(I - 3) = Format(Text1(I - 3).Text, FormatoDescuento)
    '            Text3(i - 6) = Format(Text3(i - 6).Text, "000")
                Text1(I + 3).Text = Format(Text1(I + 3).Text, FormatoImporte)
            End If
        End If
    Next I
End Sub



Private Sub ComprobarDatosTotales()
Dim I As Byte

    For I = 22 To 25
        Text1(I).Text = ComprobarCero(Text1(I).Text)
    Next I
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
    If Me.Check1.Value = 1 Then 'si esta contabilizada
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
        If numasien <> "" Then LEtra = LEtra & vbCrLf & "N� asiento: " & numasien
        LEtra = LEtra & vbCrLf & vbCrLf & "�Continuar?"
        
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
    If Me.Check1.Value = 1 Then 'si esta contabilizada
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
Dim Rs As ADODB.Recordset

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
                    
                    Set Rs = New ADODB.Recordset
                    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not Rs.EOF Then 'where para la factura
                        cad = " WHERE codtipom='" & Rs!codtipom & "' AND numfactu= " & Rs!NumFactu & " AND fecfactu=" & DBSet(Rs!FecFactu, "F")
                    Else
                        cad = " WHERE numfactu=-1"
                    End If
                    Rs.Close
                    Set Rs = Nothing
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


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim I As Byte
    
    Combo1.Clear
    
    Sql = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom in  ('FPC','FRP')"
    
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
    'Cadena para seleccion N� de Factura
    '---------------------------------------------------
    
        'Cod Tipo Movimiento
        devuelve = "{" & NombreTabla & ".codtipom}='" & Text1(1).Text & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        'N� Factura
        devuelve = "{" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        
        
        'cODTIPOA
        devuelve = "{scafaccli1.codtipoa}=" & DBSet(Data3.Recordset!codtipoa, "T")
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        'Numalbar
        devuelve = "{scafaccli1.numalbar}=" & DBSet(Data3.Recordset!NumAlbar, "N")
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        
        
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
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(4).Text, "N")
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
                .Titulo = "Albar�n facturado"
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
    If vParamAplic.ContabilidadNueva Then
        cad = "Select * from cobros where numserie='" & LEtra & "'"
        cad = cad & " AND numfactu =" & Codfaccl
        cad = cad & " AND fecfactu =" & DBSet(Fecha, "F")
    Else
        cad = "Select * from scobro where numserie='" & LEtra & "'"
        cad = cad & " AND codfaccl =" & Codfaccl
        cad = cad & " AND fecfaccl =" & DBSet(Fecha, "F")
    End If
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
                    If vParamAplic.ContabilidadNueva Then
                        If DBLet(vR!transfer, "N") = 1 Then
                            cad = "Esta en una transferencia"
                        Else
                           If DBLet(vR!impcobro, "N") > 0 Then cad = "Esta parcialmente cobrado: " & vR!impcobro
                        
                            
                                    'Si hubeira que poner mas coas iria aqui
                        End If 'transfer
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
                    End If
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
            cad = cad & "�Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then ComprobarCobroArimoney = True
        End If
    Else
        ComprobarCobroArimoney = True
    End If
            
EComprobarCobroArimoney:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set vR = Nothing
End Function
'uff me entere que no podr� cogerme la baja alli y aqui no...
'grrrrrrrr asi que la semana del 02/08 tendr� que apechugar
'por ahi me hago un certificado de 48hs y otro d�a voy al medico
'asi que mira tendr� posibilidades de no trabajar toda la semana
'como esta, asi que nada ya veremos que tal...
'por ahora tendr� q aguantarme y en todo caso pedir la baja por
'riesgo de embarazo en los dos trabajos y arreglar para continuar
'trabajando aqui pero desde casa...
'aunque empiezo la semana 20 que es cuando te dan de baja
'empiezo esa semana el 16/10/2010 asi que si aqui me renovaron
'pedir� la baja por riesgo y a currar desde casa a hablar por tel
'y a venir una vez por semana para reunirme con ellos y ver
'trabajos... si no les viene mal, seria una buena idea...
'por lo menos para poder pillar la baja en los dos y luego
'estar en ono hasta la vuelta de la maternidad mas las vacaciones
'que seria para agosto cuando me quedara solo con esto
'q ya estariam'os en el pisico nuevo y con el bebuco de 5 meses
'jijijijiji q mocion... a veces tengo la sensaci�n que sera Iker
'jajaja no se porque... ya no digo la patita, porque pienso en
'que puede ser caracol... y tengo ese sentimiento raro que sera
'un varon... aunque a mi me gustaria una nena, pero es igual
'mientras sea sanito...
'si las cosas salen bien ya tendremos a la patita mas adelante
'ufff tengo un calor hoy, toy sudando un monton... y no huelo
'muy bien, grrrrr eso que anoche me ba�e y hoy taba limpita
'pero me parece que me equivoque con la blusa que me puse, la de
'hilo del corte ingl�s, y es muy pesada y calurosa, asi que nada
'cuando llegue a casa me la cambiar� y me pondr� algo mas fresco
'solo que quedan 3 jornadas y media para mis vacacionessssss
'jejeje son las 11.08hs y ya he hecho bastante en el programa
'aunque estoy esperando una contestaci�n de manolo para poder
'seguir... mientras descanso un poquete... miguel angel me dijo
'que el programa de la competencia tiene un fallo con lo del iva
'y estaba super contento porque el nuestro no, jijijijiji
'mejor ya veremos que tal, esperemos que todo salga bien...
'quiero de verdad seguir en este trabajo y adem�s tener a mi bebe
'y poder continuar con media jornada m�s tiempo, y cuidar a mi
'bebe mientras aman trabaja y entrena...
'ayer mi angelito hizo macarrones con tomate y atun para hoy, es
'increible, no podria ser mejor de lo que es, viene a verme a
'estar media hora conmigo en la cena, y me trae lo que le pido sin
'problemas, la verdad si pidiera alguien mejor me dirian que no
'existe, lo amo con locura, jamas me arrepentire de estar con el
'fue la mejor decision de mundo... me alegra muchisimo, saber que
'por fin encontre el amor que siempre busque... y lo mejor �l me
'ama como yo a �l...
'ufff quiero ir a casa ya a cambiarme, me siento sudada... grrrr
'mmm tengo frioooo, el aire esta a 19 afuera hace mucho calor
'pero aqui tengo frio... ufff que mal no?
'tengo noniiii, uff quiero q sea viernes a las 23.30 y este llendo
'a casita a dormir.... porfiiii madre mia que sue�o que tengo no
'puedo mas snif snif snif snif, pasado ma�ana entro en la 7�
'semana y cuando vuelva de vacaciones ya esta en la 8� ahhhhh q
'mocion, y en nada a ver a nuestro bebucoooo, faltan 29 d�as...

