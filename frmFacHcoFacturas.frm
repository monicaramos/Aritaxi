VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFacHcoFacturas2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas Socios"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   13545
   Icon            =   "frmFacHcoFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   13545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   158
      Top             =   90
      Width           =   3195
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   159
         Top             =   180
         Width           =   2805
         _ExtentX        =   4948
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
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3360
      TabIndex        =   156
      Top             =   90
      Width           =   885
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   150
         TabIndex        =   157
         Top             =   180
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir Albarán"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4290
      TabIndex        =   154
      Top             =   90
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
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
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
      Left            =   8580
      TabIndex        =   153
      Top             =   300
      Width           =   1605
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
      TabIndex        =   147
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   7785
      Visible         =   0   'False
      Width           =   6045
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   127
      Top             =   855
      Width           =   13095
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
         Left            =   1440
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
         Tag             =   "Nombre Cliente|T|N|||scafac|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   330
         Width           =   4770
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
         Tag             =   "Cod. Socio|N|N|0|999999|scafac|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   330
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
         Left            =   1440
         TabIndex        =   1
         Tag             =   "Tipo Factura|T|N|||scafac|codtipom||S|"
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
         Left            =   2790
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Factura|F|N|||scafac|fecfactu|dd/mm/yyyy|S|"
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
         Tag             =   "Nº Factura|N|N|||scafac|numfactu|0000000|S|"
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
         Left            =   4440
         TabIndex        =   4
         Tag             =   "Contabilizado|N|N|0|1|scafac|intconta||N|"
         Top             =   285
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
         Left            =   6240
         TabIndex        =   131
         Top             =   330
         Width           =   585
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6855
         ToolTipText     =   "Buscar cliente"
         Top             =   360
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
         Left            =   2790
         TabIndex        =   130
         Top             =   120
         Width           =   1275
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
         TabIndex        =   129
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
         Left            =   1440
         TabIndex        =   128
         Top             =   120
         Width           =   1095
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
      Height          =   5400
      Left            =   150
      TabIndex        =   34
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1725
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   9525
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmFacHcoFacturas.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameCliente"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameFactura"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmFacHcoFacturas.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameToolAux"
      Tab(1).Control(1)=   "txtAux(11)"
      Tab(1).Control(2)=   "cmdaux"
      Tab(1).Control(3)=   "txtAux(10)"
      Tab(1).Control(4)=   "txtAux(9)"
      Tab(1).Control(5)=   "Text3(15)"
      Tab(1).Control(6)=   "Text3(14)"
      Tab(1).Control(7)=   "txtAux3(2)"
      Tab(1).Control(8)=   "txtAux3(1)"
      Tab(1).Control(9)=   "txtAux3(0)"
      Tab(1).Control(10)=   "txtAux(5)"
      Tab(1).Control(11)=   "txtAux(3)"
      Tab(1).Control(12)=   "txtAux(2)"
      Tab(1).Control(13)=   "txtAux(1)"
      Tab(1).Control(14)=   "txtAux(0)"
      Tab(1).Control(15)=   "cmdObserva"
      Tab(1).Control(16)=   "Text2(3)"
      Tab(1).Control(17)=   "Text3(3)"
      Tab(1).Control(18)=   "Text3(4)"
      Tab(1).Control(19)=   "Text3(5)"
      Tab(1).Control(20)=   "Text3(7)"
      Tab(1).Control(21)=   "Text3(6)"
      Tab(1).Control(22)=   "Text3(8)"
      Tab(1).Control(23)=   "Text2(0)"
      Tab(1).Control(24)=   "Text3(0)"
      Tab(1).Control(25)=   "Text2(1)"
      Tab(1).Control(26)=   "Text3(1)"
      Tab(1).Control(27)=   "Text2(2)"
      Tab(1).Control(28)=   "Text3(2)"
      Tab(1).Control(29)=   "txtAux(4)"
      Tab(1).Control(30)=   "txtAux(6)"
      Tab(1).Control(31)=   "txtAux(7)"
      Tab(1).Control(32)=   "txtAux(8)"
      Tab(1).Control(33)=   "DataGrid1"
      Tab(1).Control(34)=   "DataGrid2"
      Tab(1).Control(35)=   "FrameObserva"
      Tab(1).Control(36)=   "imgBuscar(6)"
      Tab(1).Control(37)=   "imgBuscar(9)"
      Tab(1).Control(38)=   "imgBuscar(8)"
      Tab(1).Control(39)=   "Label1(40)"
      Tab(1).Control(40)=   "Label1(22)"
      Tab(1).Control(41)=   "Label1(18)"
      Tab(1).Control(42)=   "Label1(6)"
      Tab(1).Control(43)=   "Label1(2)"
      Tab(1).Control(44)=   "Label1(21)"
      Tab(1).Control(45)=   "Label1(24)"
      Tab(1).Control(46)=   "Label1(23)"
      Tab(1).Control(47)=   "Label1(9)"
      Tab(1).Control(48)=   "imgBuscar(7)"
      Tab(1).ControlCount=   49
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   -74760
         TabIndex        =   151
         Top             =   2640
         Width           =   1500
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   152
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
         Height          =   315
         Index           =   11
         Left            =   -70920
         MaxLength       =   9
         TabIndex        =   148
         Tag             =   "Nº Bultos|N|S|0||slifac|numbultos|#,###,##0|N|"
         Text            =   "numbultos"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdaux 
         Caption         =   "+"
         Height          =   320
         Left            =   -65520
         TabIndex        =   121
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
         Height          =   315
         Index           =   10
         Left            =   -65280
         MaxLength       =   15
         TabIndex        =   143
         Tag             =   "Nº Lote|T|S|||slifac|numlote||N|"
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
         Height          =   315
         Index           =   9
         Left            =   -66120
         MaxLength       =   30
         TabIndex        =   120
         Tag             =   "Cod. Proveedor|N|N|||slifac|codprovex|0||"
         Text            =   "prove"
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
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
         Index           =   15
         Left            =   -63840
         MaxLength       =   7
         TabIndex        =   133
         Tag             =   "Nº Terminal|N|S|||scafac1|numtermi||N|"
         Text            =   "Text1 7"
         Top             =   2040
         Visible         =   0   'False
         Width           =   885
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
         Index           =   14
         Left            =   -62640
         MaxLength       =   7
         TabIndex        =   132
         Tag             =   "Nº Venta|N|S|||scafac1|numventa|0000000|N|"
         Text            =   "Text1 7"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1185
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
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   126
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
         Height          =   315
         Index           =   1
         Left            =   -73920
         MaxLength       =   15
         TabIndex        =   125
         Tag             =   "Nº Albaran|N|N|||scafac1|numalbar|0000000|N|"
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
         Height          =   315
         Index           =   0
         Left            =   -74640
         MaxLength       =   7
         TabIndex        =   124
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
         Height          =   315
         Index           =   5
         Left            =   -69360
         MaxLength       =   5
         TabIndex        =   115
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
         Height          =   315
         Index           =   3
         Left            =   -71760
         MaxLength       =   12
         TabIndex        =   113
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
         Height          =   315
         Index           =   2
         Left            =   -72840
         MaxLength       =   12
         TabIndex        =   112
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
         Height          =   315
         Index           =   1
         Left            =   -73680
         MaxLength       =   12
         TabIndex        =   111
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
         Height          =   315
         Index           =   0
         Left            =   -74640
         MaxLength       =   12
         TabIndex        =   110
         Tag             =   "Almacen|N|N|0|999|slifac|codalmac|000|N|"
         Text            =   "almacen"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdObserva 
         Height          =   375
         Left            =   -71040
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   520
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
         Index           =   3
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   102
         Text            =   "Text2"
         Top             =   2160
         Width           =   3525
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
         Index           =   3
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   30
         Tag             =   "Cod. Envío|N|N|0|999|scafac1|codenvio|000|N|"
         Text            =   "Text1"
         Top             =   2160
         Width           =   660
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
         Index           =   4
         Left            =   -63840
         MaxLength       =   7
         TabIndex        =   96
         Tag             =   "Nº Oferta|N|S|||scafac1|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   1800
         Visible         =   0   'False
         Width           =   885
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
         Index           =   5
         Left            =   -62640
         MaxLength       =   10
         TabIndex        =   95
         Tag             =   "Fecha Oferta|F|S|||scafac1|fecofert|dd/mm/yyyy|N|"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1185
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
         Index           =   7
         Left            =   -63240
         MaxLength       =   10
         TabIndex        =   94
         Tag             =   "Fecha Pedido|F|S|||scafac1|fecpedcl|dd/mm/yyyy|N|"
         Top             =   960
         Visible         =   0   'False
         Width           =   1185
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
         Index           =   6
         Left            =   -64200
         MaxLength       =   7
         TabIndex        =   93
         Tag             =   "Nº Pedido|N|S|||scafac1|numpedcl|0000000|N|"
         Text            =   "Text1 7"
         Top             =   960
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   -61920
         MaxLength       =   10
         TabIndex        =   92
         Tag             =   "Semana Entrega|N|S|||scafac1|sementre||N|"
         Top             =   960
         Visible         =   0   'False
         Width           =   705
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
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   91
         Text            =   "Text2"
         Top             =   480
         Width           =   3525
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
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   27
         Tag             =   "Trabajador Albaran|N|N|0|9999|scafac1|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   660
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
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   90
         Text            =   "Text2"
         Top             =   1040
         Visible         =   0   'False
         Width           =   3525
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
         Index           =   1
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   28
         Tag             =   "Trabajador pedido|N|S|0|9999|scafac1|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   1040
         Visible         =   0   'False
         Width           =   660
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
         Index           =   2
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   89
         Text            =   "Text2"
         Top             =   1600
         Width           =   3525
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
         Index           =   2
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   29
         Tag             =   "Preparador materia|N|N|0|9999|scafac1|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   1600
         Width           =   660
      End
      Begin VB.Frame FrameFactura 
         Height          =   2130
         Left            =   300
         TabIndex        =   59
         Top             =   2580
         Width           =   12495
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
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   139
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva3re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   525
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
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   138
            Tag             =   "Importe IVA 1|N|S|||scafac|imporiv3re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   1485
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
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   137
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva2re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1335
            Width           =   525
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
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   136
            Tag             =   "Importe IVA 1|N|S|||scafac|imporiv2re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1335
            Width           =   1485
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
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   135
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
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
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   134
            Tag             =   "Importe IVA 1|N|S|||scafac|imporiv1re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
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
            Left            =   9750
            MaxLength       =   15
            TabIndex        =   84
            Tag             =   "Total Factura|N|N|||scafac|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1350
            Width           =   1785
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
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   79
            Tag             =   "Importe IVA 3|N|S|||scafac|imporiv3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   1485
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
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   78
            Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   525
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
            Left            =   2220
            MaxLength       =   4
            TabIndex        =   77
            Tag             =   "Cod. IVA 3|N|S|0|9999|scafac|codigiv3|0000|N|"
            Text            =   "Text1 7"
            Top             =   1725
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
            Index           =   34
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   76
            Tag             =   "Base Imponible 3|N|S|||scafac|baseimp3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   1485
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
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   75
            Tag             =   "Importe IVA 2|N|S|||scafac|imporiv2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1335
            Width           =   1485
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
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   74
            Tag             =   "% IVA 2|N|S|0|99.90|scafac|porciva2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1335
            Width           =   525
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
            Left            =   2220
            MaxLength       =   4
            TabIndex        =   73
            Tag             =   "Cod. IVA 2|N|S|0|9999|scafac|codigiv2|0000|N|"
            Text            =   "Text1 7"
            Top             =   1335
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
            Index           =   33
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   72
            Tag             =   "Base Imponible 2 |N|S|||scafac|baseimp2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1335
            Width           =   1485
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
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   71
            Tag             =   "Importe IVA 1|N|N|||scafac|imporiv1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
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
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   70
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
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
            Left            =   2220
            MaxLength       =   4
            TabIndex        =   69
            Tag             =   "Cod. IVA 1|N|S|0|9999|scafac|codigiv1|0000|N|"
            Text            =   "Text1 7"
            Top             =   960
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
            Index           =   32
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   68
            Tag             =   "Base Imponible 1|N|N|||scafac|baseimp1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
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
            TabIndex        =   63
            Text            =   "Text1 7"
            Top             =   345
            Width           =   1485
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
            TabIndex        =   62
            Tag             =   "Imp. Dto Gn|N|N|||scafac|impdtogr|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   375
            Width           =   1365
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
            TabIndex        =   61
            Tag             =   "Imp. Dto PP|N|N|||scafac|impdtopp|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   375
            Width           =   1365
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
            TabIndex        =   60
            Tag             =   "Imp.Bruto|N|N|||scafac|brutofac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   360
            Width           =   1485
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
            Left            =   7560
            TabIndex        =   142
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "%RE"
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
            Left            =   6720
            TabIndex        =   141
            Top             =   720
            Width           =   495
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
            Left            =   5340
            TabIndex        =   140
            Top             =   720
            Width           =   1215
         End
         Begin VB.Line Line1 
            X1              =   2040
            X2              =   2040
            Y1              =   1020
            Y2              =   2040
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
            Left            =   570
            TabIndex        =   123
            Top             =   1320
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "%IVA"
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
            Left            =   4560
            TabIndex        =   122
            Top             =   720
            Width           =   615
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
            Left            =   9750
            TabIndex        =   87
            Top             =   1110
            Width           =   1800
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
            Left            =   9390
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   83
            Top             =   720
            Width           =   1455
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
            TabIndex        =   82
            Top             =   240
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
            TabIndex        =   81
            Top             =   240
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
            TabIndex        =   80
            Top             =   240
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
            Left            =   5790
            TabIndex        =   67
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
            TabIndex        =   66
            Top             =   120
            Width           =   1245
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
            TabIndex        =   65
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
            TabIndex        =   64
            Top             =   120
            Width           =   1725
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
         Height          =   315
         Index           =   4
         Left            =   -70200
         MaxLength       =   12
         TabIndex        =   114
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
         Height          =   315
         Index           =   6
         Left            =   -68640
         MaxLength       =   5
         TabIndex        =   116
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
         Height          =   315
         Index           =   7
         Left            =   -67920
         MaxLength       =   30
         TabIndex        =   117
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
         Height          =   315
         Index           =   8
         Left            =   -67320
         MaxLength       =   12
         TabIndex        =   119
         Tag             =   "Importe|N|N|0||slifac|importel|#,###,###,##0.00|N|"
         Text            =   "Importe"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Socios"
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
         Height          =   2235
         Left            =   300
         TabIndex        =   36
         Top             =   360
         Width           =   12705
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
            Left            =   5700
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "IBAN|T|S|||scafac|iban||N|"
            Text            =   "Text1 7"
            Top             =   1785
            Width           =   645
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
            Left            =   9900
            MaxLength       =   10
            TabIndex        =   149
            Tag             =   "Aportacion|N|S|||scafac|portes|#,##0.00|N|"
            Text            =   "Portes"
            Top             =   1800
            Visible         =   0   'False
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
            Index           =   45
            Left            =   10950
            MaxLength       =   10
            TabIndex        =   144
            Tag             =   "Aportacion|N|S|||scafac|aportacion|#,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1485
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
            Left            =   8430
            MaxLength       =   10
            TabIndex        =   24
            Tag             =   "Cuenta Bancaria|T|S|||scafac|cuentaba|0000000000|N|"
            Text            =   "0000000000"
            Top             =   1785
            Width           =   1365
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
            Left            =   7980
            MaxLength       =   2
            TabIndex        =   23
            Tag             =   "Digito Control|T|S|||scafac|digcontr|00|N|"
            Text            =   "Text1 7"
            Top             =   1785
            Width           =   405
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
            Left            =   7110
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "Sucursal|N|S|0|9999|scafac|codsucur|0000|N|"
            Text            =   "0000"
            Top             =   1785
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
            Index           =   18
            Left            =   6420
            MaxLength       =   4
            TabIndex        =   21
            Tag             =   "Banco|N|S|0|9999|scafac|codbanco|0000|N|"
            Text            =   "Text1 7"
            Top             =   1785
            Width           =   645
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
            Left            =   1335
            MaxLength       =   20
            TabIndex        =   13
            Tag             =   "Refere. Cliente|T|S|||scafac1|referenc|||"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   1800
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
            Left            =   7740
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   15
            Tag             =   "Direccion/Dpto.|T|S|||scafac|nomdirec||N|"
            Text            =   "Text1"
            Top             =   285
            Visible         =   0   'False
            Width           =   3405
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
            Left            =   7155
            MaxLength       =   3
            TabIndex        =   14
            Tag             =   "Direccion/Dpto.|N|S|0|999|scafac|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   285
            Visible         =   0   'False
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
            Left            =   1335
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||scafac|proclien||N|"
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
            Left            =   1335
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "CPostal|T|N|||scafac|codpobla||N|"
            Text            =   "Text15"
            Top             =   1020
            Width           =   810
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
            Left            =   2175
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Población|T|N|||scafac|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1020
            Width           =   3465
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
            Left            =   3495
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "teléfono Cliente|T|S|||scafac|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   285
            Width           =   2145
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
            Left            =   1335
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "NIF Cliente|T|N|||scafac|nifclien||N|"
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
            Left            =   7155
            MaxLength       =   4
            TabIndex        =   16
            Tag             =   "Cod. Agente|N|N|0|9999|scafac|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   645
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
            Left            =   7740
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   42
            Text            =   "Text2"
            Top             =   645
            Width           =   3405
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
            Left            =   7155
            MaxLength       =   3
            TabIndex        =   17
            Tag             =   "Forma de Pago|N|N|0|999|scafac|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   1020
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
            Left            =   7740
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   38
            Text            =   "Text2"
            Top             =   1020
            Width           =   3405
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
            Height          =   315
            Index           =   16
            Left            =   12090
            MaxLength       =   5
            TabIndex        =   18
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scafac|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   270
            Width           =   525
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
            Height          =   315
            Index           =   17
            Left            =   12090
            MaxLength       =   5
            TabIndex        =   19
            Tag             =   "Descuento General|N|N|0|99.90|scafac|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   630
            Width           =   525
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
            Left            =   1335
            MaxLength       =   35
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||scafac|domclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   645
            Width           =   4305
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
            Left            =   5700
            TabIndex        =   150
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Aportación"
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
            Left            =   10980
            TabIndex        =   145
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6870
            ToolTipText     =   "Buscar agente"
            Top             =   675
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
            Left            =   8460
            TabIndex        =   58
            Top             =   1560
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
            Left            =   7980
            TabIndex        =   57
            Top             =   1560
            Width           =   285
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
            Left            =   7050
            TabIndex        =   56
            Top             =   1560
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
            Left            =   6420
            TabIndex        =   55
            Top             =   1560
            Width           =   645
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
            TabIndex        =   49
            Top             =   1800
            Width           =   1185
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
            Left            =   5700
            TabIndex        =   48
            Top             =   285
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   3
            Left            =   6870
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   315
            Visible         =   0   'False
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
            TabIndex        =   47
            Top             =   1410
            Width           =   945
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
            TabIndex        =   46
            Top             =   1020
            Width           =   945
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
            Left            =   2595
            TabIndex        =   45
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
            TabIndex        =   44
            Top             =   285
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1065
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
            Left            =   5700
            TabIndex        =   43
            Top             =   645
            Width           =   855
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
            Left            =   5700
            TabIndex        =   41
            Top             =   1020
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
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
            Index           =   25
            Left            =   11160
            TabIndex        =   40
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
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
            Index           =   26
            Left            =   11160
            TabIndex        =   39
            Top             =   630
            Width           =   1035
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   6870
            ToolTipText     =   "Buscar forma de pago"
            Top             =   1050
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
            TabIndex        =   37
            Top             =   645
            Width           =   915
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacHcoFacturas.frx":0044
         Height          =   2025
         Left            =   -74760
         TabIndex        =   54
         Top             =   3270
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   3572
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
         Bindings        =   "frmFacHcoFacturas.frx":0059
         Height          =   1945
         Left            =   -74760
         TabIndex        =   88
         Top             =   520
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
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   2655
         Left            =   -74760
         TabIndex        =   103
         Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
         Top             =   2670
         Width           =   12735
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
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   108
            Tag             =   "Observación 5|T|S|||scafac1|observa5||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1800
            Width           =   8940
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
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   107
            Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1410
            Width           =   8940
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
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   106
            Tag             =   "Observación 3|T|S|||scafac1|observa3||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1020
            Width           =   8940
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
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   105
            Tag             =   "Observación 2|T|S|||scafac1|observa2||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   630
            Width           =   8940
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
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   104
            Tag             =   "Observación 1|T|S|||scafac1|observa1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   240
            Width           =   8940
         End
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   -69120
         ToolTipText     =   "Buscar trabajador"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -69120
         ToolTipText     =   "Buscar forma de envio"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -69120
         ToolTipText     =   "Buscar trabajador"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
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
         Index           =   40
         Left            =   -63840
         TabIndex        =   101
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
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
         Index           =   22
         Left            =   -62400
         TabIndex        =   100
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
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
         Left            =   -63240
         TabIndex        =   99
         Top             =   615
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
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
         Left            =   -64230
         TabIndex        =   98
         Top             =   615
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Sem. Entrega"
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
         Left            =   -62040
         TabIndex        =   97
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Trab.Albaran"
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
         Left            =   -70560
         TabIndex        =   53
         Top             =   525
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo  Envío"
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
         Index           =   24
         Left            =   -70560
         TabIndex        =   52
         Top             =   2160
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Prep.Material"
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
         Index           =   23
         Left            =   -70560
         TabIndex        =   51
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "Trab.Pedido"
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
         Left            =   -70560
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   -69120
         ToolTipText     =   "Buscar trabajador"
         Top             =   1080
         Visible         =   0   'False
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
      Index           =   16
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   118
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   7410
      Visible         =   0   'False
      Width           =   7125
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Index           =   0
      Left            =   150
      TabIndex        =   32
      Top             =   7185
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
         TabIndex        =   33
         Top             =   180
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
      Left            =   12090
      TabIndex        =   26
      Top             =   7470
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
      Left            =   10800
      TabIndex        =   25
      Top             =   7470
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
      Left            =   12090
      TabIndex        =   31
      Top             =   7470
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
      TabIndex        =   146
      Top             =   7785
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliación Línea"
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
      TabIndex        =   35
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
         Shortcut        =   ^A
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
Attribute VB_Name = "frmFacHcoFacturas2"
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
Private WithEvents frmHcoFac2Pre As frmFacHcoFacturas2Prev 'Form para busquedas
Attribute frmHcoFac2Pre.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmC As frmGesSocios 'Form Mto Clientes
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


Private Sub Check1_Click()
    If Modo = 1 Then CheckCadenaBusqueda Check1, BuscaChekc
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer

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
'            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            
            PonerModo 2
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
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
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        lblIndicador.Caption = "Preparando bus."
        lblIndicador.Refresh
        LimpiarCampos
        LimpiarDataGrids
        DoEvents
        
        CadenaConsulta = "Select scafac.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
        If publicidad Then
            CadenaConsulta = CadenaConsulta & " where codtipom='FPC'"
        Else
            CadenaConsulta = CadenaConsulta & " where codtipom not in ('FLI','FPS','FCE','FCN','FAC','FPC')"
        End If
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
    If FactContabilizada3(EnTesoreria) Then
        TerminaBloquear
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
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
'    DeVarios = EsClienteVarios(Text1(4).Text)
'    BloquearDatosCliente (DeVarios)
End Sub


Private Sub BotonModificarLinea()
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
    
'    txtAux(9).Text = DataGrid1.Columns(16).Text
'    txtAux2(9).Text = DataGrid1.Columns(17).Text
'    'num lote
'    txtAux(10).Text = DataGrid1.Columns(19).Text
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
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
                If jj <> 9 And jj <> 10 Then
                    If jj = 4 Or (jj >= 6 And jj <= 10) Then
                        txtAux(jj).Height = DataGrid1.RowHeight
                        txtAux(jj).top = alto
                        txtAux(jj).visible = b
                    End If
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
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim Cad As String
Dim EstaEnTesoreria As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada3(EstaEnTesoreria) Then Exit Sub
    
    Cad = "Cabecera de Facturas." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Factura:            "
    Cad = Cad & vbCrLf & "Tipo:  " & Text1(1).Text
    Cad = Cad & vbCrLf & "Nº Fact.:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(2).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
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
        Me.FrameToolAux.visible = False
        Me.cmdObserva.Picture = frmppal.ImgListComun1.ListImages(18).Picture
'        CargarICO Me.cmdObserva, "volver.ico"
        Me.cmdObserva.ToolTipText = "volver lineas albaran"
        BloqueaText3
    Else
        Me.DataGrid1.visible = True
        Me.FrameToolAux.visible = True
        Me.FrameObserva.visible = False
        Me.cmdObserva.Picture = frmppal.ImgListComun1.ListImages(41).Picture
'        CargarICO Me.cmdObserva, "message.ico"
        Me.cmdObserva.ToolTipText = "ver observaciones albaran"
    End If
    SSTab1_Click 0
End Sub

Private Sub BloqueaText3()
Dim i As Byte
Dim b As Boolean

    'bloquear los Text3 que son las lineas de scafac1
    b = Modo <> 4 And Modo <> 1
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


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

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
    On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        If ModificaLineas <> 1 Then
            Text2(16).Text = DBLet(Data2.Recordset.Fields!Ampliaci, "T")
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
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(5).Image = 4   'Modificar
'        .Buttons(6).Image = 5   'Borrar
'        .Buttons(9).Image = 10 'Mto Lineas
'        .Buttons(10).Image = 16 'Imprimir
'        .Buttons(11).Image = 40 'Imprimir albaran
'        .Buttons(13).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'    End With
    
    With Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun1
        'ASignamos botones
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2 'Ver Todos
        .Buttons(1).Image = 3 'Añadir
        .Buttons(2).Image = 4 'Modificar
        .Buttons(3).Image = 5 'Eliminar
        .Buttons(8).Image = 16 'Imprimir
    End With
    
    With Me.Toolbar5
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun1
        .Buttons(1).Image = 40 'Imprimir albaran
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun1
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
    
    
    'cargar icono de observaciones de los albaranes de factura
    Me.cmdObserva.Picture = frmppal.ImgListComun1.ListImages(41).Picture
'    CargarICO Me.cmdObserva, "message.ico"
    Me.FrameObserva.visible = False
    Me.cmdObserva.ToolTipText = "ver observaciones albaran"
    
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
        txtAux(9).Tag = "Cod. centro coste|T|S|||slifac|codccost|||"
        Label1(46).Caption = "Centro coste"
    Else
        txtAux(9).Tag = "Cod. Proveedor|N|N|||slifac|codprovex|0||"
        Label1(46).Caption = "Proveedor"
    End If
        
        
    '## A mano
    NombreTabla = "scafac"
    NomTablaLineas = "slifac" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafac.codtipom, scafac.numfactu, scafac.fecfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Dim T1 As Single
    T1 = Timer
    
    'Mariela 01/07/2010 consulto si estoy llamando desde publicidad
    If publicidad Then
        CadenaConsulta = "Select * from " & NombreTabla & " where codtipom='FPC'"
    Else
        CadenaConsulta = "Select * from " & NombreTabla & " where codtipom not in ('FLI','FPS','FCE','FCN','FAC','FPC')"
    End If
    
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        'como tenemos aqui el nº de albaran, buscar a que factura corresponde
        'en la scafac1
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura
'        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        'CadenaConsulta = CadenaConsulta & " where numfactu=-1"
        'Cambio sugerido por Msoler
        'mariela 01/07/2010
        If Not publicidad Then CadenaConsulta = CadenaConsulta & " and codtipom is null and numfactu is null and fecfactu is null "
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
'Form Mantenimiento de Socios
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


Private Sub frmHcoFac2Pre_DatoSeleccionado(CadenaSeleccion As String)
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
            Set frmC = New frmGesSocios
            frmC.DatosADevolverBusqueda = "0"
            Screen.MousePointer = vbDefault
            frmC.Show vbModal
            Set frmC = Nothing
            indice = 5
            PonerFoco Text1(indice)
            
        Case 1 'NIF para cliente de Varios
'            Set frmCV = New frmFacClientesV
'            frmCV.DatosADevolverBusqueda = "0"
'            frmCV.Show vbModal
'            Set frmCV = Nothing
'            Indice = 6
'            PonerFoco Text1(Indice)
            
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
    BotonMtoLineas 1, "Facturas"
End Sub


Private Sub mnModificar_Click()

    If vUsu.Nivel > 0 Then
        MsgBox "No tiene permiso para realizar la accion", vbExclamation
        Exit Sub
    End If

    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then
                If BloqueaLineasFac Then BotonModificarLinea
            End If
        End If
         
    Else   'Modificar Pedido
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
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
    Sql = "select * FROM scafac1 "
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
    Sql = "select * FROM slifac "
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
    
    
    '--- David.  No se pq referencia NO lleva tag. Si han puesto algo lo paso a la cadena de busqueda
    
    If CadB <> "" Then CadB = CadB & " and scafac.codtipom not in ('FLI','FPS','FCE','FCN','FAC','FPC')"
    
    '--- Laura 12/01/2007
    Text1(5).Text = cadAux
    '---
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
'        cadB = cadB & " and scafac.codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los FAV
        CadenaConsulta = "select scafac.* from " & NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY scafac.codtipom,scafac.numfactu,scafac.fecfactu " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)

'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    Cad = ""
    If EsCabecera Then
'        Cad = Cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
'        Cad = Cad & ParaGrid(Text1(0), 15, "Nº Factura")
'        Cad = Cad & ParaGrid(Text1(2), 15, "Fecha Fac.")
'        Cad = Cad & ParaGrid(Text1(4), 10, "Cliente")
'        Cad = Cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
'      If publicidad Then
'        Tabla = NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom='FPC' and scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
'      Else
'        Tabla = NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
'      End If
'        'CadenaConsulta = "select scafac.* from " & NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
'        'CadenaConsulta = CadenaConsulta & " WHERE " & cadB & " GROUP BY scafac.codtipom,scafac.numfactu,scafac.fecfactu " & Ordenacion
'
'        Titulo = "Facturas"
'        devuelve = "0|1|2|"


        Set frmHcoFac2Pre = New frmFacHcoFacturas2Prev
    
        frmHcoFac2Pre.DatosADevolverBusqueda = "0|1|2|3|"
        
        If CadB <> "" Then CadB = CadB & " and "
        CadB = CadB & " scafac.codtipom not in ('FLI','FPS','FCE','FCN','FAC','FPC')  "
        
        frmHcoFac2Pre.cWhere = CadB
        frmHcoFac2Pre.Show vbModal
    
        Set frmHcoFac2Pre = Nothing

        Exit Sub
    Else
        If vParamAplic.Departamento Then
            Titulo = "Dptos Cliente: "
            Desc = "Dpto."
        Else
            Titulo = "Direc. Cliente: "
            Desc = "Direc."
        End If
        Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
        Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N||15·"
        Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35·"
        Tabla = "sdirec"
        devuelve = "0|1|"
    End If

    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = CadB
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
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafac1 de la factura seleccionada
Dim b As Boolean
Dim b2 As Boolean

    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafac1
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
'    Label1(2).visible = Not (b And b2)
'    Text3(8).visible = Not (b And b2)
'    'OFERTA
'    Text3(4).visible = Not b
'    Text3(5).visible = Not b
'    'VENTA
'    Text3(14).visible = b
'    Text3(15).visible = b
    
    
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
Dim i As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    For i = 0 To txtAux.Count - 1
        Text1(i).BackColor = vbWhite
    Next i



    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
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
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
    Me.Check1.Enabled = (Modo = 1)
    
    b = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(3), b 'referencia
    
    'Importes siempre bloqueados, excepto para busquedas. ivas y aportacion tb bloqueado
    For i = 22 To 45
        BloquearTxt Text1(i), (Modo <> 1)
    Next i
    'Campo B.Imp y Imp. IVA siempre en azul
    BloquearTxt Text1(25), True
    Text1(25).BackColor = &HFFFFC0
    
    If Modo <> 1 Then
        Text1(35).BackColor = &HFFFFC0
        Text1(36).BackColor = &HFFFFC0
        Text1(37).BackColor = &HFFFFC0
'    Text1(38).BackColor = &HC0C0FF    'Total factura
        Text1(38).BackColor = &HFFFFC0 '&HC0FFC0
    End If
    
    'bloquear los Text3 que son las lineas de scafac1
    BloqueaText3
    If Modo = 1 Then
        'Busqueda. Habilitamos numero pedido y fecha pedido
'        BloquearTxt Text3(6), False
'        BloquearTxt Text3(7), False
'        BloquearTxt Text3(16), False
    End If
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    BloquearTxt txtAux(8), True
    BloquearTxt txtAux(10), True
    BloquearTxt txtAux(11), True
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los albaranes de factura)
    For i = 0 To txtAux3.Count - 1
        BloquearTxt txtAux3(i), (Modo <> 1)
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


    Me.Combo1.visible = (Modo = 1)

    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    For i = 0 To 5
        Me.imgBuscar(i).Enabled = b
    Next i
    For i = 6 To 9
        Me.imgBuscar(i).Enabled = b And (Modo <> 1)
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

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
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
    
    'comprobamos datos OK de la tabla scafac
    b = CompForm(Me, 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim i As Byte
Dim Aux As String

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For i = 0 To txtAux.Count - 1
        If i = 4 Or i = 6 Or i = 7 Then
            If txtAux(i).Text = "" Then
                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
            
            
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
'        Case 1
'            BotonAnyadirLinea
        Case 2
            BotonModificarLinea
'        Case 3
'            BotonEliminarLinea
    End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5: mnBuscar_Click  'Buscar
        Case 6: BotonVerTodos  'Todos
            

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
        Sql = "UPDATE slifac SET "
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
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(0)|T|Alm.|500|;S|txtAux(1)|T|Artículo|1600|;S|txtAux(2)|T|Nombre Art.|3300|;"
            tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|1000|;S|txtAux(11)|T|Bultos|700|;S|txtAux(4)|T|Precio|1600|;S|txtAux(5)|T|OP|400|;S|txtAux(6)|T|Dto 1|600|;S|txtAux(7)|T|Dto 2|600|;S|txtAux(8)|T|Importe|1340|;"
            'TRAZA
'            tots = tots & "S|txtAux(9)|T|Prov.|750|;S|cmdaux|B|||;S|txtAux(10)|T|Nom. prov.|2000|;"
'            If vEmpresa.TieneAnalitica Then
'                'codprove,nomprove, codccost
'                tots = tots & "N||||0|;S|cmdaux|B|||;N||||0|;S|txtAux(9)|T|CCoste|750|;"
'
'            Else
'                tots = tots & "S|txtAux(9)|T|Prov.|750|;S|cmdaux|B|||;N||||0|;N||||0|;"
'            End If
            'numlote
'            tots = tots & "S|txtAux(10)|T|Nº Lote|1300|;"
            
            
            arregla tots, DataGrid1, Me, 350
            DataGrid1.Columns(9).Alignment = dbgRight
            DataGrid1.Columns(10).Alignment = dbgRight
            DataGrid1.Columns(12).Alignment = dbgCenter
            DataGrid1.Columns(13).Alignment = dbgRight
            DataGrid1.Columns(14).Alignment = dbgRight
            DataGrid1.Columns(15).Alignment = dbgRight
            'ocultar las q no usamos
            DataGrid1.Columns(16).visible = False
            DataGrid1.Columns(17).visible = False
            DataGrid1.Columns(18).visible = False
            DataGrid1.Columns(19).visible = False
            
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


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnImprimirAlbaran_Click
    End Select
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
        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(8), 1
    End If
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


Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim cContaFra As cContabilizarFacturas
Dim MenError As String
    
    On Error GoTo FinEliminar

    b = False
    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
    ConnConta.BeginTrans
        
    'Eliminar en las tablas de la Contabilidad
    '------------------------------------------
    LEtra = ObtenerLetraSerie(Data1.Recordset!codtipom)
    
    '[Monica]27/09/2012: si dejamos borrar la factura he de devolver el stock
    If Data1.Recordset!codtipom = "FAV" Then
        MenError = "Restableciendo stocks de almacen."
        b = ReestablecerStock
    End If
    

'[Monica]18/02/2011: No contabilizamos las facturas
'    Set cContaFra = New cContabilizarFacturas

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
    
        'Lineas de facturas (slifac)
        conn.Execute "Delete from slifac " & Sql
    
        'Lineas de cabeceras de albaranes de la factura
        conn.Execute "Delete from scafac1 " & Sql
        
        'Eliminar los vencimientos
        conn.Execute "Delete from svenci " & Sql
        
        'Cabecera de facturas (scafac)
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
'        If LEtra <> "" Then
''            Preparao para eliminar
'            If cContaFra.EstablecerValoresInciales(ConnConta) Then
'                SQL = CStr(Data1.Recordset!FecFactu)
'                cContaFra.FijarNumeroFactura CLng(Data1.Recordset!NumFactu), Year(Data1.Recordset!FecFactu), LEtra
'            End If
'        End If
        'De ARIGES
        conn.CommitTrans
        ConnConta.CommitTrans

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

Private Function InicializarCStock(ByRef RS As ADODB.Recordset, ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
Dim Sql As String
Dim Rs2 As ADODB.Recordset

'On Error Resume Next
On Error Resume Next

    Sql = "select * from scafac1 where codtipom = " & DBSet(RS!codtipom, "T") & " and numfactu =  " & DBSet(RS!NumFactu, "N")
    Sql = Sql & " and fecfactu = " & DBSet(RS!FecFactu, "F") & " and numalbar = " & DBSet(RS!NumAlbar, "N")

    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs2.EOF Then Rs2.MoveFirst


    vCStock.tipoMov = TipoM 'Movimiento de Entrada o Salida
    vCStock.DetaMov = Rs2!codtipoa '"ALC=Albaran de Compra"
    vCStock.Fechamov = Rs2!FechaAlb
    vCStock.Trabajador = CLng(Rs2!CodTraba) 'En smoval guardamos el Proveedor
    vCStock.Documento = Format(Rs2!NumAlbar, "0000000")
    
    If ModificaLineas = 1 Or (ModificaLineas = 2 And TipoM = "E") Then '1=Insertar, 2=Modificar
        vCStock.codArtic = RS!codArtic
        vCStock.codAlmac = CInt(RS!codAlmac)
        vCStock.Cantidad = CSng(ComprobarCero(RS!Cantidad))
        vCStock.Importe = CCur(ComprobarCero(RS!ImporteL))
    Else
        vCStock.codArtic = RS!codArtic
        vCStock.codAlmac = CInt(RS!codAlmac)
        vCStock.Cantidad = CSng(RS!Cantidad)
        vCStock.Importe = CCur(RS!ImporteL)
    End If
    If ModificaLineas = 1 Then
         vCStock.LineaDocu = CInt(ComprobarCero(RS!numlinea))
    Else
        vCStock.LineaDocu = CInt(RS!numlinea)
    End If
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function


Private Function ReestablecerStock() As Boolean
Dim vCStock As CStock
Dim b As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset


    On Error GoTo ERestablecer
    
    ReestablecerStock = False
    b = True
    
    Sql = "select * from slifac where " & ObtenerWhereCP(False) & " order by numalbar, numlinea "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Para cada linea de albaran reestablecer el stock
    While (Not RS.EOF) And b
        Set vCStock = New CStock
        If InicializarCStock(RS, vCStock, "E") Then
            'Actualiza el stock en salmac y borra de smoval
            If Not vCStock.DevolverStock2() Then b = False
        Else
            b = False
        End If
        RS.MoveNext
        Set vCStock = Nothing
    Wend
    
ERestablecer:
    If Err.Number <> 0 Then b = False
    ReestablecerStock = b
End Function




Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
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
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    If Opcion = 1 Then
        Sql = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,numbultos, precioar, origpre, dtoline1, dtoline2, importel ,codprovex, nomprove,codccost,numlote"
        Sql = Sql & " FROM slifac left join sprove on codprovex=codprove " 'lineas de factura
    ElseIf Opcion = 2 Then
        Sql = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb, numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa  "
        Sql = Sql & " FROM scafac1 " 'cabeceras albaranes de la factura
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
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean
Dim i As Integer
Dim bAux As Boolean

    b = (Modo = 2) 'Or (Modo = 5 And ModificaLineas = 0)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = (Modo = 2)
    Me.mnEliminar.Enabled = (Modo = 2)
        
    b = (Modo = 2)
'    'Mantenimiento lineas
'    Toolbar1.Buttons(9).Enabled = b
'    Me.mnLineas.Enabled = b
    'Imprimir
    Toolbar1.Buttons(8).Enabled = b
    Me.mnImprimir.Enabled = b
    Toolbar5.Buttons(1).Enabled = b
    mnImprimirAlbaran.Enabled = b
    
    b = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    
    b = (Modo = 2)
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = False
        If Not Data2.Recordset Is Nothing Then
            If b Then bAux = (b And Me.Data2.Recordset.RecordCount > 0)
        End If
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = False
    Next i
    
End Sub



Private Sub PonerDatosCliente(CodClien As String, Optional nifClien As String)
'Dim vCliente As CCliente
Dim Observaciones As String
Dim b As Boolean
Dim Sql As String

    On Error GoTo EPonerDatos
    
    If CodClien = "" Then
        LimpiarDatosCliente
        Exit Sub
    End If
    
    Sql = "select * from sclien where codclien=" & CodClien
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
'    Set vCliente = New CCliente
    'si se ha modificado el cliente volver a cargar los datos
'    If vCliente.Existe(codClien) Then
    If Not miRsAux.EOF Then
        'If vCliente.LeerDatos(codClien) Then
            'si el cliente esta bloqueado salimos
            'If vCliente.ClienteBloqueado Then
            If ClienteBloqueado(miRsAux!codsitua) Then
                If Modo = 3 Then
                    b = True
                ElseIf Modo = 4 Then
                     If (Val(Text1(4).Text) <> Val(Data1.Recordset!CodClien)) Then b = True
                End If
                If b Then
                    LimpiarDatosCliente
                    'Set vCliente = Nothing
                    Exit Sub
                End If
            End If
            
'            EsDeVarios = vCliente.EsClienteVarios(Text1(4).Text)
'            EsDeVarios = vCliente.DeVarios
'            BloquearDatosCliente (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el cliente no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!CodClien) Then
                    'Set vCliente = Nothing
                    Exit Sub
                End If
            End If
        
        
'            If Actualizar = False And EsDeVarios = False Then Exit Sub
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = Format(miRsAux!CodClien, "000000")
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = miRsAux!nomclien  'Nom clien
                Text1(8).Text = miRsAux!domclien
                Text1(9).Text = miRsAux!codpobla
                Text1(10).Text = miRsAux!pobclien
                Text1(11).Text = miRsAux!proclien
                Text1(6).Text = miRsAux!nifClien
                Text1(7).Text = DBLet(miRsAux!telclie1, "T")
            End If
            
            'insertar
            'If Modo = 3 Then Text1(15).Text = vCliente.ForPago

            Observaciones = DBLet(miRsAux!observac)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del socio"
            End If
                
            'cuenta bancaria
            Text1(18).Text = miRsAux!codbanco
            FormateaCampo Text1(18)
            Text1(19).Text = miRsAux!codsucur
            FormateaCampo Text1(19)
            Text1(20).Text = miRsAux!digcontr
            Text1(21).Text = miRsAux!cuentaba
            
            'Comprobar si el cliente tiene cobros pendientes
            'ComprobarCobrosCliente2 codClien, Text1(1).Text
        
    Else
        LimpiarDatosCliente
    End If
    'Set vCliente = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub

Private Function ClienteBloqueado(Situacion As Integer) As Boolean
Dim Tipo As String
Dim devuelve As String

    On Error GoTo EBloqueado
    ClienteBloqueado = False
    
    If Situacion <> 0 Then '0: situacion normal
        Tipo = "tipositu"
        devuelve = DevuelveDesdeBDNew(conAri, "ssitua", "nomsitua", "codsitua", CStr(Situacion), "N", Tipo)
        
        If Tipo = "1" Then 'Cliente Bloqueado por Situación Especial.
            MsgBox UCase("Cliente Bloqueado por: ") & vbCrLf & devuelve, vbInformation, "Situación Especial del Cliente."
            ClienteBloqueado = True
        Else
            MsgBox devuelve, vbInformation, "Situación Especial del Cliente."
        End If
    End If
    
EBloqueado:
    If Err.Number <> 0 Then Err.Clear
End Function

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
    If (OpcionListado = 53) Then
        If Text1(1).Text = "FAZ" Then
            'Factura B
            indRPT = 30
        Else
            indRPT = 12 'Facturas Clientes
        End If
    Else
        '-----------------------------------------------
        indRPT = 18 'Facturas Clientes TPV
    End If
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, ImprimeDirecto, pPdfRpt) Then Exit Sub
      
      
'--[Monica]arituclo de reciclado pasa a ser reutilizado como articulo de gastos de administracion
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
     
     
        
        devuelve = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
     
     
         With frmImprimir
                'Nuevo. Febrero 2010
                .outClaveNombreArchiv = devuelve & Format(Text1(0).Text, "000")
                .outCodigoCliProv = Text1(4).Text
                .outTipoDocumento = 100
                
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

    cadImpresion = "{scafac.codtipom}='" & Text1(1).Text & "' and {scafac.numfactu}=" & Text1(0).Text
    Sql = cadImpresion & " and {scafac.fecfactu}=" & DBSet(Text1(2).Text, "F")
    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(CDate(Text1(2).Text)) & "," & Month(CDate(Text1(2).Text)) & "," & Day(CDate(Text1(2).Text)) & ")"
    
    If Not HayRegParaInforme("scafac", Sql) Then Exit Sub
    
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
'    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(RSVenta!fecventa) & "," & Month(RSVenta!fecventa) & "," & Day(RSVenta!fecventa) & ")"
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
    'comprobar datos OK de la scafac1
     b = CompForm(Me, 2) 'Comprobar formato datos ok de la cabecera alb: opcion=2
    If Not b Then Exit Function
    
    Sql = "UPDATE scafac1 SET codenvio=" & Text3(3).Text & ", "
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
'si se ha modificado la linea de slifac, añadir a la transaccion la modificación de la linea y recalcular
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
        
        'modificamos la scafac
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario(Me, 1)
        
'        If bol Then
'            'Si es cliente de varios actualizar datos cliente en tabla:sclvar
'            MenError = "Modificando datos cliente varios"
'            bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
'        End If
        
        If bol Then
            MenError = "Modificando albaranes de factura"
            'modificar la tabla: scafac1
            bol = ModificaAlbxFac
            
            If bol And recalcular Then 'si se ha modificado la factura
                MenError = "Actualizando en Tesoreria"
                'borrar los vencimientos de aritaxi.svenci
                'y eliminar de tesoreria conta.scobros los registros de la factura(si existen en Tesoreria)
                
                'Eliminar los vencimientos
                '----------------------------------------
                Sql = ObtenerWhereCP(True)
                conn.Execute "Delete from svenci " & Sql
                
                'Eliminar de Tesoreria
                '----------------------------------------
'                SQL = ObtenerLetraSerie(Text1(1).Text)
'                SQL = "SELECT COUNT(*) FROM scobro WHERE numserie='" & SQL & "' and codfaccl=" & Text1(0).Text
'                SQL = SQL & " AND fecfaccl=" & DBSet(Text1(2).Text, "F")
'
'                If RegistrosAListar(SQL, conConta) Then
                    'antes de Eliminar en las tablas de la Contabilidad
                Set vFactura = New CFactura
                If vFactura.LeerDatos(Text1(1).Text, Text1(0).Text, Text1(2).Text) Then
                Else
                  bol = False
                End If
              
                If bol Then
                    'Eliminar de la scobro
                    If vParamAplic.ContabilidadNueva Then
                        Sql = " numserie='" & vFactura.LetraSerie & "' AND numfactu=" & Data1.Recordset.Fields!NumFactu
                        Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                        ConnConta.Execute "Delete from cobros WHERE " & Sql
                    Else
                        Sql = " numserie='" & vFactura.LetraSerie & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
                        Sql = Sql & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                        ConnConta.Execute "Delete from scobro WHERE " & Sql
                    End If
                    bol = True

                    'Volvemos a Insertar los Vencimientos de la Factura. Tabla: svenci
                    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
                    If bol Then
                        vFactura.Agente = Text1(14).Text
                        'comentamos hasta tener claro con manolo tema de contabilización de facturas
                        'de cuotas.
                        bol = vFactura.InsertarEnTesoreria("", MenError)
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
Dim i As Integer
Dim vFactu As CFactura
Dim FacOK As Boolean

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 22 To 38
         Text1(i).Text = ""
    Next i
    
    
    Set vFactu = New CFactura
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(16).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(17).Text))
    vFactu.Cliente = Text1(4).Text
    
    If vFactu.CalcularDatosFactura(False, ObtenerWhereCP(False), NombreTabla, NomTablaLineas) Then
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
    If Me.Check1.Value = 1 Then 'si esta contabilizada
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
Dim Cad As String
Dim RS As ADODB.Recordset

    On Error Resume Next

    Cad = ""
    If Me.DesdeFichaCliente Then
        '
        Cad = " and codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
        
    Else
        'Tengo YA el codigo de la factura
                '******************************************************
                'laura: esto se puede comentar, ya no hay movimiento FTI en la smoval
                If hcoCodTipoM = "FTI" Then
                    'no hay albaran directamente va a factura de ticket
                    
                    'ver si lo encontramos como factura: codtipom, numfactu,fecfactu
                    Cad = "SELECT COUNT(*) FROM scafac "
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
                    Cad = "SELECT codtipom,numfactu,fecfactu FROM scafac1 "
                    Cad = Cad & " WHERE codtipoa=" & DBSet(hcoCodTipoM, "T") & " AND numalbar=" & hcoCodMovim & " AND fechaalb=" & DBSet(hcoFechaMov, "F")
                    
                    Set RS = New ADODB.Recordset
                    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not RS.EOF Then 'where para la factura
                        Cad = " WHERE codtipom='" & RS!codtipom & "' AND numfactu= " & RS!NumFactu & " AND fecfactu=" & DBSet(RS!FecFactu, "F")
                    Else
                        Cad = " WHERE numfactu=-1"
                    End If
                    RS.Close
                    Set RS = Nothing
                End If
    
    End If
    ObtenerSelFactura = Cad
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
Dim RS As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    Combo1.Clear
    
    '[Monica]11/02/2011: todo tipo de facturas excepto las de liquidacion,publicidad y cuotas de socio
    '                    y las facturas de cliente FAC y FPC
    Sql = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom LIKE 'F%' and codtipom not in ('FLI','FPS','FCE','FCN','FAC','FPC')"
    
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
        devuelve = "{scafac1.codtipoa}=" & DBSet(Data3.Recordset!codtipoa, "T")
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        'Numalbar
        devuelve = "{scafac1.numalbar}=" & DBSet(Data3.Recordset!NumAlbar, "N")
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
        '[Monica]--20/12/2010
        'devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(4).Text, "N")
        devuelve = "1"
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
Dim Cad As String


On Error GoTo EComprobarCobroArimoney
    ComprobarCobroArimoney = False
    Set vR = New ADODB.Recordset
    If vParamAplic.ContabilidadNueva Then
        Cad = "Select * from cobros where numserie='" & LEtra & "'"
        Cad = Cad & " AND numfactu =" & Codfaccl
        Cad = Cad & " AND fecfactu =" & DBSet(Fecha, "F")
    Else
        Cad = "Select * from scobro where numserie='" & LEtra & "'"
        Cad = Cad & " AND codfaccl =" & Codfaccl
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
