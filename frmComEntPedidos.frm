VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmComEntPedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos Proveedor"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   12090
   Icon            =   "frmComEntPedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   8
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   133
      Text            =   "nom ccoste"
      Top             =   6320
      Visible         =   0   'False
      Width           =   5805
   End
   Begin VB.Frame Frame2 
      Height          =   910
      Left            =   120
      TabIndex        =   82
      Top             =   420
      Width           =   11775
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   7855
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Nombre Proveedor|T|N|||scappr|nomprove||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   520
         Width           =   3645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   7020
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Cod. Proveedor|N|N|0|999999|scappr|codprove|000000|N|"
         Text            =   "Text1"
         Top             =   520
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   7020
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Realizado Por|N|N|0|9999|scappr|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   160
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   7855
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   85
         Text            =   "Text2"
         Top             =   160
         Width           =   3645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Pedido|F|N|||scappr|fecpedpr|dd/mm/yyyy|N|"
         Top             =   420
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "N� Pedido|N|S|0||scappr|numpedpr|0000000|S|"
         Text            =   "Text1 7"
         Top             =   420
         Width           =   1125
      End
      Begin VB.CheckBox chkRestoPed 
         Caption         =   "Resto de Pedido"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Tag             =   "Resto de Pedido|N|N|||scappr|restoped||N|"
         Top             =   420
         Width           =   1575
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6735
         Picture         =   "frmComEntPedidos.frx":000C
         ToolTipText     =   "Buscar proveedor"
         Top             =   520
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   87
         Top             =   520
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Realizado Por"
         Height          =   255
         Index           =   21
         Left            =   5640
         TabIndex        =   86
         Top             =   160
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   6735
         Picture         =   "frmComEntPedidos.frx":010E
         ToolTipText     =   "Buscar trabajador"
         Top             =   160
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Ped."
         Height          =   255
         Index           =   14
         Left            =   1560
         TabIndex        =   84
         Top             =   220
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2475
         Picture         =   "frmComEntPedidos.frx":0210
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "N� Pedido"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   83
         Top             =   220
         Width           =   735
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   16
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   46
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   5940
      Visible         =   0   'False
      Width           =   6885
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   5775
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   32
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10530
      TabIndex        =   19
      Top             =   5880
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9360
      TabIndex        =   18
      Top             =   5880
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3600
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
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
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas Pedido"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Albaran"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Pedido"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   0
         Left            =   9000
         MaxLength       =   15
         TabIndex        =   125
         Text            =   "TOTAL"
         Top             =   100
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   50
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   124
         Text            =   "Text1 7"
         Top             =   80
         Width           =   1530
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6480
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   5640
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4425
      Left            =   120
      TabIndex        =   35
      Top             =   1350
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7805
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
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmComEntPedidos.frx":029B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtAux(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtAux(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAux(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAux(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAux(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAux(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAux(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FrameCliente"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAux(8)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdAux(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmComEntPedidos.frx":02B7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameHco"
      Tab(1).Control(1)=   "FrameDirFactura"
      Tab(1).Control(2)=   "FrameDirMercancia"
      Tab(1).Control(3)=   "Text1(21)"
      Tab(1).Control(4)=   "Text1(20)"
      Tab(1).Control(5)=   "Text1(19)"
      Tab(1).Control(6)=   "Text1(18)"
      Tab(1).Control(7)=   "Text1(17)"
      Tab(1).Control(8)=   "Label1(45)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Totales"
      TabPicture(2)   =   "frmComEntPedidos.frx":02D3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameFactura"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   2
         Left            =   11040
         TabIndex        =   132
         ToolTipText     =   "Buscar centro coste"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   45
         Tag             =   "centro coste"
         Text            =   "cc"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame FrameHco 
         Caption         =   "Datos  Eliminaci�n"
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
         Height          =   1970
         Left            =   -67200
         TabIndex        =   126
         Top             =   2350
         Width           =   3495
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   24
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   27
            Top             =   230
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   25
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   600
            Width           =   780
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   25
            Left            =   75
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   128
            Text            =   "Text2"
            Top             =   940
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   26
            Left            =   135
            MaxLength       =   30
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   1560
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   26
            Left            =   675
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   127
            Text            =   "Text2"
            Top             =   1560
            Width           =   2685
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   131
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
            Height          =   255
            Index           =   38
            Left            =   75
            TabIndex        =   130
            Top             =   670
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   960
            Picture         =   "frmComEntPedidos.frx":02EF
            ToolTipText     =   "Buscar trabajador"
            Top             =   660
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   129
            Top             =   1360
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1080
            Picture         =   "frmComEntPedidos.frx":03F1
            ToolTipText     =   "Buscar incidencia"
            Top             =   1320
            Width           =   240
         End
      End
      Begin VB.Frame FrameFactura 
         Height          =   3300
         Left            =   -74640
         TabIndex        =   92
         Top             =   720
         Width           =   10575
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   109
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   108
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   107
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   106
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   105
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   104
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   5017
            MaxLength       =   5
            TabIndex        =   103
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   570
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   46
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   102
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   101
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   100
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   5017
            MaxLength       =   5
            TabIndex        =   99
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   570
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   98
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   97
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   96
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   5017
            MaxLength       =   5
            TabIndex        =   95
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   570
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   48
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   94
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   49
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   93
            Text            =   "Text1 7"
            Top             =   2640
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   27
            Left            =   5760
            TabIndex        =   123
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   122
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   23
            Left            =   2160
            TabIndex        =   121
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   22
            Left            =   3960
            TabIndex        =   120
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   18
            Left            =   5760
            TabIndex        =   119
            Top             =   360
            Width           =   1215
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
            TabIndex        =   118
            Top             =   480
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
            TabIndex        =   117
            Top             =   480
            Width           =   135
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
            TabIndex        =   116
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   7560
            TabIndex        =   115
            Top             =   1230
            Width           =   1335
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
            Index           =   8
            Left            =   7320
            TabIndex        =   114
            Top             =   1320
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   4320
            X2              =   7320
            Y1              =   1065
            Y2              =   1065
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
            TabIndex        =   113
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL PEDIDO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   39
            Left            =   5640
            TabIndex        =   112
            Top             =   2660
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   5040
            TabIndex        =   111
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
            Height          =   255
            Index           =   42
            Left            =   4320
            TabIndex        =   110
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.Frame FrameDirFactura 
         Caption         =   "Direc. Factura"
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
         Height          =   1900
         Left            =   -69000
         TabIndex        =   72
         Top             =   420
         Width           =   5175
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   1005
            MaxLength       =   30
            TabIndex        =   21
            Tag             =   "Direc. Factura|N|S|0|999|scappr|coddiref|000|N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   2
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   81
            Text            =   "Text2"
            Top             =   360
            Width           =   3510
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   24
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   76
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1425
            Width           =   2565
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   22
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   75
            Text            =   "Text15"
            Top             =   1065
            Width           =   630
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   23
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   74
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1065
            Width           =   3405
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   21
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   73
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   720
            Width           =   4050
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   5
            Left            =   720
            Picture         =   "frmComEntPedidos.frx":04F3
            ToolTipText     =   "Buscar direcci�n"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   80
            Top             =   1425
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   79
            Top             =   1035
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   78
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Cod.Dir"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame FrameDirMercancia 
         Caption         =   "Direc. Mercancia"
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
         Height          =   1900
         Left            =   -74760
         TabIndex        =   62
         Top             =   420
         Width           =   5175
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   17
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   67
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   720
            Width           =   4050
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   19
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   66
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1065
            Width           =   3405
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   18
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   65
            Text            =   "Text15"
            Top             =   1065
            Width           =   630
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   20
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   64
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1425
            Width           =   2565
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   1005
            MaxLength       =   30
            TabIndex        =   20
            Tag             =   "Direc. Mercancia|N|S|0|999|scappr|coddirea|000|N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   15
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   63
            Text            =   "Text2"
            Top             =   360
            Width           =   3510
         End
         Begin VB.Label Label1 
            Caption         =   "Cod.Dir"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   71
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   70
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   69
            Top             =   1035
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Top             =   1425
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   4
            Left            =   720
            Picture         =   "frmComEntPedidos.frx":05F5
            ToolTipText     =   "Buscar direcci�n"
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame FrameCliente 
         Height          =   1770
         Left            =   240
         TabIndex        =   51
         Top             =   310
         Width           =   11415
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   1
            Left            =   7635
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   90
            Text            =   "Text2"
            Top             =   560
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   23
            Left            =   6930
            MaxLength       =   30
            TabIndex        =   13
            Tag             =   "Solicitado Por|N|S|0|9999|scappr|codtrab1|0000|N|"
            Text            =   "Text1"
            Top             =   550
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   0
            Left            =   7635
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   88
            Text            =   "Text2"
            Top             =   190
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   22
            Left            =   6930
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Cliente|N|S|0|999999|scappr|codclien|000000|N|"
            Text            =   "Text1"
            Top             =   190
            Width           =   660
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   16
            Left            =   6600
            MaxLength       =   25
            TabIndex        =   15
            Tag             =   "Tipo Portes|T|S|||scappr|tipoporte||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwww"
            Top             =   1380
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1170
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Provincia|T|N|||scappr|proprove||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1300
            Width           =   2565
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   9
            Tag             =   "CPostal|T|N|||scappr|codpobla||N|"
            Text            =   "Text15"
            Top             =   910
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1820
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "Poblaci�n|T|N|||scappr|pobprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   910
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3375
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "tel�fono Proveedor|T|S|||scappr|telprove||N|"
            Text            =   "12345678911234567899"
            Top             =   190
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   6
            Tag             =   "NIF Proveedor|T|N|||scappr|nifprove||N|"
            Text            =   "123456789"
            Top             =   190
            Width           =   1230
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   6930
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "Forma de Pago|N|N|0|999|scappr|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   910
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   7635
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   53
            Text            =   "Text2"
            Top             =   910
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   9315
            MaxLength       =   7
            TabIndex        =   16
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaped|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1390
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   10200
            MaxLength       =   7
            TabIndex        =   17
            Tag             =   "Descuento General|N|N|0|99.90|scaped|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1390
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1170
            MaxLength       =   35
            TabIndex        =   8
            Tag             =   "Domicilio|T|N|||scappr|domprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   550
            Width           =   4170
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   6645
            Picture         =   "frmComEntPedidos.frx":06F7
            ToolTipText     =   "Buscar trabajador"
            Top             =   560
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Solicitado por"
            Height          =   255
            Index           =   6
            Left            =   5660
            TabIndex        =   91
            Top             =   560
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   6645
            Picture         =   "frmComEntPedidos.frx":07F9
            ToolTipText     =   "Buscar cliente"
            Top             =   210
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Para Cliente"
            Height          =   255
            Index           =   4
            Left            =   5660
            TabIndex        =   89
            Top             =   190
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   900
            Picture         =   "frmComEntPedidos.frx":08FB
            ToolTipText     =   "Buscar proveedor varios"
            Top             =   210
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Atenci�n de "
            Height          =   255
            Index           =   3
            Left            =   5640
            TabIndex        =   61
            Top             =   1380
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   900
            Picture         =   "frmComEntPedidos.frx":09FD
            ToolTipText     =   "Buscar poblaci�n"
            Top             =   940
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   60
            Top             =   1300
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   59
            Top             =   910
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
            Height          =   255
            Index           =   19
            Left            =   2565
            TabIndex        =   58
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   57
            Top             =   190
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5660
            TabIndex        =   56
            Top             =   910
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
            Height          =   255
            Index           =   25
            Left            =   9240
            TabIndex        =   55
            Top             =   1210
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   10080
            TabIndex        =   54
            Top             =   1210
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   6645
            Picture         =   "frmComEntPedidos.frx":0AFF
            ToolTipText     =   "Buscar forma de pago"
            Top             =   930
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   52
            Top             =   550
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   50
         ToolTipText     =   "Buscar art�culo"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   49
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   40
         Tag             =   "Nombre Art�culo"
         Text            =   "nomArtic"
         Top             =   3960
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   9360
         MaxLength       =   12
         TabIndex        =   47
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   8880
         MaxLength       =   30
         TabIndex        =   44
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   8280
         MaxLength       =   5
         TabIndex        =   43
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   7320
         MaxLength       =   12
         TabIndex        =   42
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   3960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   6000
         MaxLength       =   16
         TabIndex        =   41
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   39
         Tag             =   "C�digo Art�culo"
         Text            =   "Artic Artic Artic5"
         Top             =   3900
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   15
         TabIndex        =   38
         Tag             =   "C�digo Almacen"
         Text            =   "codalmac"
         Top             =   3900
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   21
         Left            =   -74760
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "Observaci�n 5|T|S|||scappr|observa5||N|"
         Top             =   3960
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   20
         Left            =   -74760
         MaxLength       =   80
         TabIndex        =   25
         Tag             =   "Observaci�n 4|T|S|||scappr|observa4||N|"
         Top             =   3690
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   19
         Left            =   -74760
         MaxLength       =   80
         TabIndex        =   24
         Tag             =   "Observaci�n 3|T|S|||scappr|observa3||N|"
         Top             =   3420
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   18
         Left            =   -74760
         MaxLength       =   80
         TabIndex        =   23
         Tag             =   "Observaci�n 2|T|S|||scappr|observa2||N|"
         Top             =   3150
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   17
         Left            =   -74760
         MaxLength       =   80
         TabIndex        =   22
         Tag             =   "Observaci�n 1|T|S|||scappr|observa1||N|"
         Top             =   2880
         Width           =   7485
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComEntPedidos.frx":0C01
         Height          =   2025
         Left            =   240
         TabIndex        =   48
         Top             =   2130
         Width           =   11415
         _ExtentX        =   20135
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
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   45
         Left            =   -74760
         TabIndex        =   37
         Top             =   2640
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10530
      TabIndex        =   30
      Top             =   5880
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblF 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   135
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Centro coste"
      Height          =   255
      Index           =   46
      Left            =   2400
      TabIndex        =   134
      Top             =   6320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliaci�n L�nea"
      Height          =   255
      Index           =   35
      Left            =   2400
      TabIndex        =   36
      Top             =   5760
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
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
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
      Begin VB.Menu mnGenAlbaran 
         Caption         =   "&Generar Albaran"
         HelpContextID   =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnImpPedido 
         Caption         =   "&Imprimir Pedido"
         Shortcut        =   ^I
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
Attribute VB_Name = "frmComEntPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)


Public MostrarDatos As String  'Para ver un dato enconcreto
Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schppr, y solo en modo de consulta
                              
                              
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmProv As frmComProveedores  'Form Mto Proveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmProveV As frmComProveV  'Form Mto Proveedores Varios
Attribute frmProveV.VB_VarHelpID = -1
Private WithEvents frmDir As frmComDirecciones
Attribute frmDir.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmArt As frmAlmArticulos   'Form Articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientes 'form mantenimiento clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmInc As frmIncidencias  'form mantenimiento incidencias eliminacion
Attribute frmInc.VB_VarHelpID = -1

Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar n� Series
Attribute frmNSerie.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoOfer 'Listados
Attribute frmList.VB_VarHelpID = -1


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
'   6.- Cargar cantidad servidas al Generar Albaran no completo (Pedido --> Albaran)
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------


Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean


'Codigo tipo de movimiento en funci�n del valor en la tabla de par�metros: stipom
Dim CodTipoMov As String


Dim EsDeVarios As Boolean 'Si el Proveedor mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String
Private CadenaSQL As String 'Para crear consulta de Generar Albaran a partir del Pedido

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla de Cabecera
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1
Dim btnAnyadir As Byte

'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1
Dim btnPrimero As Byte


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Dim gridCargado As Boolean 'Saber si el grid esta cargado cuando se ejecuta DataGrid1_RowColChange

Dim AlbCompleto As Boolean 'Si se va a servir el Pedido Completo (slialb.cantidad=sliped.cantidad)
                            'o se va a servir una parte (slialb.cantidad=sliped.servidas)

Dim PulsadoMas2 As Boolean

'================================================================================

Private Sub cmdAceptar_Click()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR Cabecera Pedido
            If DatosOk Then
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(CodTipoMov) Then
                    Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
                    Sql = CadenaInsertarDesdeForm(Me)
                    If Sql <> "" Then
                        If InsertarPedido(Sql, vTipoMov) Then
                            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                            PonerCadenaBusqueda
                            PonerModo 2
                            'Ponerse en Modo Insertar Lineas
                            BotonMtoLineas 1, "Pedidos"
                            BotonAnyadirLinea
                        End If
                    End If
                    FormateaCampo Text1(0)
                End If
                Set vTipoMov = Nothing
            End If
            Me.SSTab1.Tab = 0
            
        Case 4  'MODIFICAR Cabecera Pedido
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    'Actualizar los datos del Proveedor si es de varios
                    ActualizarProveVarios Text1(4).Text, Text1(6).Text
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'sliped'
            If ModificaLineas = 1 Then 'INSERTAR lineas Pedidos
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea Then
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    BotonAnyadirLinea
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                End If
                Me.DataGrid1.Enabled = True
            End If
            CalcularDatosFactura
            
        Case 6 'Pasar Pedido a Albaran
            If BLOQUEADesdeFormulario(Me) Then GenerarAlbaran
            TerminaBloquear
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            PonerFoco txtAux(Index)
            
        Case 1 'Busqueda de Cod. Artic
            Set frmArt = New frmAlmArticulos
            frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en modo busqueda
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco txtAux(Index)
            
        Case 2 'COD. CENTRO DE COSTE
            If vEmpresa.TieneAnalitica Then
                'centro de coste
                AbrirForm_CentroCoste
                PonerFoco txtAux(8)
            End If
    End Select
    
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
            BloquearTxt Text2(16), True
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            Else
                ModificaLineas = 0
            End If
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            
        Case 6 'Insertar servidas en Generar Albaran (Pedido --> Albaran)
            If MsgBox("Desea cancelar la introducci�n de unidades del pedido?", vbQuestion + vbYesNo) = vbYes Then
                TerminaBloquear
                InicializarServidas
                PonerModo 2
                CargaTxtAuxServidas False, False
                CargaGrid DataGrid1, Data2, True, False
                
            Else
                PonerFoco Me.txtAux(3)
            End If
    End Select
End Sub


Private Sub BotonAnyadir()
'A�adir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vac�a los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo A�adir Linea
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
    'Poner el Almacen por defecto del Trabajador
    txtAux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")
    
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    BloquearTxt Text2(16), False
    
    ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste
    'si contab. analitica por trabajador traer su centro de coste
    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
        txtAux(8).Text = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
        Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
    End If
    ' ----
    
    
    PonerFoco txtAux(1)
    Me.DataGrid1.Enabled = False
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla & " " & Ordenacion
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
'Prepara el Form para Modificar la cabecera de Pedidos (tabla: scaped)
Dim DeVarios As Boolean
Dim Sql As String
On Error GoTo EModificar

    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(1)
            
    If EsDeVarios Then
        If Data1.Recordset.EOF Then Exit Sub
        Sql = " SELECT * FROM sprvar WHERE nifprove='" & Data1.Recordset!nifProve & "' FOR UPDATE "
        conn.Execute Sql
    End If
    
     'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = EsProveedorVarios(Text1(4).Text)
    BloquearDatosProve (DeVarios)
    
EModificar:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificarLinea()
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String
On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If Data2.Recordset.EOF Then Exit Sub
    vWhere = ObtenerWhereCP(False) & " and numlinea=" & Data2.Recordset!numlinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    'A�adiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(2), True 'campo nombre articulo
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Pedidos (scaped)
' y los registros correspondientes de las tablas de lineas (sliped)
Dim cad As String
Dim vTipoMov As CTiposMov
Dim NumPedElim As Long 'Numero del Pedido que se ha Eliminado

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    cad = "Cabecera de Pedidos Compras." & vbCrLf
    cad = cad & "--------------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Pedido:            "
    cad = cad & vbCrLf & "N�:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Proveedor:  " & Format(Text1(4).Text, "000000") & " - " & Text1(5).Text
    cad = cad & vbCrLf & vbCrLf & " �Desea Eliminarlo? "
       
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumPedElim = Data1.Recordset.Fields(0).Value
        
        CadenaSQL = ""
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 81
        frmList.Show vbModal
        Set frmList = Nothing
    
        If CadenaSQL = "" Then Exit Sub
        cad = ""
        cad = DBSet(RecuperaValor(CadenaSQL, 1), "F") & " as fechelim,"
        cad = cad & RecuperaValor(CadenaSQL, 2) & " as trabelim,"
        cad = cad & DBSet(RecuperaValor(CadenaSQL, 3), "T") & " as codincid"
        CadenaSQL = cad
        
        
        If Not Eliminar() Then Exit Sub
        PosicionarDataTrasEliminar
        
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, NumPedElim
        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Pedido", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea Del Pedido. (Tabla: sliped)
Dim Sql As String
On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
            
    ModificaLineas = 3 'Eliminar
    Sql = "�Seguro que desea eliminar la l�nea del Pedido?     "
    Sql = Sql & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    Sql = Sql & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    Sql = Sql & vbCrLf & "Art�culo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        Sql = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        
        Sql = Sql & " and numlinea=" & Data2.Recordset!numlinea
        conn.Execute Sql
        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataTrasEliminar Data2, NumRegElim
        CalcularDatosFactura
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub BotonGenerarAlbaran()
    'Pasar una Pedido a Albaran
Dim Resp As Byte

    'Comprobar que hay un Pedido seleccionado
    If Text1(0).Text = "" Then Exit Sub
        
    'Comprobar que hay lineas
    Resp = 1
    If Not (Data2.Recordset Is Nothing) Then
        If Not Data2.Recordset.EOF Then
            If Not IsNull(Data2.Recordset!numlinea) Then Resp = 0
        End If
    End If
    If Resp = 1 Then
        MsgBox "Pedido sin lineas", vbExclamation
        Exit Sub
    End If
    'Preguntar si se Recibe el pedido completo o no
    Resp = MsgBox("�Recibir el pedido completo?", vbYesNoCancel + vbQuestion)
    If Resp = vbCancel Then Exit Sub
    
    If Resp = vbYes Then 'RECIBIR EL PEDIDO COMPLETO
        AlbCompleto = True
        Screen.MousePointer = vbHourglass

        GenerarAlbaran
        TerminaBloquear
        
    ElseIf Resp = vbNo Then 'RECIBIR PEDIDO INCOMPLETO
        AlbCompleto = False
        Me.SSTab1.Tab = 0
        TerminaBloquear
        'Si no se va a servir completo Mostrar lineas para que se indiquen las Servidas
        MsgBox "Introduzca la cantidad  a recibir para cada l�nea.", vbInformation
        Modo = 6
        gridCargado = False
        Me.cmdAceptar.visible = True
        Me.cmdCancelar.visible = True
        PonerModoOpcionesMenu Modo
        CargaGrid DataGrid1, Data2, True, True
        CargaTxtAuxServidas True, True
        PrimeraVez = True
    Else
        TerminaBloquear
    End If

End Sub





Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        cmdRegresar.Caption = "Regresar"
    Else 'Se llama desde alg�n Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
        'DAVID. Pongo a pi�on el numero de pedido. YA NO SE UTILIZA
        'cad = Data1.Recordset.Fields(0)
        'RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1
    
    If Modo = 6 And gridCargado Then '6: Pasar Pedido a Albaran no Completo (Introducir las servidas)
        CargaTxtAuxServidas True, True
        txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
    End If
    
    If Modo = 5 Then 'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            'Poner descripcion de ampliacion lineas
            Text2(16).Text = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numpedpr", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
        Else
            Text2(16).Text = ""
        End If
        
        '- centro de coste
        ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste familia
        If Not Data2.Recordset.EOF And vEmpresa.TieneAnalitica Then
            Me.txtAux(8).Text = DBLet(Data2.Recordset!CodCCost, "T")
            Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
        Else
            txtAux2(8).Text = ""
        End If
    End If
    Exit Sub
    
Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
    If MostrarDatos <> "" Then
        MostrarDatos = ""
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim SelectInicial As String

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 19
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Bot�n Buscar
        .Buttons(2).Image = 2   'Bot�n Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 26 'Generar Albaran
        
        .Buttons(14).Image = 16 'Imprimir Pedido
        .Buttons(16).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
    End With
    Me.SSTab1.Tab = 0
      
    LimpiarCampos   'Limpia los campos TextBox
    cmdAux(2).Tag = -1
          
    '## A mano
     Me.FrameHco.visible = EsHistorico
    
    
    'Si no lleva datosdvolverbusqueda
    
    If Not EsHistorico Then
        NombreTabla = "scappr"
        NomTablaLineas = "slippr" 'Tabla lineas de Pedido
        Me.Caption = "Pedidos Proveedores"
        Ordenacion = " ORDER BY numpedpr "

    Else
        NombreTabla = "schppr"
        NomTablaLineas = "slhppr"
        CargarTagsHco Me, "scappr", NombreTabla
        'Estos campos solo estan en la tabla del hist�rico
        Text1(24).Tag = "Fecha Eliminaci�n|F|N|||" & NombreTabla & "|fechelim|dd/mm/yyyy|N|"
        Text1(25).Tag = "Trabajador Eliminaci�n|N|N|0|9999|" & NombreTabla & "|trabelim|0000|N|"
        Text1(26).Tag = "Incidencia elim.|T|N|||" & NombreTabla & "|codincid||N|"
        Me.Caption = "Hist�rico Pedidos Proveedores"
        Ordenacion = " ORDER BY numpedpr,fecpedpr "
    End If
    
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    If MostrarDatos = "" Then
        CodTipoMov = "-1"
    Else
        CodTipoMov = MostrarDatos
    End If
    Data1.RecordSource = "Select * from " & NombreTabla & "  WHERE numpedpr= " & CodTipoMov
    Data1.Refresh
    
    Me.Tag = "" 'Para que no carge los datos
 
    If MostrarDatos = "" Then
        PonerModo 0
    Else
        PonerModo 2
    End If
    
    
    CodTipoMov = "PEC"
    VieneDeBuscar = False
    
    
    
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True

    'Poner los grid sin apuntar a nada
    If MostrarDatos = "" Then LimpiarDataGrids
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkRestoPed.Value = 0
    
    Text3(0).Text = "TOTAL"
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    conn.Execute "DELETE FROM tmpnseries WHERE codusu=" & vUsu.Codigo
    'DatosADevolverBusqueda2 = "
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If Val(cmdAux(2).Tag) > 0 Then
            'Llama desde boton busqueda centros de coste
            ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste familia
            Me.txtAux(8).Text = RecuperaValor(CadenaDevuelta, 1)
            Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
        Else
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
                cadB = cadB & " and " & Aux
            End If
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Text1(22).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod cliente
    FormateaCampo Text1(22)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom clien
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim indice As Byte
Dim devuelve As String

    indice = 9
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'poblacion
    Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve)
    'provincia
    Text1(indice + 2).Text = devuelve
End Sub


Private Sub frmDir_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Direcciones
Dim indice As Byte
    indice = CByte(Me.imgBuscar(0).Tag)
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Direccion
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Direc

    CargarDatosDirec Text1(indice).Text, indice
End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim indice As Byte
    indice = CByte(Me.imgFecha(0).Tag) + 1
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim indice As Byte

    indice = 12
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de incidencias
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod incidencia
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2) 'nom incidencia
End Sub

Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'Aqui devuelve los valores que se introducen en el Listado
'para pasar de Pedido a Albaran, o para pasar al historico
    
    CadenaSQL = CadenaSeleccion
End Sub



Private Sub frmNSerie_CargarNumSeries()
'Insertar un registro en la tabla "sserie" por cada uno de los
'N� de Serie introducidos en la Tabla Temporal
Dim RStmp As ADODB.Recordset
Dim RSalb As ADODB.Recordset
Dim Sql As String
Dim i As Byte
Dim b As Boolean
    
    On Error GoTo EInsertar

    
    Sql = "SELECT slialp.codartic, numlinea, cantidad "
    Sql = Sql & " FROM slialp INNER JOIN sartic on slialp.codartic=sartic.codartic "
    Sql = Sql & " WHERE numalbar=" & DBSet(Me.cmdAux(1).Tag, "T") & " and fechaalb=" & DBSet(Me.cmdAux(0).Tag, "F") & " and "
    Sql = Sql & "slialp.codprove=" & Text1(4).Text
    Sql = Sql & " And nseriesn = 1 "
    Sql = Sql & " ORDER BY codartic, numlinea "

    Set RSalb = New ADODB.Recordset
    RSalb.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RSalb.EOF 'Para cada linea del ALbaran
        'Recuperar los N� Serie de ese articulo cargados en la Temporal
        'Seleccionar los n� de serie cargados en la temporal: tmpnseries
        Sql = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
        Sql = Sql & " AND codartic=" & DBSet(RSalb!codArtic, "T")
        Sql = Sql & " ORDER BY codartic, numlinea "
        Set RStmp = New ADODB.Recordset
        RStmp.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        'If Not RStmp.EOF Then RStmp.MoveFirst
        'Intentar asignar un N� serie al total de cantidad del articulo
        
        b = True
        For i = 1 To RSalb!Cantidad
            If Not RStmp.EOF Then
                InsertarNSerie RStmp!numSerie, RStmp!codArtic, RSalb!numlinea
                RStmp.MoveNext
            End If
        Next i
        RStmp.Close
        Set RStmp = Nothing
        RSalb.MoveNext
    Wend
    RSalb.Close
    Set RSalb = Nothing
EInsertar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Insertando N� Serie", Err.Description
End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores
Dim indice As Byte

    indice = 4
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Prove
    FormateaCampo Text1(indice)
End Sub

Private Sub frmProveV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento Proveedores varios
Dim indice As Byte

    indice = 6
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'nif Prove
    Text1(indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'nom Prove
    PonerDatosProveVario Text1(indice).Text
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim indice As Byte

    indice = Val(Me.imgBuscar(0).Tag)
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
    FormateaCampo Text1(indice)
    If indice = 23 Then indice = 1
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0 'Cod. Proveedor
            indice = 4
            Set frmProv = New frmComProveedores
            frmProv.DatosADevolverBusqueda = "0"
            frmProv.Show vbModal
            Set frmProv = Nothing
            
        Case 1 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            indice = 9
            VieneDeBuscar = True
            
        Case 2, 8 'Realizada Por Trabajador
            If Index = 2 Then
                indice = 3
            Else
                indice = 23
            End If
            Me.imgBuscar(0).Tag = indice
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 3 'Forma de Pago
            indice = 12
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 4, 5 'Direccion
            If Index = 4 Then indice = 15
            If Index = 5 Then indice = 2
            Me.imgBuscar(0).Tag = indice
            Set frmDir = New frmComDirecciones
            frmDir.DatosADevolverBusqueda = "0"
            frmDir.Show vbModal
            Set frmDir = Nothing
            
        Case 6 'NIF de Proveedores VARIOS
            indice = 6
            Set frmProveV = New frmComProveV
            frmProveV.DatosADevolverBusqueda = "0"
            frmProveV.Show vbModal
            Set frmProveV = Nothing
            
        Case 7 'Cliente
            indice = 22
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0"
            frmCli.Show vbModal
            Set frmCli = Nothing
            
        Case 10 'Incidencias
            indice = 26
            Set frmInc = New frmIncidencias
            frmInc.DatosADevolverBusqueda = "0"
            frmInc.Show vbModal
            Set frmInc = Nothing
    End Select
    
    PonerFoco Text1(indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   indice = Index + 1
   Me.imgFecha(0).Tag = Index
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub

Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
         Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub mnGenAlbaran_Click()
    'bloqueamos el pedido y lo pasamos a Albaran
    If BLOQUEADesdeFormulario(Me) Then BotonGenerarAlbaran
End Sub


Private Sub mnImpPedido_Click()
'Imprime un Pedido
       frmListadoOfer.NumCod = Text1(0).Text    'N� de Pedido
       frmListadoOfer.CodClien = Text1(4).Text 'Cod.Proveedor
       If EsHistorico Then
            AbrirListadoOfer (56) '59: Informe de Pedidos Compras (Historico)
            frmListadoOfer.FecEntre = Text1(1).Text
       Else
            AbrirListadoOfer (55) '55: Informe de Pedidos Compras
       End If
End Sub

Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Pedidos"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Pedido
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'A�adir lineas
         BotonAnyadirLinea
    Else 'A�adir Cabecera de Pedidos
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
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


'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ind As Integer
Dim b As Boolean
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
    
    If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
        b = False
        If Text1(Index).Text = "" Then
            b = True
        Else
            If Text1(Index).SelLength = Len(Text1(Index).Text) Then b = True
        End If
        If b Then
                Ind = -1
                Select Case Index
                Case 2
                    Ind = 5
                Case 3
                    Ind = 2
                Case 4
                    Ind = 0
                Case 6
                    Ind = 6
                Case 9
                    Ind = 1
                Case 12
                    Ind = 3
                Case 15
                    Ind = 4
                Case 22, 23
                    Ind = Index - 15
                End Select
                If Ind >= 0 Then
                    PulsadoMas2 = True
                    PulsarTeclaMas True, Ind
                End If
            End If
        End If
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
Dim i As Byte
        
        
    'Han pulsado el mas
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        Text1(Index).Text = ""
        Exit Sub
    End If
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
       
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha Oferta, Fecha Entrega
            PonerFormatoFecha Text1(Index)
        
        Case 3, 23 'Cod Trabajador
            i = Index
            If Index = 23 Then i = 1
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(i).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba", "el Trabajador")
            Else
                Text2(i).Text = ""
            End If
            
        Case 4 'Cod. Prove
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Busqueda
                    Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
                Else ' cargar datos de Tabla sprove
                    PonerDatosProveedor (Text1(Index).Text)
                End If
            Else
                LimpiarDatosProve
            End If
            
         Case 6 'NIF
            If Not EsDeVarios Or Modo <> 3 Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(6).Text = Data1.Recordset!nifProve Then
                    Exit Sub
                End If
            End If
            PonerDatosProveVario (Text1(Index).Text)
             
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
            
        Case 12 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 13, 14 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then   'Tipo 4: Decimal(4,2)
                 If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 15, 2 'Cod. Direccion
            If PonerFormatoEntero(Text1(Index)) Then
                Me.imgBuscar(0).Tag = Index
                If Not CargarDatosDirec(Text1(Index).Text, CByte(Index)) Then
                    PonerFoco Text1(Index)
                End If
            Else
                LimpiarDatosDirec CByte(Index)
            End If
            
        Case 22 'cod.cliente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(0).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
            Else
                Text2(0).Text = ""
            End If
            
        Case 21
            If Me.ActiveControl.Name = "SSTab1" Then PonerFocoBtn Me.cmdAceptar
            
        Case 26 'cod Incidencia de eliminacion
            If EsHistorico Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sincid", "nomincid")
                If Not (Text2(Index).Text = "" And Text1(Index).Text <> "") Then
                    PonerFocoBtn Me.cmdAceptar
                Else
                    PonerFoco Text1(Index)
                End If
            End If
    End Select
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
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
'    If EsCabecera Then
        cad = cad & ParaGrid(Text1(0), 15, "N� Pedido")
        cad = cad & ParaGrid(Text1(1), 20, "Fecha Ped.")
        cad = cad & ParaGrid(Text1(4), 15, "Proveedor")
        cad = cad & ParaGrid(Text1(5), 50, "Nombre Prov.")
        Tabla = NombreTabla
        Titulo = "Pedidos Compras"
        If EsHistorico Then
            Titulo = "Hist�rico de Pedidos"
            devuelve = "0|1|"
        Else
            Titulo = "Pedidos"
            devuelve = "0|"
        End If
'    End If
    
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
        frmB.vConexionGrid = conAri 'Conexi�n a BD: Aritaxi
'        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
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
            PonerFoco Text1(0)
        End If
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


Private Sub PonerCamposLineas()
'Carga las Pesta�as con las tablas de lineas del Trabajador seleccionado para mostrar
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slippr
    CargaGrid DataGrid1, Data2, True

    Screen.MousePointer = vbDefault
    Exit Sub
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Realizado por
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    Text2(12).Text = PonerNombreDeCod(Text1(12), conAri, "sforpa", "nomforpa")
    'Cliente para
    Text2(0).Text = PonerNombreDeCod(Text1(22), conAri, "sclien", "nomclien")
    'Solicitado por
    Text2(1).Text = PonerNombreDeCod(Text1(23), conAri, "straba", "nomtraba", "codtraba")
    
    'Poner las direcciones
    CargarDatosDirec Text1(15).Text, 15
    CargarDatosDirec Text1(2).Text, 2
    
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Pedidos
    
    If EsHistorico Then
        'poner datos de eliminacion
        Text2(25).Text = PonerNombreDeCod(Text1(25), conAri, "straba", "nomtraba", "codtraba")
        Text2(26).Text = PonerNombreDeCod(Text1(26), conAri, "sincid", "nomincid", "codincid")
    End If
    
    CalcularDatosFactura 'rellenar campos pesta�a de totales
    
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
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    lblF.Caption = ""
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    If Modo = 6 Then Me.lblIndicador.Caption = "Insertar Cant. Servidas"
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos

        cmdRegresar.visible = False
 
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    'Campo Numero de Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    b = (Modo <> 1)
    BloquearTxt Text1(0), b, True
       
    'datos cliente siempre bloqueados hasta que sea de varios
    If Modo = 3 Then
        EsDeVarios = False
        BloquearDatosProve (EsDeVarios)
    End If
       
       
    '-----  Datos Totales de Factura siempre bloqueado
    For i = 33 To 50
        BloquearTxt Text3(i), True
    Next i
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    Text3(46).BackColor = &HFFFFC0
    Text3(47).BackColor = &HFFFFC0
    Text3(48).BackColor = &HFFFFC0
    Text3(49).BackColor = &HC0C0FF    'Tatal factura
    Text3(50).BackColor = &HC0C0FF    'Tatal factura
    '---------------------------------------------------
       
       
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    BloquearTxt Text2(16), (Modo <> 5)
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = b
    Next i
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    Me.imgBuscar(1).visible = False
           
    'Modo Linea de Ofertas. Poner el campo ampliacion linea
    Me.Label1(35).visible = (Modo = 5)
    Me.Text2(16).visible = (Modo = 5)
    BloquearTxt Text2(16), True
    
    ' ---- [20/10/2009] [LAURA] : a�adir del centro de coste
    Me.Label1(46).visible = (vEmpresa.TieneAnalitica) And (Modo = 5)
    Me.txtAux2(8).visible = (vEmpresa.TieneAnalitica) And (Modo = 5)
    BloquearTxt txtAux2(8), True
    
    
       
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
       
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprueba si los datos de la cabecera son correctos antes de Insertar o Modificar el
'Pedido
Dim b As Boolean
On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
            
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
'Comprueba si los datos de una linea son correctos antes de Insertar o Modificar
'una linea del Pedido
Dim b As Boolean
'Dim devuelve As String
Dim i As Byte
Dim vArtic As CArticulo

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    'Comprobar que los campos NOT NULL tienen valor
    For i = 0 To txtAux.Count - 1
        If txtAux(i).Text = "" Then
            If i = 8 And vEmpresa.TieneAnalitica = False Then
                'no hace nada pq puede ser nulo
            Else
                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
        
    'Comprobar que existe el articulo en el almacen seleccionado
    Set vArtic = New CArticulo
    vArtic.Codigo = txtAux(1).Text
    If Not vArtic.ExisteEnAlmacen(txtAux(0).Text) Then
        b = False
        PonerFoco txtAux(1)
    End If
    Set vArtic = Nothing
    
'    devuelve = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", txtAux(1).Text, "T", , "codalmac", txtAux(0).Text, "N")
'    If devuelve = "" Then
'        MsgBox "No existen unidades del Art�culo: " & txtAux(1).Text & "  en el Almacen: " & txtAux(0).Text, vbExclamation
'        b = False
'        PonerFoco txtAux(1)
'    End If
    
    DatosOkLinea = b
    Exit Function
    
EDatosOkLinea:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Ampliacion linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Ampliaci�n linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
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
        Case 10  'Lineas
            mnLineas_Click
        Case 11 'Generar Albaran
            mnGenAlbaran_Click
        Case 14 'Imprimir Pedido
             mnImpPedido_Click
        Case 16    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
       
    J = Val(Me.mnGenAlbaran.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenAlbaran.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub
    
    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slipre
Dim Sql As String
Dim numlinea As String, vWhere As String
On Error GoTo EInsertarLinea

    InsertarLinea = False
    Sql = ""

    If DatosOkLinea() Then 'Lineas de Pedidos
        'Conseguir el siguiente numero de linea
        vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
        
        Sql = "INSERT INTO " & NomTablaLineas
        Sql = Sql & "(numpedpr,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, recibida, precioar, dtoline1, dtoline2, importel,codccost) "
        Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        Sql = Sql & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        Sql = Sql & DBSet(txtAux(3).Text, "N") & ", 0,"
        Sql = Sql & DBSet(txtAux(4).Text, "S") & "," & DBSet(txtAux(5).Text, "N") & ", "   'Mayo 2009   La "N" es ahora una "S"
        Sql = Sql & DBSet(txtAux(6).Text, "N") & ", " 'Dto 2
        Sql = Sql & DBSet(txtAux(7).Text, "N") & "," 'importe
        Sql = Sql & DBSet(txtAux(8).Text, "T", "S") 'centro coste
        Sql = Sql & ")"
    End If
    
    If Sql <> "" Then
        conn.Execute Sql
        InsertarLinea = True
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Pedido" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim Sql As String
On Error GoTo EModificarLinea

    ModificarLinea = False
    Sql = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        Sql = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
        Sql = Sql & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        Sql = Sql & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
        Sql = Sql & "precioar= " & DBSet(txtAux(4).Text, "S") & ", "   'MAYO 2009.  La "N" es ahora una "S"
        Sql = Sql & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
        Sql = Sql & "importel= " & DBSet(txtAux(7).Text, "N") & ", "
        Sql = Sql & "codccost= " & DBSet(txtAux(8).Text, "T", "S")
        Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
    End If
    
    If Sql <> "" Then
        conn.Execute Sql
        ModificarLinea = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Pedido" & vbCrLf & Err.Description
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
        Me.lblIndicador.Caption = "L�neas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    
    'Habilitar las opciones correctas del menu seg�n Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu seg�n Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean, Optional conServidas As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim b As Boolean
Dim Sql As String

On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    Sql = MontaSQLCarga(enlaza, conServidas)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez

    If conServidas Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
    CargaGrid2 vDataGrid, vData, conServidas
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not b
    PrimeraVez = False
    gridCargado = True
    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, Optional conServidas As Boolean)
Dim i As Byte
On Error GoTo ECargaGrid

    vData.Refresh
    
    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False
    i = 1
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                i = i + 1
                vDataGrid.Columns(i).Caption = "Alm."
                If conServidas Then
                    vDataGrid.Columns(i).Width = 450
                Else
                    vDataGrid.Columns(i).Width = 500
                End If
                vDataGrid.Columns(i).NumberFormat = "000"
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Articulo"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 1600
                Else
                    vDataGrid.Columns(i).Width = 1700
                End If
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Desc. Art�culo"
                If conServidas Then
                    If vEmpresa.TieneAnalitica Then
                        vDataGrid.Columns(i).Width = 3100
                    Else
                        vDataGrid.Columns(i).Width = 3200
                    End If
                Else
                    If vEmpresa.TieneAnalitica Then
                        vDataGrid.Columns(i).Width = 3400
                    Else
                        vDataGrid.Columns(i).Width = 3600
                    End If
                End If
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Ampl. L�nea"
                vDataGrid.Columns(i).Width = 7980
                vDataGrid.Columns(i).visible = False
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Cantidad"
                vDataGrid.Columns(i).Width = 900
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoImporte
                
                i = i + 1
                If conServidas Then
                    'Cargar el grid con la columna de cantidad servida
                    vDataGrid.Columns(i).Caption = "Recibidas"
                    vDataGrid.Columns(i).Width = 800
                    vDataGrid.Columns(i).Alignment = dbgRight
                    vDataGrid.Columns(i).NumberFormat = FormatoImporte
                    i = i + 1
                End If
                vDataGrid.Columns(i).Caption = "Precio"
                If vEmpresa.TieneAnalitica Then
                    vDataGrid.Columns(i).Width = 1100
                Else
                    vDataGrid.Columns(i).Width = 1200
                End If
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoPrecio2   'Mayo 2009
                
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Dto.1"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 550
                Else
                    vDataGrid.Columns(i).Width = 600
                End If
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoDescuento
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Dto.2"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 550
                Else
                    vDataGrid.Columns(i).Width = 600
                End If
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoDescuento
            
                i = i + 1
                vDataGrid.Columns(i).Caption = "Importe"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 1100
                Else
                    vDataGrid.Columns(i).Width = 1300
                End If
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoImporte
                
                
                '---- [19/10/2009] [LAURA] : se a�ade el centro de coste
                i = i + 1
                If vEmpresa.TieneAnalitica Then
                    vDataGrid.Columns(i).Caption = "CCoste"
                    If conServidas Then
                        vDataGrid.Columns(i).Width = 650
                    Else
                        vDataGrid.Columns(i).Width = 700
                    End If
                Else
                    vDataGrid.Columns(i).visible = False
                End If
                
'                vDataGrid.Columns(i).Alignment = dbgRight
'                vDataGrid.Columns(i).NumberFormat = FormatoImporte
    End Select

    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    vDataGrid.HoldFields
    
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 1 'TextBox
            txtAux(i).Top = 290
            txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        cmdAux(2).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                BloquearTxt txtAux(i), False
            Next i
        Else 'Vamos a modificar
            For i = 0 To txtAux.Count - 1
                If i < 3 Then
                    txtAux(i).Text = DataGrid1.Columns(i + 2).Text
                Else
                    txtAux(i).Text = DataGrid1.Columns(i + 3).Text
                End If
                txtAux(i).Locked = False
            Next i
        End If
        
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(7), True
        
        
        ' ---- [20/10/2009] [LAURA] : a�adir centro de coste
        BloquearTxt txtAux(8), Not (vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 2)
        Me.cmdAux(2).Enabled = Not txtAux(8).Locked
        Me.cmdAux(2).visible = Me.cmdAux(2).Enabled
        ' ----
        
        

        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 20)
        
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(2).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        cmdAux(2).Height = DataGrid1.RowHeight
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(2).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(3).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(4).Width - 10
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(6).Width - 10
        'Precio, Dto1, Dto2, Precio
        For i = 4 To txtAux.Count - 1
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 10
            txtAux(i).Width = DataGrid1.Columns(i + 3).Width - 10
        Next i
        
        cmdAux(2).Left = txtAux(i - 1).Left + txtAux(i - 1).Width - cmdAux(2).Width
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux.Count - 2
            txtAux(i).visible = visible
        Next i
        txtAux(8).visible = visible And vEmpresa.TieneAnalitica
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    End If
End Sub


Private Sub CargaTxtAuxServidas(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
'Carga el TxtAux(3) con el campo RECIBIDAS de la tabla slippr
Dim alto As Single
Dim i As Byte

    i = 3
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux(i).Top = 290
        txtAux(i).visible = visible
        txtAux(i).BackColor = vbWhite
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            txtAux(i).Text = ""
            BloquearTxt txtAux(i), False
            txtAux(i).BackColor = &H80000013
        End If
      
        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 230
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        txtAux(i).Top = alto
        txtAux(i).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cantidad servida
        alto = DataGrid1.Left + 330 + DataGrid1.Columns(2).Width + DataGrid1.Columns(3).Width
        alto = alto + DataGrid1.Columns(4).Width + DataGrid1.Columns(6).Width
        txtAux(i).Left = alto
        txtAux(i).Width = DataGrid1.Columns(7).Width - 15
        
        'Los ponemos Visibles o No
        '--------------------------
        txtAux(i).visible = visible
        PonerFoco txtAux(i)
    End If
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
    If Index = 3 Or Index = 4 Then
        If Modo <> 6 Then
            If Index = 3 Then
                lblF.Caption = "Ver articulo"
            Else
                lblF.Caption = "Ver precio"
            End If
        End If
    Else
        lblF.Caption = ""
    End If
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Modo <> 6 Then
        KEYpress KeyAscii
    Else 'Pasar el Pedido a Albaran
        If KeyAscii = 13 Then 'ENTER
            PonerServidas True
'            ConseguirFoco txtAux(3), Modo
        End If
    End If
End Sub




Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Modo <> 6 Then 'Pasar de Pedido a Albaran
        ' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
        If KeyCode = 113 Then
            If Index = 3 Then AbrirForm_Articulos
            If Index = 4 And txtAux(1).Text <> "" Then
                frmListadoPrecios.Opcion = 0
                frmListadoPrecios.CadenaPasoDatos = txtAux(1).Text & "|" & Text1(4).Text & "|"
                frmListadoPrecios.Show vbModal
            End If
        
        
         
        
        Else
          If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
                If Index < 2 Or Index = 8 Then  'Para los que tienen busqueda
                    If Modo = 5 And ModificaLineas = 1 Then
                        If txtAux(Index).Text = "" Then
                            PulsadoMas2 = True
                            KeyCode = 0
                
                            PulsarTeclaMas False, Index
                        End If
                    End If
                End If
            End If
        End If
    Else 'Modo lineas
        Select Case KeyCode
            Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Row > 0 Then
                    DataGrid1.Row = DataGrid1.Row - 1
                    CargaTxtAuxServidas True, True
                Else
                    If Data2.Recordset.BOF Then
                        PonerFoco txtAux(3)
                    Else
                        gridCargado = False
                        Data2.Recordset.MovePrevious
                        gridCargado = True
                        If Data2.Recordset.BOF Then Data2.Recordset.MoveFirst
                         If DataGrid1.Row > 0 Then
                            DataGrid1.Row = DataGrid1.Row - 1
                            CargaTxtAuxServidas True, True

                        End If
                    End If
                End If
                txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
                ConseguirFoco txtAux(3), Modo
                
            Case 40 'Desplazamiento Flecha Hacia Abajo
'                If DataGrid1.Row < Data2.Recordset.RecordCount - 1 Then
'                    DataGrid1.Row = DataGrid1.Row + 1
'                    CargaTxtAuxServidas True, True
'                Else
'                    PonerFocoBtn Me.cmdAceptar
'                End If
'                txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
'                ConseguirFoco txtAux(3), Modo
                
                PonerServidas True
        End Select
    End If
End Sub



Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String
'Dim vPrecio As CPreciosCom
Dim TipoDto As Byte
Dim b As Boolean


    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtAux(Index).Text = ""
        Exit Sub
    End If

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 0 'Cod ALMACEN
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(Index)

        Case 1 'Cod. ARTICULO
            If txtAux(1).Text = "" Then
                txtAux(2).Text = ""
                Exit Sub
            End If
            
            If txtAux(0).Text = "" Then
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If
            
            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, , , devuelve) Then
                '---- [20/10/2009] [LAURA] : a�adir centro de coste
                If Not vEmpresa.TieneAnalitica Then
                    txtAux(8).Text = ""
                ElseIf vParamAplic.ModoAnalitica = 1 Then 'por familia
                    txtAux(8).Text = devuelve
                    Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
                End If
                '----
                
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
            Else
                PonerFoco txtAux(Index)
            End If
            
            
'            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov) Then
'                If txtAux(2).Locked Then PonerFoco txtAux(3)
'                'Si es articulo de varios podemos modificar la descripci�n del articulo, sino bloqueamos.
''                If Not EsArticuloVarios(txtAux(Index).Text) Then
''                    BloquearTxt txtAux(2), True
''                Else
''                    BloquearTxt txtAux(2), False
''                    PonerFoco txtAux(2)
''                End If
'            Else
'                PonerFoco txtAux(Index)
'            End If
            
        Case 2 'Desc. Articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Comprobar si hay suficiente stock
                If (Modo = 5) And (ModificaLineas = 1 Or (ModificaLineas = 2 And txtAux(4).Text = "")) Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    ObtenerPrecioCompra
                    
'                    Set vPrecio = New CPreciosCom
'                    If vPrecio.Leer(txtAux(1).Text, Text1(4).Text) Then
'                        If vPrecio.ComprobarCantidad(CInt(txtAux(3).Text)) Then
'                            txtAux(4).Text = vPrecio.ObtenerPrecio(Text1(1).Text)
'                            PonerFormatoDecimal txtAux(4), 2
'                            txtAux(5).Text = vPrecio.Descuento1
'                            PonerFormatoDecimal txtAux(5), 4
'                            txtAux(6).Text = vPrecio.Descuento2
'                            PonerFormatoDecimal txtAux(6), 4
'                        Else
'                            PonerFoco txtAux(Index)
'                        End If
'                    End If
'                    Set vPrecio = Nothing
                End If
            End If
            
        Case 4 'Precio
            PonerFormatoDecimal_Single txtAux(Index), 9 'Tipo 9: Decimal(10,5)parametros
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
            
        Case 8 'COD. CENTRO DE COSTE
            ' ---- [20/10/2009] [LAURA]: a�adir centro de coste a la linea
            If txtAux(Index).Text = "" Then
                 txtAux2(Index).Text = ""
            ElseIf vEmpresa.TieneAnalitica Then
                'centro de coste
                ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste familia
                Me.txtAux2(Index).Text = PonerNombreCCoste(Me.txtAux(Index))
            End If
    End Select
    
    If Modo = 5 Then
         If (Index = 3 Or Index = 4 Or Index = 5 Or Index = 6) Then 'Cant., Precio, Dto1, Dto2
'            If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'            If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
            If txtAux(1).Text = "" Then Exit Sub
            TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
            txtAux(7).Text = CalcularImporteSng(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, TipoDto)
            PonerFormatoDecimal txtAux(7), 1
        End If
    End If
End Sub



Private Sub ObtenerPrecioCompra()
Dim vPrecio As CPreciosCom
Dim cad As String

    On Error GoTo EPrecios
    
    Set vPrecio = New CPreciosCom
    If vPrecio.Leer(txtAux(1).Text, Text1(4).Text) Then
        If vPrecio.ComprobarCantidad(CInt(txtAux(3).Text)) Then
            txtAux(4).Text = vPrecio.ObtenerPrecio(Text1(1).Text)    'FALTARA QUE DEVUELVE 5 decimales
'            PonerFormatoDecimal txtAux(4), 2
            txtAux(5).Text = vPrecio.Descuento1
'            PonerFormatoDecimal txtAux(5), 4
            txtAux(6).Text = vPrecio.Descuento2
'            PonerFormatoDecimal txtAux(6), 4
        Else
            PonerFoco txtAux(3)
            Exit Sub
        End If
    Else
        'Obtener el ult. precio de compra de ese articulo (sartic)
        cad = DevuelveDesdeBDNew(conAri, "sartic", "preciouc", "codartic", txtAux(1).Text, "T")
        If cad <> "" Then
            txtAux(4).Text = cad
            txtAux(5).Text = "0"
            txtAux(6).Text = "0"
        End If
    End If
    PonerFormatoDecimal_Single txtAux(4), 9   '10,5
    PonerFormatoDecimal txtAux(5), 4
    PonerFormatoDecimal txtAux(6), 4
    
    Set vPrecio = Nothing
    
EPrecios:
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub BotonMtoLineas(numTab As Integer, cad As String)
        Me.SSTab1.Tab = numTab
        TituloLinea = cad
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
        DataGrid1_RowColChange 1, 1
End Sub


Private Function Eliminar() As Boolean
Dim b As Boolean
Dim vWhere As String
On Error GoTo FinEliminar

        conn.BeginTrans
         vWhere = ObtenerWhereCP(False)

'        If opt = 1 Then 'ELIMINAR
'            b = EliminarPedido(Data1.Recordset!numpedpr)
'        Else 'Pasar al HISTORICO
            b = ActualizarElTraspaso("", vWhere, CodTipoMov, CadenaSQL)
'        End If
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Pedido"
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ning�n registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = ObtenerWhereCP(False)
         vWhere = Replace(vWhere, NombreTabla & ".", "")
         If SituarData(Data1, vWhere, Indicador) Then
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


Private Sub PosicionarDataTrasEliminar()
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    If SituarDataTrasEliminar(Data1, NumRegElim) Then
        PonerCampos
    Else
        LimpiarCampos
        LimpiarDataGrids
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP(conW As Boolean) As String
'Obtiene la where de la Clave Primaria de la tabla de Cabecera: scaped
Dim Sql As String
On Error Resume Next
    Sql = ""
    If conW Then Sql = " WHERE "
    Sql = Sql & NombreTabla & ".numpedpr= " & Val(Text1(0).Text)
    If EsHistorico Then Sql = Sql & " AND " & NomTablaLineas & ".fecpedpr=" & DBSet(Text1(1).Text, "F")
    ObtenerWhereCP = Sql
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Optional conServidas As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data2
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    Sql = "SELECT numpedpr, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, "
    If conServidas Then Sql = Sql & "recibida, "
'    SQL = SQL & "precioar, origpre, dtoline1, dtoline2,importel "
    Sql = Sql & "precioar, dtoline1, dtoline2,importel,codccost "
    Sql = Sql & " FROM " & NomTablaLineas
    If enlaza Then
        Sql = Sql & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        If EsHistorico Then Sql = Sql & " and fecpedpr='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        Sql = Sql & " WHERE numpedpr = -1"
    End If
    Sql = Sql & " Order by numpedpr, numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean

        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0) And Not EsHistorico
        Me.mnNuevo.Enabled = (b Or Modo = 0) And Not EsHistorico
        'Modificar
        Toolbar1.Buttons(6).Enabled = b And Not EsHistorico
        Me.mnModificar.Enabled = b And Not EsHistorico
        'eliminar
        Toolbar1.Buttons(7).Enabled = b And Not EsHistorico
        Me.mnEliminar.Enabled = b And Not EsHistorico
            
        b = (Modo = 2) And Not EsHistorico
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = (Modo = 2)
        Me.mnLineas.Enabled = (Modo = 2)
        'Generar Albaran desde Pedido
        Toolbar1.Buttons(11).Enabled = b
        Me.mnGenAlbaran.Enabled = b
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub


Private Function CargarDatosDirec(CodDirec As String, indice As Byte) As Boolean
'Direcciones Propias
Dim RS As ADODB.Recordset
Dim devuelve As String
Dim b As Boolean
On Error GoTo ECargarProve

    b = False
    If CodDirec <> "" Then
        devuelve = "Select nomdirec, domdirec, codpobla, pobdirec, prodirec "
        devuelve = devuelve & " FROM sdirpr Where coddirec=" & Val(CodDirec)
        
        Set RS = New ADODB.Recordset
        RS.Open devuelve, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RS.EOF Then
            Text1(indice).Text = Format(CodDirec, "000")
            Text2(indice).Text = RS.Fields!nomdirec 'Nom Direccion
            If indice = 2 Then
                indice = 21
            Else
                indice = 17
            End If
            Text2(indice).Text = RS.Fields!domdirec 'Domicilio
            Text2(indice + 1).Text = RS.Fields!codpobla
            Text2(indice + 2).Text = RS.Fields!pobdirec
            Text2(indice + 3).Text = RS.Fields!prodirec
            b = True
        Else
            MsgBox "No existe la direcci�n: " & Text1(indice).Text, vbInformation
            LimpiarDatosDirec (indice)
        End If
        RS.Close
        Set RS = Nothing
    Else
        LimpiarDatosDirec (indice)
        b = True
    End If
    
    CargarDatosDirec = b
    
ECargarProve:
    If Err.Number <> 0 Then CargarDatosDirec = False
End Function


Private Sub LimpiarDatosDirec(indice As Byte)
    Text2(indice).Text = ""
    If indice = 2 Then
        indice = 21
    Else
        indice = 17
    End If
    Text2(indice).Text = "" 'Domicilio
    Text2(indice + 1).Text = "" 'cpostal
    Text2(indice + 2).Text = "" 'poblacion
    Text2(indice + 3).Text = "" 'provincia
End Sub


Private Function InsertarPedido(vSQL As String, vTipoMov As CTiposMov) As Boolean
'Insertar la Cabecera de un Pedido, tabla: scaped
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe un Pedido con ese contador y si existe lo incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numpedpr", "numpedpr", Text1(0).Text, "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Pedidos (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    
    'Actualizar los datos del proveedor si es de varios
    If EsDeVarios Then
        'Si es cliente de varios actualizar datos cliente en tabla:sclvar
        MenError = "Modificando datos proveedor varios."
        bol = ActualizarProveVarios(Text1(4).Text, Text1(6).Text)
    End If
    
    MenError = "Error al actualizar el contador del Pedido."
    vTipoMov.IncrementarContador (CodTipoMov)

EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Pedido." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarPedido = True
        Else
            conn.RollbackTrans
            InsertarPedido = False
        End If
End Function


Private Sub LimpiarDatosProve()
'Limpia los campos del Form con datos del Proveedor
Dim i As Byte

    For i = 4 To 14
        Text1(i).Text = ""
    Next i
End Sub
    


Private Function InicializarCStockAlbar(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String, Optional ByRef RS As ADODB.Recordset) As Boolean
'Para comprobar stock al pasar de Pedido a Albaran de Venta
Dim TipoDto As Byte
On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = "ALC"
    vCStock.Trabajador = CLng(Text1(4).Text) 'En codigope ponemos el Proveedor
    vCStock.Documento = Text1(0).Text
    vCStock.codArtic = RS!codArtic
    vCStock.codAlmac = CInt(RS!codAlmac)
    
    If AlbCompleto Then
        vCStock.Cantidad = CSng(RS!Cantidad)
        If RS.Fields.Count > 3 Then 'Si no se selecciona el campo importe de la tabla es que solo vamos a comprobar stock y no se necesita
            vCStock.Importe = CCur(RS!ImporteL)
        End If
    Else
        vCStock.Cantidad = CSng(RS!recibida)
        'Si se va a Insertar en alguna linea obtener el importe
        'Si solo vamos a comprobar stock no hace falta el importe
        If RS.Fields.Count > 4 Then
            TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Me.Data1.Recordset!codProve, "N")
            vCStock.Importe = CCur(CalcularImporte(RS!recibida, RS!precioar, RS!dtoline1, RS!dtoline2, TipoDto))
        End If
    End If
    
    vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStockAlbar = False
    Else
        InicializarCStockAlbar = True
    End If
End Function


Private Function PasarPedidoAAlbaran(NumAlb As String, FechaAlb As String) As Boolean
'OUT -> numalb: Devuelve el N� de albaran asignado al pedido
Dim bol As Boolean
Dim MenError As String
Dim Sql As String
Dim RS As ADODB.Recordset
Dim vWhere As String
Dim cProve As CProveedor

    On Error GoTo EGenPedido

    bol = False
            
    
    'Aqui empieza transaccion
    conn.BeginTrans
    
    'Insertar en tablas de Albaranes Proveedor el Pedido  (scaalp, slialp)
    bol = InsertarAlbaran(MenError, NumAlb)
    
    
    
    
    
    
    'Para cada linea del pedido:
    ' Actualizar precio medio ponderado del articulo
    ' Actualizar precio y fecha ultima compra del articulo
    ' Actualizar Stock en salmac (entrada de stock), e introducir movimiento en smoval
    If bol Then
        MenError = "Actualizando Stocks"
        bol = InsertarMovStock(NumAlb, FechaAlb)
    End If
    
    If bol Then
        'Actualizar la ult.fecha de compra del Proveedor
        MenError = "Actualizando ultima fecha compra en Proveedor."
        Set cProve = New CProveedor
        bol = cProve.ActualizaFechaUltCompra(Text1(4).Text, FechaAlb)
        Set cProve = Nothing
        
'        If bol Then
'            'Actualizar ult. fecha de compra y el precio ult compra de los articulos del Albaran
'            MenError = "Actualizando ultima fecha compra en Art�culos."
'            SQL = "numalbar=" & DBSet(NumAlb, "T") & " and fechaalb=" & DBSet(FechaAlb, "F") & " and slialp.codprove=" & Text1(4).Text
'            bol = ActualizarUltFechaCom(SQL)
'        End If
    End If
    
    
    If bol Then
        If AlbCompleto Then  'Si se inserta Albaran
            'Borrar el Pedido de las tablas de Pedidos (scaped, sliped)
            MenError = "Eliminando cabecera y lineas del Pedido."
            bol = EliminarPedido(CLng(Text1(0).Text))
        Else
            'Actualizar la cantidad=cantidad-recibida y recibida= 0 en slippr
            bol = ActualizarPedido()
            'Marcar Resto de pedido: restoped=1
            If bol Then bol = ActualizarCabPedido
        End If
    End If
    
    
    
    If bol Then
        'si se ha generado correctamente el ALBARAN ver si hay alguna l�nea que tiene
        'el art�culo con control de n� de lote y pedir los n� de lotes.
        
    End If
    
    
    
    
    If bol Then
        'Se ha generado correctamente el ALBARAN y vemos si tiene N� Series
'        FechaAlb = RecuperaValor(CadenaSQL, 3)
        'Comprobar si Hay N� SERIE en compras y Mostrar
        'ventana para pedir los N� Serie de la cantidad introducida si lo requiere algun articulo
        ComprobarNSeriesLineas NumAlb, FechaAlb
        
        
        If Not AlbCompleto Then
            'Eliminar las filas del pedido que se servieron completas (slippr)
            MenError = "Eliminando lineas pedidido servidas completas."
            Sql = "DELETE FROM " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND cantidad=0"
            conn.Execute Sql
            
            'Comprobar que si no quedan lineas en el pedido se elimine la cabecera del pedido
            MenError = "Eliminando cabecera del pedido."
            Sql = "select codalmac, codartic FROM " & NomTablaLineas & " WHERE numpedpr=" & Text1(0).Text
            Set RS = New ADODB.Recordset
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RS.EOF Then 'No hay lineas de pedido --> Eliminar la cabecera
                Sql = "DELETE FROM " & NombreTabla & " WHERE numpedcl=" & Text1(0).Text
                conn.Execute Sql
            End If
            RS.Close
            Set RS = Nothing
        End If
        bol = True
    End If
    
    
EGenPedido:
    If Err.Number <> 0 Then
'        MenError = "Pasando Pedido a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
'        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        PasarPedidoAAlbaran = True
    Else
        conn.RollbackTrans
        PasarPedidoAAlbaran = False
        MenError = "Pasando Pedido a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
    End If
End Function


Private Function InsertarAlbaran(MenError As String, NumAlb As String) As Boolean
'Devuelve el mensaje de error si se produce
Dim bol As Boolean
Dim vSQL As String
Dim FechaAlb As String
Dim TrabAlb As String

    On Error GoTo EInsertarAlbaran
    
    bol = False
    InsertarAlbaran = bol
    
    NumAlb = RecuperaValor(CadenaSQL, 2)
    FechaAlb = RecuperaValor(CadenaSQL, 3)
    TrabAlb = RecuperaValor(CadenaSQL, 1)
    
    vSQL = "INSERT INTO scaalp (numalbar, fechaalb, codprove, nomprove, domprove, codpobla, pobprove, proprove, nifprove, telprove, codforpa, codtraba, codtrab1, dtoppago, dtognral, observa1, observa2, observa3, observa4, observa5, numpedpr, fecpedpr)"
    vSQL = vSQL & " SELECT " & DBSet(NumAlb, "T") & " as numalbar, " & DBSet(FechaAlb, "F") & " as fechaalb, "
    vSQL = vSQL & "codprove, nomprove, domprove, codpobla, pobprove, proprove, nifprove, telprove, codforpa, "
    vSQL = vSQL & TrabAlb & " as codtraba,codtraba as codtrab1, dtoppago, dtognral, observa1, observa2, observa3, observa4, observa5, numpedpr, fecpedpr "
    vSQL = vSQL & " FROM " & NombreTabla & " WHERE numpedpr=" & Text1(0).Text

    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Albaranes Proveedor (scaalp)."
    conn.Execute vSQL, , adCmdText
    
    'Insertar Lineas Albaran desde Pedido
    MenError = "Error al insertar en la tabla Lineas de Albaran (slialp)."
    If Not InsertarLineasAlbaran(NumAlb, FechaAlb) Then Exit Function
    
    bol = True
    
EInsertarAlbaran:
        If Err.Number <> 0 Then
            bol = False
            MenError = MenError & vbCrLf & Err.Description
        End If
        If bol Then
            InsertarAlbaran = True
        Else
            InsertarAlbaran = False
        End If
End Function


Private Function InsertarLineasAlbaran(NumAlb As String, FechaAlb As String) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
'IN -> TipoM, numAlb
Dim Sql As String
Dim RS As ADODB.Recordset
Dim ImpLinea As String
Dim TipoDto As Byte
On Error GoTo EInsertarLinAlb

    If AlbCompleto Then
        'Insertar en la tabla de Albaran, los registros seleccionados de la tabla de Pedidos
        Sql = ""
        Sql = "SELECT " & DBSet(NumAlb, "T") & " as numalbar, " & DBSet(FechaAlb, "F") & " as fechaalb, " & Val(Text1(4).Text) & " as codprove, numlinea, codartic, codalmac, nomartic, ampliaci, "
        Sql = Sql & "cantidad, precioar, dtoline1, dtoline2, importel,codccost "
        Sql = Sql & " FROM " & NomTablaLineas & " WHERE numpedpr=" & Val(Text1(0).Text)
        Sql = "INSERT INTO slialp (numalbar, fechaalb, codprove, numlinea, codartic, codalmac, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,codccost) " & Sql
        conn.Execute Sql, , adCmdText
    Else
        Sql = "select * from " & NomTablaLineas
        Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not RS.EOF 'Para cada linea de pedido insertar una de albaran si recibidas >0
            If RS!recibida > 0 Then
                TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
                ImpLinea = CalcularImporte(RS!recibida, RS!precioar, RS!dtoline1, RS!dtoline2, TipoDto)
                Sql = "INSERT INTO slialp (numalbar, fechaalb, codprove, numlinea,codartic, codalmac, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,codccost) "
                Sql = Sql & " VALUES(" & DBSet(NumAlb, "T") & ", " & DBSet(FechaAlb, "F") & ", " & Val(Text1(4).Text) & ", " & RS!numlinea & ", "
                Sql = Sql & DBSet(RS!codArtic, "T") & "," & RS!codAlmac & ", " & DBSet(RS!NomArtic, "T") & ", " & DBSet(RS!Ampliaci, "T") & ", "
                Sql = Sql & DBSet(RS!recibida, "N") & ", " & DBSet(RS!precioar, "S") & ", " & DBSet(RS!dtoline1, "N") & ", " & DBSet(RS!dtoline2, "N") & ", "
                Sql = Sql & DBSet(ImpLinea, "N") & ","
                Sql = Sql & DBSet(RS!CodCCost, "T", "S") & ")"
                conn.Execute Sql, , adCmdText
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End If
    
EInsertarLinAlb:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasAlbaran = False
        MuestraError Err.Number, "Insertar lineas albaran.", Err.Description
    Else
        InsertarLineasAlbaran = True
    End If
End Function


Private Function EliminarPedido(numPed As Long) As Boolean
'Eliminar las lineas y la Cabecera de un Pedido. Tablas: scaped, sliped
Dim Sql As String
On Error GoTo EEliminarPed

     Sql = " WHERE  numpedpr=" & numPed

    'Lineas de Pedido
    conn.Execute "Delete from " & NomTablaLineas & Sql
        
    'Cabecera
    conn.Execute "Delete from " & NombreTabla & Sql

EEliminarPed:
    If Err.Number <> 0 Then
        EliminarPedido = False
    Else
        EliminarPedido = True
    End If
End Function


Private Function ActualizarPedido() As Boolean
'Actualiza la tabla de lineas de pedido (sliped)
'cantidad=cantidad-servidas y servidas=0
Dim Sql As String
Dim RS As ADODB.Recordset
Dim ImpLinea As String
Dim TipoDto As Byte

    On Error GoTo EActPedido

    Sql = "select numlinea, codalmac, codartic, cantidad, recibida, precioar, dtoline1, dtoline2 from " & NomTablaLineas
    Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF 'Para cada linea
        TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
        ImpLinea = CalcularImporte(RS!Cantidad - RS!recibida, RS!precioar, RS!dtoline1, RS!dtoline2, TipoDto)
        Sql = "UPDATE " & NomTablaLineas & " SET cantidad=cantidad-recibida, recibida=0, importel=" & DBSet(ImpLinea, "N")
'        SQL = SQL & " WHERE codalmac=" & RS!codAlmac & " AND codartic='" & RS!codArtic & "'"
        Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        Sql = Sql & " AND numlinea=" & RS!numlinea
        conn.Execute Sql
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
EActPedido:
    If Err.Number <> 0 Then
        ActualizarPedido = False
    Else
        ActualizarPedido = True
    End If
End Function


Private Function ActualizarCabPedido() As Boolean
Dim Sql As String
On Error Resume Next

    Sql = "UPDATE " & NombreTabla & " SET restoped=1 " & ObtenerWhereCP(True)
    conn.Execute Sql
    If Err.Number <> 0 Then
        ActualizarCabPedido = False
    Else
        ActualizarCabPedido = True
    End If
End Function


Private Function InsertarMovStock(NumAlb As String, FechaAlb As String) As Boolean
Dim vCStock As CStock
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim cArt As CArticulo

    On Error Resume Next

    InsertarMovStock = False
    
    Set vCStock = New CStock
    b = True
    
    Sql = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Sql = "select * from " & NomTablaLineas & Sql
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    vCStock.Fechamov = FechaAlb
    
    'para cada linea del Pedido Insertar en smoval y Actualizar Stock en salmac
    While (Not RS.EOF) And b
        If InicializarCStockAlbar(vCStock, "E", CStr(RS!numlinea), RS) Then
            vCStock.Documento = NumAlb
            If vCStock.Cantidad <> 0 Then
                '==== Laura 22/09/2006
                '-- antes de actualizar el stock calculamos el precio medio ponderado del articulo
                Set cArt = New CArticulo
                If cArt.LeerDatos(vCStock.codArtic) Then
                    '17 Junio 2009
                    'Si la cantidad es negativa no actualiza ni precio medio ponderado NI fecha ult compra
                    If vCStock.Cantidad >= 0 Then
                    
                        'Laura 19/12/2006: Calcular precio_med_pond con el precio con los descuentos,e.d. importe/cantidad
                        'If Not cArt.ActualizarPrecioMedPond(CCur(vCStock.Cantidad), CCur(RS!precioar)) Then b = False
                        If Not cArt.ActualizarPrecioMedPond(CCur(vCStock.Cantidad), Round2(CCur(vCStock.Importe) / CCur(vCStock.Cantidad), 4)) Then b = False
                        
                        '--actualizar fecha y precio ultima compra del articulo
                        'Laura 19/12/2006: actualizar precio_ult_compra con el precio con los descuentos,e.d. importe/cantidad
                        'If Not cArt.ActualizarUltFechaCompra(vCStock.Fechamov, CStr(RS!precioar)) Then b = False
                        If Not cArt.ActualizarUltFechaCompra(vCStock.Fechamov, Round2(CCur(vCStock.Importe) / CCur(vCStock.Cantidad), 4)) Then b = False

                    End If 'De cantidad >=0
                End If
                Set cArt = Nothing
                '====
            
            
                'en actualizar stock comprobamos si el articulo tiene control de stock
                b = vCStock.ActualizarStock
            End If
        Else
            b = False
        End If
        RS.MoveNext
    Wend
    Set vCStock = Nothing
    RS.Close
    Set RS = Nothing
    
    InsertarMovStock = b
    
End Function


Private Sub ImprimirAlbaran(Opcion As Integer, NumAlbar As String)
End Sub


Private Function ActualizarServidas() As Boolean
Dim Sql As String
On Error Resume Next

    Sql = "UPDATE " & NomTablaLineas & " SET recibida= " & DBSet(txtAux(3).Text, "N")
    Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
    conn.Execute Sql
    
    If Err.Number <> 0 Then
        ActualizarServidas = False
    Else
        ActualizarServidas = True
    End If
End Function


Private Sub PonerServidas(HaciaAlante As Boolean)
Dim NumFila As Integer
Dim cadMen As String

'    NumFila = DataGrid1.Row
    NumFila = Data2.Recordset.AbsolutePosition
    If PonerFormatoDecimal(txtAux(3), 1) Then  'Tipo 1: Decimal(12,2)
        If CCur(txtAux(3).Text) > Data2.Recordset!Cantidad Then
            cadMen = "La cantidad a Recibir no puede ser superior a la del Pedido."
            MsgBox cadMen, vbExclamation
            PonerFoco txtAux(3)
            Exit Sub
        End If
    End If
    ActualizarServidas
    CargaGrid2 DataGrid1, Data2, True
'    DataGrid1.Row = NumFila
    SituarDataPosicion Data2, CLng(NumFila), ""
    If HaciaAlante Then MoverSigRegistro
End Sub




Private Sub MoverSigRegistro()
    On Error GoTo EMover
    
    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.AbsolutePosition <= Data2.Recordset.RecordCount - 1 Then
        DataGrid1.Row = DataGrid1.Row + 1
        CargaTxtAuxServidas True, True
    Else
        PonerFocoBtn Me.cmdAceptar
    End If
    txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
    PonerFoco txtAux(3)
    ConseguirFocoLin txtAux(3)
    txtAux(3).Refresh
EMover:
    If Err.Number <> 0 Then MuestraError Err.Description, "Mover registro.", Err.Description
End Sub





Private Sub GenerarAlbaran()
Dim numPed As Long 'N� Pedido
Dim NumAlb As String 'N� Albaran
Dim FechaAlb As String 'Fecha del Albaran
Dim Sql As String
Dim RS As ADODB.Recordset
Dim b As Boolean

    NumRegElim = Data1.Recordset.AbsolutePosition
    numPed = Data1.Recordset!numpedpr
    
    'pedir por pantalla: el operador, N� albaran y fecha albaran
    Set frmList = New frmListadoOfer
    frmList.OpcionListado = 57
    CadenaSQL = ""
    frmList.Show vbModal
    Set frmList = Nothing
    
    If CadenaSQL = "" Then Exit Sub
    FechaAlb = RecuperaValor(CadenaSQL, 3)
    
    
    'Antes de pasar el pedido al albaran nos guardamos los articulos cuyo precio_compra
    'se han modificado para preguntar despues si se quiere actualizar precios_venta
    'hay q guardarlo antes de pasar pedido a albaran ya q aqui se actualiza el precio_ult_compra
    '-- Laura 19/12/2006: calcular precio_med_pond con el precio aplicados los descuentos, ed. importe/cantidad
    ' Iremos cambiando el numero de decimales poc a poco ANTES era un 4
    Sql = "SELECT slippr.codartic,sartic.nomartic,round(slippr.importel/slippr.cantidad," & PrecioDecimales & ")"
    Sql = Sql & " as precioar,sartic.preciouc,sum(cantidad) "
    Sql = Sql & " FROM slippr INNER JOIN sartic ON slippr.codartic=sartic.codartic "
    'SQL = SQL & " WHERE numpedpr=" & numPed & " and (slippr.precioar<>sartic.preciouc)"
    Sql = Sql & " WHERE numpedpr=" & numPed & " and (round(slippr.importel/slippr.cantidad,4)<>sartic.preciouc)"
    'seleccionar solo de las que se vayan a recibir
    If Not AlbCompleto Then Sql = Sql & " and slippr.recibida>0 "
    Sql = Sql & " group by slippr.codartic,slippr.precioar,sartic.preciouc "
    Sql = Sql & " Having Sum(Cantidad) > 0"
    b = ObtenerRSprecios(RS, Sql)
    
    
    
    If PasarPedidoAAlbaran(NumAlb, FechaAlb) Then
        MsgBox "Se ha generado correctamente el Albaran: " & NumAlb, vbInformation
                
'        FechaAlb = RecuperaValor(CadenaSQL, 3)
'        'Comprobar si Hay N� SERIE en compras y Mostrar
'        'ventana para pedir los N� Serie de la cantidad introducida si lo requiere algun articulo
'        ComprobarNSeriesLineas NumAlb, FechaAlb

        PonerModo 2
        
        
        'comprobar si hay lineas de art�culos cuyo precio_ultima_compra
        'se ha modificado y preguntar si que quieren actualizar los precio_venta
        '--------------------------------------------------------
        If b Then
            While Not RS.EOF
                Sql = "Se ha modificado el precio �ltima compra del art�culo:" & vbCrLf
                Sql = Sql & RS!codArtic & ":  " & RS!NomArtic & vbCrLf
                Sql = Sql & vbCrLf & "�Desea actualizar los precios de venta?"
                If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                    'Comprobar que el art�culo tiene margen comercial
                    If ArticuloTieneMargen(RS!codArtic) Then
                        'Aplicar margen comercial a los precios
                        'Modificar precios de venta en articulo y tarifas
                        frmComActPrecios.parCodArtic = RS!codArtic
                        frmComActPrecios.parNomArtic = RS!NomArtic
                        frmComActPrecios.Show vbModal
                    End If
                End If
                RS.MoveNext
            Wend
            RS.Close
            Set RS = Nothing
        End If
        
        
        
        
        If AlbCompleto Then
            'Se habra eliminado el pedido de (scaped, sliped)
            PosicionarDataTrasEliminar
        Else
            Sql = DevuelveDesdeBDNew(conAri, "scappr", "numpedpr", "numpedpr", Text1(0).Text, "N")
            If Sql = "" Then 'Ya no existe le pedido lo hemos eliminado
                PosicionarDataTrasEliminar
            Else
                PosicionarData
                PonerCampos
                CargaGrid DataGrid1, Data2, True, False
            End If
            CargaTxtAuxServidas False, False
        
            'Eliminar las filas del pedido que se servieron completas (slippr)
'            SQL = "DELETE FROM " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND cantidad=0"
'            Conn.Execute SQL
            
            'Comprobar que si no quedan lineas en el pedido se elimine la cabecera del pedido
'            SQL = "select codalmac, codartic FROM " & NomTablaLineas & " WHERE numpedpr=" & numPed
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            If RS.EOF Then 'No hay lineas de pedido --> Eliminar la cabecera
'                SQL = "DELETE FROM " & NombreTabla & " WHERE numpedcl=" & numPed
'                Conn.Execute SQL
'                PosicionarDataTrasEliminar
'            Else 'Quedan lineas en el pedido --> Actualizar las lineas
'                PosicionarData
'                PonerCampos
'                CargaGrid DataGrid1, Data2, True, False
'            End If
'            RS.Close
'            Set RS = Nothing
'            CargaTxtAuxServidas False, False
        End If
        Screen.MousePointer = vbDefault
        
'        Imprimer albaran si se solicit�
'        If ImprimeAlb Then
'            ImprimirAlbaran 45, NumAlb
'        End If
    Else 'Si no se ha pasado el Pedido a Albaran
        
    End If
End Sub


Private Sub InicializarServidas()
'Pone el campo servidas a 0 en la tabla lineas de pedido (sliped)
Dim Sql As String
    On Error Resume Next
    Sql = "UPDATE " & NomTablaLineas & " SET recibida= 0 "
    Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    conn.Execute Sql
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ComprobarNSeriesLineas(NumAlb As String, FechaAlb As String)
'Al pasar de PEDIDO a ALBARAN
'control de N� Series si hay algun articulo en las lineas de pedido que requiere N� de serie
'y hay control de N� de serie en compras pedirlos
Dim Sql As String
Dim RSLineas As ADODB.Recordset
Dim cadWHERE As String
        
    If vParamAplic.NumSeries Then 'So control de N� Series en COMPRAS
        cadWHERE = " WHERE numalbar=" & DBSet(NumAlb, "T") & " AND "
        cadWHERE = cadWHERE & " fechaalb=" & DBSet(FechaAlb, "F") & " AND "
        cadWHERE = cadWHERE & " slialp.codprove=" & Text1(4).Text
        
        'Seleccionamos aquellas lineas de albaran que tienen N� de Serie
        Sql = "SELECT slialp.codartic, sum(cantidad) as cantidad, slialp.numlinea "
        Sql = Sql & " FROM slialp INNER JOIN sartic on slialp.codartic=sartic.codartic "
        Sql = Sql & cadWHERE & " And nseriesn = 1 "
        Sql = Sql & " GROUP BY codartic ORDER BY Codartic "
    
        Set RSLineas = New ADODB.Recordset
        RSLineas.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RSLineas.EOF Then
            'Comprobar si NO Hay N� SERIE en Compras y si no se realizo alli
            'Mostrar ahora ventana para pedir los N� Serie de la cantidad introducida
            Me.cmdAux(1).Tag = NumAlb
            Me.cmdAux(0).Tag = FechaAlb
            PedirNSeries RSLineas
        End If
        RSLineas.Close
        Set RSLineas = Nothing
    End If
End Sub


Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
On Error GoTo EPedirNSeries
        
        'Visualizar en pantalla el Grid, y rellenar los N� Serie
        PedirNSeriesGnral RS, True

        Set frmNSerie = New frmRepCargarNSerie
        frmNSerie.DeVentas = False 'Se llama desde Alb. de Venta
        frmNSerie.Show vbModal
        Set frmNSerie = Nothing
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Function InsertarNSerie(numSerie As String, codArtic As String, numlinea As String) As Boolean
'Inserta o Actualiza en la tabla sserie, si al pasar Pedido -> Albaran
'existen lineas con control de N� Serie
'Dim CadValues As String, cadValuesU As String
Dim devuelve As String
Dim NumAlbar As String
Dim nSerie As CNumSerie
Dim b As Boolean

    On Error GoTo EInsertarNS

    Set nSerie = New CNumSerie
    
    '[Monica]24/10/2012: si es de recompra pasamos los datos de venta al hco y los limpiamos
    '                    para ello primero leo lo que hay
    devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", numSerie, "T", NumAlbar, "codartic", codArtic, "T")
    If devuelve <> "" Then
        nSerie.LeerDatos numSerie, codArtic
    End If
    ' hasta aqui
    
    
    nSerie.numSerie = numSerie
    nSerie.Articulo = codArtic
    nSerie.Proveedor = CInt(Text1(4).Text)
    nSerie.NumAlbProve = Me.cmdAux(1).Tag
    nSerie.fechaCom = Me.cmdAux(0).Tag
    nSerie.NumLinAlbPr = numlinea
    'calculamos la fecha de fin garantia para el articulo comprado
    nSerie.ObtenFechaFinGarantia codArtic, Me.cmdAux(0).Tag
    
    'Comprobar si existe en la tabla sserie
    NumAlbar = "numalbpr" 'N� albaran de Compra
    devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", numSerie, "T", NumAlbar, "codartic", codArtic, "T")
    If devuelve <> "" Then 'EXISTE en tabla sserie
        If NumAlbar = "" Then
            b = nSerie.ActualizarNumSerie(False)
        Else
            b = True
        End If
        
        '[Monica]24/10/2012: si es de recompra pasamos los datos de venta al hco y los limpiamos
        If b And nSerie.Cliente <> 0 Then b = nSerie.RecuperarParaVenta(Text1(1).Text, True)
        
    Else
        b = nSerie.InsertarNumSerie
    End If
    Set nSerie = Nothing
    
EInsertarNS:
    If Err.Number <> 0 Then b = False
    If Not b Then
        InsertarNSerie = False
    Else
        InsertarNSerie = True
    End If
End Function



Private Sub PonerDatosProveedor(codProve As String, Optional nifProve As String)
'lee de la tabla de proveedores y pone los valores
Dim vProve As CProveedor
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If codProve = "" Then
        LimpiarDatosProve
        Exit Sub
    End If

    Set vProve = New CProveedor
    'si se ha modificado el proveedor volver a cargar los datos
    If vProve.Existe(codProve) Then
        If vProve.LeerDatos(codProve) Then
            'NUEVO. Situacion proveedor
            If vProve.ProveedorBloqueado Then
                LimpiarDatosProve
                Set vProve = Nothing
                PonerFoco Text1(4)
                Exit Sub
            End If
            EsDeVarios = vProve.DeVarios
            BloquearDatosProve (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el proveedor no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!codProve) Then
                    Set vProve = Nothing
                    Exit Sub
                End If
            End If
        
            Text1(4).Text = vProve.Codigo
            FormateaCampo Text1(4)
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vProve.Nombre  'Nom prove
                Text1(8).Text = vProve.Domicilio
                Text1(9).Text = vProve.CPostal
                Text1(10).Text = vProve.Poblacion
                Text1(11).Text = vProve.Provincia
                Text1(6).Text = vProve.NIF
                Text1(7).Text = DBLet(vProve.TfnoAdmon, "T")
            End If
            
            If Modo = 3 Then 'insertar
                Text1(12).Text = vProve.ForPago
                Text2(12).Text = PonerNombreDeCod(Text1(12), conAri, "sforpa", "nomforpa")
                Text1(13).Text = Format(vProve.DtoPPago, FormatoDescuento)
                Text1(14).Text = Format(vProve.DtoGnral, FormatoDescuento)
            End If

            Observaciones = DBLet(vProve.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del proveedor"
            End If
        End If
    Else
        LimpiarDatosProve
        PonerFoco Text1(4)
    End If
    Set vProve = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
End Sub


Private Sub PonerDatosProveVario(nifProve As String)
'Poner el los campos Text el valor del cliente
Dim vProve As CProveedor
Dim b As Boolean
   
    If nifProve = "" Then Exit Sub
   
    Set vProve = New CProveedor
    b = vProve.LeerDatosProveVario(nifProve)
    
    If b Then
        Text1(5).Text = vProve.Nombre   'Nom proveedor
        Text1(8).Text = vProve.Domicilio
        Text1(9).Text = vProve.CPostal
        Text1(10).Text = vProve.Poblacion
        Text1(11).Text = vProve.Provincia
        Text1(7).Text = DBLet(vProve.TfnoAdmon, "T")
    End If
    Set vProve = Nothing
End Sub


Private Sub BloquearDatosProve(bol As Boolean)
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea proveedor de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(6).visible = bol 'NIF
        Me.imgBuscar(6).Enabled = bol 'NIF
        Me.imgBuscar(1).Enabled = bol 'poblacion
        
        For i = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(i), Not bol
        Next i
    End If
End Sub


Private Function ActualizarProveVarios(Prove As String, NIF As String) As Boolean
Dim vProve As CProveedor

    On Error GoTo EActualizarCV

    ActualizarProveVarios = False
    
    Set vProve = New CProveedor
    If EsProveedorVarios(Prove) Then
        vProve.NIF = NIF
        vProve.Nombre = Text1(5).Text
        vProve.Domicilio = Text1(8).Text
        vProve.CPostal = Text1(9).Text
        vProve.Poblacion = Text1(10).Text
        vProve.Provincia = Text1(11).Text
        vProve.TfnoAdmon = Text1(7).Text
        'Actualiza la tabla de proveedores varios con los datos que tenemos
        vProve.ActualizarProveV (NIF)
    End If
    Set vProve = Nothing
    
    ActualizarProveVarios = True
    
EActualizarCV:
    If Err.Number <> 0 Then
        ActualizarProveVarios = False
    Else
        ActualizarProveVarios = True
    End If
End Function


Private Sub CalcularDatosFactura()
Dim i As Byte
Dim cadWHERE As String
Dim vFactu As CFacturaCom

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 33 To 50
         Text3(i).Text = ""
    Next i
    
    cadWHERE = ObtenerWhereCP(False)
    
    Set vFactu = New CFacturaCom
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(13).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(14).Text))
    If vFactu.CalcularDatosFactura(cadWHERE, NombreTabla, NomTablaLineas) Then
        Text3(33).Text = vFactu.BrutoFac
        Text3(34).Text = vFactu.ImpPPago
        Text3(35).Text = vFactu.ImpGnral
        Text3(36).Text = vFactu.BaseImp
        Text3(37).Text = QuitarCero(vFactu.TipoIVA1)
        Text3(38).Text = QuitarCero(vFactu.TipoIVA2)
        Text3(39).Text = QuitarCero(vFactu.TipoIVA3)
        Text3(40).Text = vFactu.PorceIVA1
        Text3(41).Text = vFactu.PorceIVA2
        Text3(42).Text = vFactu.PorceIVA3
        Text3(43).Text = vFactu.BaseIVA1
        Text3(44).Text = vFactu.BaseIVA2
        Text3(45).Text = vFactu.BaseIVA3
        Text3(46).Text = vFactu.ImpIVA1
        Text3(47).Text = vFactu.ImpIVA2
        Text3(48).Text = vFactu.ImpIVA3
        Text3(49).Text = vFactu.TotalFac
        Text3(50).Text = vFactu.BaseImp
       
        FormatoDatosTotales
        
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
End Sub



Private Sub FormatoDatosTotales()
Dim i As Byte

    For i = 33 To 36
        If i = 34 Or i = 35 Then Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
    
    'Desglose B.Imponible por IVA
    For i = 43 To 45
        If Text3(i).Text <> "" Then
             If CSng(Text3(i).Text) = 0 And Text3(i - 6).Text = "" Then
                Text3(i).Text = QuitarCero(Text3(i).Text)
                Text3(i - 3).Text = QuitarCero(Text3(i - 3).Text)
                Text3(i - 6).Text = QuitarCero(Text3(i - 6).Text)
                Text3(i + 3).Text = QuitarCero(Text3(i + 3).Text)
            Else
                Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
                Text3(i - 3) = Format(Text3(i - 3).Text, FormatoDescuento)
    '            Text3(i - 6) = Format(Text3(i - 6).Text, "000")
                Text3(i + 3).Text = Format(Text3(i + 3).Text, FormatoImporte)
            End If
        End If
    Next i
    
    'Los Totales
    For i = 49 To 50
'        Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
End Sub




Private Function ActualizarUltFechaCom(cadW As String) As Boolean
''Actualiza la ultima fecha de compra y el ult. precio de compra
''en el articulo, poniendo los valores del albaran de compra
'Dim SQL As String
'Dim RS As ADODB.Recordset
'
'    On Error GoTo EActualizaFecha
'
'    SQL = "select distinct numalbar,fechaalb,slialp.codartic,max(slialp.precioar) as precioar , sartic.ultfecco "
'    SQL = SQL & " from slialp INNER JOIN sartic ON slialp.codartic=sartic.codartic "
''    SQL = SQL & " where numalbar='K2500088' and fechaalb='2005-10-06' and slialp.codprove=21"
'    SQL = SQL & " WHERE " & cadW
'    SQL = SQL & " and (fechaalb>ultfecco or isnull(ultfecco))"
'    SQL = SQL & " group by numalbar,fechaalb,slialp.codartic "
'    SQL = SQL & " order by codartic "
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not RS.EOF
'        SQL = "UPDATE sartic SET ultfecco=" & DBSet(RS!FechaAlb, "F") & ", preciouc=" & DBSet(RS!precioar, "N")
'        SQL = SQL & " WHERE codartic=" & DBSet(RS!codArtic, "T")
'        Conn.Execute SQL
'        RS.MoveNext
'    Wend
'    RS.Close
'    Set RS = Nothing
'
'EActualizaFecha:
'    If Err.Number <> 0 Then
'        ActualizarUltFechaCom = False
'    Else
'        ActualizarUltFechaCom = True
'    End If
End Function



Private Function ObtenerRSprecios(ByRef RS As ADODB.Recordset, cadSQL As String) As Boolean
    On Error GoTo ErrRS
    Set RS = New ADODB.Recordset
    RS.Open cadSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ObtenerRSprecios = True
    Exit Function
    
ErrRS:
    ObtenerRSprecios = False
    If Not RS Is Nothing Then Set RS = Nothing
    MuestraError Err.Number, "Cargando RS precios ultima compra.", Err.Description
End Function




Private Sub AbrirForm_CentroCoste()
    Screen.MousePointer = vbHourglass
    cmdAux(2).Tag = "2"

    Set frmB = New frmBuscaGrid
    frmB.vCampos = "Codigo|cabccost|codccost|T||20�Descripci�n|cabccost|nomccost|T||70�"
    frmB.vTabla = "cabccost"
    frmB.vSQL = ""
    HaDevueltoDatos = False
    '###A mano
    frmB.vDevuelve = "0|1|"
    frmB.vTitulo = "Centros de coste"
    frmB.vselElem = 0
    frmB.vConexionGrid = conConta
    
    frmB.Show vbModal
    Set frmB = Nothing
    cmdAux(2).Tag = "-1"
End Sub



' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
Private Sub AbrirForm_Articulos()
    If Trim(txtAux(1).Text) = "" Then Exit Sub
    
    Set frmArt = New frmAlmArticulos
    frmArt.DeConsulta = True
    frmArt.DatosADevolverBusqueda2 = "::" & Trim(txtAux(1).Text)  'DevNombreSQL(Data2.Recordset!codarti1)
    frmArt.parNumTAb = 6
    frmArt.Show vbModal
    Set frmArt = Nothing
End Sub
' -----


'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, Index As Integer)

    If InsertandoCabecera Then
        If imgBuscar(Index).visible Then imgBuscar_Click Index
        
    Else
        'Lineas
        If Index = 8 Then Index = 2
        cmdAux_Click Index
        
        
    End If
        
End Sub


