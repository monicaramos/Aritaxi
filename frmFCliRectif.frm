VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFCliRectif 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes Rectificativos Clientes"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   12810
   Icon            =   "frmFCliRectif.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   170
      Top             =   60
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   171
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
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
      Left            =   3720
      TabIndex        =   168
      Top             =   60
      Width           =   1845
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   150
         TabIndex        =   169
         Top             =   180
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Factura "
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Marcar Facturar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Imprimir Portes"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5610
      TabIndex        =   166
      Top             =   60
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   167
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
      Left            =   8250
      TabIndex        =   165
      Top             =   150
      Width           =   1605
   End
   Begin VB.TextBox Text3 
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
      Index           =   56
      Left            =   10860
      MaxLength       =   15
      TabIndex        =   164
      Text            =   "Text1 7"
      Top             =   480
      Width           =   1530
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   0
      Left            =   9300
      MaxLength       =   15
      TabIndex        =   163
      Text            =   "BASE IMP."
      Top             =   495
      Width           =   1490
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
      Index           =   9
      Left            =   3840
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   152
      Text            =   "nom ccoste"
      Top             =   7545
      Visible         =   0   'False
      Width           =   5805
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
      Left            =   2370
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   48
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   7140
      Visible         =   0   'False
      Width           =   7275
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   6915
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
      Left            =   11400
      TabIndex        =   50
      Top             =   7020
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
      Left            =   10080
      TabIndex        =   49
      Top             =   7020
      Width           =   1135
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   240
      Top             =   6360
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   960
      Top             =   6360
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
      Height          =   5070
      Left            =   120
      TabIndex        =   34
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1815
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   8943
      _Version        =   393216
      Style           =   1
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
      TabPicture(0)   =   "frmFCliRectif.frx":000C
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
      Tab(0).Control(5)=   "txtAux(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAux(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAux(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FrameCliente"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAux(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtAux(9)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdAux(9)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAux(10)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtAux(11)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "FrameToolAux"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmFCliRectif.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameFacRec"
      Tab(1).Control(1)=   "Text1(39)"
      Tab(1).Control(2)=   "FrameHco"
      Tab(1).Control(3)=   "Text2(28)"
      Tab(1).Control(4)=   "Text1(27)"
      Tab(1).Control(5)=   "Text2(27)"
      Tab(1).Control(6)=   "Text1(2)"
      Tab(1).Control(7)=   "Text1(25)"
      Tab(1).Control(8)=   "Text1(26)"
      Tab(1).Control(9)=   "Text1(24)"
      Tab(1).Control(10)=   "Text1(23)"
      Tab(1).Control(11)=   "Text1(38)"
      Tab(1).Control(12)=   "Text1(18)"
      Tab(1).Control(13)=   "Text1(19)"
      Tab(1).Control(14)=   "Text1(20)"
      Tab(1).Control(15)=   "Text1(21)"
      Tab(1).Control(16)=   "Text1(22)"
      Tab(1).Control(17)=   "Text1(28)"
      Tab(1).Control(18)=   "Frame3"
      Tab(1).Control(19)=   "Label1(45)"
      Tab(1).Control(20)=   "Label1(23)"
      Tab(1).Control(21)=   "imgBuscar(8)"
      Tab(1).Control(22)=   "Label1(9)"
      Tab(1).Control(23)=   "imgBuscar(7)"
      Tab(1).Control(24)=   "Label1(12)"
      Tab(1).Control(25)=   "Label1(11)"
      Tab(1).Control(26)=   "Label1(10)"
      Tab(1).Control(27)=   "Label1(5)"
      Tab(1).Control(28)=   "Label1(3)"
      Tab(1).ControlCount=   29
      TabCaption(2)   =   "Totales"
      TabPicture(2)   =   "frmFCliRectif.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameFactura"
      Tab(2).ControlCount=   1
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   180
         TabIndex        =   172
         Top             =   2250
         Width           =   1500
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   173
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
      Begin VB.Frame FrameFacRec 
         Caption         =   "Datos Factura a rectificar "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -66750
         TabIndex        =   133
         Top             =   420
         Width           =   3615
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
            Index           =   37
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   138
            Tag             =   "Tipo Mov. Factura|T|S|||scaalbcli|codtipmf||N|"
            Top             =   330
            Width           =   1785
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
            Index           =   36
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   136
            Tag             =   "Nº. Factura|N|S|0||scaalbcli|numfactu|0000000|N|"
            Top             =   750
            Width           =   1785
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
            Index           =   35
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   134
            Tag             =   "Fecha Factura|F|S|||scaalbcli|fecfactu|dd/mm/yyyy|N|"
            Top             =   1170
            Width           =   1785
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Mov."
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
            Left            =   120
            TabIndex        =   139
            Top             =   360
            Width           =   1065
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
            Index           =   46
            Left            =   120
            TabIndex        =   137
            Top             =   780
            Width           =   1185
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
            Index           =   44
            Left            =   120
            TabIndex        =   135
            Top             =   1200
            Width           =   1245
         End
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   11
         Left            =   12600
         MaxLength       =   15
         TabIndex        =   154
         Text            =   "numlote"
         Top             =   3960
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
         Index           =   10
         Left            =   7440
         MaxLength       =   5
         TabIndex        =   41
         Tag             =   "Bultos"
         Text            =   "12345"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   9
         Left            =   12360
         TabIndex        =   150
         ToolTipText     =   "Buscar proveedor"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
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
         Left            =   11640
         MaxLength       =   6
         TabIndex        =   47
         Tag             =   "proveedor"
         Text            =   "codc"
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   39
         Left            =   -71160
         MaxLength       =   7
         TabIndex        =   141
         Tag             =   "Nº Venta|N|S|||scaalbcli|numventa|0000000|N|"
         Text            =   "Text1 7"
         Top             =   3000
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Frame FrameHco 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   122
         Top             =   2310
         Width           =   4455
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
            Index           =   33
            Left            =   795
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   130
            Text            =   "Text2"
            Top             =   1560
            Width           =   3525
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
            Index           =   33
            Left            =   135
            MaxLength       =   30
            TabIndex        =   129
            Text            =   "Text1"
            Top             =   1560
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
            Height          =   315
            Index           =   32
            Left            =   795
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   127
            Text            =   "Text2"
            Top             =   840
            Width           =   3525
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
            Index           =   32
            Left            =   135
            MaxLength       =   30
            TabIndex        =   126
            Text            =   "Text1"
            Top             =   840
            Width           =   660
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
            Index           =   31
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   124
            Top             =   240
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   1080
            ToolTipText     =   "Buscar incidencia"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
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
            Left            =   120
            TabIndex        =   131
            Top             =   1320
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1080
            ToolTipText     =   "Buscar trabajador"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
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
            Index           =   38
            Left            =   120
            TabIndex        =   128
            Top             =   615
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Eliminación"
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
            Index           =   37
            Left            =   120
            TabIndex        =   125
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Height          =   255
            Index           =   29
            Left            =   360
            TabIndex        =   123
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame FrameFactura 
         Height          =   3300
         Left            =   -74040
         TabIndex        =   91
         Top             =   600
         Width           =   10575
         Begin VB.TextBox Text3 
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
            Index           =   54
            Left            =   6840
            MaxLength       =   15
            TabIndex        =   147
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1485
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
            Index           =   51
            Left            =   6240
            MaxLength       =   5
            TabIndex        =   146
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   555
         End
         Begin VB.TextBox Text3 
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
            Index           =   53
            Left            =   6840
            MaxLength       =   15
            TabIndex        =   145
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1485
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
            Index           =   50
            Left            =   6240
            MaxLength       =   5
            TabIndex        =   144
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   555
         End
         Begin VB.TextBox Text3 
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
            Index           =   52
            Left            =   6840
            MaxLength       =   15
            TabIndex        =   143
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
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
            Index           =   49
            Left            =   6240
            MaxLength       =   5
            TabIndex        =   142
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   555
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
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   108
            Text            =   "Text1 7"
            Top             =   615
            Width           =   1485
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
            Index           =   34
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   107
            Text            =   "Text1 7"
            Top             =   615
            Width           =   1365
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
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   106
            Text            =   "Text1 7"
            Top             =   615
            Width           =   1365
         End
         Begin VB.TextBox Text3 
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
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   105
            Text            =   "Text1 7"
            Top             =   615
            Width           =   1485
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
            Index           =   43
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   104
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1365
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
            Index           =   37
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   103
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   675
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
            Index           =   40
            Left            =   4200
            MaxLength       =   5
            TabIndex        =   102
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   675
         End
         Begin VB.TextBox Text3 
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
            Index           =   46
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   101
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1275
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
            Index           =   44
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   100
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1365
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
            Index           =   38
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   99
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   675
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
            Index           =   41
            Left            =   4200
            MaxLength       =   5
            TabIndex        =   98
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   675
         End
         Begin VB.TextBox Text3 
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
            Index           =   47
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   97
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1275
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
            Index           =   45
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   96
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1365
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
            Index           =   39
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   95
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   675
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
            Index           =   42
            Left            =   4200
            MaxLength       =   5
            TabIndex        =   94
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   675
         End
         Begin VB.TextBox Text3 
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
            Index           =   48
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   93
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1275
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   55
            Left            =   6240
            MaxLength       =   15
            TabIndex        =   92
            Text            =   "Text1 7"
            Top             =   2640
            Width           =   2085
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
            Index           =   48
            Left            =   6240
            TabIndex        =   149
            Top             =   1200
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. RE"
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
            Left            =   6960
            TabIndex        =   148
            Top             =   1200
            Width           =   975
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
            Index           =   28
            Left            =   2760
            TabIndex        =   121
            Top             =   1200
            Width           =   1485
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
            Index           =   27
            Left            =   240
            TabIndex        =   120
            Top             =   360
            Width           =   1305
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
            Index           =   22
            Left            =   2160
            TabIndex        =   119
            Top             =   360
            Width           =   1635
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
            Index           =   18
            Left            =   3960
            TabIndex        =   118
            Top             =   360
            Width           =   1425
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
            Index           =   2
            Left            =   5760
            TabIndex        =   117
            Top             =   360
            Width           =   1755
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
            TabIndex        =   116
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
            TabIndex        =   115
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
            TabIndex        =   114
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
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
            Left            =   5040
            TabIndex        =   113
            Top             =   1200
            Width           =   735
         End
         Begin VB.Line Line1 
            X1              =   2040
            X2              =   8040
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
            TabIndex        =   112
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL ALBARAN"
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
            Height          =   285
            Index           =   39
            Left            =   4200
            TabIndex        =   111
            Top             =   2655
            Width           =   1860
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
            Left            =   4320
            TabIndex        =   110
            Top             =   1200
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Cod."
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
            Index           =   42
            Left            =   2040
            TabIndex        =   109
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   28
         Left            =   -72600
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   82
         Text            =   "Text2"
         Top             =   3960
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   27
         Left            =   -73320
         MaxLength       =   30
         TabIndex        =   24
         Tag             =   "Trabajador pedido|N|S|0|9999|scaalbcli|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   3480
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   27
         Left            =   -72600
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   80
         Text            =   "Text2"
         Top             =   3480
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -72240
         MaxLength       =   10
         TabIndex        =   78
         Tag             =   "Semana Entrega|N|S|||scaalbcli|sementre||N|"
         Top             =   2880
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   25
         Left            =   -74880
         MaxLength       =   7
         TabIndex        =   75
         Tag             =   "Nº Pedido|N|S|||scaalbcli|numpedcl|0000000|N|"
         Text            =   "Text1 7"
         Top             =   3000
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   26
         Left            =   -73680
         MaxLength       =   10
         TabIndex        =   74
         Tag             =   "Fecha Pedido|F|S|||scaalbcli|fecpedcl|dd/mm/yyyy|N|"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   24
         Left            =   -70200
         MaxLength       =   10
         TabIndex        =   70
         Tag             =   "Fecha Oferta|F|S|||scaalbcli|fecofert|dd/mm/yyyy|N|"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   23
         Left            =   -71160
         MaxLength       =   7
         TabIndex        =   69
         Tag             =   "Nº Oferta|N|S|||scaalbcli|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   2880
         Visible         =   0   'False
         Width           =   885
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
         Left            =   8880
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   43
         Tag             =   "Descuento 1"
         Text            =   "OF"
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame FrameCliente 
         Height          =   1950
         Left            =   200
         TabIndex        =   54
         Top             =   315
         Width           =   11820
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
            Index           =   29
            Left            =   7620
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   162
            Text            =   "Text2"
            Top             =   165
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
            Index           =   29
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   161
            Tag             =   "Cod. Envío|N|N|0|999|scaalbcli|codenvio|000|N|"
            Text            =   "Text1"
            Top             =   165
            Width           =   690
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
            Left            =   7800
            MaxLength       =   7
            TabIndex        =   20
            Tag             =   "Descuento General|N|S|||scaalbcli|aportacion|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1020
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   34
            Left            =   11460
            MaxLength       =   30
            TabIndex        =   21
            Tag             =   "Cant. Km|N|S|0|99999|scaalbcli|cantidkm||N|"
            Text            =   "Text1"
            Top             =   480
            Visible         =   0   'False
            Width           =   950
         End
         Begin VB.CheckBox chkFacturarKm 
            Caption         =   "Facturar Km"
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
            Left            =   3990
            TabIndex        =   23
            Tag             =   "Facturar Km|N|N|||scaalbcli|facturkm||N|"
            Top             =   1320
            Visible         =   0   'False
            Width           =   1575
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
            Height          =   315
            Index           =   13
            Left            =   1305
            MaxLength       =   20
            TabIndex        =   14
            Tag             =   "Referencia Cliente|T|S|||scaalbcli|referenc||N|"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   1560
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   66
            Tag             =   "Direccion/Dpto.|T|S|||scaped|nomdirec||N|"
            Text            =   "Text2"
            Top             =   165
            Visible         =   0   'False
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "Direccion/Dpto.|N|S|0|999|scaalbcli|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   165
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
            Height          =   315
            Index           =   11
            Left            =   1305
            MaxLength       =   30
            TabIndex        =   13
            Tag             =   "Provincia|T|N|||scaalbcli|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1209
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
            Left            =   1305
            MaxLength       =   6
            TabIndex        =   11
            Tag             =   "CPostal|T|N|||scaalbcli|codpobla||N|"
            Text            =   "Text15"
            Top             =   861
            Width           =   780
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
            Left            =   2085
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Población|T|N|||scaalbcli|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   861
            Width           =   3435
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
            TabIndex        =   9
            Tag             =   "teléfono Cliente|T|S|||scaalbcli|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   135
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
            Left            =   1305
            MaxLength       =   15
            TabIndex        =   8
            Tag             =   "NIF Cliente|T|N|||scaalbcli|nifclien||N|"
            Text            =   "123456789"
            Top             =   135
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
            Index           =   17
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   16
            Tag             =   "Cod. Agente|N|N|0|9999|scaalbcli|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   513
            Width           =   690
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
            Index           =   17
            Left            =   7620
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   61
            Text            =   "Text2"
            Top             =   513
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
            Index           =   14
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   17
            Tag             =   "Forma de Pago|N|N|0|999|scaalbcli|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   861
            Width           =   690
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
            Left            =   7620
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   56
            Text            =   "Text2"
            Top             =   861
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
            Left            =   6885
            MaxLength       =   7
            TabIndex        =   18
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaalbcli|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1209
            Width           =   690
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
            Left            =   6885
            MaxLength       =   7
            TabIndex        =   19
            Tag             =   "Descuento General|N|N|0|99.90|scaalbcli|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   690
         End
         Begin VB.ComboBox cboFacturacion 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9840
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Tag             =   "Tipo Facturación|N|N|||scaalbcli|tipofact||N|"
            Top             =   1500
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
            Index           =   8
            Left            =   1305
            MaxLength       =   35
            TabIndex        =   10
            Tag             =   "Domicilio|T|N|||scaalbcli|domclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   513
            Width           =   4215
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   6600
            ToolTipText     =   "Buscar forma de envio"
            Top             =   165
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cod.Envío"
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
            Index           =   24
            Left            =   5700
            TabIndex        =   160
            Top             =   180
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "APORTACION TERMINAL"
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
            Index           =   49
            Left            =   7560
            TabIndex        =   151
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Km a facturar"
            Height          =   255
            Index           =   43
            Left            =   11460
            TabIndex        =   132
            Top             =   255
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Ref.Cliente"
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
            TabIndex        =   73
            Top             =   1560
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   6
            Left            =   1065
            ToolTipText     =   "Buscar población"
            Top             =   870
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
            TabIndex        =   68
            Top             =   165
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   6600
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   165
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
            TabIndex        =   67
            Top             =   1215
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
            TabIndex        =   65
            Top             =   855
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
            Left            =   2565
            TabIndex        =   64
            Top             =   165
            Width           =   945
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
            TabIndex        =   63
            Top             =   165
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1065
            ToolTipText     =   "Buscar cliente varios"
            Top             =   165
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
            TabIndex        =   62
            Top             =   510
            Width           =   765
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   6600
            ToolTipText     =   "Buscar agente"
            Top             =   516
            Width           =   240
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
            TabIndex        =   60
            Top             =   855
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.Pago"
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
            Left            =   5700
            TabIndex        =   59
            Top             =   1215
            Width           =   1215
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
            Left            =   5715
            TabIndex        =   58
            Top             =   1575
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Facturación"
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
            Left            =   9840
            TabIndex        =   57
            Top             =   1200
            Width           =   1725
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6600
            ToolTipText     =   "Buscar forma de pago"
            Top             =   867
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
            TabIndex        =   55
            Top             =   510
            Width           =   885
         End
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   53
         ToolTipText     =   "Buscar artículo"
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
         TabIndex        =   52
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
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
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   39
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   3960
         Visible         =   0   'False
         Width           =   3165
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
         Left            =   10560
         MaxLength       =   12
         TabIndex        =   46
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
         Left            =   9960
         MaxLength       =   30
         TabIndex        =   45
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   3960
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
         Index           =   6
         Left            =   9360
         MaxLength       =   5
         TabIndex        =   44
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
         Left            =   8040
         MaxLength       =   12
         TabIndex        =   42
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   3960
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
         Index           =   3
         Left            =   6120
         MaxLength       =   16
         TabIndex        =   40
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
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   38
         Tag             =   "Código Artículo"
         Text            =   "Artic Artic Artic5"
         Top             =   3900
         Visible         =   0   'False
         Width           =   1455
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
         Index           =   0
         Left            =   360
         MaxLength       =   15
         TabIndex        =   37
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3900
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFCliRectif.frx":0060
         Height          =   2040
         Left            =   195
         TabIndex        =   51
         Top             =   2895
         Width           =   12020
         _ExtentX        =   21193
         _ExtentY        =   3598
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
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   38
         Left            =   -70200
         MaxLength       =   10
         TabIndex        =   140
         Tag             =   "Nº terminal|N|S|||scaalbcli|numtermi||N|"
         Top             =   2880
         Visible         =   0   'False
         Width           =   705
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
         Index           =   18
         Left            =   -74880
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "Observación 1|T|S|||scaalbcli|observa01||N|"
         Top             =   765
         Width           =   7845
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
         Index           =   19
         Left            =   -74880
         MaxLength       =   80
         TabIndex        =   27
         Tag             =   "Observación 2|T|S|||scaalbcli|observa02||N|"
         Top             =   1125
         Width           =   7845
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
         Index           =   20
         Left            =   -74880
         MaxLength       =   80
         TabIndex        =   28
         Tag             =   "Observación 3|T|S|||scaalbcli|observa03||N|"
         Top             =   1485
         Width           =   7845
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
         Left            =   -74880
         MaxLength       =   80
         TabIndex        =   29
         Tag             =   "Observación 4|T|S|||scaalbcli|observa04||N|"
         Top             =   1845
         Width           =   7845
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
         Index           =   22
         Left            =   -74880
         MaxLength       =   80
         TabIndex        =   30
         Tag             =   "Observación 5|T|S|||scaalbcli|observa05||N|"
         Top             =   2205
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   28
         Left            =   -73230
         MaxLength       =   30
         TabIndex        =   25
         Tag             =   "Preparador Material|N|N|0|9999|scaalbcli|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   1950
         Width           =   660
      End
      Begin VB.Frame Frame3 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   156
         Top             =   840
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CheckBox chkDocArchi 
            Caption         =   "Documento archivado"
            Height          =   375
            Left            =   3120
            TabIndex        =   159
            Tag             =   "Docar|N|N|||scaalbcli|docarchiv||N|"
            Top             =   1200
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   41
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   158
            Tag             =   "Fecha envio|F|S|||scaalbcli|fecenvio|dd/mm/yyyy|N|"
            Top             =   1200
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   40
            Left            =   1320
            ToolTipText     =   "Buscar fecha"
            Top             =   1200
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha envio"
            Height          =   195
            Index           =   52
            Left            =   360
            TabIndex        =   157
            Top             =   1200
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   45
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Preparador Material"
         Height          =   255
         Index           =   23
         Left            =   -74880
         TabIndex        =   83
         Top             =   1800
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -73455
         ToolTipText     =   "Buscar trabajador"
         Top             =   1800
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Pedido"
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   81
         Top             =   1335
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   -73455
         ToolTipText     =   "Buscar trabajador"
         Top             =   1335
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sem. Entrega"
         Height          =   255
         Index           =   12
         Left            =   -72645
         TabIndex        =   79
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   11
         Left            =   -74880
         TabIndex        =   77
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         Height          =   255
         Index           =   10
         Left            =   -73905
         TabIndex        =   76
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
         Height          =   255
         Index           =   5
         Left            =   -70200
         TabIndex        =   72
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
         Height          =   255
         Index           =   3
         Left            =   -71160
         TabIndex        =   71
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
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
      Left            =   11400
      TabIndex        =   31
      Top             =   7020
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   84
      Top             =   780
      Width           =   12435
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
         Left            =   6960
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Cod. Cliente|N|N|0|999999|scaalbcli|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   540
         Width           =   825
      End
      Begin VB.TextBox Text1 
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
         Index           =   5
         Left            =   7780
         MaxLength       =   40
         TabIndex        =   7
         Tag             =   "Nombre Cliente|T|N|||scaalbcli|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   555
         Width           =   4500
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
         Left            =   6960
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Realizada Por|N|N|0|9999|scaalbcli|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   180
         Width           =   825
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
         Left            =   7800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   89
         Text            =   "Text2"
         Top             =   180
         Width           =   4500
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
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Tipo Albaran|T|N|||scaalbcli|codtipom||S|"
         Text            =   "Text3"
         Top             =   435
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
         Index           =   1
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Albaran|F|N|||scaalbcli|fechaalb|dd/mm/yyyy|N|"
         Top             =   420
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
         Tag             =   "Nº Albaran|N|S|0||scaalbcli|numalbar|0000000|S|"
         Text            =   "Text1 7"
         Top             =   435
         Width           =   1035
      End
      Begin VB.CheckBox chkFacturar 
         Caption         =   "Facturar"
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
         Left            =   3990
         TabIndex        =   4
         Tag             =   "Facturar|N|N|||scaalbcli|factursn||N|"
         Top             =   420
         Width           =   1215
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6660
         ToolTipText     =   "Buscar cliente"
         Top             =   570
         Width           =   240
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
         Left            =   5595
         TabIndex        =   90
         Top             =   555
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Realizada Por"
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
         Left            =   5595
         TabIndex        =   88
         Top             =   225
         Width           =   1050
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   6660
         ToolTipText     =   "Buscar trabajador"
         Top             =   210
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. Alb."
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
         Left            =   2640
         TabIndex        =   87
         Top             =   180
         Width           =   945
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3630
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Albaran"
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
         Index           =   50
         Left            =   240
         TabIndex        =   86
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
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
         Left            =   1365
         TabIndex        =   85
         Top             =   180
         Width           =   1005
      End
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
      Left            =   9720
      TabIndex        =   155
      Top             =   7500
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Centro coste"
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
      Index           =   51
      Left            =   2400
      TabIndex        =   153
      Top             =   7545
      Visible         =   0   'False
      Width           =   1335
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
      Top             =   6900
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Shortcut        =   ^I
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
Attribute VB_Name = "frmFCliRectif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran  de Venta de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALV,ALR,ALS)

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schalb, y solo en modo de consulta
                              
Public RecuperarFactu As Boolean 'si esta recuperando facturas al generar las facturas no coger contaror
                                 'pedirlas por teclado
                                 
Public AlbAvisoGenerado As Long 'Cuando desde aviso cierro reparacion, creo un albaran y llamo a este form
                                'Entonces lo cargo el albaran y lo meto insertando lineas
                                
'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmC As frmGesSocios 'Form Mto Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCV As frmGesSocios  'Form Mto Clientes Varios
Attribute frmCV.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmA As frmFacAgentesCom   'Form Mto Agentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmArt As frmAlmArticulos   'Form Articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmFE As frmFacFormasEnvio  'Form Formas de Envio
Attribute frmFE.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientes2
Attribute frmCli.VB_VarHelpID = -1



Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar nº Series
Attribute frmNSerie.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Form Mensajes
Attribute frmMen.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoOfer
Attribute frmList.VB_VarHelpID = -1

Private WithEvents frmProv As frmComProveedores
Attribute frmProv.VB_VarHelpID = -1

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


Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas


Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String 'Para el ORDER BY de la consulta
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Dim PorCaja As Boolean
'Para Saber si se ha salido con precio caja y hay que calcular el importe de la
'linea aplicando el precio de la caja. Si PorCaja=false se aplicaca el precio de unidad

Dim Precio As String 'Precio de la linea de Articulo

Dim cadList As String 'cadena para pasar al historico
Dim motivo As String 'cadena para el motivo si es factura Rectificativa


Dim PulsadoMas2 As Boolean

Dim txtAnterior As String

Dim ClienteConTasaReciclado As Boolean  'Cuando pasamos a las lineas pondremos esta variab


'PORTES
' Tipo fontenas
Dim KilosAnteriores As Currency
Dim RutaCliente As Integer
Dim ZonaCliente As Integer


Dim AlmacenLineas As Integer
'Dim ElArticulo As String

Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Check1_Click()

End Sub

Private Sub chkDocArchi_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkDocArchi_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkFacturar_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkFacturar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkFacturarKm_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2
Dim numlinea As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                InsertarCabecera
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificarCabAlbaran Then
                    TerminaBloquear
                    
                    UpdateaNomDirec
                    
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'slialbcli'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea(numlinea, False) Then
                    'Comprobar si el Articulo tiene control de Nº de Serie
                    ComprobarNSeriesLineas numlinea
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    BotonAnyadirLinea
                End If
                
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    numlinea = Data2.Recordset!numlinea
                    'Comprobar si el Articulo tiene control de Nº de Serie
                    ComprobarNSeriesLineas numlinea
                    TerminaBloquear
                    NumRegElim = Val(Data2.Recordset!numlinea)
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    PosicionarData2
                    ModificaLineas = 0
'                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                    BloquearTxt Text2(9), True
                End If
                PonerModo 2
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
                
                Me.DataGrid1.Enabled = True
            End If
            CalcularDatosFactura
    End Select
    Screen.MousePointer = vbDefault
    Text1(30).visible = True
'    Combo1.visible = False
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function ModificarCabAlbaran() As Boolean
Dim b As Boolean
Dim Sql As String

    On Error GoTo EModificaAlb
    conn.BeginTrans
    
    'Si es cliente de varios actualizar datos cliente en tabla:sclvar
'    b = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    b = True
    If b Then
        b = ModificaDesdeFormulario(Me, 1)
        If b Then
            Sql = "UPDATE scaalbcli SET nomdirec=" & DBSet(Text2(12).Text, "T") & " WHERE codtipom=" & DBSet(CodTipoMov, "T") & " and numalbar=" & Data1.Recordset!NumAlbar
            conn.Execute Sql
        End If

        If b Then
            'comprobar si se ha cambiado el cliente
            'o si se ha cambiado la fecha del albaran
            'If (CInt(Me.Data1.Recordset!CodClien) <> CInt(Text1(4).Text)) Or (CDate(Data1.Recordset!FechaAlb) <> CDate(Text1(1).Text)) Then
            'DAVID.   No es un CINT. Tiene que ser un clng o val
            If (Val(Me.Data1.Recordset!CodClien) <> Val(Text1(4).Text)) Or (CDate(Data1.Recordset!FechaAlb) <> CDate(Text1(1).Text)) Then
                'si hay numeros de serie en ese albaran, actualizamos el cliente
                'al nuevo cliente
                Sql = "UPDATE sserie SET codclien=" & DBSet(Text1(4).Text, "N") & ","
                Sql = Sql & " fechavta=" & DBSet(Text1(1).Text, "F")
                Sql = Sql & " WHERE codtipom='" & CodTipoMov & "'" & " AND numalbar=" & Data1.Recordset!NumAlbar & " and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
                conn.Execute Sql
                
                'Modificar el cliente en la smoval
                Sql = "UPDATE smoval SET codigope=" & DBSet(Text1(4).Text, "N") & ","
                Sql = Sql & " fechamov=" & DBSet(Text1(1).Text, "F")
                Sql = Sql & ", horamovi= concat(" & DBSet(Text1(1).Text, "F") & ",' ',hour(horamovi),':',minute(horamovi),':',second(horamovi))"
                Sql = Sql & " WHERE detamovi='" & CodTipoMov & "'"
                Sql = Sql & " AND document='" & Text1(0).Text & "'"
                Sql = Sql & " and fechamov=" & DBSet(Data1.Recordset!FechaAlb, "F")
                conn.Execute Sql
            End If
        End If
    End If
    
EModificaAlb:
    If Err.Number <> 0 Then b = False
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    ModificarCabAlbaran = b
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar cabecera Albaran.", Err.Description
End Function




Private Sub cmdAux_Click(Index As Integer)
Dim b As Boolean

    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            
        Case 1 'Busqueda de Cod. Artic
            b = True
            If CodTipoMov = "ART" Or CodTipoMov = "ARC" Or CodTipoMov = "ARP" Then
                If MsgBox("¿Desea traer líneas de la factura que va a rectificar?", vbQuestion + vbYesNo) = vbYes Then
                
                    'si es Albaran de Factura rectificativa cargar un listview con todas las
                    'lineas de la factura y marcar las que queremos seleccionar para
                    'cargarlas en las lineas del Albaran rectificativo
                    b = False
                    Set frmMen = New frmMensajes
                    frmMen.cadWHERE = " codtipom=" & DBSet(Text1(37).Text, "T") & " and numfactu=" & Text1(36).Text & " and fecfactu=" & DBSet(Text1(35).Text, "F")
                    frmMen.OpcionMensaje = 111 'Lineas Factura a Rectificar
                    frmMen.Show vbModal
                    Set frmMen = Nothing
                    CargaGrid Me.DataGrid1, Me.Data2, True
                    cmdCancelar_Click
                End If
            End If
            
            If b Then
                Set frmArt = New frmAlmArticulos
                frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en Modo busqueda
                frmArt.DeConsulta = True
                frmArt.Show vbModal
                Set frmArt = Nothing
'                txtAux_LostFocus (1)
            End If
            
    Case 9 'CENTRO COSTE/ PROVEEDOR
        If vEmpresa.TieneAnalitica Then
            'centro de coste
            AbrirForm_CentroCoste
        Else
            Set frmProv = New frmComProveedores
            frmProv.DatosADevolverBusqueda = "1"
            frmProv.Show vbModal
            Set frmProv = Nothing
        End If
    End Select
    PonerFoco txtAux(Index)
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
            BloquearTxt Text2(9), True
            DataGrid1.Columns(4).Caption = "Artículo"
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
'            PonerBotonCabecera True
            PonerModo 2
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Me.DataGrid1.Enabled = True
    
    End Select
    Text1(30).visible = True
'    Combo1.visible = False
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String
Dim Cad As String
Dim RS As ADODB.Recordset

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'ocultamos el campo de codtipom y mostramos el combo para que elija q
    'tipo de albaran quiere crear
'    Text1(30).visible = False
'    Combo1.visible = True
    
    'Si es Albaran para factura RECTIFICATIVA pedir la Factura que se va
    'a Rectificar y si existe en el historico, tabla "scafac", entonces dejamos
    'que inserte el Albaran Rectificativo, si no salimos
    If CodTipoMov = "ART" Or CodTipoMov = "ARP" Then
        cadList = ""

        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 225
        frmList.Show vbModal
        Set frmList = Nothing
        If cadList = "" Then Exit Sub
        
        'cargar los datos de la factura recuperada en el formulario
        NomTraba = "select codtipom as codtipmf,numfactu,fecfactu,codclien,nomclien,domclien,scafac.codpobla,pobclien,proclien,nifclien,telclien,"
        NomTraba = NomTraba & "coddirec,nomdirec,scafac.codagent,nomagent,scafac.codforpa, nomforpa,dtoppago,dtognral "
        NomTraba = NomTraba & " from (scafac inner join sforpa on scafac.codforpa=sforpa.codforpa) "
        NomTraba = NomTraba & " inner join sagent on scafac.codagent=sagent.codagent where " & cadList
        
        If CodTipoMov = "ARP" Then NomTraba = Replace(NomTraba, "scafac", "scafaccli")
        
        Set RS = New ADODB.Recordset
        RS.Open NomTraba, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        PonerModo 3
        
        If Not RS.EOF Then
            Text1(4).Text = RS!CodClien
            FormateaCampo Text1(4)
            Text1(5).Text = RS!nomclien
            Text1(6).Text = RS!nifClien
            Text1(7).Text = DBLet(RS!telclien, "T")
            Text1(8).Text = RS!domclien
            Text1(9).Text = RS!codpobla
            Text1(10).Text = RS!pobclien
            Text1(11).Text = DBLet(RS!proclien, "T")
            Text1(12).Text = DBLet(RS!CodDirec, "T")
            FormateaCampo Text1(12)
            Text2(12).Text = DBLet(RS!nomdirec, "T")
            Text1(14).Text = RS!codforpa
            FormateaCampo Text1(14)
            Text2(14).Text = RS!nomforpa
            Text1(15).Text = DBLet(RS!DtoPPago, "N")
            FormateaCampo Text1(15)
            Text1(16).Text = DBLet(RS!DtoGnral, "N")
            FormateaCampo Text1(16)
            Text1(17).Text = DBLet(RS!codagent, "T")
            FormateaCampo Text1(17)
            Text2(17).Text = RS!nomagent
            Text1(37).Text = RS!codtipmf
            Text1(36).Text = DBLet(RS!NumFactu, "N")
            FormateaCampo Text1(36)
            Text1(35).Text = RS!FecFactu
            
            'Observacion 1   'DAVID
            'Text1(18).Text = "RECTIFICA A FACTURA: " & RS!codtipmf & ", " & RS!NumFactu & ", " & RS!FecFactu
            
'            If Text1(30).Text = "ARP" Then
                Text1(18).Text = RS!NumFactu & ", " & RS!FecFactu
                'Observacion 2
                Text1(19).Text = motivo
'            Else
'
'                Dim SqlNew As String
'                Dim RSNew As ADODB.Recordset
'                Set RSNew = New ADODB.Recordset
'
'                SqlNew = "select observa1, observa2, observa3 from scafaccli1 where " & Replace(cadList, "scafac", "scafaccli1")
'                RSNew.Open SqlNew, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If Not RSNew.EOF Then
'                    Text1(18).Text = DBLet(RSNew!observa1)
'                    Text1(19).Text = DBLet(RSNew!observa2)
'                    Text1(20).Text = DBLet(RSNew!observa3)
'                End If
'
'
'                Text1(21).Text = RS!NumFactu & ", " & RS!FecFactu
'                Text1(22).Text = motivo
'
'            End If
            NomTraba = "tipofact"
            'cad = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(4).Text, "N", NomTraba)
            
            NomTraba = "0"
            Cad = "0"
            If Cad = "0" Then BloquearDatosCliente (False)
            
            'Memorizo cad con codtipom
            Cad = RS!codtipmf
            
            'recuperamos el tipo de facturacion del cliente
            Me.cboFacturacion.ListIndex = CInt(NomTraba)
        Else
            Cad = "N" 'para que la busqueda de despues no de error
        End If
        RS.Close
        
        'DAVID
        'Para que meta la letra de serie, NO el tipo moviemiento
        RS.Open "SELECT * FROM stipom WHERE codtipom='" & Cad & "'"
        If Not RS.EOF Then Cad = DBLet(RS!LetraSer, "T")
        RS.Close
        If Cad = "" Then Cad = CodTipoMov
'        If Text1(30).Text = "ARP" Then
            Text1(18).Text = "RECTIFICA A FACTURA: " & Cad & ", " & Text1(18).Text
'        Else
'            Text1(21).Text = "RECTIFICA A FACTURA: " & cad & ", " & Text1(21).Text
'        End If
        
        
        Set RS = Nothing
    Else
        'Añadiremos el boton de aceptar y demas objetos para insertar
        PonerModo 3
    End If
    
    NomTraba = ""
    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

    'El preparador del material lo hacemos tb al trabajador actual
    Text1(28).Text = Text1(3).Text
    Text2(28).Text = Text2(3).Text

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    Text1(30).Text = CodTipoMov
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
'    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    'Poner el Almacen por defecto del Trabajador
    'txtAux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    txtAux(0).Text = Format(AlmacenLineas, "000")
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    Text2(9).Text = ""
    BloquearTxt Text2(16), False
    BloquearTxt Text2(9), True
    
    ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
    'si contab. analitica por trabajador traer su centro de coste
    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
        txtAux(9).Text = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
        Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
    End If
    
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
        Text1(0).BackColor = vbLightBlue 'vbYellow
'        Text1(30).visible = False
'        Combo1.visible = True
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue 'vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
    
    If Me.EsHistorico = False Then
        'Hacer busquedar del tipo de movimiento de albaran en el que estamos
        Text1(30).Text = CodTipoMov
        BloquearTxt Text1(30), True
    End If
End Sub


Private Sub BotonVerTodos()
Dim Cad As String
'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        Cad = " codtipom='" & CodTipoMov & "'"
        MandaBusquedaPrevia Cad
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla
        If EsHistorico = False Then
            CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & CodTipoMov & "'"
        End If
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
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

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4

    PonerFoco Text1(1)
   
    'Si es Cliente de Varios no se pueden modificar sus datos
'    DeVarios = EsClienteVarios(Text1(4).Text)
'    BloquearDatosCliente (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If Data2.Recordset.EOF Then Exit Sub
    
    
    
    'bloqueamos el registro a modificar
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas) & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    CargaTxtAux True, False
    
    'si es factura rectificativa y es una linea de la factura que rectificamos
    'solo podremos modificar la cantidad el resto de campos bloqueados
    If CodTipoMov = "ART" Or CodTipoMov = "ARC" Then '(Albaran Rectificativo)
        vWhere = "codtipom='" & Text1(37).Text & "' and numfactu=" & Text1(36).Text & " and fecfactu=" & DBSet(Text1(35).Text, "F")
        vWhere = vWhere & " and codartic=" & DBSet(txtAux(1).Text, "T")
        vWhere = "SELECT COUNT(*) FROM slifac WHERE " & vWhere
        If RegistrosAListar(vWhere) > 0 Then
            'modificamos una linea de factura a rectificar y solo podemos modificar cantidad
            BloquearTxt txtAux(0), True
            BloquearTxt txtAux(1), True
            BloquearTxt txtAux(2), True
            BloquearTxt txtAux(4), True
            BloquearTxt txtAux(6), True
            BloquearTxt txtAux(7), True
            Me.cmdAux(0).Enabled = False
            Me.cmdAux(1).Enabled = False
        End If
    End If
    
    
    ' ---- [21/10/2009] [LAURA]: añadir campo centro de coste por trabajador
    'si contab. analitica por trabajador traer su centro de coste
    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
        txtAux(9).Text = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
        Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
    End If
    
    
    
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
'    PonerBotonCabecera False
    
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt Text2(9), True 'Campo nomprove
    BloquearTxt txtAux(2), True
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False

EModificarLinea:
    
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Mantenimientos (scaman)
' y los registros correspondientes de las tablas de lineas (sliman y slima1)
Dim Cad As String
Dim NumAlbElim As Long

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    Cad = "Cabecera de Albaranes." & vbCrLf
    Cad = Cad & "------------------------------------       " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Albaran:            "
    Cad = Cad & vbCrLf & "Tipo:  " & Text1(30).Text
    Cad = Cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Text1(1).Text
'    cad = cad & vbCrLf & "Cliente:  " & Format(Text1(4).Text, "000000") & " - " & Text1(5).Text
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
      
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Abrir frame de informes para pedir datos antes de grabar en el historico
        cadList = ""
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 80
        frmList.Show vbModal
        Set frmList = Nothing
        If cadList = "" Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumAlbElim = Data1.Recordset.Fields(1).Value
        CodTipoMov = Text1(30).Text
        
        If Not Eliminar(NumAlbElim) Then
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            PosicionarDataTrasEliminar
        End If
        
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim Sql As String

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
        
    If vParamAplic.ArtPortes <> "" Then
        'NO deberia eliminar el aticulo
        
    End If
    
    ModificaLineas = 3 'Eliminar
    Sql = "¿Seguro que desea eliminar la línea de Albaran?     "
    Sql = Sql & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    Sql = Sql & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    Sql = Sql & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        If EliminarLinea Then
            ModificaLineas = 0
            CargaGrid2 DataGrid1, Data2
            SituarDataTrasEliminar Data2, NumRegElim
            CalcularDatosFactura
        End If
'        CancelaADODC
    End If
    PonerModo 2
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String
Dim Port As Integer      'Port: para saber si ha metido/Modificado el articulo de portes

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
    
        'Fontenas
        
        If vParamAplic.ArtPortes <> "" Then
            'Si lleva portes haremos varias cosas
            Port = HacerAccionesPortes
            CargaGrid DataGrid1, Data2, True
            Set miRsAux = Nothing
        End If
    
        ' ---- [15/09/2009] (LAURA)
        DescuentosCantidad ""
        ' ----
    
        PonerModo 2
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            If Port = 0 Then    'Si  ha metido/modifgicado portes no hago nada (port>0)
            
                'Enero 2010
                'para que no se vuelva a la primera linea
                'DeseleccionaGrid DataGrid1
                'DataGrid1.Bookmark = 1
            Else
                Data2.Recordset.MoveLast  'El ultimo es el porte
            End If
        End If
        
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


Private Sub DataGrid1_DblClick()
    If Modo = 2 Then
        If Not Data2.Recordset.EOF Then AbrirForm_Articulos DBLet(Data2.Recordset!codArtic, "T")
    End If
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Modo = 2 And KeyCode = 113 Then
        If Not Data2.Recordset.EOF Then AbrirForm_Articulos DBLet(Data2.Recordset!codArtic, "T")
    End If
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Ayuda de Etiqueta de precio de salida de la Función de Precios
    If Data2.Recordset Is Nothing Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        Me.DataGrid1.ToolTipText = ""
        If X > 7660 And X < 7950 Then
            If IsNull(Me.Data2.Recordset!origpre) Then Exit Sub
            Select Case DataGrid1.Columns(9).Value
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
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim RS As ADODB.Recordset
Dim Sql As String

    On Error GoTo Error1

    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
        '- ampliacion lineas
        Sql = "select ampliaci from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " and numlinea=" & Data2.Recordset!numlinea
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            Text2(16).Text = DBLet(RS.Fields(0).Value, "T")
        End If
        RS.Close
        Set RS = Nothing
        
        If vEmpresa.TieneAnalitica Then
            '- centro de coste
            ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
            Me.txtAux(9).Text = DBLet(Data2.Recordset!CodCCost, "T")
            Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
        Else
            '- nombre proveedor
            Text2(9).Text = DBLet(Me.Data2.Recordset!nomprove, "T")
        End If
    Else
        Text2(16).Text = ""
        Text2(9).Text = ""
    End If
    Exit Sub
    
Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    If AlbAvisoGenerado > 0 Then
        PonerCadenaBusqueda
        'Simulo que pulsa lineas
        mnLineas_Click
        
        'Simulo que le da a insertar nueva
        mnNuevo_Click
        
        'AlbAvisoGenerado
        AlbAvisoGenerado = 0
    End If
        
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF And Modo <> 5 Then PonerCadenaBusqueda
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
'    btnAnyadir = 5
'    btnPrimero = 19
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'        .Buttons(10).Image = 10 'Mto Lineas Ofertas
'        .Buttons(11).Image = 33 'Nº Serie si lineas con articulos de control Nº serie
'        .Buttons(12).Image = 26 'GEnerar factura
'        .Buttons(13).Image = 30 'Marcar a facturar
'
'        .Buttons(14).Image = 27 'Imprimir portes
'        .Buttons(15).Image = 16 'Imprimir Pedido
''
'        .Buttons(16).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'
'        If vParamAplic.ArtPortes = "" Then
'            .Buttons(14).Style = tbrSeparator
'            .Buttons(14).ToolTipText = ""
'        Else
'            .Buttons(14).Style = tbrDefault
'            .Buttons(14).ToolTipText = "Imprimir portes"
'        End If
'    End With

    With Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun1
        'ASignamos botones
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2 'Ver Todos
        .Buttons(1).Image = 3 'Añadir
        .Buttons(2).Image = 4 'Modificar
        .Buttons(3).Image = 5 'Eliminar
        .Buttons(8).Image = 16 'Imprimir
    End With
    
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun1
        .Buttons(1).Image = 26 'GEnerar factura
        .Buttons(2).Image = 30 'Marcar a facturar
        If vParamAplic.ArtPortes = "" Then
            .Buttons(3).Style = tbrSeparator
            .Buttons(3).ToolTipText = ""
        Else
            .Buttons(3).Style = tbrDefault
            .Buttons(3).ToolTipText = "Imprimir portes"
        End If
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun1
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With

    With Me.ToolAux(0)
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
    End With




    Me.SSTab1.Tab = 0
    
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    
    Me.imgFecha(0).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Me.imgFecha(40).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    
    LimpiarCampos   'Limpia los campos TextBox
    
    CargarComboFacturacion
    VieneDeBuscar = False
    CodTipoMov = hcoCodTipoM
    
    If CodTipoMov = "ALR" Then
        Me.Caption = "Albaranes Reparación"
        Label1(3).visible = False
        Label1(5).visible = False
        Text1(23).visible = False
        Text1(24).visible = False
        Label1(12).visible = False
        Text1(2).visible = False
        'Captions
        Label1(11).Caption = "Nº Repa."
        Label1(10).Caption = "Fecha repara."
        Text1(24).visible = False
        'Terminal
        Text1(38).visible = False
        Text1(39).visible = False
        
        
    Else
        Label1(11).Caption = "Nº Pedido"
        Label1(10).Caption = "Fecha pedido"
    End If
   
    'Comprobar si es Departamento o Direccion
    If vParamAplic.Departamento Then
        Me.Label1(1).Caption = "Dpto."
    Else
        Me.Label1(1).Caption = "Direc."
    End If
    
    ' ---- [19/10/2009] [LAURA] : añadir centro de coste a la linea
    If vEmpresa.TieneAnalitica Then
        cmdAux(9).ToolTipText = "Buscar centro coste"
        txtAux(9).Tag = "centro coste"
        Label1(51).Caption = "Centro coste"
    Else
        Label1(51).Caption = "Proveedor"
    End If
    cmdAux(9).Tag = -1
        
        
    '## A mano
    Me.FrameHco.visible = EsHistorico
    Me.FrameFacRec.visible = (CodTipoMov = "ART" Or CodTipoMov = "ARC" Or CodTipoMov = "ARP")
    
    
    
    'Aportacion a terminal
    Label1(49).visible = hcoCodTipoM = "ALM" And vParamAplic.ctaAportacion <> ""
    Text1(40).visible = hcoCodTipoM = "ALM" And vParamAplic.ctaAportacion <> ""
    
    
    If Not EsHistorico Then
        NombreTabla = "scaalbcli"
        NomTablaLineas = "slialbcli" 'Tabla lineas de Albaranes
        Ordenacion = " ORDER BY codtipom, numalbar "
        
        If CodTipoMov = "ALV" Then
            Me.Caption = "Albaranes Socios"
        ElseIf CodTipoMov = "ALM" Then
            Me.Caption = "Albaranes de Mostrador"
            If Me.RecuperarFactu Then Me.Caption = "Albaranes de Mostrador (Recuperar facturas)"
        ElseIf CodTipoMov = "ART" Then
            Me.Caption = "Albaranes Rectificativos"
        ElseIf CodTipoMov = "ALS" Then
            Me.Caption = "Albaranes de Cuotas"
        ElseIf CodTipoMov = "ARC" Then
            Me.Caption = "Albaranes Rectificativos de Cuotas"
        End If
    Else
        NombreTabla = "schalb"
        NomTablaLineas = "slhalb"
        CargarTagsHco Me, "scaalbcli", NombreTabla
        'Estos campos solo estan en la tabla del histórico
        Text1(31).Tag = "Fecha Eliminación|F|N|||schalb|fechelim|dd/mm/yyyy|N|"
        Text1(32).Tag = "Trabajador Eliminación|N|N|0|9999|schalb|trabelim|0000|N|"
        Text1(33).Tag = "Incidencia elim.|T|N|||schalb|codincid||N|"
        Me.Caption = "Histórico Albaranes Socios"
        Ordenacion = " ORDER BY codtipom, numalbar,fechaalb "
    End If
 
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If AlbAvisoGenerado > 0 Then hcoCodMovim = AlbAvisoGenerado
        
    'Lo que hacia antes, todo normal
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        CadenaConsulta = CadenaConsulta & " where numalbar=-1"
    End If

    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
            Text1(0).BackColor = vbLightBlue 'vbYellow
        End If
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
    End If
'    CargaCombo
    
End Sub

'Private Sub CargaCombo()
'Dim RS As ADODB.Recordset
'Dim SQL As String
'Dim i As Byte
'
'    Combo1.Clear
'
'    SQL = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom LIKE 'AL%' and tipodocu=33"
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not RS.EOF
'        Combo1.AddItem RS!codtipom & "-" & RS!nomtipom
'        Combo1.ItemData(Combo1.NewIndex) = i
'        i = i + 1
'        RS.MoveNext
'    Wend
'    RS.Close
'    Set RS = Nothing
'End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.cboFacturacion.ListIndex = -1
    Me.chkFacturar.Value = 0
    Me.chkFacturarKm.Value = 0
    Me.chkDocArchi.Value = 0
    
    Text3(0).Text = "BASE IMP."
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    AlbAvisoGenerado = 0   'por si acaso
End Sub




Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Agentes
Dim indice As Byte
    indice = 17
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Agente
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom agente
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
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If EsCabecera Then 'Llama desde VerTodos del Form
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(30), CadenaDevuelta, 1)
            CadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
            CadB = CadB & " and " & Aux
            
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
                CadB = CadB & " and " & Aux
            End If
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 2), "0000000")
            
        ElseIf Val(cmdAux(9).Tag) > 0 Then
            'Llama desde boton busqueda centros de coste
            ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
            Me.txtAux(9).Text = RecuperaValor(CadenaDevuelta, 1)
            Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
        Else
            'Llama desde Prismatico Direcciones/Departamentos
            Text1(12).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            Text2(12).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
Dim encontrado
'Form Mantenimiento de Clientes
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Clien
'    encontrado = DevuelveDesdeBD(conAri, "nomsocio", "ssocio", "codsocio", Text1(4).Text, "T")
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 2)
    HaDevueltoDatos = True
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    HaDevueltoDatos = True
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim indice As Byte
Dim devuelve As String

    indice = 9
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve) 'Poblacion
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

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim indice As Byte
    indice = CByte(Me.imgFecha(0).Tag) + 1
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
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
    indice = 14
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'devuelve los datos necesarios para grabar en la tabla del historico
' o para recuperar una factura que vamos a Rectificar

    cadList = ""
    
    If frmList.OpcionListado = 225 Then  'Factura Rectificativa
        If CadenaSeleccion <> "" Then
            'codtipom
            cadList = " codtipom='" & RecuperaValor(CadenaSeleccion, 1) & "' and numfactu="
            'numfactu
            cadList = cadList & RecuperaValor(CadenaSeleccion, 2) & " and fecfactu="
            'fecfactu
            cadList = cadList & DBSet(RecuperaValor(CadenaSeleccion, 3), "F")
            
            'campos observaciones
            motivo = "MOTIVO: " & RecuperaValor(CadenaSeleccion, 4)
        End If
        
    Else 'Para recoger los Datos de Eliminacion que se introdujeron
        cadList = DBSet(RecuperaValor(CadenaSeleccion, 1), "F") & " as fechelim,"
        cadList = cadList & RecuperaValor(CadenaSeleccion, 2) & " as trabelim,"
        cadList = cadList & DBSet(RecuperaValor(CadenaSeleccion, 3), "T") & " as codincid"
    End If
End Sub


Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mensaje de Nº de Serie disponibles
'En cadena seleccion estan concatenados los seleccionados

    If frmMen.OpcionMensaje = 11 Or frmMen.OpcionMensaje = 111 Then
        'En cadenaseleccion tenemos la WHERE que selecciona las lineas de la factura
        'que nos queremos traer para generar un albaran de rectificacion
        'Insertaremos estas lineas en la tabla slialbcli, y luego se podran eliminar,modificar,etc. (son de apoyo)
         InsertarLineasFactu (CadenaSeleccion)
    Else
        If Text1(30).Text = "ART" Then
            'Albaran de factura rectificativa
            If Not QuitarNumSeriesAlbVenta(CadenaSeleccion) Then MsgBox "Los nº de serie a rectificar no se han actualizado correctamente.", vbExclamation
        Else
            If Not AsignarNumSeriesAlbVenta(CadenaSeleccion) Then
                MsgBox "Los nº de serie del albaran no se han actualizado correctamente.", vbExclamation
            End If
        End If
    End If
End Sub


Private Sub frmNSerie_CargarNumSeries()
Dim CadValues As String, cadValuesU As String
Dim devuelve As String
Dim TieneMan As String * 1

    'Estamos en VENTAS e insertamos datos venta vacios
    If ModificaLineas = 4 Then
        CargarNumSeries
    Else
        'Viene de insertar Nº de series al insertar una linea

        'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
        TieneMan = "0"
'        devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
'        'El cliente tiene Mantenimientos
'        If devuelve <> "" Then TieneMan = "1"
        
        'cadena para INSERT
        'Estamos en VENTAS e insertamos datos de Cliente
        CadValues = ""
        CadValues = CadValues & Text1(4).Text & ", " & DBSet(Text1(12).Text, "T") & ", " & TieneMan & ", " & DBSet(devuelve, "T") & ", "
        CadValues = CadValues & ValorNulo & ", " & ValorNulo & ", " 'Fecha ult. Repar y Fin Garantia
        'Datos Venta
        CadValues = CadValues & DBSet(Text1(30).Text, "T") & ", " & ValorNulo & ", '" & Format(Text1(1).Text, FormatoFecha) & "', " & Text1(0).Text & ", " & Me.cmdAux(0).Tag & ", "
        'Rellenar los datos COMPRA del Proveedor a NULO
        CadValues = CadValues & ValorNulo & ", " & ValorNulo & ", " & ValorNulo & ", " & ValorNulo
        
        'cadena para UPDATE
        cadValuesU = " codclien=" & Text1(4).Text & ", coddirec=" & DBSet(Text1(12).Text, "T") & ", "
        cadValuesU = cadValuesU & " tieneman=" & TieneMan & ", nummante=" & DBSet(devuelve, "T") & ", codtipom=" & DBSet(Text1(30).Text, "T")
        cadValuesU = cadValuesU & ", fechavta='" & Format(Text1(1).Text, FormatoFecha) & "' "
        cadValuesU = cadValuesU & ", numalbar=" & Text1(0).Text & ", numline1=" & Me.cmdAux(0).Tag
        InsertarNSeries txtAux(1).Text, CadValues, cadValuesU, True
    End If
End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(9).Text = RecuperaValor(CadenaSeleccion, 1)
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
    TerminaBloquear
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            Set frmCli = New frmFacClientes2
            frmCli.DatosADevolverBusqueda = "0|1|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            indice = 5
            PonerFoco Text1(indice)
            
            
            
        Case 1 'NIF para cliente de Varios
            Set frmCV = New frmGesSocios
            frmCV.DatosADevolverBusqueda = "0"
            frmCV.Show vbModal
            Set frmCV = Nothing
            indice = 6
            
        Case 2 'Cod. Direc.
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
             
        Case 3, 7, 8 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            If Index = 7 Then
                indice = 27
            ElseIf Index = 8 Then
                indice = 28
            Else
                indice = Index
            End If
            Me.imgBuscar(3).Tag = indice
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 4 'Forma de Pago
            indice = 14
            PonerFoco Text1(indice)
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 5 'Agente
            indice = 17
            PonerFoco Text1(indice)
            Set frmA = New frmFacAgentesCom
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 6 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            indice = 9
            VieneDeBuscar = True
            
        Case 9 'Cod Envio
            indice = 29
            PonerFoco Text1(indice)
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0"
            frmFE.Show vbModal
            Set frmFE = Nothing
    End Select
    
    PonerFoco Text1(indice)
    Screen.MousePointer = vbDefault
    
    If Modo = 4 Then
        If Not BLOQUEADesdeFormulario(Me) Then cmdCancelar_Click
    End If
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
    Else   'Eliminar Albaran
         BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
'Imprimir Albaran
    BotonImprimir 45, False '45: Informe de Albaranes
End Sub


Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Albaranes"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar albaran
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
    Else 'Añadir Cabecera
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


Private Sub SSTab1_Click(PreviousTab As Integer)
    Me.Label1(35).visible = Me.SSTab1.Tab = 0
    Me.Text2(16).visible = Me.SSTab1.Tab = 0
    Me.Label1(51).visible = (Modo = 5) And (vEmpresa.TieneAnalitica) And SSTab1.Tab = 0
    Me.Text2(9).visible = (Modo = 5) And (vEmpresa.TieneAnalitica) And Me.SSTab1.Tab = 0
    
End Sub



Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True        'Cod. Postal
    
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    txtAnterior = Text1(Index).Text
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
   
    If Not (Index = 30 And Modo = 1) Then ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ind As Integer
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
    
     If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
    
        If Text1(Index).Text = "" Then
            Ind = -1
            Select Case Index
            Case 3
                Ind = 3
            Case 4
                Ind = 0
            Case 6
                Ind = 1
            Case 9
                Ind = 6
            Case 12
                Ind = 2
            Case 17
                Ind = 5
            Case 14
                Ind = 4
            Case 27, 28, 29
                Ind = Index - 20
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
Dim campo As String
        
        
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
          
          
    'Por si no ha cambiado nada
    If txtAnterior = Text1(Index).Text Then Exit Sub
          
    
          
          
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 41 'Fecha Albaran,fecenvio
                If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 3, 27, 28 'Cod Vendedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
                If Index = 3 And Modo = 3 Then
                    Text1(28).Text = Text1(Index).Text
                    Text2(28).Text = Text2(Index).Text
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Cod. ssocio
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Modo=1 Busqueda
                    Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
                Else 'If Modo = 3 Then 'Modo Insertar
                    'si es ART-Albaran de factura Rectificativa ya he cargado los
                    'datos de la factura
'                    If CodTipoMov <> "ART" Then
                        PonerDatosCliente (Text1(Index).Text)
'                    Else
'                        campo = "nomsocio"
'                        devuelve = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(4).Text, "N", campo)
'                        If campo <> Text1(5).Text Then PonerDatosCliente Text1(Index).Text
'                    End If
                End If
            Else
                LimpiarDatosCliente
            End If
            
        Case 6 'NIF
'            If Not EsDeVarios Then Exit Sub
'            'si no se ha modificado el nif del cliente no hacer nada (Modo 4=Modificar)
'            If (Modo = 4) Then
'                If (Text1(6).Text = Data1.Recordset!nifClien) Then Exit Sub
'            End If
'            PonerDatosClienteVario (Text1(Index).Text)
                     
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
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            
            'Comprobar que el cliente seleccionada tiene esa direccion
            If PonerDptoEnCliente Then
                'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
            Else
                PonerFoco Text1(Index)
            End If
            
        Case 13 'Referencia Obligatoria
            If Trim(Text1(4).Text) <> "" Then ComprobarRefObligatoria
            
        Case 14 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 15, 16 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then   'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 17 'Cod. Agente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 29 'Cod envio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "senvio", "nomenvio")
            Else
                Text2(Index).Text = ""
            End If
        Case 40
            PonerFormatoDecimal Text1(Index), 3
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    'Poner el valor del combo Tipos de Movimiento Asociado
'    If Me.cboTipomov.ListIndex <> -1 Then
'        Text1(30).Text = ObtenerCodTipom
'    End If

    CadB = ObtenerBusqueda(Me, False)
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
'        If Me.EsHistorico = False Then
'            cadB = cadB & " and codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los ALV
'        End If
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
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
        Cad = Cad & ParaGrid(Text1(30), 10, "Tipo Alb.")
        Cad = Cad & ParaGrid(Text1(0), 15, "Nº Albaran")
        Cad = Cad & ParaGrid(Text1(1), 15, "Fecha Ped.")
        Cad = Cad & ParaGrid(Text1(4), 10, "Cliente")
        Cad = Cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
        Tabla = NombreTabla
        Titulo = "Albaranes"
        
        If EsHistorico Then
            Titulo = "Histórico de Albaranes"
            devuelve = "0|1|2|"
        Else
            Titulo = "Albaranes"
            devuelve = "0|1|"
        End If
    Else
        If vParamAplic.Departamento Then
            Titulo = "Dptos Cliente: "
            Desc = "Dpto."
        Else
            Titulo = "Direc. Cliente: "
            Desc = "Direc."
        End If
        Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
        Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|15·"
        Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||55·"
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
        frmB.vDevuelve = devuelve
'        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri  'Conexión a BD: Aritaxi
        If Not EsCabecera Then frmB.Label1.FontSize = 11
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
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            Me.cboFacturacion.ListIndex = -1
            PonerFoco Text1(kCampo)
            Text1(0).BackColor = vbLightBlue 'vbYellow
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
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar

    Screen.MousePointer = vbHourglass
    On Error GoTo EPonerLineas

    'Datos de la tabla slipre
    CargaGrid DataGrid1, Data2, True

    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim b As Boolean

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
     'Si es un Albaran de Ticket visualizamos unos datos y sino otros
    b = (Data1.Recordset!EsTicket = 1)
    Me.Toolbar1.Buttons(11).Enabled = (Not b) And (Not EsHistorico)
    

    If hcoCodTipoM <> "ALR" And hcoCodTipoM <> "ARC" Then
'        'sem. entrega pedido
'        Label1(12).visible = Not b
'        Text1(2).visible = Not b
'        'num oferta
'        Text1(23).visible = Not b And hcoCodTipoM <> "ALR"
'        'fecha oferta
'        Text1(24).visible = Not b
'        'nº terminal
'        Text1(38).visible = b
'        'nº venta
'        Text1(39).visible = b
    
    
        If b Then
        'El albaran se genero a partir de un ticket
            Me.Label1(11).Caption = "Nº Ticket"
            Me.Label1(10).Caption = "Fecha Ticket"
            Me.Label1(9).Caption = "Trabajador Ticket"
        
            'ocultamos los datos de la oferta
            Me.Label1(3).Caption = "Nº Venta"
            Label1(5).Caption = "Nº Terminal"
        Else
            Me.Label1(11).Caption = "Nº Pedido"
            Me.Label1(10).Caption = "Fecha Pedido"
            Me.Label1(9).Caption = "Trabajador Pedido"
    
            'Mostramos los datos de la oferta
            Me.Label1(3).Caption = "Nº Oferta"
            Label1(5).Caption = "Fecha Oferta"
        End If
        
    End If
    Text1(30).visible = True
'    Combo1.visible = False
    PonerCamposForma Me, Data1
    
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba", "codtraba")
    Text2(27).Text = PonerNombreDeCod(Text1(27), conAri, "straba", "nomtraba", "codtraba")
    Text2(28).Text = PonerNombreDeCod(Text1(28), conAri, "straba", "nomtraba", "codtraba")
    Text2(29).Text = PonerNombreDeCod(Text1(29), conAri, "senvio", "nomenvio")
    Text2(12).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
    Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    If EsHistorico Then
        'poner datos de eliminacion
        Text2(32).Text = PonerNombreDeCod(Text1(32), conAri, "straba", "nomtraba", "codtraba")
        Text2(33).Text = PonerNombreDeCod(Text1(33), conAri, "sincid", "nomincid", "codincid")
    End If
    
   
    
    
    CalcularDatosFactura
    
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

    lblF.Caption = ""

    'Actualiza Iconos Insertar,Modificar,Eliminar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
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
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    'Campo Nº Albaran y Tipo Movim. siempre bloqueado, excepto si estamos en modo de busqueda
    b = (Modo <> 1)
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(30), b
    'Bloquear los campos de Oferta
    If Text1(23).visible Then
        BloquearTxt Text1(23), b
        BloquearTxt Text1(24), b
    End If
    'Bloquear los campos de Pedido
    For i = 25 To 27
        BloquearTxt Text1(i), b
    Next i
    BloquearTxt Text1(2), b
    'bloquea los datos de venta del TPV (si hay)
    If Text1(38).visible Then
        BloquearTxt Text1(38), b
        BloquearTxt Text1(39), b
    End If
    
    'Bloquea los campos de Factura (si visibles, ed, si es Rectificativa)
    For i = 35 To 37
        BloquearTxt Text1(i), b
    Next i
  
    '-----  Datos Totales de Factura siempre bloqueado
    For i = 33 To 56
        BloquearTxt Text3(i), True
    Next i
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    For i = 46 To 48
        Text3(i).BackColor = &HFFFFC0
        Text3(i + 6).BackColor = &HFFFFC0
    Next i
    
    
    'Campos total Factura en verde
    Text3(55).BackColor = &HC0FFC0
    Text3(56).BackColor = &HC0FFC0    'Tatal factura
    '---------------------------------------------------
    
    b = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    Me.cboFacturacion.Enabled = b
    Me.chkFacturar.Enabled = b
    Me.chkFacturarKm.Enabled = b
    Me.chkDocArchi.Enabled = b
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    BloquearTxt Text2(16), (Modo <> 5)
    BloquearTxt Text2(9), (Modo <> 5)
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    Me.imgFecha(0).Enabled = b
    Me.imgFecha(40).Enabled = b
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    Me.imgBuscar(1).visible = False
    Me.imgBuscar(7).Enabled = (Modo = 1)
              
    'Modo Linea de Albaranes
    '- poner visible ampliacion linea
    BloquearTxt Text2(16), True
    '- poner visible nombre proveedor linea
    BloquearTxt Text2(9), True
    SSTab1_Click 0
      
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
Dim devuelve As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    'Asignarle el valor del Combo Tipo de Movimiento al texto oculto text1(30)
'    Text1(30).Text = ObtenerCodTipom
    
    Text1(28).Text = Text1(3).Text
    
    
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    'Comprobar si la referencia del cliente es obligatoria que tenga valor
'     If Trim(Text1(4).Text) <> "" Then
'        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "referobl", "codclien", Text1(4).Text, "N")
'        If devuelve = "1" And Text1(13).Text = "" Then 'Referencia Obligatoria
'            MsgBox "La Referencia del Cliente es Obligatoria.", vbInformation
'            PonerFoco Text1(13)
'            B = False
'        End If
'    End If
    If Not b Then Exit Function
    
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea(ByRef vCStock As CStock) As Boolean
Dim b As Boolean
Dim i As Byte
Dim Aux As String

    
    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For i = 0 To txtAux.Count - 2
        If txtAux(i).Text = "" And (i <> 5 And i <> 10) Then
            'El campo 5= origpre puede ser nulo (en alb.repar)
            MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
            b = False
            PonerFoco txtAux(i)
            Exit Function
        End If
    Next i
    
    RecalcularImportes txtAux(8), True, txtAux(3), txtAux(4), txtAux(6), txtAux(7)
    
'    '-- [19/10/2009] [LAURA]
'    '-- comprobar centro de coste
'    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 2 Then
'        b = (txtAux(8).Text <> "")
'    End If
    
    'Comprobar si se mueve mover stock (hay stock, o si no hay pero no control de stock)
    If vCStock.MueveStock Then
        b = vCStock.MoverStock(False, False)
    End If
    DatosOkLinea = b
    Exit Function
    
EDatosOkLinea:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
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

Private Sub Text2_LostFocus(Index As Integer)
    If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
End Sub


Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    BotonMtoLineas 0, "Albaranes"
    
    Select Case Button.Index
        Case 1: mnNuevo_Click 'Nuevo
        Case 2: mnModificar_Click  'Modificar
        Case 3: mnEliminar_Click  'Borrar
    End Select
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 5: mnBuscar_Click  'Buscar
        Case 6: BotonVerTodos  'Todos
            
        Case 1: mnNuevo_Click 'Nuevo
        Case 2: mnModificar_Click  'Modificar
        Case 3: mnEliminar_Click  'Borrar
            
        Case 8: mnImprimir_Click 'Imprimir Albaran
            
'        Case 10: mnLineas_Click  'Lineas
'        Case 11: BotonNSeries 'Control Nº Series
'
'        Case 12 'Generar Factura Mostrador
'                'o Factura Rectificativa (FRT)
'            If Me.RecuperarFactu Then
'                BotonRecuperarFactura
'            Else
'
'                'Septiebmre2009
'                If Data2.Recordset Is Nothing Then Exit Sub
'                If Data2.Recordset.RecordCount = 0 Then
'                    MsgBox "No tiene lineas de albarán", vbExclamation
'                    Exit Sub
'                End If
'
'
'                'procedimiento normal
'                If Data1.Recordset!codtipom = "ART" Then
'                    'Comprobar nº serie de las facturas rectificativas
'                    DevolverNumSeries
'                End If
'
'
'                'Comprobar que esta marcada para facturar
''                If Data1.Recordset!codTipoM <> "ALM" Then Exit Sub
'                If Me.chkFacturar.Value = 1 Then
'                    NumRegElim = Data1.Recordset.AbsolutePosition
'
'                    'Facturacion de Albaran de Mostrador
'                    frmListadoPed.CodClien = CodTipoMov  'utilizamos esta vble para pasarle el tipo de movimiento
'                    frmListadoPed.NumCod = Text1(0).Text  'utilizamos esta vble para pasarle el nº albaran
'                    frmListadoPed.EstaRecupFact = False
'                    frmListadoPed.Escliente = True
'                    AbrirListadoPed (222)
'
'                    PosicionarDataTrasEliminar
'                Else
'                    MsgBox "El Albaran no esta marcado para facturar", vbInformation
'                End If
'            End If
'
'        Case 13
'            'DAVID
'            'Marca los albaranes que esten como NO facturar a facturar
'            Screen.MousePointer = vbHourglass
'            MarcarAlbaranes
'            Screen.MousePointer = vbDefault
'
'        Case 14
'            If vParamAplic.ArtPortes = "" Then Exit Sub
'
'            BotonImprimir 45, True
'
'        Case 15: mnImprimir_Click 'Imprimir Albaran
'        Case 16: mnSalir_Click   'Salir
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

  
'DesdeRecuperaParaRectificativa:  Para que no inserte el punto verde
Private Function InsertarLinea(numlinea As String, DesdeRecuperaParaRectificativa As Boolean) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialbcli
Dim Sql As String
Dim vWhere As String
Dim b As Boolean
Dim vCStock As CStock
Dim ImpReciclado As Single
Dim DentroTRANS As Boolean

    InsertarLinea = False
    Sql = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    Me.cmdAux(0).Tag = numlinea 'Aqui almaceno el Nº linea que acabo de Insertar
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S", numlinea) Then Exit Function
    
    If DatosOkLinea(vCStock) Then 'Lineas de Albaranes
        'Inserta en tabla "slialbcli"
        Sql = "INSERT INTO " & NomTablaLineas
        Sql = Sql & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,numbultos,precioar, dtoline1, dtoline2, importel, origpre,codprovex,numlote,codccost) "
        Sql = Sql & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        Sql = Sql & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        '- cantidad,numbultos
        Sql = Sql & DBSet(txtAux(3).Text, "N") & ", " & DBSet(txtAux(10).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(4).Text, "N") & ", " & DBSet(txtAux(6).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(7).Text, "N") & ","
        Sql = Sql & DBSet(txtAux(8).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(5).Text, "T", "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            '- codprove,numlote,codccost
            Sql = Sql & "0," & DBSet(txtAux(11).Text, "T", "S") & "," & DBSet(txtAux(9).Text, "T", "S") & ")"
        Else
            '- codprove,numlote,codccost
            Sql = Sql & DBSet(txtAux(9).Text, "N", "N") & "," & DBSet(txtAux(11).Text, "T", "S") & "," & ValorNulo & ")"
        End If
        
        '-
'        sql = sql & DBSet(txtAux(11).Text, "T", "S") & ")"
     Else
        Exit Function
     End If
    
    If Sql <> "" Then
        On Error GoTo EInsertarLinea
        conn.BeginTrans
        DentroTRANS = True
        
        'insertar la linea
        conn.Execute Sql
        
        'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
        'en actualizar stock comprobamos si el articulo tiene control de stock
        b = vCStock.ActualizarStock(False, True)
        
        
        
        
        'Si ha actualizado el sctock
        If b Then
            If ClienteConTasaReciclado And Not DesdeRecuperaParaRectificativa Then
                If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                    'Insertamos la linea del reciclado

'--[Monica]arituclo de reciclado pasa a ser reutilizado como articulo de gastos de administracion
'                    vWhere = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtReciclado, "T")
'                    SQL = "INSERT INTO " & NomTablaLineas
'                    SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,  precioar,"
'                    SQL = SQL & "dtoline1, dtoline2, importel, origpre) "
'                    SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & numlinea + 1 & ", " & Val(txtAux(0).Text) & ","
'                    SQL = SQL & DBSet(vParamAplic.ArtReciclado, "T") & "," & DBSet(vWhere, "T") & ", Null, "
'                    SQL = SQL & DBSet(txtAux(3).Text, "N") & "," 'Cantidad. La misma
'                    SQL = SQL & DBSet(ImpReciclado, "N") & ",0,0,"
'                    'Importe linea
'                    ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
'                    SQL = SQL & DBSet(ImpReciclado, "N") & ", 'A')"
'                    conn.Execute SQL
                        
                    
                End If 'articulo con sunida reciclado
            End If  'Cliente con tasa reciclado
        End If 'ok actualiza stock
        
        
    
    End If
    Set vCStock = Nothing
    
    
    
    If b Then
        conn.CommitTrans
        InsertarLinea = True
        AlmacenLineas = CInt(txtAux(0).Text)
        
        ' ---- [15/09/2009] (LAURA)
'        'Miramos en los descuentos
'        'Hacer sdesca
'        ElArticulo = txtAux(1).Text
'        DescuentosCantidad ElArticulo
        ' ----
        
    Else
        conn.RollbackTrans
         InsertarLinea = False
    End If
    Exit Function
    
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
        If DentroTRANS Then conn.RollbackTrans
        MuestraError Err.Number, "Insertar Lineas Albaran" & vbCrLf & Err.Description
    End If

End Function




Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialbcli
Dim Sql As String
Dim vCStock As CStock
Dim b As Boolean
Dim ImpReciclado As Single

    On Error GoTo EModificarLinea

    ModificarLinea = False
    Sql = ""
    
    '## LAURA 15/11/2006
    'si se ha modificado el articulo eliminar de la smoval y reestablecer stock
    'Inicilizar la clase para Actualizar los stocks
    
    
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S") Then Exit Function
    
    '#### LAURA 15/11/2006
    conn.BeginTrans
        
    If DatosOkLinea(vCStock) Then
        
        
'        Set vCStock = New CStock
        'iniciamos la clase con los valores anteriores para deshacer lo q insertamos antes
        b = InicializarCStock(vCStock, "E")
        If b Then
            b = vCStock.DevolverStock2 'eliminamos de smoval y devolvemos stock valores anteriores
            If b Then
                'si se ha modificado el articulo
                If CStr(Data2.Recordset!codArtic) <> txtAux(1).Text Then
                    'si la linea tenia numero de serie vaciar los campos correspondien al albaran venta
                    Sql = "UPDATE sserie SET codclien=" & ValorNulo & ",codtipom=" & ValorNulo & ", fechavta=" & ValorNulo & ",numalbar=" & ValorNulo & ",numline1=" & ValorNulo
                    Sql = Sql & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " and codtipom='" & CodTipoMov & "' and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
                    Sql = Sql & " AND numalbar=" & Data1.Recordset!NumAlbar & " AND numline1=" & Data2.Recordset!numlinea
                    conn.Execute Sql
                End If
            End If
            'ahora leemos los valores nuevos
            If b Then b = InicializarCStock(vCStock, "S")
            'insertamos en smoval y actualizamos stock a los valores nuevos
            vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text))
            If b Then b = vCStock.ActualizarStock(False, True)
    
            'actualizar la linea de Albaran
            If b Then
                Sql = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
                Sql = Sql & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
                Sql = Sql & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", numbultos=" & DBSet(txtAux(10).Text, "N") & ","
                Sql = Sql & "precioar= " & DBSet(txtAux(4).Text, "N") & ", " 'precio
                Sql = Sql & "dtoline1= " & DBSet(txtAux(6).Text, "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N") & ", "
                Sql = Sql & "importel= " & DBSet(txtAux(8).Text, "N") & ", " 'Importe
                Sql = Sql & "origpre=" & DBSet(txtAux(5).Text, "T", "S") & ","
                ' ---- [19/10/2009] [LAURA] : añadir centro de coste a la linea
                If vEmpresa.TieneAnalitica Then
                    Sql = Sql & "codccost=" & DBSet(txtAux(9).Text, "T", "S") & ","
                Else
                    Sql = Sql & "codprovex=" & DBSet(txtAux(9).Text, "N", "N") & ","
                End If
                Sql = Sql & "numlote=" & DBSet(txtAux(11).Text, "T", "S") & ""
                
                Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
                conn.Execute Sql
                
                
                'Llegado aqui, si tiene Punto verde(tasa ecologica)
                'Y el cliente tiene tasa recliclado
                If ClienteConTasaReciclado Then
                    If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                        
                       'Si el articulo siguiente es PV entoces lo updatearemos
                       Sql = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea"
                       'QUITO EL WHERE
                       Sql = Mid(Sql, 8)
                       NumRegElim = Val(DBLet(Data2.Recordset!numlinea, "N")) + 1
                       Sql = DevuelveDesdeBD(conAri, "codartic", "slialbcli", Sql, CStr(NumRegElim))
                       'En SQL tengo el codarti de la linea SIGUIENTE
                       'SI es punto verde de parametros, supondremos que esta vinculado con la linea que estamos modificando

'--[Monica]arituclo de reciclado pasa a ser reutilizado como articulo de gastos de administracion
'                       If SQL = vParamAplic.ArtReciclado Then
'
'                            SQL = "UPDATE " & NomTablaLineas & " SET "
'                            SQL = SQL & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
'                            SQL = SQL & "precioar= " & DBSet(ImpReciclado, "N") & ", " 'precio
'                            ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
'                            SQL = SQL & "importel= " & DBSet(ImpReciclado, "N")  'Importe
'                            'WHERE
'                            SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & NumRegElim
'                            conn.Execute SQL
'                      End If  'linea siguiente con codarti=puntoverde
                    End If  'articulo con reciclado
                End If ' de cliente con tasa reciclado
                
            End If
'        If SQL <> "" Then
'
'            vCStock.Cantidad = CSng(txtAux(3).Text)
'            b = vCStock.ModificarStock(Data2.Recordset!Cantidad)
'        End If
        End If
    End If
    Set vCStock = Nothing
    
EModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Albaran" & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        ModificarLinea = True
        
        ' ---- [15/09/2009] (LAURA)
'        If txtAux(1).Text = Data2.Recordset!codArtic Then
'            ElArticulo = Data2.Recordset!codArtic
'        Else
'            'Son distintos. Que recalcule todo
'            ElArticulo = ""
'        End If
'        DescuentosCantidad ElArticulo
        ' ----
        
    Else
        conn.RollbackTrans
         ModificarLinea = False
    End If
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
        Me.cmdRegresar.Cancel = True
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.cmdCancelar.Cancel = True
    End If
    
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Sql As String
    
    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    CargaGrid2 vDataGrid, vData
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    vDataGrid.ScrollBars = dbgAutomatic
    PrimeraVez = False
    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Byte
    
    On Error GoTo ECargaGrid

    vData.Refresh

    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False
    vDataGrid.Columns(2).visible = False

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
            i = 3
            vDataGrid.Columns(i).Caption = "Alm."
            vDataGrid.Columns(i).Width = 670
            vDataGrid.Columns(i).NumberFormat = "000"
            
            i = i + 1 '4
            vDataGrid.Columns(i).Caption = "Articulo"
            vDataGrid.Columns(i).Width = 1700
            i = i + 1 '5
            vDataGrid.Columns(i).Caption = "Desc. Artículo"
            vDataGrid.Columns(i).Width = 3500

            i = 6
            vDataGrid.Columns(i).visible = False
            i = 7
            vDataGrid.Columns(i).Caption = "Cantidad"
            vDataGrid.Columns(i).Width = 1050
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
            i = 8
            vDataGrid.Columns(i).visible = False
                
            i = i + 1 '9
            vDataGrid.Columns(i).Caption = "Precio"
            vDataGrid.Columns(i).Width = 1150
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoPrecio
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "OP"
            vDataGrid.Columns(i).Width = 350
            vDataGrid.Columns(i).Alignment = dbgCenter
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "Dto. 1"
            vDataGrid.Columns(i).Width = 700
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoDescuento
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "Dto. 2"
            vDataGrid.Columns(i).Width = 700
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoDescuento
                
            i = i + 1
            vDataGrid.Columns(i).Caption = "Importe linea"
            vDataGrid.Columns(i).Width = 1500
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
           
            i = i + 1
            vDataGrid.Columns(i).visible = False
            i = i + 1
            vDataGrid.Columns(i).visible = False
            
            If vEmpresa.TieneAnalitica Then
                i = i + 1
                vDataGrid.Columns(i).visible = False
            End If
                

                
    End Select

    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
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
        cmdAux(9).visible = visible
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
                    txtAux(i).Text = DataGrid1.Columns(i + 3).Text
                ElseIf i = 3 Then
                    txtAux(i).Text = DataGrid1.Columns(i + 4).Text
                ElseIf i >= 4 And i < 9 Then
                    txtAux(i).Text = DataGrid1.Columns(i + 5).Text
                ElseIf i = 9 Then
                    txtAux(i).Text = DataGrid1.Columns(14).Text
                ElseIf i = 10 Then
                    'txtAux(i).Text = DataGrid1.Columns(8).Text
                ElseIf i > 10 Then
                    ' ---- [19/10/2009] [LAURA] : centro de coste si hay conta analitica
'                    If vEmpresa.TieneAnalitica Then
'                        'txtAux(i).Text = DataGrid1.Columns(i + 4).Text
'                        txtAux(i).Text = DataGrid1.Columns(i + 3).Text
'                    Else
'                        'txtAux(i).Text = DataGrid1.Columns(i + 5).Text
'                        txtAux(i).Text = DataGrid1.Columns(i + 2).Text
'                    End If
                    
                End If
                txtAux(i).Locked = False
            Next i
        End If
        
        cmdAux(0).Enabled = True
        cmdAux(1).Enabled = True
'        cmdAux(9).Enabled = True
               
        'El Campo de Origen del precio se actualiza por programa al modificar el precio
        BloquearTxt txtAux(5), True
        'El campo Importe es calculado y lo bloqueamos.
        '##### Laura Recuperar facturas ALZIRA
'        BloquearTxt txtAux(8), True
        BloquearTxt txtAux(8), Not (Me.RecuperarFactu)
        '#####
       
        
        
        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For i = 0 To txtAux.Count - 1
            If i < 9 Then
                txtAux(i).Top = alto
                txtAux(i).Height = DataGrid1.RowHeight
            End If
        Next i
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
'        cmdAux(9).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
'        cmdAux(9).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(3).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 50
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 10
        txtAux(1).Width = DataGrid1.Columns(4).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width + 20
        txtAux(2).Width = DataGrid1.Columns(5).Width - 20
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 20
        txtAux(3).Width = DataGrid1.Columns(7).Width - 20
        txtAux(4).Left = txtAux(3).Left + txtAux(3).Width + 20
        txtAux(4).Width = DataGrid1.Columns(9).Width - 20
        
        'OP, Dto1, Dto2, Precio, (codProve/codccost)
        For i = 5 To 8
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 20
            txtAux(i).Width = DataGrid1.Columns(i + 5).Width - 20
        Next i
        
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux.Count - 1
            If i < 9 Then
                txtAux(i).visible = visible
            End If
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
'        cmdAux(9).visible = visible
    End If
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
'        Case 1: BotonNSeries 'Control Nº Series
            
        Case 1 'Generar Factura Mostrador
                'o Factura Rectificativa (FRT)
            If Me.RecuperarFactu Then
                BotonRecuperarFactura
            Else
            
                'Septiebmre2009
                If Data2.Recordset Is Nothing Then Exit Sub
                If Data2.Recordset.RecordCount = 0 Then
                    MsgBox "No tiene lineas de albarán", vbExclamation
                    Exit Sub
                End If
            
            
                'procedimiento normal
                If Data1.Recordset!codtipom = "ART" Then
                    'Comprobar nº serie de las facturas rectificativas
                    DevolverNumSeries
                End If
                    
                    
                'Comprobar que esta marcada para facturar
'                If Data1.Recordset!codTipoM <> "ALM" Then Exit Sub
                If Me.chkFacturar.Value = 1 Then
                    NumRegElim = Data1.Recordset.AbsolutePosition
                    
                    'Facturacion de Albaran de Mostrador
                    frmListadoPed.CodClien = CodTipoMov  'utilizamos esta vble para pasarle el tipo de movimiento
                    frmListadoPed.NumCod = Text1(0).Text  'utilizamos esta vble para pasarle el nº albaran
                    frmListadoPed.EstaRecupFact = False
                    frmListadoPed.Escliente = True
                    AbrirListadoPed (222)
                    
                    PosicionarDataTrasEliminar
                Else
                    MsgBox "El Albaran no esta marcado para facturar", vbInformation
                End If
            End If
            
        Case 2
            'DAVID
            'Marca los albaranes que esten como NO facturar a facturar
            Screen.MousePointer = vbHourglass
            MarcarAlbaranes
            Screen.MousePointer = vbDefault
            
        Case 3
            If vParamAplic.ArtPortes = "" Then Exit Sub
                
            BotonImprimir 45, True
    
    End Select

End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub TxtAux_Change(Index As Integer)
    If Index = 4 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
        txtAux(5).Text = "M"
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
    '[Monica]13/09/2012: Solo permitimos ver el articulo, no los precios
    If Index = 3 Then LabelAyudatxtAux Index, lblF
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub 'campo almacen y flecha arriba
    
    If Index < 2 Or Index = 9 Then  'Para los que tienen busqueda
    
        
            'Insertando linea albaran
            
            If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
                
                If Modo = 5 And ModificaLineas = 1 Then
                    If txtAux(Index).Text = "" Then
                        PulsadoMas2 = True
                        KeyCode = 0
                
                        PulsarTeclaMas False, Index
                    End If
                End If
            Else
                'Ha pulsado F2
                If KeyCode = 113 Then Me.DataGrid1.Columns(4).Caption = "EAN"
            End If
    
        
    ' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
    ElseIf KeyCode = 113 Then AccionesF2 Index
    ' ----
    End If
    KEYdown KeyCode
End Sub

Private Sub AccionesF2(Index As Integer)
    If Index = 3 Then
        AbrirForm_Articulos txtAux(1).Text
    Else
'        If Index = 4 Then AbrirConsultaPrecio Text1(4).Text, txtAux(1).Text
            
    End If
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String, cadMen As String
Dim codTarif As String
Dim CPrecioFact As CPreciosFact
Dim NumCajas As Integer, RestoUnid As Integer
Dim OrigP As String 'De donde viene el precio
Dim Cantidad As String
Dim vCStock As CStock
Dim b As Boolean
Dim okArticulo As Boolean
Dim DtoPermitido As Boolean
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtAux(Index).Text = Mid(txtAux(Index).Text, 1, Len(txtAux(Index).Text) - 1)
        Exit Sub
    End If
    
    Select Case Index
        Case 0 'Cod ALMACEN
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(Index)

        Case 1 'Cod. ARTICULO
            If txtAux(Index).Text = "" Then
                txtAux(2).Text = ""
                Exit Sub
            End If
        
            If txtAux(0).Text = "" Then
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If
        
            devuelve = ""
            If ModificaLineas = 2 Then
                If Not Data2.Recordset.EOF Then devuelve = Data2.Recordset!codArtic
            End If
            Cantidad = txtAux(9).Text
            
            If Me.DataGrid1.Columns(4).Caption = "EAN" Then
                'Ha pulsado F2, para meter, en lugar del codigo del articulo, el EAN
                okArticulo = PonerArticuloEan(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, devuelve, , Cantidad)
            Else
                okArticulo = PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, devuelve, , Cantidad)
            End If
            If Not okArticulo Then
                If Me.DataGrid1.Columns(4).Caption = "EAN" Then txtAux(1).Text = ""
                PonerFoco txtAux(Index)
            Else
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
                
                
                '---- [20/10/2009] [LAURA] : añadir centro de coste
                If Not vEmpresa.TieneAnalitica Then
                    'Si  ha cambiado el articulo, el proveedor
                    If txtAux(9).Text = "" Then
                        txtAux(9).Text = Cantidad
                        'Fuerzo el lostfocus para que carge el proveedor
                        txtAux_LostFocus 9
                    End If
                ElseIf vParamAplic.ModoAnalitica = 1 Then 'Por familia
                    txtAux(9).Text = Cantidad
                    Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
                End If
                '----
                
                
'                'Si  ha cambiado el articulo, el proveedor
'                If txtAux(9).Text = "" Then
'                    txtAux(9).Text = Cantidad
'                    'Fuerzo el lostfocus para que carge el proveedor
'                    txtAux_LostFocus 9
'                End If
            End If
        
        Case 2 'Nombre Articulo
           If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Si es factura rectifica la cantidad solo puede ser negativa
                If CodTipoMov = "ART" Or CodTipoMov = "ARC" Then
                    If CCur(txtAux(Index)) >= 0 Then
                        MsgBox "En facturas rectificativas la cantidad debe ser negativa.", vbExclamation
                        PonerFoco txtAux(Index)
                        Exit Sub
                    End If
                End If
            
                'Comprobar si hay suficiente stock
                Set vCStock = New CStock
                If Not InicializarCStock(vCStock, "S") Then Exit Sub
                If vCStock.MueveStock Then 'Comprobar si el articulo mueve stock: tiene control de stock y no es instalacion
                  If Not vCStock.MoverStock(False, False, False) Then
                    PonerFoco txtAux(Index)
                    Set vCStock = Nothing
                    Exit Sub
                  End If
                End If
                
                b = False
                If Modo = 5 Then 'Modo lineas
                    'Comprobar si el articulo se vende por cajas antes de entrar a la función
                    devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", txtAux(1).Text, "T")
                    
                    If devuelve <> "" Then
                        '- obtener el nº bultos: cantidad/unids.caja
                        txtAux(10).Text = CalcularNumBultos(CCur(txtAux(3).Text), CInt(devuelve))
                    End If
                    
                    If ModificaLineas = 1 Then 'insertar linea
                        b = True
                    ElseIf ModificaLineas = 2 Then 'modificar linea
                        If Data2.Recordset!codArtic <> txtAux(1).Text Then b = True
                    End If
                End If
                
                If b Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    If devuelve <> "" Then
'                        '- obtener el nº bultos: cantidad/unids.caja
'                        txtAux(10).Text = CalcularNumBultos(CCur(txtAux(3).Text), CInt(devuelve))
                        
                    
                        Set CPrecioFact = New CPreciosFact
                        'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                        'precio de caja, y otra linea con el resto unidades un precio unidad
                        Cantidad = txtAux(Index).Text
                        
                        
                        'Mayo 2009
                        'Si este parametro esta a FALSE, siempre cojera precio ud
                        If vParamAplic.CajasCompletas Then
                            NumCajas = CPrecioFact.ObtenerNumCajas(Cantidad, devuelve)
                            RestoUnid = CInt(ComprobarCero(Cantidad)) - NumCajas * CInt(devuelve)
                        Else
                            NumCajas = 0
                            RestoUnid = 0
                        End If
                        
                        CPrecioFact.CodigoClien = Text1(4).Text
                        
                        'Obtenemos la Tarifa del Cliente
                        'AHORA ESTA DENTRO DE LA CLASE
                        'codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                        
'[Monica]12/09/2012: he quitado lo de fijartarifaactividad y he puesto a piñon estas dos instrucciones
                        CPrecioFact.CodigoLista2 = 0
                        CPrecioFact.CodActividad = 0
'                        CPrecioFact.FijarTarifaActividad
                        CPrecioFact.CodigoArtic = txtAux(1).Text
                        
                        PorCaja = (NumCajas > 0)
                        Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP)
                        'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
                        'Ya que a regresado con pvp del Articulo
                        If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
                            cadMen = "El Artículo puede venderse por Cajas (" & devuelve & "uds. por Caja)." & vbCrLf
                            cadMen = cadMen & vbCrLf & "Inserte dos Lineas:   "
                            cadMen = cadMen & vbCrLf & "   Linea 1:  " & NumCajas * CInt(devuelve) & " uds a Precio Caja"
                            cadMen = cadMen & vbCrLf & "   Linea 2:  " & CInt(Cantidad) - NumCajas * CInt(devuelve) & " uds a Precio Unidad"
                            MsgBox cadMen, vbInformation
                        Else
                            If (txtAux(4).Text = "") Or (txtAux(4).Text <> "" And ModificaLineas = 2 And b) Then
                                txtAux(4).Text = Precio
                                txtAux(5).Text = OrigP 'De donde viene el precio
                            End If
                            PonerFormatoDecimal txtAux(4), 2
                            If txtAux(6).Text = "" Then txtAux(6).Text = CPrecioFact.Descuento1
                            PonerFormatoDecimal txtAux(6), 4
                            If txtAux(7).Text = "" Then txtAux(7).Text = CPrecioFact.Descuento2
                            PonerFormatoDecimal txtAux(7), 4
                        End If
                        
                        'Si tiene dto permitido
                        If Not CPrecioFact.DtoPermitido Then
                            txtAux(6).Text = "0"
                            txtAux(7).Text = "0"
                            txtAux(6).Enabled = False
                            txtAux(7).Enabled = False
                        End If
                        
'                            PonerFoco txtAux(Index + 1)
                        Set CPrecioFact = Nothing
                    End If
                End If
                Set vCStock = Nothing
            End If
            
            
        Case 4 'Precio
             If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
                If ModificaLineas = 1 Then
                    If CSng(txtAux(Index).Text) <> CSng(ComprobarCero(Precio)) Then txtAux(5).Text = "M"
                End If
            End If
            
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
            
            
        Case 9
            ' ---- [19/10/2009] [LAURA]: añadir centro de coste a la linea
            If txtAux(9).Text = "" Then
                 Text2(9).Text = ""
            Else
                If vEmpresa.TieneAnalitica Then
                    'centro de coste
                    ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
                    Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
            
                Else
                    'Cod proveeee
'                    If txtAux(9).Text = "" Then
'                        devuelve = ""
'                    Else
                        If Not IsNumeric(txtAux(9).Text) Then
                            MsgBox "Campo proveedor debe ser numérico", vbExclamation
                            devuelve = ""
                        Else
                                
                            devuelve = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", txtAux(9).Text, "N")
                            If devuelve = "" Then MsgBox "No existe el proveedor: " & txtAux(9).Text, vbExclamation
                        End If
                        If devuelve = "" Then
                            txtAux(9).Text = ""
                            PonerFoco txtAux(9)
                        End If
'                    End If
                    Text2(9).Text = devuelve
                End If
            End If
        
           

           
    End Select
    
     If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
'        If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'        If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
        If txtAux(1).Text = "" Then Exit Sub
        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(8), 1
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
'    Me.SSTab1.Tab = numTab
'    TituloLinea = Cad
    ModificaLineas = 0
    
'        If vParamAplic.ArtReciclado <> "" Then
'            ClienteConTasaReciclado = Val(DevuelveDesdeBD(conAri, "tasareciclado", "sclien", "codclien", Text1(4).Text)) = 1
'        Else
'            ClienteConTasaReciclado = False
'        End If
    
        
    If vParamAplic.ArtPortes <> "" Then
        KilosAnteriores = SumaKilosLineas
    End If
    
    PonerModo 5
'    PonerBotonCabecera True
    
    AlmacenLineas = -1
    PonerUltAlmacen
    
End Sub


Private Function Eliminar(NumAlbElim As Long) As Boolean
Dim Sql As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim MenError As String

    On Error GoTo FinEliminar
    conn.BeginTrans
    
    Sql = ObtenerWhereCP(False)
    
    'Reestablecer el stock en la tabla salmac a partir de todas las lineas del albaran
    MenError = "Restableciendo stocks de almacen."
    b = ReestablecerStock
    
    
    If b Then
        'eliminamos de albaranes y pasamos al historico
        b = ActualizarElTraspaso(MenError, Sql, CodTipoMov, cadList)
        
        If b Then
            MenError = "Actualizando numeros de serie."
            'Actualizar los posibles num. serie de ese albaran. vaciar los campos
            Sql = "UPDATE  sserie SET codclien=" & ValorNulo & ", codtipom=" & ValorNulo & ","
            Sql = Sql & " fechavta=" & ValorNulo & ", numalbar=" & ValorNulo & ", numline1=" & ValorNulo
            Sql = Sql & " WHERE codtipom='" & CodTipoMov & "' AND numalbar=" & Data1.Recordset!NumAlbar & " AND fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
            conn.Execute Sql
            
            
            'Devolvemos contador, si no estamos actualizando
            Set vTipoMov = New CTiposMov
            b = CBool(vTipoMov.DevolverContador(CodTipoMov, NumAlbElim))
            Set vTipoMov = Nothing
        End If
    End If
        
FinEliminar:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, MenError, Err.Description
    End If
    If Not b Then
        conn.RollbackTrans
    Else
        conn.CommitTrans
    End If
    Eliminar = b
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         vWhere = Replace(vWhere, NombreTabla & ".", "")
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
'         If SituarDataGral(Data1, Text1(30).Text, "T", Text1(0).Text, "N", Indicador) Then
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
    
    Sql = " " & NombreTabla & ".codtipom= '" & Text1(30).Text & "' and " & NombreTabla & ".numalbar= " & Val(Text1(0).Text)
    If EsHistorico Then Sql = Sql & " AND " & NombreTabla & ".fechaalb=" & DBSet(Text1(1).Text, "F")
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    
    'Enero 2008. David
    'Para la trazabilidad con repescto al codproveedor en las lineas
    Sql = "SELECT codtipom, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,numbultos, precioar, origpre, dtoline1, dtoline2, importel "
    If vEmpresa.TieneAnalitica Then
        Sql = Sql & ",codccost"
    Else
        Sql = Sql & ",codprovex,nomprove"
    End If
    'Mariela 09/07/2010
    'SQL = SQL & ",numlote "
    Sql = Sql & " FROM " & NomTablaLineas
    'traza
    If vEmpresa.TieneAnalitica = False Then
        Sql = Sql & " LEFT JOIN sprove on codprovex=codprove "
    End If
    
    If enlaza Then
        Sql = Sql & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        If EsHistorico Then Sql = Sql & " and fechaalb='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        Sql = Sql & " WHERE numalbar = -1"
    End If
    Sql = Sql & " Order by codtipom, numalbar, numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

        b = ((Modo = 2) Or (Modo = 5 And ModificaLineas = 0))
        'Insertar
        Toolbar1.Buttons(1).Enabled = (b Or Modo = 0) And Not EsHistorico
        Me.mnNuevo.Enabled = (b Or Modo = 0) And Not EsHistorico
        'Modificar
        Toolbar1.Buttons(2).Enabled = b And Not EsHistorico
        Me.mnModificar.Enabled = b And Not EsHistorico
        'eliminar
        Toolbar1.Buttons(3).Enabled = b And Not EsHistorico
        Me.mnEliminar.Enabled = b And Not EsHistorico
            
        b = (Modo = 2) And Not EsHistorico
        'Mantenimiento lineas
'        Toolbar1.Buttons(10).Enabled = b
'        Me.mnLineas.Enabled = b
'
'        'Nº Series
'        Toolbar1.Buttons(11).Enabled = b And Not EsHistorico
        
        'Generar Factura
        'DAVID###
        'Antes:
        'Toolbar1.Buttons(12).Enabled = b And (CodTipoMov = "ALM" Or CodTipoMov = "ART")
        'Ahora.  Cualquier tipo se puede generar la factura
        Toolbar5.Buttons(1).Enabled = b
        
        'Imprimir
        Toolbar1.Buttons(8).Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
        Me.mnImprimir.Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
'        Toolbar1.Buttons(14).Enabled = Toolbar1.Buttons(15).Enabled And vParamAplic.ArtPortes <> ""
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub

Private Sub CargarComboFacturacion()
'### Combo Tipo Facturación
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Factura Colectiva, 1-Factura x Albaran

    cboFacturacion.Clear
    cboFacturacion.AddItem "Factura Colectiva"
    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 0

    cboFacturacion.AddItem "Factura x Albaran"
    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 1

End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numalbar", "codtipom", Text1(30).Text, "T", , "numalbar", Text1(0).Text, "N")
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
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del cliente si es de varios
    If EsDeVarios Then
        'Si es cliente de varios actualizar datos cliente en tabla:sclvar
        MenError = "Modificando datos cliente varios"
        bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    End If
           
    If bol Then
        'Actualizar el campo fechamov (ult. movimiento) de la tabla de clientes (sclien)
        
        MenError = "Actualizando Fecha Movimiento del Cliente."
'[Monica]--
'        bol = ActualizarFecMovCliente
        
        MenError = "Error al actualizar el contador del Pedido."
    '    bol = vTipoMov.IncrementarContador("REG")
        vTipoMov.IncrementarContador (CodTipoMov)
    End If
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarOferta = True
        Else
            conn.RollbackTrans
            InsertarOferta = False
        End If
End Function


Private Sub LimpiarDatosCliente()
Dim i As Byte

    For i = 4 To 17
        Text1(i).Text = ""
    Next i
    Text2(12).Text = ""
    Text2(14).Text = ""
    Text2(17).Text = ""
    Me.cboFacturacion.ListIndex = -1
End Sub
    

Private Function EliminarLinea() As Boolean
Dim vCStock As CStock
Dim Sql As String
Dim b As Boolean
Dim ImpReciclado As Single



    EliminarLinea = False
    
    'Construir la SQL para eliminar la linea de la tabla "slialbcli"
    Sql = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Sql = Sql & " and numlinea=" & Data2.Recordset!numlinea
    
    
    
    'Inicilizar la clase para Actualizar los stocks
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "E") Then Exit Function
    
    On Error GoTo EEliminarLinea
    
    conn.BeginTrans
    conn.Execute Sql 'Eliminar linea
    b = vCStock.DevolverStock2
    Set vCStock = Nothing

    ' ---- [15/09/2009] (LAURA)
    'El articulo
'    ElArticulo = Data2.Recordset!codArtic
    ' ----
    
    If b Then
        'Ha borrado la linea y ha devuelvto correctamente el sctock
                   'Llegado aqui, si tiene Punto verde(tasa ecologica)
                'Y el cliente tiene tasa recliclado
                If ClienteConTasaReciclado Then
                    Sql = CStr(Data2.Recordset!codArtic)
                    If ArticuloConTasaReciclado(Sql, ImpReciclado) Then
                        
                       'Si el articulo siguiente es PV entoces lo updatearemos
                       Sql = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea"
                       'QUITO EL WHERE
                       Sql = Mid(Sql, 8)
                       NumRegElim = Val(DBLet(Data2.Recordset!numlinea, "N")) + 1
                       Sql = DevuelveDesdeBD(conAri, "codartic", "slialbcli", Sql, CStr(NumRegElim))
                       'En SQL tengo el codarti de la linea SIGUIENTE
                       'SI es punto verde de parametros, supondremos que esta vinculado con la linea que estamos modificando

'--[Monica]arituclo de reciclado pasa a ser reutilizado como articulo de gastos de administracion
'                       If SQL = vParamAplic.ArtReciclado Then
'
'                            SQL = "DELETE FROM " & NomTablaLineas
'                            'WHERE
'                            SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & NumRegElim
'                            conn.Execute SQL
'                      End If  'linea siguiente con codarti=puntoverde
                    End If  'articulo con reciclado
                End If ' de cliente con tasa reciclado
                
    End If


    'si la linea tenia numero de serie vaciar los campos correspondien al albaran venta
    Sql = "UPDATE sserie SET codclien=" & ValorNulo & ",codtipom=" & ValorNulo & ", fechavta=" & ValorNulo & ",numalbar=" & ValorNulo & ",numline1=" & ValorNulo
    Sql = Sql & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " and codtipom='" & CodTipoMov & "' and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
    Sql = Sql & " AND numalbar=" & Data1.Recordset!NumAlbar & " AND numline1=" & Data2.Recordset!numlinea
    conn.Execute Sql
    
EEliminarLinea:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Linea Albaran " & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        EliminarLinea = True
        
        ' ---- [15/09/2009] (LAURA)
'        DescuentosCantidad ElArticulo
        ' ----
        
        
    Else
        conn.RollbackTrans
         EliminarLinea = False
    End If
End Function


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
    On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = CodTipoMov
    vCStock.Trabajador = CLng(Text1(4).Text) 'guardamos el cliente del albaran
    vCStock.Documento = Text1(0).Text 'Nº Albaran
    vCStock.Fechamov = Text1(1).Text 'Fecha del Albaran
    
    '1=Insertar, 2=Modificar
    If ModificaLineas = 1 Or (ModificaLineas = 2 And TipoM = "S") Then
        vCStock.codArtic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        If ModificaLineas = 1 Then '1=Insertar
            vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text))
        Else '2=Modificar(Debe haber en stock la diferencia)
            If Data2.Recordset!codArtic = txtAux(1).Text Then
                vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text)) - Data2.Recordset!Cantidad
            Else
                vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text))
            End If
        End If
        vCStock.Importe = CCur(ComprobarCero(txtAux(8).Text))
    Else
        vCStock.codArtic = Data2.Recordset!codArtic
        vCStock.codAlmac = CInt(Data2.Recordset!codAlmac)
        vCStock.Cantidad = CSng(Data2.Recordset!Cantidad)
        vCStock.Importe = CCur(Data2.Recordset!ImporteL)
    End If
    If ModificaLineas = 1 Then
         vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    Else
        vCStock.LineaDocu = CInt(Data2.Recordset!numlinea)
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

    On Error GoTo ERestablecer
    
    ReestablecerStock = False
    b = True
    If Data2.Recordset.RecordCount > 0 Then
       Data2.Refresh
       Data2.Recordset.MoveFirst
    
       'Para cada linea de albaran reestablecer el stock
       While (Not Data2.Recordset.EOF) And b
           Set vCStock = New CStock
           If InicializarCStock(vCStock, "E", Data2.Recordset!numlinea) Then
               'Actualiza el stock en salmac y borra de smoval
               If Not vCStock.DevolverStock2() Then b = False
           Else
               b = False
           End If
           Data2.Recordset.MoveNext
           Set vCStock = Nothing
       Wend
    End If
    
ERestablecer:
    If Err.Number <> 0 Then b = False
    ReestablecerStock = b
End Function


Private Sub BotonImprimir(OpcionListado As Byte, EsInformePortes As Boolean)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim ImpresionDirecta As Boolean
    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Albaran para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    If (OpcionListado = 45) Then
        If EsInformePortes Then
            'Es el de portes
             indRPT = 34
        Else
            'ALBARANES
            If hcoCodTipoM = "ALZ" Then
                indRPT = 29   'Albaranes B
            ElseIf hcoCodTipoM = "ALR" Then
                indRPT = 36
            ElseIf hcoCodTipoM = "ALS" Then
                indRPT = 39
            Else
                If EsHistorico Then
                    indRPT = 11 'Hist. Albaranes clientes
                Else
                    indRPT = 10 'Albaran Clientes
                End If
            End If
        End If
    End If
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, ImpresionDirecta, pPdfRpt) Then Exit Sub
   
    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    cadParam = cadParam & "pCodUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
    
    'PORTES
    cadParam = cadParam & "vPortes=""" & vParamAplic.ArtPortes & """|"
    numParam = numParam + 1

'--[Monica]arituclo de reciclado pasa a ser reutilizado como articulo de gastos de administracion
'    'PUNTO VERDE
'    cadParam = cadParam & "PuntoVerde=""" & vParamAplic.ArtReciclado & """|"
'    numParam = numParam + 1
    
    'Si se imprimen importes y/o
'[Monica]--
'    devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", Text1(4).Text, "N")
    devuelve = ""
    If devuelve = "" Then devuelve = "0"
    ' 0 "Todo"
    ' 1 "Cantidad y Precio"
    ' 2 "Cantidad"
    cadParam = cadParam & "Albarcon=" & devuelve & "|"
    numParam = numParam + 1
    
    
    'Nombre fichero .rpt a Imprimir
    If Not ImpresionDirecta Then
        frmImprimir.NombreRPT = nomDocu
        frmImprimir.NombrePDF = pPdfRpt
    End If
        
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Albaran
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Cod Tipo Movimiento
        devuelve = "{" & NombreTabla & ".codtipom}='" & CodTipoMov & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Nº Albaran
        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        cadSelect = cadFormula
        
        If EsHistorico Then
            'El campo fecha tambien es clave primaria
            devuelve = Text1(1).Text
            devuelve = "{" & NombreTabla & ".fechaalb}=Date(" & Year(devuelve) & "," & Month(devuelve) & "," & Day(devuelve) & ")"
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            
            devuelve = "{" & NombreTabla & ".fechaalb}='" & Format(Text1(1).Text, FormatoFecha) & "'"
            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        End If
        
    End If
   
    '=========================================================================
    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
    'que se aplica a ese cliente
'[Monica]--
'    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(4).Text, "N")
    devuelve = "1"
    If devuelve <> "" Then
        cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
        numParam = numParam + 1
    End If

        
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    devuelve = NombreTabla & " INNER JOIN " & NomTablaLineas & " ON "
    devuelve = devuelve & NombreTabla & ".codtipom=" & NomTablaLineas & ".codtipom AND " & NombreTabla & ".numalbar= " & NomTablaLineas & ".numalbar "
    If EsHistorico Then devuelve = devuelve & " AND " & NombreTabla & ".fechaalb= " & NomTablaLineas & ".fechaalb "
    If Not HayRegParaInforme(devuelve, cadSelect) Then Exit Sub
    
    
    If ImpresionDirecta Then
        'Imrpimie directamente. Tipo 4tonda.  -----------
        If MsgBox("¿Imprimir el albarán?", vbQuestion + vbYesNo) = vbYes Then ImprimirDirectoAlb cadSelect
    Else
        With frmImprimir
            'Febrero 2010
            If indRPT = 34 Then
                .outTipoDocumento = 0
            Else
                .outTipoDocumento = 4
                .outClaveNombreArchiv = Text1(30).Text & Text1(0).Text
                .outCodigoCliProv = CLng(Text1(4).Text)
            End If
            
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = OpcionListado
            If indRPT = 34 Then
                .Titulo = "Portes albaran "
            Else
                .Titulo = "Albaran de Socio"
            End If
            .ConSubInforme = True
            .Show vbModal
        End With
    End If
End Sub


Private Sub MostrarNSeries(ByRef RSLineas As ADODB.Recordset, Optional Dif As String, Optional cadSel As String)
'Si los Nº de serie se introdujeron en ALBARAN COMPRAS se muestran
'los Nº de serie de los articulos comprados y se seleccionan tantos como cantidad de la linea
'Dif: si se ha modificado la cantidad pasamos la difencia con lo que habia
Dim Sql As String
Dim Campos As String

    On Error GoTo EMostrarNSeries

    If Text1(30).Text = "ART" Then
        Sql = MostrarNSeriesGnral(RSLineas, Campos, True)
    Else
        Sql = MostrarNSeriesGnral(RSLineas, Campos)
    End If
    
   If Sql <> "" Then
        Set frmMen = New frmMensajes
        frmMen.cadWHERE = Sql
        
        If Dif <> "" Then
            Sql = " WHERE (codtipom=" & DBSet(CodTipoMov, "T") & " and "
            Sql = Sql & " numalbar=" & Text1(0).Text & " and numline1=" & Data2.Recordset!numlinea & ")"
            frmMen.cadWHERE2 = Dif & "|" & Sql & "|"
        Else
            If cadSel <> "" Then
                'seleccionar lineas de nº serie de la factura a rectificar
                frmMen.cadWHERE2 = cadSel
            Else
                frmMen.cadWHERE2 = ""
            End If
        End If
        frmMen.OpcionMensaje = 4 'Nº Series Articulo
        frmMen.vCampos = Campos
        frmMen.Show vbModal
        Set frmMen = Nothing
    End If
    
EMostrarNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
Dim Sql As String

    On Error GoTo EPedirNSeries

        Sql = "El artículo tienen control de Nº de Serie." & vbCrLf & vbCrLf
        Sql = Sql & "Introduzca los Nº De Serie." & vbCrLf
        MsgBox Sql, vbInformation
        PedirNSeriesGnral RS, False
        
        Set frmNSerie = New frmRepCargarNSerie
        frmNSerie.DeVentas = True 'Se llama desde Alb. de Venta
        frmNSerie.Show vbModal
        Set frmNSerie = Nothing
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String
    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonMtoLineas 0, "Albaranes"
                BotonAnyadirLinea
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
    End If
    Set vTipoMov = Nothing
    Me.SSTab1.Tab = 0
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub ComprobarNSeriesLineas(numlinea As String)
'Al pasar de PEDIDO a ALBARAN
'control de Nº Series si hay algun articulo en las lineas de pedido que requiere Nº de serie
'Si NO se realiza control Nº series en compras pedirlos ahora
'Si se realiza control Nº Series en compras verificar que efectivamente estan introducidos
'y mostrarlos para seleccionarlos
Dim Sql As String
Dim RSLineas As ADODB.Recordset
Dim cadWHERE As String
Dim Dif As Single

    'Comprobar si el Articulo tiene control de Nº de Serie
    Sql = DevuelveDesdeBDNew(conAri, "sartic", "nseriesn", "codartic", txtAux(1).Text, "T")
    
    If Sql = "1" Then 'Hay NºSerie para el Articulo
        'si estamos insertando
        If Modo = 5 Then
            If ModificaLineas = 1 Then 'Insertar
                'Comprobar que la cantidad comprada es >0
                If ComprobarCero(txtAux(3).Text) <= 0 Then Exit Sub
            ElseIf ModificaLineas = 2 Then 'Modificar
                'si se ha modificado la cantidad, habrá que quitar algun nº serie
                'de los seleccionado o anyadir alguno mas
                Dif = CSng(txtAux(3).Text) - CSng(Data2.Recordset!Cantidad)
                If Dif = 0 Then Exit Sub
                If Text1(30).Text = "ART" Then Exit Sub
'                    Dif = CSng(Data2.Recordset!Cantidad) - CSng(txtAux(3).Text)
            End If
        End If
        
        cadWHERE = " WHERE codtipom=" & DBSet(CodTipoMov, "T") & " and "
        cadWHERE = cadWHERE & " numalbar=" & Text1(0).Text & " and numlinea=" & numlinea
    
        'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
        Sql = "SELECT slialbcli.codartic, sum(cantidad) as cantidad, numlinea "
        Sql = Sql & " FROM slialbcli INNER JOIN sartic on slialbcli.codartic=sartic.codartic "
        Sql = Sql & cadWHERE & " And nseriesn = 1 "
        Sql = Sql & " GROUP BY codartic ORDER BY Codartic "

        Set RSLineas = New ADODB.Recordset
        RSLineas.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Me.cmdAux(1).Tag = Text1(0).Text 'Num Albaran
        Me.cmdAux(0).Tag = numlinea 'Num Linea
        
        'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
        'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
        If Not vParamAplic.NumSeries And ModificaLineas = 1 Then
            PedirNSeries RSLineas
        Else 'Se realizo contro en COMPRAS, Mostramos los Nº y seleccionamos
            If ModificaLineas = 1 Then 'Insertando la linea
                'Comprobar que efectivamente estan en tabla sserie los NºSerie del Articulo
                ' y que no esten asignados ya a otro albaran de venta
                Sql = " select distinct count(numserie) from sserie where codartic=" & DBSet(txtAux(1).Text, "T") & " and (numalbar='' or isnull(numalbar))"
                '=== Laura 17/01/2007
                'y que no este asignados a una factura
                Sql = Sql & " and (numfactu='' or isnull(numfactu))"
                '===
                If RegistrosAListar(Sql) = 0 Then 'No hay Nº de Serie para elegir
                    PedirNSeries RSLineas
                Else
                    MostrarNSeries RSLineas
                End If
            ElseIf ModificaLineas = 2 Then
                Sql = " select distinct count(numserie) from sserie " & Replace(cadWHERE, "numlinea", "numline1")
                If RegistrosAListar(Sql) > 0 Then
                    MostrarNSeries RSLineas, CStr(Dif)
                End If
            End If
        End If

        RSLineas.Close
        Set RSLineas = Nothing
    End If
End Sub


Private Sub BotonNSeries()
Dim cadWHERE As String, Sql As String
Dim RSLineas As ADODB.Recordset

    If Me.Data1.Recordset!EsTicket Then
        MsgBox "Albaranes provenientes de Ticket no tienen control de Nº Serie.", vbInformation
        Exit Sub
    End If

    'Si es Albaran para Factura rectificativa (ART) o albaran para Factura rectificativa de cuota (ARC)
    If CodTipoMov = "ART" Or CodTipoMov = "ARC" Then
'      'Si es una Factura Venta(FAV) generada desde un ticket del TPV entonces
'      'no hay numseries
'      SQL = DevuelveDesdeBDNew(conAri, "scafac1", "codtipoa", "codtipom", Data1.Recordset!codtipmf, "T", , "numfactu", Data1.Recordset!NumFactu, "N", "fecfactu", Data1.Recordset!FecFactu, "F")
'      If SQL = "FTI" Then
'        MsgBox "Facturas provenientes de Ticket no tienen control de Nº Serie.", vbInformation
'        Exit Sub
'      Else
        Exit Sub
'      End If
    End If
    
    
    
    ModificaLineas = 4

    cadWHERE = " WHERE codtipom='" & Text1(30).Text & "'"
    cadWHERE = cadWHERE & " and numalbar=" & Text1(0).Text
    
    'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
    Sql = "SELECT numlinea, slialbcli.codartic, sum(cantidad) as cantidad "
    Sql = Sql & " FROM slialbcli INNER JOIN sartic on slialbcli.codartic=sartic.codartic "
    Sql = Sql & cadWHERE & " And nseriesn = 1 "
    
    'Pudioera ser que tuvieran un mismo articulo wen dos lineas, y por lo tanto
    'el articulo tendria numeros de sr asociados a distintas lineas
    'por lo tanto hay que agrupar por numlinea TB
    Sql = Sql & " GROUP BY codartic,numlinea ORDER BY Codartic "
    

    Set RSLineas = New ADODB.Recordset
    RSLineas.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RSLineas.EOF Then
        'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
        'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
        PedirNSeriesT RSLineas
    Else
        MsgBox "No hay ninguna linea de Articulo con Control de Nº Serie", vbInformation
    End If
    RSLineas.Close
    Set RSLineas = Nothing
    ModificaLineas = 0
End Sub


Private Sub PedirNSeriesT(ByRef RS As ADODB.Recordset)
Dim RSseries As ADODB.Recordset
Dim Sql As String
Dim linea As Integer

    On Error GoTo EPedirNSeries


        PedirNSeriesGnral RS, False
        RS.MoveFirst
        While Not RS.EOF
            linea = 0
            'Cargar los Nº de serie asignados
            Sql = "SELECT numserie, codartic,nummante FROM sserie "
            Sql = Sql & " WHERE codtipom='" & Text1(30).Text & "' and "
            Sql = Sql & "numalbar=" & Text1(0).Text
            Sql = Sql & " and numline1=" & RS!numlinea
            Sql = Sql & " ORDER BY codartic "
            Set RSseries = New ADODB.Recordset
            RSseries.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RSseries.EOF
                linea = linea + 1
                Sql = "UPDATE tmpnseries SET numserie=" & DBSet(RSseries!numSerie, "T")
                Sql = Sql & ", nummante = " & DBSet(RSseries!nummante, "T")
                Sql = Sql & " WHERE codartic=" & DBSet(RS!codArtic, "T")
                Sql = Sql & " and numlinealb=" & RS!numlinea
                Sql = Sql & " and numlinea=" & linea
                conn.Execute Sql
                RSseries.MoveNext
            Wend
            RS.MoveNext
        Wend
        RSseries.Close
        Set RSseries = Nothing
        Set frmNSerie = New frmRepCargarNSerie
        frmNSerie.DeVentas = True 'Se llama desde Alb. de Venta
        frmNSerie.NumAlb = Text1(0).Text
        frmNSerie.Show vbModal
        Set frmNSerie = Nothing
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub CargarNumSeries()
'Insertar un registro en la tabla "sserie" por cada uno de los
'Nº de Serie introducidos en la Tabla Temporal o actualizarlo
Dim RStmp As ADODB.Recordset
Dim Sql As String
Dim b As Boolean

    On Error GoTo ECargar
    
    conn.BeginTrans
    
    'Limpiar primero los Nº de serie asignados al ALV y luego volver a cargarlos
    Sql = "UPDATE sserie SET codtipom=" & ValorNulo & ", numalbar=" & ValorNulo & ", fechavta="
    Sql = Sql & ValorNulo & ", numline1=" & ValorNulo
    'Enero 2010
    'Tambien reestablezco los valores de tieneman y numeromantenimiento
     Sql = Sql & ", TieneMan=0 , NumMante= " & ValorNulo
    
    Sql = Sql & " WHERE codtipom=" & DBSet(Text1(30).Text, "T") & " and numalbar=" & Text1(0).Text & " AND year(fechavta)=" & Year(Text1(1).Text)
    conn.Execute Sql
    
    'Recuperar los Nº Serie de ese articulo cargados en la Temporal
    'Seleccionar los nº de serie cargados en la temporal: tmpnseries
    Sql = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
    Sql = Sql & " ORDER BY codartic, numlinealb, numlinea "
    Set RStmp = New ADODB.Recordset
    RStmp.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
    b = True
    While Not RStmp.EOF And b
        b = InsertarNSerie(RStmp!numSerie, RStmp!codArtic, RStmp!numlinealb, DBLet(RStmp!nummante, "T"))
        RStmp.MoveNext
    Wend
    RStmp.Close
    Set RStmp = Nothing
    
ECargar:
    If Err.Number <> 0 Then b = False
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
End Sub


Private Function InsertarNSerie(numSerie As String, codArtic As String, numlinea As String, nummante As String) As Boolean
'Inserta o Actualiza en la tabla sserie, si al pasar Pedido -> Albaran
'existen lineas con control de Nº Serie
Dim devuelve As String
Dim TieneMan As Boolean
Dim NumAlbar As String
Dim nSerie As CNumSerie
Dim b As Boolean

    On Error GoTo EInsertarNSerie


    'Enero 2010
    'AHora si tiene mantenimiento lo habra indicado en la introduccion de numero de serie
    '
    ''Comprobar que el cliente tiene mantenimientos en esa direc/dpto
    'TieneMan = "0"
    'devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    ''El cliente tiene Mantenimientos
    'If devuelve <> "" Then TieneMan = "1"
    nummante = Trim(nummante)
    TieneMan = nummante <> ""



    Set nSerie = New CNumSerie
    nSerie.numSerie = numSerie
    nSerie.Articulo = codArtic
    
    nSerie.Cliente = CLng(Text1(4).Text)
    nSerie.DirDpto = Text1(12).Text
    nSerie.conMante = TieneMan
    nSerie.tipoMov = CodTipoMov
    nSerie.FechaVta = Text1(1).Text
    nSerie.NumAlbaran = Text1(0).Text
    nSerie.NumLinAlb = numlinea
    nSerie.ObtenFechaFinGarantia codArtic, Text1(1).Text
    nSerie.nummante = nummante   'Si ha indicado el numero de mantenimiento
    
    
    

    
    'Comprobar si existe en la tabla sserie
     NumAlbar = "numalbar" 'Nº albaran de Venta
     devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", numSerie, "T", NumAlbar, "codartic", codArtic, "T")
     If devuelve <> "" Then 'EXISTE en tabla sserie
        If NumAlbar = "" Then b = nSerie.ActualizarNumSerie(True)
     Else
        b = nSerie.InsertarNumSerie
    End If
    InsertarNSerie = True
    Set nSerie = Nothing
    
EInsertarNSerie:
    If Err.Number <> 0 Then b = False
    If b Then
        InsertarNSerie = True
    Else
        InsertarNSerie = False
    End If
End Function




Private Sub PosicionarDataTrasEliminar()
Dim HayDatos As Boolean
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    HayDatos = SituarDataTrasEliminar(Data1, NumRegElim)
    If HayDatos Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            Data1.Recordset.MoveLast
            If Data1.Recordset.EOF Then HayDatos = False
        End If
    End If
    If HayDatos Then
        PonerCampos
    Else
        LimpiarCampos
        LimpiarDataGrids
        PonerModo 0
    End If
End Sub


Private Sub PonerDatosCliente(CodClien As String, Optional nifClien As String)
'Dim vCliente As CCliente
Dim Observaciones As String
Dim Sql As String
    
    On Error GoTo EPonerDatos
    
    If CodClien = "" Then
        LimpiarDatosCliente
        Exit Sub
    End If
    
    Sql = "select * from scliente where codclien=" & Text1(4).Text
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Set vCliente = New CCliente
    
    'si se ha modificado el cliente volver a cargar los datos
    If Not miRsAux.EOF Then
            
            Text1(4).Text = miRsAux!CodClien
            FormateaCampo Text1(4)
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = miRsAux!nomclien
                Text1(8).Text = miRsAux!domclien
                Text1(9).Text = miRsAux!codpobla
                Text1(10).Text = miRsAux!pobclien
                Text1(11).Text = miRsAux!proclien
                Text1(6).Text = miRsAux!nifClien
                Text1(7).Text = DBLet(miRsAux!telclie1, "T")
            End If

            If Modo = 3 Or Modo = 4 Then 'insertar
                Text1(14).Text = vParamAplic.PorDefecto_Zona
                Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
                Text1(15).Text = 0
                Text1(16).Text = 0
                Text1(17).Text = vParamAplic.PorDefecto_Agente
                Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
                Text1(34).Text = 0
                Text1(29).Text = vParamAplic.PorDefecto_Envio
                Text2(29).Text = PonerNombreDeCod(Text1(29), conAri, "senvio", "nomenvio")
            End If
            Observaciones = DBLet(miRsAux!observac)
            If Len(Trim(Observaciones)) <> 0 Then
                MsgBox Observaciones, vbInformation, "Observaciones del Cliente"
            End If
    Else
        LimpiarDatosCliente
    End If
    
    If Not miRsAux.EOF Then
    End If
EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Socio", Err.Description
End Sub


Private Sub PonerDatosClienteVario(nifClien As String)
Dim vCliente As CCliente
Dim b As Boolean
   
    If nifClien = "" Then Exit Sub
   
    Set vCliente = New CCliente
    b = vCliente.LeerDatosCliVario(nifClien)
    If b Then Text1(5).Text = vCliente.Nombre         'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
'    If Not b Then PonerFoco Text1(6)
    Set vCliente = Nothing
End Sub


Private Sub BloquearDatosCliente(bol As Boolean)
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol
        Me.imgBuscar(1).Enabled = bol
        Me.imgBuscar(6).Enabled = bol
        
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


Private Function ActualizarFecMovCliente() As Boolean
Dim vCliente As CCliente
Dim b As Boolean

    On Error GoTo EActFecha

    ActualizarFecMovCliente = False
    Set vCliente = New CCliente
    vCliente.Codigo = Text1(4).Text
    b = vCliente.ActualizaUltFecMovim(Text1(1).Text)
    Set vCliente = Nothing
    
EActFecha:
    If Err.Number <> 0 Then b = False
    ActualizarFecMovCliente = b
End Function


Private Sub CalcularDatosFactura()
Dim i As Integer
Dim cadWHERE As String, Sql As String
Dim vFactu As CFactura
Dim EsAnterior As Boolean

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 33 To 56
         Text3(i).Text = ""
    Next i

    'Comprobar que hay lineas de albaran para calcular totales
    cadWHERE = ObtenerWhereCP(False)
    Sql = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadWHERE, NombreTabla, NomTablaLineas)
    If RegistrosAListar(Sql) = 0 Then Exit Sub
    
    Set vFactu = New CFactura
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(15).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(16).Text))
    vFactu.Cliente = Text1(4).Text
    If hcoCodTipoM = "ALZ" Then vFactu.codtipom = "ALZ"
    EsAnterior = CDate(Text1(1).Text) < CDate(vParamAplic.FecCambioIva)
    If vFactu.CalcularDatosFactura(False, cadWHERE, NombreTabla, NomTablaLineas, EsAnterior) Then
        Text3(33).Text = vFactu.BrutoFac
        Text3(34).Text = vFactu.ImpPPago
        Text3(35).Text = vFactu.ImpGnral
        Text3(36).Text = vFactu.BaseImp
        Text3(37).Text = vFactu.TipoIVA1
        Text3(38).Text = vFactu.TipoIVA2
        Text3(39).Text = vFactu.TipoIVA3
        Text3(40).Text = vFactu.PorceIVA1
        Text3(41).Text = vFactu.PorceIVA2
        Text3(42).Text = vFactu.PorceIVA3
        Text3(43).Text = vFactu.BaseIVA1
        Text3(44).Text = vFactu.BaseIVA2
        Text3(45).Text = vFactu.BaseIVA3
        Text3(46).Text = vFactu.ImpIVA1
        Text3(47).Text = vFactu.ImpIVA2
        Text3(48).Text = vFactu.ImpIVA3
        Text3(55).Text = vFactu.TotalFac
        Text3(56).Text = vFactu.BaseImp
        
        'Recargos de equivalencia
        Text3(49).Text = vFactu.PorceIVA1RE
        Text3(50).Text = vFactu.PorceIVA2RE
        Text3(51).Text = vFactu.PorceIVA3RE
        Text3(52).Text = vFactu.ImpIVA1RE
        Text3(53).Text = vFactu.ImpIVA2RE
        Text3(54).Text = vFactu.ImpIVA3RE
        
        FormatoDatosTotales
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
End Sub


Private Function FormatoDatosTotales()
Dim i As Byte

    For i = 33 To 36
        Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
    
    For i = 49 To 54
        Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
    'Desglose B.Imponible por IVA
    For i = 43 To 45
        If Text3(i).Text <> "" Then
             If CSng(Text3(i).Text) = 0 Then
                Text3(i).Text = QuitarCero(Text3(i).Text)
                Text3(i - 3).Text = QuitarCero(Text3(i - 3).Text)
                Text3(i - 6).Text = QuitarCero(Text3(i - 6).Text)
                '---- Laura: Modificado 27/09/2006
'                Text3(i + 3).Text = QuitarCero(Text3(i).Text)
                Text3(i + 3).Text = QuitarCero(Text3(i + 3).Text)
                '----
            Else
                Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
                Text3(i - 3) = Format(Text3(i - 3).Text, FormatoDescuento)
                Text3(i + 3).Text = Format(Text3(i + 3).Text, FormatoImporte)
            End If
        End If
    Next i
    
    'TOTALES
    Text3(55).Text = Format(Text3(55).Text, FormatoImporte)
    Text3(56).Text = Format(Text3(56).Text, FormatoImporte)
End Function



Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.Codigo = Text1(4).Text
    'si existe el departamento para el cliente
    If vClien.DptoCliente(Text1(12).Text, NomDpto) Then
        Text2(12).Text = NomDpto
        PonerDptoEnCliente = True
    Else
        PonerDptoEnCliente = False
    End If
    Set vClien = Nothing
End Function


Private Sub ComprobarRefObligatoria()
Dim vClien As CCliente

    Set vClien = New CCliente
    vClien.Codigo = Text1(4).Text
    If vClien.TieneRefObligatoria(Text1(13).Text) Then
        If Text1(13).Text = "" Then PonerFoco Text1(13)
    End If
    Set vClien = Nothing
End Sub



 Private Sub InsertarLineasFactu(cadWHERE)
'cadSerie = "INSERT INTO slialbcli(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
'cadSerie = cadSerie & " SELECT '" & Text1(30).Text & "' as codtipom," & Text1(0).Text & " as numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre FROM slifac WHERE " & CadenaSeleccion
 Dim RS As ADODB.Recordset
 Dim Sql As String
 Dim i As Integer
 Dim cadI As String
 Dim numlin As String
 
    On Error GoTo EInsFactu
    Screen.MousePointer = vbHourglass
    
    If cadWHERE <> "" Then
        'Obtenemos el numero de linea a insertar
'        SQL = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'        SQL = SugerirCodigoSiguienteStr("slialbcli", "numlinea", SQL)
'        i = Int(SQL)
    
        cadI = ""
    
        Sql = "SELECT * FROM slifac WHERE " & cadWHERE
        
        If hcoCodTipoM = "ARP" Then Sql = Replace(Sql, "slifac", "slifaccli")
    
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            txtAux(0).Text = RS!codAlmac
            txtAux(1).Text = RS!codArtic
            txtAux(2).Text = RS!NomArtic
            Text2(16).Text = DBLet(RS!Ampliaci, "T")
'            Text2(9).Text = DBLet(RS!nomprove, "T")
            txtAux(3).Text = CStr(RS!Cantidad * -1)
            txtAux(4).Text = RS!precioar
            txtAux(5).Text = DBLet(RS!origpre, "T")
            txtAux(6).Text = RS!dtoline1
            txtAux(7).Text = RS!dtoline2
            txtAux(8).Text = CStr(RS!ImporteL * -1)
            
            ' ---- [21/10/2009] [LAURA] : se añade el centro de coste
            If Not vEmpresa.TieneAnalitica Then
                txtAux(9).Text = DBLet(RS!codprovex, "N")
                Text2(9).Text = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", CStr(RS!codprovex), "N")
            Else
                txtAux(9).Text = DBLet(RS!CodCCost, "T")
                Me.Text2(9).Text = PonerNombreCCoste(Me.txtAux(9))
            End If
            
            'para no tener que traer ahora el proveedor pongo en txt(10) un texto
'            txtAux(10).Text = "*"
'            Text2(9).Text = "*"
            
            'numbultos
            txtAux(10).Text = CStr(RS!numbultos * -1)
            'numlote
            txtAux(11).Text = DBLet(RS!numlote, "T")
            
            If InsertarLinea(numlin, True) Then
            
            End If
            
'            SQL = "('" & Text1(30).Text & "'," & Text1(0).Text & "," & i & ","  'codtipoa,numalbar,numlinea
'            SQL = SQL & DBSet(RS!codAlmac, "N") & "," & DBSet(RS!codArtic, "T") & "," & DBSet(RS!NomArtic, "T") & "," & DBSet(RS!ampliaci, "T") & ","
'            SQL = SQL & DBSet(RS!cantidad * -1, "N") & "," & DBSet(RS!precioar, "N") & "," & DBSet(RS!dtoline1, "N") & "," & DBSet(RS!dtoline2, "N") & ","
'            SQL = SQL & DBSet(RS!ImporteL * -1, "N") & "," & DBSet(RS!origpre, "T") & ")"
'            If cadI = "" Then
'                cadI = SQL
'            Else
'                cadI = cadI & "," & SQL
'            End If
'            i = i + 1
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        
        CalcularDatosFactura
        
'        If cadI <> "" Then
'            SQL = "INSERT INTO slialbcli(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) VALUES "
'            SQL = SQL & cadI
'            Conn.Execute SQL
'        End If
    End If
    Screen.MousePointer = vbDefault
    
EInsFactu:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Lineas Factura", Err.Description
    End If
End Sub


'Private Sub InsertarLineasFactu_old(cadWHERE)
''cadSerie = "INSERT INTO slialbcli(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
''cadSerie = cadSerie & " SELECT '" & Text1(30).Text & "' as codtipom," & Text1(0).Text & " as numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre FROM slifac WHERE " & CadenaSeleccion
' Dim RS As ADODB.Recordset
' Dim SQL As String
' Dim i As Integer
' Dim cadI As String
'
'    On Error GoTo EInsFactu
'    Screen.MousePointer = vbHourglass
'
'    If cadWHERE <> "" Then
'        'Obtenemos el numero de linea a insertar
'        SQL = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'        SQL = SugerirCodigoSiguienteStr("slialbcli", "numlinea", SQL)
'        i = Int(SQL)
'
'        cadI = ""
'
'        SQL = "SELECT * FROM slifac WHERE " & cadWHERE
'
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        While Not RS.EOF
'            SQL = "('" & Text1(30).Text & "'," & Text1(0).Text & "," & i & ","  'codtipoa,numalbar,numlinea
'            SQL = SQL & DBSet(RS!codAlmac, "N") & "," & DBSet(RS!codArtic, "T") & "," & DBSet(RS!NomArtic, "T") & "," & DBSet(RS!ampliaci, "T") & ","
'            SQL = SQL & DBSet(RS!Cantidad * -1, "N") & "," & DBSet(RS!precioar, "N") & "," & DBSet(RS!dtoline1, "N") & "," & DBSet(RS!dtoline2, "N") & ","
'            SQL = SQL & DBSet(RS!ImporteL * -1, "N") & "," & DBSet(RS!origpre, "T") & ")"
'            If cadI = "" Then
'                cadI = SQL
'            Else
'                cadI = cadI & "," & SQL
'            End If
'            i = i + 1
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'
'        If cadI <> "" Then
'            SQL = "INSERT INTO slialbcli(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) VALUES "
'            SQL = SQL & cadI
'            Conn.Execute SQL
'        End If
'    End If
'    Screen.MousePointer = vbDefault
'
'EInsFactu:
'    If Err.Number <> 0 Then
'        Screen.MousePointer = vbDefault
'        MuestraError Err.Number, "Lineas Factura", Err.Description
'    End If
'End Sub



Private Function AsignarNumSeriesAlbVenta(cadSel As String) As Boolean
Dim i As Integer
Dim cant As Integer
Dim cadSerie As String
Dim nSerie As CNumSerie
Dim devuelve As String
Dim b As Boolean
    
    'Para cada valor empipado actualizar la tabla sserie
    
    
    cant = CInt(ComprobarCero(txtAux(3).Text))
    
    On Error GoTo ErrorNSerie
    conn.BeginTrans
        
    If ModificaLineas = 2 Then 'Venimos de modificar la cantidad de una linea
        'Borramos los numeros de serie que tenia asignada la linea del albaran
        Set nSerie = New CNumSerie
        nSerie.tipoMov = CodTipoMov
        nSerie.NumAlbaran = Text1(0).Text
        nSerie.NumLinAlb = ComprobarCero(Me.cmdAux(0).Tag)
        b = nSerie.BorrarNumSeriesAlbVta
        Set nSerie = Nothing
    Else
        b = True
    End If
        
    If b Then
        Set nSerie = New CNumSerie
        nSerie.Articulo = txtAux(1).Text
        nSerie.Cliente = CLng(Text1(4).Text)
        nSerie.DirDpto = Text1(12).Text
        nSerie.tipoMov = CodTipoMov
        nSerie.FechaVta = Text1(1).Text
        If nSerie.FechaVta <> "" Then
            devuelve = DevuelveDesdeBDNew(conAri, "sartic", "garantia", "codartic", txtAux(1).Text, "T")
            nSerie.FinGarantia = CStr(CDate(nSerie.FechaVta) + CInt(ComprobarCero(devuelve)))
        End If
        nSerie.NumAlbaran = Text1(0).Text
        nSerie.NumLinAlb = ComprobarCero(Me.cmdAux(0).Tag)
                
        For i = 1 To cant
            cadSerie = RecuperaValor(cadSel, i + 1)
            If cadSerie <> "" Then
                nSerie.numSerie = cadSerie
                If nSerie.ActualizarNumSerie(True) = False And b Then b = False
            End If
        Next i
        Set nSerie = Nothing
    End If
ErrorNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizar tabla Nº Series", Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    AsignarNumSeriesAlbVenta = b
End Function




Private Sub DevolverNumSeries()
Dim Sql As String
Dim cadWHERE As String
Dim RS As ADODB.Recordset

    On Error GoTo EDevNumSerie
        
    cadWHERE = ObtenerWhereCP(True)
    Sql = "select slialbcli.codartic,abs(cantidad) as cantidad,numlinea"
    Sql = Sql & " from slialbcli inner join scaalbcli on slialbcli.codtipom=scaalbcli.codtipom and scaalbcli.numalbar=slialbcli.numalbar "
    Sql = Sql & " inner join sserie on slialbcli.codartic=sserie.codartic and sserie.numfactu=scaalbcli.numfactu and sserie.codclien=scaalbcli.codclien "
    '-- LAURA: 02/07/2007
'    SQL = SQL & " inner join scafac1 on scafac1.codtipom=scaalbcli.codtipmf and scafac1.numfactu=scaalbcli.numfactu and scafac1.fecfactu=scaalbcli.fecfactu "
'    SQL = SQL & " inner join sserie on scafac1.codtipoa=sserie.codtipom and scafac1.numalbar=sserie.numalbar and scafac1.fechaalb=sserie.fechavta "
    Sql = Sql & cadWHERE & " and scaalbcli.numfactu=" & CStr(Me.Data1.Recordset!NumFactu)
'    If Me.Data1.Recordset!codtipmf = "FAV" Then SQL = SQL & " AND codtipom='ALV'"
    '--

    
    
    
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Hay articulos con nº de serie en las lineas del albaran rectificativo
    'que hay que quitar de los nº de serie que tenia asignados
    'estamos devolviendo nº serie y pedimos los que vamos a devolver y a estos
    'le limpiamos los campos de venta de la tabla sserie
    If Not RS.EOF Then
        Sql = "select sserie.numserie, sserie.codartic, sartic.nomartic"
        Sql = Sql & " from slialbcli inner join scaalbcli on slialbcli.codtipom=scaalbcli.codtipom and scaalbcli.numalbar=slialbcli.numalbar "
        '-- LAURA: 02/07/2007
'        SQL = SQL & " inner join scafac1 on scafac1.codtipom=scaalbcli.codtipmf and scafac1.numfactu=scaalbcli.numfactu and scafac1.fecfactu=scaalbcli.fecfactu "
'        SQL = SQL & " inner join sserie on scafac1.codtipoa=sserie.codtipom and scafac1.numalbar=sserie.numalbar and scafac1.fechaalb=sserie.fechavta "
        Sql = Sql & " inner join sserie on slialbcli.codartic=sserie.codartic and sserie.numfactu=scaalbcli.numfactu  and sserie.codclien=scaalbcli.codclien "
        '--
        Sql = Sql & " inner join sartic on sserie.codartic=sartic.codartic "
        Sql = Sql & cadWHERE & " and scaalbcli.numfactu=" & CStr(Me.Data1.Recordset!NumFactu)
    
        MostrarNSeries RS, , Sql
    End If
    RS.Close
    Set RS = Nothing
    
EDevNumSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizando Nº Serie.", Err.Description
    End If
End Sub




Private Function QuitarNumSeriesAlbVenta(cadSel As String) As Boolean
Dim i As Integer
Dim numSerie As String
Dim codArtic As String
Dim nSerie As CNumSerie
Dim Grupo As String
Dim b As Boolean
    
    'Para cada valor empipado actualizar la tabla sserie
   
    On Error GoTo ErrorNSerie
    
    b = True
    While cadSel <> ""
        i = InStr(1, cadSel, "·")
        If i > 0 Then
            Grupo = Mid(cadSel, 1, i - 1)
            cadSel = Mid(cadSel, i + 1, Len(cadSel))
            If Grupo <> "" Then
                codArtic = RecuperaValor(Grupo, 1)
                numSerie = RecuperaValor(Grupo, 2)
                
                Set nSerie = New CNumSerie
                nSerie.numSerie = numSerie
                nSerie.Articulo = codArtic
                b = b And nSerie.ActualizarNumSerie(True)
                Set nSerie = Nothing
            End If
        End If
    Wend
   
ErrorNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizar tabla Nº Series", Err.Description
        Set nSerie = Nothing
        b = False
    End If
    QuitarNumSeriesAlbVenta = b
End Function


Private Sub BotonRecuperarFactura()
'Genera una factura a partir del Albaran de Mostrador
'pero sin coger contador de factura lo pide en un form

    'Comprobar que esta marcada para facturar
    If Me.chkFacturar.Value = 1 Then
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        'Facturacion de Albaran de Mostrador
        frmListadoPed.CodClien = CodTipoMov  'utilizamos esta vble para pasarle el tipo de movimiento
        frmListadoPed.NumCod = Text1(0).Text  'utilizamos esta vble para pasarle el nº albaran
        frmListadoPed.EstaRecupFact = Me.RecuperarFactu
        AbrirListadoPed (222)
        
        PosicionarDataTrasEliminar
    Else
        MsgBox "El Albaran no esta marcado para facturar", vbInformation
    End If
End Sub


Private Sub MarcarAlbaranes()

        'Lanzara un desde hasta y los marcara
        frmListado.NumCod = hcoCodTipoM
        CadenaDesdeOtroForm = ""
        AbrirListado 82
        If CadenaDesdeOtroForm = "OK" Then
            'OK. Cambiadas las marcas. Refrescamos y situamos
            Screen.MousePointer = vbHourglass
            DoEvents
            PonerCadenaBusqueda
            PosicionarData
            Screen.MousePointer = vbDefault
        End If
        
End Sub

'FALTA###
'No se el porque del importe
Private Function SumaKilosLineas(Optional ImporteL As Currency) As Currency
Dim C As String
    On Error GoTo ESumaKilosLineas
    SumaKilosLineas = 0
    Set miRsAux = New ADODB.Recordset
    C = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    C = C & " AND slialbcli.codartic=sartic.codartic"
    C = C & " AND slialbcli.codartic <> " & DBSet(vParamAplic.ArtPortes, "T")
    C = "select sum(cantidad*pesoarti),sum(importel) from slialbcli,sartic " & C
    
    
    'El enlzace
    
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        SumaKilosLineas = DBLet(miRsAux.Fields(0), "N")
        ImporteL = DBLet(miRsAux.Fields(1), "N")
    End If
    miRsAux.Close
    
    
    'Fijo la zona y la ruta del cliente
    
    RutaCliente = -1
    ZonaCliente = -1
    C = "Select codzonas,codrutas from sclien where codclien = " & Val(Text1(4).Text)
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        ZonaCliente = DBLet(miRsAux!codzonas, "N")
        RutaCliente = DBLet(miRsAux!codrutas, "N")
    End If
    miRsAux.Close
    
    
ESumaKilosLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function

'Si devuelve cero nada
'si devuelve >0 marcara la linea de portes
Private Function HacerAccionesPortes() As Integer
Dim ImporteLineas As Currency
Dim KilosAhora As Currency
Dim C As String
Dim CodEnvio As Integer
Dim PrecioKilo As Currency
Dim DtoPorte As Currency
Dim DesdeKilo As Currency
Dim ImporteL_Portes As Currency
Dim codCCoste As String 'centro de coste

    HacerAccionesPortes = 0
    KilosAhora = SumaKilosLineas(ImporteLineas)
    
    
    ' Si no cambia los kilos me salgo
    '-----------------------------------------------
    'If KilosAhora = KilosAnteriores Then Exit Function
    If Data2.Recordset.EOF Then Exit Function
    
    If MsgBox("Desea recalcular los portes?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    
    
    Set miRsAux = New ADODB.Recordset
    
    
    If ZonaCliente > 0 Then
        'Ha encontrado la zona /ruta. Miro en sportes
        C = "select sporte.codenvio,nomenvio,PrecioKg,desdekgs from sporte,senvio where sporte.codenvio=senvio.codenvio "
        C = C & " AND codcentr = " & ZonaCliente
        'Los kilos  hastakgs
        C = C & " AND desdekgs <= " & TransformaComasPuntos(CStr(KilosAhora))
        C = C & " AND hastakgs >= " & TransformaComasPuntos(CStr(KilosAhora))
        C = C & " group by sporte.codenvio"
        miRsAux.Open C, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        NumRegElim = 0
        CodEnvio = 0
        If Not miRsAux.EOF Then
            'Por si acaso hay mas de uno
            CadenaDesdeOtroForm = ""
            While Not miRsAux.EOF
                CodEnvio = miRsAux!CodEnvio
                PrecioKilo = miRsAux!preciokg
                DesdeKilo = DBLet(miRsAux!DesdeKgs, "N")
                NumRegElim = NumRegElim + 1
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & miRsAux!CodEnvio & "<" & miRsAux!nomenvio & "<" & miRsAux!preciokg & "<" & DBLet(miRsAux!DesdeKgs, "N") & "|"
                miRsAux.MoveNext
            Wend
            
            
            If NumRegElim > 1 Then
                'Mostraremos un form para que seleccione la opcion correspondiente
                frmVarios.Opcion = 3
                frmVarios.Show vbModal
                If CadenaDesdeOtroForm <> "" Then
                    C = RecuperaValor(CadenaDesdeOtroForm, 1)
                    CodEnvio = Val(C)
                    
                    C = RecuperaValor(CadenaDesdeOtroForm, 3)
                    PrecioKilo = CCur(C)
                    
                    DesdeKilo = CCur(RecuperaValor(CadenaDesdeOtroForm, 4))
                End If
            Else
                    CadenaDesdeOtroForm = Replace(CadenaDesdeOtroForm, "<", "|")
                    CadenaDesdeOtroForm = RecuperaValor(CadenaDesdeOtroForm, 2)
            
            End If
            
        End If
        miRsAux.Close
        
        
        'Dto en portes
        DtoPorte = 0
        ImporteL_Portes = 0
        If RutaCliente = 1 Or RutaCliente = 3 Or RutaCliente = 4 Then DtoPorte = vParamAplic.AbonoKilos
        If RutaCliente = 1 Or RutaCliente = 2 Then PrecioKilo = 0
        If RutaCliente = 4 And ImporteLineas < vParamAplic.ImporteMinimo Then 'importe pedido menor que importe minimo todo a cero(preciokilo, dtokilo)
               PrecioKilo = 0
               DtoPorte = 0
               ImporteL_Portes = 0
        Else
            If RutaCliente = 4 Then ImporteL_Portes = PrecioKilo
        End If
        
        If DesdeKilo = 1 Then
            If RutaCliente <> 4 Then
                ImporteL_Portes = PrecioKilo
                KilosAhora = 1
            End If
        Else
            ImporteL_Portes = (PrecioKilo - DtoPorte) * KilosAhora
        End If
        If RutaCliente <> 1 And ImporteL_Portes < 0 Then ImporteL_Portes = 0 'masl 090709
        
        'Ahora compruebo si tiene la linea de portes para aplicarle el importe
        C = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        C = "Select numlinea,codccost from slialbcli " & C & " and codartic ='" & vParamAplic.ArtPortes & "'"
        miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        If Not miRsAux.EOF Then
            NumRegElim = DBLet(miRsAux!numlinea, "N")
            codCCoste = DBLet(miRsAux!CodCCost, "T")
            
        Else
            codCCoste = ""
            If vEmpresa.TieneAnalitica Then
                If vParamAplic.ModoAnalitica = 0 Then
                    'Del trabajador
                    codCCoste = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", Text1(3).Text)
                    If codCCoste = "" Then MsgBox "Trabajador sin centro de coste", vbExclamation
                        
                End If
                If vParamAplic.ModoAnalitica <> 0 Or codCCoste = "" Then
                    codCCoste = DevuelveDesdeBD(conAri, "codccost", "sartic,sfamia", "sartic.codfamia=sfamia.codfamia AND codartic", vParamAplic.ArtPortes, "T")
                    If codCCoste = "" Then MsgBox "Familia sin centro de coste", vbExclamation
                End If
            End If
                
        End If
        miRsAux.Close
        
        
        'SI ya existe la borro, para que siempre aparezca al final
        If NumRegElim > 0 Then
            C = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
            C = C & " AND numlinea = " & NumRegElim
            C = "DELETE FROM  slialbcli " & C
            conn.Execute C
            Espera 0.1
            
        
        End If
        
     'If RutaCliente <> 1 And ImporteL_Portes < 0 Then ImporteL_Portes = 0 masl 090709
        
        
            'Si el precio es mayor k cero entonces SI pongo la linea
            C = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
            C = "select codtipom,codalmac,max(numlinea) from slialbcli " & C
            C = C & " GROUP BY codalmac"
            miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "NO deberia haberse producido", vbExclamation
                Exit Function
            End If
            NumRegElim = miRsAux.Fields(2) + 1
            HacerAccionesPortes = NumRegElim
    '            SQL = "INSERT INTO " & NomTablaLineas
    '            SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre,codprovex) "
    '            SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & NumRegElim & ", "
            
            C = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtPortes, "T")
            C = DevNombreSQL(C)
            C = miRsAux!codAlmac & ",'" & vParamAplic.ArtPortes & "','" & C & "','"
            
            C = "INSERT INTO " & NomTablaLineas & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre,codprovex,codccost)" & _
                "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & NumRegElim & ", " & C
            
            
            'Amplicacion
            C = C & CadenaDesdeOtroForm & "',"
            
            
            If RutaCliente <> 1 And RutaCliente <> 3 And RutaCliente <> 4 Then  'masl 090709
                        'Cantidad, precioar dto1 dto2
                        C = C & TransformaComasPuntos(CStr(KilosAhora)) & "," & TransformaComasPuntos(CStr(PrecioKilo))
                        C = C & "," & TransformaComasPuntos(CStr(DtoPorte)) & ",0,"
                                           
                                                                       Else 'masl 090709
                        C = C & TransformaComasPuntos(CStr(KilosAhora)) & "," & TransformaComasPuntos(CStr(DtoPorte * (-1)))
                        C = C & ",0" & ",0,"
                                                                       
            End If  'masl 090709
            
            'importel
            C = C & TransformaComasPuntos(CStr(ImporteL_Portes))
            
            'origpre,codprovex,codccost
            C = C & ",'M',0," & DBSet(codCCoste, "T") & ")"
        
        
        
            'Noviembre 2009.    Enero 2010.  SIEMPRE hay que meter la linea de portes
            'If ImporteL_Portes <> 0 Then conn.Execute C
            conn.Execute C
        
    End If
            
End Function


'Para obtener los dtos por cantidad lo que hace es a partir de un
'subtring del articulo(poscion 3 a 6) va a sdesca con la suma de la cantidad
'si en sdesca y dentro de los desde /hasta cantidad encuentra un dto lo aplica


Private Sub DescuentosCantidad(Articulo As String)
Dim Cad As String
Dim R As ADODB.Recordset
Dim NuevoDto As Currency
Dim Importe As Currency
Dim bAct As Boolean

    On Error GoTo EDescuentosCantidad
    
    If Not vParamAplic.DtoxCantidad Then Exit Sub ' ---- [14/09/2009] (LAURA)
     
    If MsgBox("¿Desea recalcular los descuentos por cantidad?", vbQuestion + vbYesNo) = vbYes Then    'masl 140909
    
        
        'Si no  tenemos portes, ni nos pasamos
    '    If vParamAplic.ArtPortes = "" Then Exit Sub
        
        
        Espera 0.2
        Set miRsAux = New ADODB.Recordset
        Set R = New ADODB.Recordset
        
        'variable articulo:
        'Si tiene valor es para no tener que recalcular todos los valores del albaran, solo los
        ' del substring() del articulo que acabamos de insertar/actualizar o eliminar
        ' Si no lleva nada recalcular los dtos para todas la lineas
        Cad = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        Cad = "select substring(codartic,3,4) raiz,sum(cantidad) suma from slialbcli " & Cad
        If Articulo <> "" Then Cad = Cad & " AND substring(codartic,3,4)= '" & Mid(Articulo, 3, 4) & "'"
        'Y origen PRECIO no es precio especial
        Cad = Cad & " AND origpre <> 'E'"
        Cad = Cad & " group by 1"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
                Cad = TransformaComasPuntos(CStr(miRsAux!Suma))
                Cad = "select * from sdesca where desdecan <=" & Cad & " and " & Cad & " <= hastacan and envagran = '"
                Cad = Cad & miRsAux!raiz & "'"
                R.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Cad = ""
                If Not R.EOF Then Cad = R!dtolinea
                R.Close
                
                
                If Cad <> "" Then
                    'OK tiene nuevo descuento
                    NuevoDto = CCur(Cad)
                    
                    'Cojo los articulos del albaran y le meto el dto
                    Cad = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
                    Cad = "select * from slialbcli " & Cad
                    '                                 a partir de la 3era posicion
                    Cad = Cad & " AND codartic like '__" & miRsAux!raiz & "%'"
                    R.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not R.EOF
                        '-- comprobar si admite descuento
                        If R!origpre = "T" Then
                            Cad = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                            Cad = DevuelveDesdeBDNew(conAri, "slista", "dtopermi", "codartic", R!codArtic, "T", , "codlista", Cad, "N")
                            bAct = (Cad = "1")
                        ElseIf R!origpre = "A" Or R!origpre = "M" Then
                            bAct = True
                        Else
                            bAct = False
                        End If
                        
                        If bAct Then
                            Cad = CalcularImporte(CStr(R!Cantidad), CStr(R!precioar), CStr(NuevoDto), CStr(R!dtoline2), vParamAplic.TipoDtos)
                            Importe = CCur(Cad)
                            Cad = "UPDATE slialbcli set dtoline1=" & TransformaComasPuntos(CStr(NuevoDto))
                            Cad = Cad & ", importel = " & TransformaComasPuntos(CStr(Importe))
                            Cad = Cad & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
                            Cad = Cad & " and numlinea = " & R!numlinea
                            conn.Execute Cad
                        End If
                        
                        'Siguiente
                        R.MoveNext
                    Wend
                    R.Close
                    
                End If
                'sig
                miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If 'masl 140909
    
EDescuentosCantidad:
    If Err.Number <> 0 Then MuestraError Err.Number, "DescuentosxCantidad"
    Set miRsAux = Nothing
    Set R = Nothing
End Sub





Private Sub AbrirForm_CentroCoste()
    Screen.MousePointer = vbHourglass
    cmdAux(9).Tag = "9"

    Set frmB = New frmBuscaGrid
    frmB.vCampos = "Codigo|cabccost|codccost|T||20·Descripción|cabccost|nomccost|T||70·"
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
    cmdAux(9).Tag = "-1"
End Sub



Private Sub PosicionarData2()
    On Error GoTo EPosicionarData2
    
    Data2.Recordset.Find "numlinea = " & NumRegElim
    If Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
    NumRegElim = 0
    Exit Sub
EPosicionarData2:
    MuestraError Err.Number
End Sub



Private Sub PonerUltAlmacen()
Dim C As String
    If AlmacenLineas < 0 Then
       If Not Data2.Recordset.EOF Then
            C = ObtenerWhereCP(True)
            C = Replace(C, NombreTabla, NomTablaLineas)
            AlmacenLineas = DevuelveUltimoAlmacen(NomTablaLineas, C)
       End If
            
       If AlmacenLineas < 0 Then
            'No hay datos todavia
            '                                                                trabajador
            C = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
            If C <> "" Then AlmacenLineas = Val(C)
        End If
    End If
End Sub



Private Sub UpdateaNomDirec()
Dim N As Integer
Dim Ol As Integer
Dim C As String

    N = -1
    If Not IsNull(Data1.Recordset!CodDirec) Then N = Data1.Recordset!CodDirec
    
    Ol = -1
    If Text1(12).Text <> "" Then Ol = CInt(Text1(12).Text)
    
    If N <> Ol Then
        If Ol < 0 Then
            C = "NULL"
        Else
            C = DBSet(Text2(12).Text, "T")
        End If
        C = "UPDATE scaalbcli set nomdirec=" & C
        C = C & " WHERE codtipom = '" & Text1(30).Text & "' AND numalbar=" & Text1(0).Text
        ejecutar C, False
    End If
End Sub





'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, Index As Integer)

    If InsertandoCabecera Then
        EsCabecera = True
        imgBuscar_Click Index
        
    Else
        'Lineas
        EsCabecera = False
        cmdAux_Click Index
        
        
    End If
        
End Sub
