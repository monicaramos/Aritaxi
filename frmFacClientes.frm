VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFacClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   ForeColor       =   &H00800000&
   Icon            =   "frmFacClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   60
      TabIndex        =   63
      Top             =   1110
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmFacClientes.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(13)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(14)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(34)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(15)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(36)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(37)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(11)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgBuscar(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgBuscar(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(17)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "imgBuscar(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "imgBuscar(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "imgBuscar(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "imgWeb"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(16)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "imgFecha(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(19)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "imgBuscar(11)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "imgBuscar(12)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(58)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(4)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(5)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(6)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(7)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(8)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(22)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "frameAdmon"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "frameComercial"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(12)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(9)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text2(9)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text2(12)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(10)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text2(11)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text2(10)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(13)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "chkClienteV"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text1(45)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(11)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text1(47)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmFacClientes.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameDptoVentas"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameDptoAdmon"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frameDptoDirec"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Direcciones"
      TabPicture(2)   =   "frmFacClientes.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ToolAux"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "FrameDirecciones"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Documentos"
      TabPicture(3)   =   "frmFacClientes.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3(0)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Text1(46)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lw1"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label3"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "imgFecha(3)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "LabelDoc"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "CRM"
      TabPicture(4)   =   "frmFacClientes.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdAccCRM(2)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdAccCRM(1)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmdAccCRM(0)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame3(1)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "lwCRM"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "LabelCRM"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Subclientes"
      TabPicture(5)   =   "frmFacClientes.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtAux1(0)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "txtAux1(1)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "cmdAux(0)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "DataGrid1"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Adodc2"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).ControlCount=   5
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74670
         MaxLength       =   6
         TabIndex        =   177
         Tag             =   "Cliente de Albaran|N|N|||scliente_albaran|codclienalb|000000||"
         Text            =   "clien"
         Top             =   2370
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73500
         TabIndex        =   176
         Text            =   "nomclien"
         Top             =   2370
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdAux 
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   -73710
         TabIndex        =   175
         ToolTipText     =   "Buscar cliente"
         Top             =   2370
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Index           =   47
         Left            =   5880
         MaxLength       =   80
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Tag             =   "Observaciones Pie Factura |T|S|||scliente|observa1|||"
         Top             =   4680
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   1560
         TabIndex        =   13
         Tag             =   "Gastos Adm.|N|S|||scliente|gasAdm|#,###,###,##0.00|N|"
         Text            =   "Tex"
         Top             =   4530
         Width           =   1365
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   2
         Left            =   -64920
         Picture         =   "frmFacClientes.frx":00B4
         Style           =   1  'Graphical
         TabIndex        =   173
         ToolTipText     =   "Eliminar"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   1
         Left            =   -64320
         Picture         =   "frmFacClientes.frx":0AB6
         Style           =   1  'Graphical
         TabIndex        =   172
         ToolTipText     =   "Impresion CRM"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   0
         Left            =   -65400
         Picture         =   "frmFacClientes.frx":1040
         Style           =   1  'Graphical
         TabIndex        =   171
         ToolTipText     =   "Acciones CRM"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4215
         Index           =   1
         Left            =   -74880
         TabIndex        =   168
         Top             =   840
         Width           =   615
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   2370
            Left            =   30
            TabIndex        =   169
            Top             =   30
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   4180
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Appearance      =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   11
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Acciones comerciales"
                  Object.Tag             =   "0"
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Llamadas"
                  Object.Tag             =   "1"
                  Style           =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Cobros"
                  Object.Tag             =   "3"
                  Style           =   2
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3855
         Index           =   0
         Left            =   -74880
         TabIndex        =   165
         Top             =   480
         Width           =   615
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   1050
            Left            =   0
            TabIndex        =   166
            Top             =   0
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1852
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Appearance      =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   10
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Llamadas"
                  Object.Tag             =   "2"
                  Style           =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Facturas"
                  Object.Tag             =   "3"
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
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   46
         Left            =   -65280
         TabIndex        =   163
         Text            =   "Text4"
         Top             =   1440
         Width           =   1455
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   4695
         Left            =   -74040
         TabIndex        =   162
         Top             =   480
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8281
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
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Left            =   -74640
         TabIndex        =   154
         Top             =   780
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "�ltimo"
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrameDirecciones 
         Caption         =   "Direcciones"
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
         Height          =   3315
         Left            =   -74760
         TabIndex        =   129
         Top             =   1500
         Width           =   10695
         Begin VB.Frame FrameCtaBanDpto 
            Height          =   840
            Left            =   5880
            TabIndex        =   155
            Top             =   2280
            Width           =   4455
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   10
               Left            =   360
               MaxLength       =   4
               TabIndex        =   140
               Tag             =   "C�digo Banco|N|S|0|9999|sdirec|codbanco|0000|N|"
               Text            =   "Text"
               Top             =   360
               Width           =   645
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   11
               Left            =   1080
               MaxLength       =   4
               TabIndex        =   141
               Tag             =   "Sucursal|N|S|0|9999|sdirec|codsucur|0000|N|"
               Text            =   "Text"
               Top             =   360
               Width           =   645
            End
            Begin VB.TextBox Text3 
               Height          =   315
               Index           =   12
               Left            =   1800
               MaxLength       =   2
               TabIndex        =   142
               Tag             =   "D�gito Control|T|S|||sdirec|digcontr|00||"
               Text            =   "Text1"
               Top             =   360
               Width           =   405
            End
            Begin VB.TextBox Text3 
               Height          =   315
               Index           =   13
               Left            =   2280
               MaxLength       =   10
               TabIndex        =   143
               Tag             =   "Cuenta Bancaria|T|S|||sdirec|cuentaba|0000000000||"
               Text            =   "Text1"
               Top             =   360
               Width           =   1605
            End
            Begin VB.Label Label1 
               Caption         =   "Banco"
               Height          =   255
               Index           =   47
               Left            =   360
               TabIndex        =   159
               Top             =   165
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Sucursal"
               Height          =   255
               Index           =   35
               Left            =   1080
               TabIndex        =   158
               Top             =   160
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "DC"
               Height          =   255
               Index           =   33
               Left            =   1800
               TabIndex        =   157
               Top             =   160
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "Cta. Bancaria"
               Height          =   255
               Index           =   20
               Left            =   2280
               TabIndex        =   156
               Top             =   160
               Width           =   975
            End
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   3
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   133
            Tag             =   "C.Postal|T|N|||sdirec|codpobla||N|"
            Text            =   "Text3"
            Top             =   1425
            Width           =   765
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   8
            Left            =   7080
            MaxLength       =   10
            TabIndex        =   138
            Tag             =   "Fax|T|S|||sdirec|faxdirec||N|"
            Text            =   "Text3"
            Top             =   1425
            Width           =   1605
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   1
            Left            =   1380
            MaxLength       =   30
            TabIndex        =   131
            Tag             =   "Nombre Direc./Dpto|T|N|||sdirec|nomdirec||N|"
            Text            =   "Text3"
            Top             =   720
            Width           =   3270
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   9
            Left            =   7080
            MaxLength       =   40
            TabIndex        =   139
            Tag             =   "e-mail|T|S|||sdirec|maidirec||N|"
            Text            =   "Text3"
            Top             =   1785
            Width           =   3270
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   6
            Left            =   7080
            MaxLength       =   30
            TabIndex        =   136
            Tag             =   "Persona Contacto|T|S|||sdirec|perdirec||N|"
            Text            =   "Text3"
            Top             =   720
            Width           =   3270
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   7
            Left            =   7080
            MaxLength       =   10
            TabIndex        =   137
            Tag             =   "Tel�fono|T|S|||sdirec|teldirec||N|"
            Text            =   "Text3"
            Top             =   1080
            Width           =   1605
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   5
            Left            =   1380
            MaxLength       =   30
            TabIndex        =   135
            Tag             =   "Provincia|T|N|||sdirec|prodirec||N|"
            Text            =   "Text3"
            Top             =   2145
            Width           =   3285
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   4
            Left            =   1380
            MaxLength       =   30
            TabIndex        =   134
            Tag             =   "Poblaci�n|T|N|||sdirec|pobdirec||N|"
            Text            =   "Text3"
            Top             =   1785
            Width           =   3285
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   2
            Left            =   1380
            MaxLength       =   30
            TabIndex        =   132
            Tag             =   "Domicilio|T|N|||sdirec|domdirec||N|"
            Text            =   "Text3"
            Top             =   1080
            Width           =   3270
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   0
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   130
            Tag             =   "C�digo Direc./Dpto|N|N|0|999|sdirec|coddirec|000|S|"
            Text            =   "Text3"
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label1 
            Caption         =   "El 0 ser� la direcci�n de facturaci�n"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   57
            Left            =   2040
            TabIndex        =   160
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Fax"
            Height          =   255
            Index           =   30
            Left            =   5880
            TabIndex        =   153
            Top             =   1425
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "E-mail"
            Height          =   255
            Index           =   10
            Left            =   5880
            TabIndex        =   152
            Top             =   1785
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Pers. Contacto"
            Height          =   255
            Index           =   27
            Left            =   5880
            TabIndex        =   151
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
            Height          =   255
            Index           =   28
            Left            =   5880
            TabIndex        =   150
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   26
            Left            =   360
            TabIndex        =   149
            Top             =   2145
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
            Height          =   255
            Index           =   25
            Left            =   360
            TabIndex        =   148
            Top             =   1785
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "C.Postal"
            Height          =   255
            Index           =   24
            Left            =   360
            TabIndex        =   147
            Top             =   1425
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   23
            Left            =   360
            TabIndex        =   146
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. Direc."
            Height          =   255
            Index           =   22
            Left            =   360
            TabIndex        =   145
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   21
            Left            =   360
            TabIndex        =   144
            Top             =   720
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1080
            ToolTipText     =   "Buscar poblaci�n"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   2
            Left            =   6720
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   1800
            Width           =   240
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   45
         Left            =   1560
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   127
         Tag             =   "Password cliente|T|N|||scliente|pasclien|||"
         Text            =   "3"
         Top             =   930
         Width           =   2220
      End
      Begin VB.CheckBox chkClienteV 
         Caption         =   "Cliente Varios"
         Height          =   195
         Left            =   3960
         TabIndex        =   4
         Tag             =   "Cliente Varios|N|N|||scliente|clivario||N|"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   13
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha de Alta|F|N|||scliente|fechaalt|dd/mm/yyyy|N|"
         Top             =   480
         Width           =   1230
      End
      Begin VB.Frame frameDptoVentas 
         Caption         =   "Datos Relacionados con Dpto. Ventas"
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
         Height          =   3015
         Left            =   -69600
         TabIndex        =   102
         Top             =   480
         Width           =   5895
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   49
            Left            =   2370
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   181
            Text            =   "Text2"
            Top             =   2640
            Width           =   3315
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   49
            Left            =   1650
            MaxLength       =   4
            TabIndex        =   180
            Tag             =   "Cod.BanPr|N|S|0|9999|scliente|codbanpr|0000|N|"
            Text            =   "Tex"
            Top             =   2640
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   38
            Left            =   5280
            MaxLength       =   1
            TabIndex        =   47
            Tag             =   "Per�odo Facturaci�n|N|N|0|9|scliente|periodof|0|N|"
            Text            =   "T"
            Top             =   1320
            Width           =   390
         End
         Begin VB.CheckBox chkReferencia 
            Caption         =   "Referencia Obligada"
            Height          =   315
            Left            =   1680
            TabIndex        =   49
            Tag             =   "Referencia obligada|N|N|||scliente|referobl||N|"
            Top             =   2280
            Width           =   1815
         End
         Begin VB.CheckBox chkPromociones 
            Caption         =   "Aplicar Promociones"
            Height          =   315
            Left            =   3765
            TabIndex        =   50
            Tag             =   "Aplicar Promociones|N|N|||scliente|promocio||N|"
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   37
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   108
            Text            =   "Text2"
            Top             =   840
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   37
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   44
            Tag             =   "Cod. Tarifa|N|N|0|999|scliente|codtarif|000|N|"
            Text            =   "Tex"
            Top             =   840
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   39
            Left            =   5280
            MaxLength       =   1
            TabIndex        =   48
            Tag             =   "Repeticiones Fact.|N|S|1|9|scliente|numrepet|#|N|"
            Text            =   "T"
            Top             =   1800
            Width           =   390
         End
         Begin VB.ComboBox cboAlbaran 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Tag             =   "Valorar albaran con|N|N|||scliente|albarcon||N|"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.ComboBox cboFacturacion 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Tag             =   "Tipo Facturaci�n|N|N|||scliente|tipofact||N|"
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   36
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   43
            Tag             =   "Cod. Agente|T|N|0|9999|scliente|codagent|0000|N|"
            Text            =   "Text"
            Top             =   360
            Width           =   645
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   36
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   103
            Text            =   "Text2"
            Top             =   360
            Width           =   3285
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   13
            Left            =   1365
            ToolTipText     =   "Buscar banco propio"
            Top             =   2640
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Banco Propio"
            Height          =   255
            Index           =   61
            Left            =   270
            TabIndex        =   182
            Top             =   2640
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Per�odo Facturaci�n"
            Height          =   255
            Index           =   46
            Left            =   3765
            TabIndex        =   110
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1395
            ToolTipText     =   "Buscar tarifa"
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. Tarifa"
            Height          =   255
            Index           =   59
            Left            =   240
            TabIndex        =   109
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Repeticiones Fact."
            Height          =   255
            Index           =   55
            Left            =   3765
            TabIndex        =   107
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Valorar Albaran con"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   106
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Facturaci�n"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   105
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. Agente"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   104
            Top             =   360
            Width           =   975
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1395
            ToolTipText     =   "Buscar agente"
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame frameDptoAdmon 
         Caption         =   "Datos Relacionados con Dpto. Administraci�n"
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
         Height          =   4815
         Left            =   -74880
         TabIndex        =   89
         Top             =   480
         Width           =   5175
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   48
            Left            =   270
            MaxLength       =   4
            TabIndex        =   36
            Tag             =   "IBAN|T|S|||scliente|iban||N|"
            Text            =   "Text"
            Top             =   3120
            Width           =   645
         End
         Begin VB.CheckBox chkCorreo 
            Caption         =   "Se le envia correo"
            Height          =   315
            Left            =   2400
            TabIndex        =   34
            Tag             =   "Referencia obligada|N|N|||scliente|enviocorreo||N|"
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CheckBox chkTasaReciclado 
            Caption         =   "Tas......"
            Height          =   315
            Left            =   2400
            TabIndex        =   35
            Tag             =   "tasareciclado|N|N|0|1|scliente|tasareciclado||N|"
            Top             =   2280
            Width           =   2445
         End
         Begin VB.ComboBox cboTipoIVA 
            Height          =   315
            ItemData        =   "frmFacClientes.frx":1A42
            Left            =   3480
            List            =   "frmFacClientes.frx":1A44
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Tag             =   "Tipo de IVA|N|N|||scliente|tipoiva||N|"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Utiliza Cta. Ventas alternativa"
            Height          =   315
            Left            =   1680
            TabIndex        =   41
            Tag             =   "Cancela abonos|N|N|||scliente|cliabono||N|"
            Top             =   3840
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   25
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   27
            Tag             =   "Dto. General|N|N|0|99.90|scliente|dtognral|#0.00||"
            Text            =   "Text1"
            Top             =   1320
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   24
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   26
            Tag             =   "Dto. Pronto Pago|N|N|0|99.90|scliente|dtoppago|#0.00||"
            Text            =   "Text1"
            Top             =   840
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   27
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   29
            Tag             =   "D�a Vto. Atrasado|N|S|0|31|scliente|diavtoat||N|"
            Text            =   "Te"
            Top             =   2280
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   28
            Left            =   3480
            MaxLength       =   2
            TabIndex        =   30
            Tag             =   "D�a Pago 1|N|S|0|99|scliente|diapago1||N|"
            Text            =   "Te"
            Top             =   840
            Width           =   405
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   35
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   97
            Text            =   "Text2"
            Top             =   4200
            Width           =   3165
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   35
            Left            =   240
            MaxLength       =   10
            TabIndex        =   42
            Tag             =   "Cta. Contable|T|N|||scliente|codmacta||N|"
            Text            =   "Text1"
            Top             =   4200
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   34
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   40
            Tag             =   "Cuenta Bancaria|T|S|||scliente|cuentaba|0000000000||"
            Text            =   "Text1"
            Top             =   3120
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   33
            Left            =   2325
            MaxLength       =   2
            TabIndex        =   39
            Tag             =   "D�gito Control|T|S|||scliente|digcontr|00||"
            Text            =   "Text1"
            Top             =   3120
            Width           =   405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   32
            Left            =   1635
            MaxLength       =   4
            TabIndex        =   38
            Tag             =   "Sucursal|N|S|0|9999|scliente|codsucur|0000|N|"
            Text            =   "Text"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   31
            Left            =   960
            MaxLength       =   4
            TabIndex        =   37
            Tag             =   "C�digo Banco|N|S|0|9999|scliente|codbanco|0000|N|"
            Text            =   "Text"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   26
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   28
            Tag             =   "Mes a no girar|N|S|0|12|scliente|mesnogir||N|"
            Text            =   "Te"
            Top             =   1800
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   29
            Left            =   3960
            MaxLength       =   2
            TabIndex        =   31
            Tag             =   "D�a de Pago 2|N|S|0|99|scliente|diapago2||N|"
            Text            =   "Te"
            Top             =   840
            Width           =   405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   30
            Left            =   4440
            MaxLength       =   2
            TabIndex        =   32
            Tag             =   "D�a Pago 3|N|S|0|99|scliente|diapago3||N|"
            Text            =   "Te"
            Top             =   840
            Width           =   405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   23
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   25
            Tag             =   "Cod. F. Pago|N|N|0|999|scliente|codforpa|000|N|"
            Text            =   "Tex"
            Top             =   360
            Width           =   645
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   23
            Left            =   2280
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   90
            Text            =   "Text2"
            Top             =   360
            Width           =   2565
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
            Height          =   255
            Index           =   60
            Left            =   270
            TabIndex        =   179
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo IVA"
            Height          =   255
            Index           =   29
            Left            =   2400
            TabIndex        =   121
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Contable"
            Height          =   255
            Index           =   51
            Left            =   240
            TabIndex        =   118
            Top             =   3960
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. General"
            Height          =   195
            Index           =   54
            Left            =   240
            TabIndex        =   101
            Top             =   1320
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Pronto Pago"
            Height          =   195
            Index           =   53
            Left            =   240
            TabIndex        =   100
            Top             =   840
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "D�a Vt. atrasado"
            Height          =   255
            Index           =   52
            Left            =   240
            TabIndex        =   99
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Mes a no girar"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   98
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1275
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   3915
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Bancaria"
            Height          =   255
            Index           =   32
            Left            =   2760
            TabIndex        =   96
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "DC"
            Height          =   255
            Index           =   50
            Left            =   2400
            TabIndex        =   95
            Top             =   2880
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   49
            Left            =   1635
            TabIndex        =   94
            Top             =   2880
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Banco"
            Height          =   255
            Index           =   48
            Left            =   960
            TabIndex        =   93
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "D�as de Pago"
            Height          =   255
            Index           =   31
            Left            =   2400
            TabIndex        =   92
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. F. Pago"
            Height          =   255
            Index           =   68
            Left            =   240
            TabIndex        =   91
            Top             =   360
            Width           =   975
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1275
            ToolTipText     =   "Buscar forma de pago"
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   10
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   86
         Text            =   "Text2"
         Top             =   4080
         Width           =   3165
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   11
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   85
         Text            =   "Text2"
         Top             =   4530
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   12
         Tag             =   "Cod. Env�o|N|S|0|999|scliente|codenvio|000|N|"
         Text            =   "Tex"
         Top             =   4080
         Width           =   645
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   12
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   84
         Text            =   "Text2"
         Top             =   4980
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   9
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   82
         Text            =   "Text2"
         Top             =   3630
         Width           =   3165
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   11
         Tag             =   "Cod.Actividad|N|N|0|999|scliente|codactiv|000|N|"
         Text            =   "Tex"
         Top             =   3630
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   12
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   14
         Tag             =   "Cod. Ruta|N|S|0|999|scliente|codrutas|000|N|"
         Text            =   "Tex"
         Top             =   4980
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame frameComercial 
         Caption         =   "Comercial"
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
         Height          =   1400
         Left            =   5880
         TabIndex        =   76
         Top             =   2040
         Width           =   5295
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   18
            Left            =   960
            MaxLength       =   30
            TabIndex        =   19
            Tag             =   "Contacto Comercial|T|S|||scliente|perclie2||N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   3990
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   19
            Left            =   960
            MaxLength       =   20
            TabIndex        =   20
            Tag             =   "Tel�fono Comercial|T|S|||scliente|telclie2||N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   1590
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   20
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   21
            Tag             =   "Fax Comercial|T|S|||scliente|faxclie2||N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   1710
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   21
            Left            =   960
            MaxLength       =   40
            TabIndex        =   22
            Tag             =   "e-mail Comercial|T|S|||scliente|maiclie2||N|"
            Text            =   "Text1"
            Top             =   960
            Width           =   3990
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   1
            Left            =   645
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Contacto"
            Height          =   255
            Index           =   44
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
            Height          =   255
            Index           =   43
            Left            =   120
            TabIndex        =   79
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Fax"
            Height          =   255
            Index           =   42
            Left            =   2880
            TabIndex        =   78
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "E-mail"
            Height          =   255
            Index           =   41
            Left            =   120
            TabIndex        =   77
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.Frame frameAdmon 
         Caption         =   "Administraci�n"
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
         Height          =   1400
         Left            =   5880
         TabIndex        =   71
         Top             =   480
         Width           =   5295
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   17
            Left            =   960
            MaxLength       =   40
            TabIndex        =   18
            Tag             =   "e-mail Admon.|T|S|||scliente|maiclie1||N|"
            Text            =   "maiclie1"
            Top             =   960
            Width           =   3990
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   16
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   17
            Tag             =   "Fax Admon.|T|S|||scliente|faxclie1||N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   1710
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   15
            Left            =   960
            MaxLength       =   20
            TabIndex        =   16
            Tag             =   "Tel�fono Admon.|T|S|||scliente|telclie1||N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   1590
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   14
            Left            =   960
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "Contacto Admon.|T|S|||scliente|perclie1||N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   3990
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   0
            Left            =   600
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "E-mail"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   75
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Fax"
            Height          =   255
            Index           =   39
            Left            =   2880
            TabIndex        =   74
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   73
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Contacto"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Index           =   22
         Left            =   5880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Tag             =   "Observaciones|T|S|||scliente|observac|||"
         Top             =   3780
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "Web|T|S|||scliente|wwwclien||N|"
         Text            =   "Text1"
         Top             =   3180
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "N.I.F.|T|N|||scliente|nifclien||N|"
         Text            =   "Text1"
         Top             =   2730
         Width           =   1725
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "Provincia|T|N|||scliente|proclien||N|"
         Text            =   "Text1"
         Top             =   2280
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   3105
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Poblaci�n|T|N|||scliente|pobclien||N|"
         Text            =   "Text1"
         Top             =   1860
         Width           =   2340
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "C.Postal|T|N|||scliente|codpobla||N|"
         Text            =   "Text1"
         Top             =   1830
         Width           =   700
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   5
         Tag             =   "Domicilio|T|N|||scliente|domclien||N|"
         Text            =   "Text1"
         Top             =   1380
         Width           =   3885
      End
      Begin VB.Frame frameDptoDirec 
         Caption         =   "Datos Relacionados con Dpto. Direcci�n"
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
         Height          =   1740
         Left            =   -69600
         TabIndex        =   111
         Top             =   3560
         Width           =   5925
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   44
            Left            =   4320
            MaxLength       =   5
            TabIndex        =   55
            Tag             =   "Distancia Km.|N|S|0|99999|scliente|kilometr||N|"
            Text            =   "Text1"
            Top             =   765
            Width           =   750
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   43
            Left            =   4320
            MaxLength       =   16
            TabIndex        =   54
            Tag             =   "L�mite cr�dito|N|S|0||scliente|limcredi|#,###,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   280
            Width           =   1470
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   40
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   51
            Tag             =   "Fecha ult. movim.|F|S|||scliente|fechamov|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   280
            Width           =   1110
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   42
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   53
            Tag             =   "Cod. Situaci�n|N|N|0|99|scliente|codsitua|00|N|"
            Text            =   "Te"
            Top             =   1230
            Width           =   645
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   42
            Left            =   2640
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   112
            Text            =   "Text2"
            Top             =   1230
            Width           =   3165
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   41
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   52
            Tag             =   "Fecha Reclamaci�n|F|S|||scliente|fechaulr|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   765
            Width           =   1110
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha ult. movim."
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   117
            Top             =   285
            Width           =   1335
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   1635
            ToolTipText     =   "Buscar fecha"
            Top             =   315
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. Situaci�n"
            Height          =   255
            Index           =   62
            Left            =   120
            TabIndex        =   116
            Top             =   1230
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   1635
            ToolTipText     =   "Buscar situaci�n"
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "L�mite Cr�dito"
            Height          =   195
            Index           =   45
            Left            =   3315
            TabIndex        =   115
            Top             =   285
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Distancia Km."
            Height          =   195
            Index           =   56
            Left            =   3315
            TabIndex        =   114
            Top             =   765
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Reclamaci�n"
            Height          =   255
            Index           =   66
            Left            =   120
            TabIndex        =   113
            Top             =   765
            Width           =   1455
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   1635
            ToolTipText     =   "Buscar fecha"
            Top             =   795
            Width           =   240
         End
      End
      Begin MSComctlLib.ListView lwCRM 
         Height          =   4335
         Left            =   -74040
         TabIndex        =   167
         Top             =   840
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7646
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacClientes.frx":1A46
         Height          =   4125
         Left            =   -74670
         TabIndex        =   178
         Top             =   630
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   -67110
         Top             =   4710
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
      Begin VB.Label Label1 
         Caption         =   "Observaciones Pie Factura"
         Height          =   240
         Index           =   58
         Left            =   5880
         TabIndex        =   174
         Top             =   4380
         Width           =   1995
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   7890
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   4380
         Width           =   240
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
         Left            =   -74040
         TabIndex        =   170
         Top             =   480
         Width           =   5745
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   7080
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   3540
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Left            =   -65160
         TabIndex        =   164
         Top             =   960
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   -64320
         Picture         =   "frmFacClientes.frx":1A5B
         ToolTipText     =   "Buscar fecha"
         Top             =   960
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
         Left            =   -65640
         TabIndex        =   161
         Top             =   480
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Password web"
         Height          =   255
         Index           =   19
         Left            =   360
         TabIndex        =   128
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1245
         ToolTipText     =   "Buscar fecha"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Alta"
         Height          =   255
         Index           =   16
         Left            =   375
         TabIndex        =   126
         Top             =   480
         Width           =   855
      End
      Begin VB.Image imgWeb 
         Height          =   255
         Left            =   1200
         Picture         =   "frmFacClientes.frx":1AE6
         Stretch         =   -1  'True
         Tag             =   "-1"
         ToolTipText     =   "Abrir web"
         Top             =   3180
         Width           =   255
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1245
         Tag             =   "-1"
         ToolTipText     =   "Buscar poblaci�n"
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1245
         ToolTipText     =   "Buscar lote"
         Top             =   4140
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1245
         ToolTipText     =   "Buscar zona"
         Top             =   4620
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Lote"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   88
         Top             =   4140
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Ruta"
         Height          =   255
         Index           =   17
         Left            =   360
         TabIndex        =   87
         Top             =   4980
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1245
         ToolTipText     =   "Buscar ruta"
         Top             =   4980
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1245
         Tag             =   "-1"
         ToolTipText     =   "Buscar actividad"
         Top             =   3660
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.  Activ."
         Height          =   255
         Index           =   5
         Left            =   375
         TabIndex        =   83
         Top             =   3660
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Gastos de Adm."
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   81
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   240
         Index           =   11
         Left            =   5880
         TabIndex        =   70
         Top             =   3540
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Web"
         Height          =   255
         Index           =   37
         Left            =   375
         TabIndex        =   69
         Top             =   3180
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.F."
         Height          =   255
         Index           =   36
         Left            =   375
         TabIndex        =   68
         Top             =   2790
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   15
         Left            =   375
         TabIndex        =   67
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Poblaci�n"
         Height          =   255
         Index           =   34
         Left            =   2370
         TabIndex        =   66
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "C. Postal"
         Height          =   255
         Index           =   14
         Left            =   375
         TabIndex        =   65
         Top             =   1890
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   13
         Left            =   375
         TabIndex        =   64
         Top             =   1470
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   580
      Left            =   120
      TabIndex        =   122
      Top             =   450
      Width           =   11415
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   7725
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Nombre Comercial|T|S|||scliente|nomcomer||N|"
         Text            =   "Text1"
         Top             =   170
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   2540
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre Cliente|T|N|||scliente|nomclien||N|"
         Text            =   "Text1"
         Top             =   170
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   670
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "C�digo Cliente|N|N|0|999999|scliente|codclien|000000|S|"
         Text            =   "Text1"
         Top             =   170
         Width           =   950
      End
      Begin VB.Label Label1 
         Caption         =   "Nom.Comercial"
         Height          =   255
         Index           =   12
         Left            =   6600
         TabIndex        =   125
         Top             =   170
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   1910
         TabIndex        =   124
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   123
         Top             =   165
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   1
      Left            =   2880
      TabIndex        =   119
      Top             =   6540
      Width           =   4575
      Begin VB.Label lblSituacion 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   120
         TabIndex        =   120
         Top             =   180
         Width           =   4395
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10200
      TabIndex        =   57
      Top             =   6645
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   59
      Top             =   6540
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   60
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10200
      TabIndex        =   58
      Top             =   6645
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9000
      TabIndex        =   56
      Top             =   6645
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5880
      Top             =   6600
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   11640
      _ExtentX        =   20532
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
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Direcciones/Departamentos"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Subclientes"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   9360
         TabIndex        =   62
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   7440
      Top             =   6690
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public DeConsulta As Boolean


Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios
Attribute frmMtoBancosPro.VB_VarHelpID = -1

Private WithEvents frmA As frmFacActividades
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmE As frmFacFormasEnvio
Attribute frmE.VB_VarHelpID = -1

Private WithEvents frmFP As frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmAC As frmFacAgentesCom
Attribute frmAC.VB_VarHelpID = -1
Private WithEvents frmT As frmFacTarifas
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmS As frmFacSituaciones
Attribute frmS.VB_VarHelpID = -1
Private WithEvents frmLLam As frmGesHisLlam
Attribute frmLLam.VB_VarHelpID = -1

Private WithEvents frmCli As frmFacClientes2
Attribute frmCli.VB_VarHelpID = -1


'Para los documentos
Private frmAlb As frmFacEntAlbaranes


Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas de Articulos x Almacen
'   6.-  Mantenimiento Lineas de Componentes de Conjuntos
'   7.-  Mantenimiento Lineas de Control de Instalaciones
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------


Private ModoFrame As Byte
'ModoFrame: 0.-Inicio, 3.-Insertar, 4.-Modificar

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1

Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas
    
Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal


'NUEVO: JULIO 2007. PARA BUSCAR POR CHECKS TB
'------------------------------------------------
Private BuscaChekc As String

Dim cadB1 As String




Private Sub cboAlbaran_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboTipoIVA_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAbonos_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkAbonos_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkClienteV_Click()
If Modo = 1 Then CheckCadenaBusqueda chkPromociones, BuscaChekc
End Sub

Private Sub chkClienteV_GotFocus()
   ConseguirfocoChk Modo
End Sub

Private Sub chkClienteV_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCorreo_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkPromociones, BuscaChekc
End Sub

Private Sub chkCorreo_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkCorreo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPromociones_Click()
 If Modo = 1 Then CheckCadenaBusqueda chkPromociones, BuscaChekc
End Sub

Private Sub chkPromociones_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkPromociones_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub chkReferencia_Click()
    
    'Buscqueda
    If Modo = 1 Then CheckCadenaBusqueda chkReferencia, BuscaChekc
    
End Sub

Private Sub chkReferencia_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkReferencia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkTasaReciclado_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkTasaReciclado, BuscaChekc
End Sub

Private Sub chkTasaReciclado_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkTasaReciclado_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAccCRM_Click(Index As Integer)
    
    'Acciones parar el CRM
    Select Case Index
    Case 1
        If Modo <> 2 Then Exit Sub
        If Data1.Recordset.EOF Then Exit Sub
        If Text1(0).Text = "" Then Exit Sub
        
        
        frmCRMImprimir.Text1 = Text1(0).Text
        frmCRMImprimir.Text2 = Text1(1).Text
        frmCRMImprimir.Show vbModal
        
    Case 0
    
        Select Case CByte(RecuperaValor(lwCRM.Tag, 1))
        Case 0
            'NUEVA, modificar o insertar acciones comerciales
            frmCRMMto.DesdeElCliente = Data1.Recordset!CodClien
            frmCRMMto.TipoPredefinido = 0  'sin tipo predefinido
            frmCRMMto.DatosADevolverBusqueda = ""   'NUEVA
            frmCRMMto.Show vbModal
        Case 1
            'NUEVA llamda EFECTUADA
            frmCRMMto.DesdeElCliente = Data1.Recordset!CodClien
            frmCRMMto.TipoPredefinido = 1  'Llamada efectuada
            frmCRMMto.DatosADevolverBusqueda = ""   'NUEVA
            frmCRMMto.Show vbModal
            
        Case 2
            'Emails
            LanzarProgramaEmails
            If MsgBox("Refrescar datos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Case 3
            'NO puede insertar nada.
            Exit Sub
        Case 4
            frmCrmObsDpto.Nuevo = True
            frmCrmObsDpto.Label2.Caption = Data1.Recordset!nomclien
            frmCrmObsDpto.Tag = Data1.Recordset!CodClien
            frmCrmObsDpto.Show vbModal
        End Select
        Me.Refresh
        DoEvents
        CargaDatosLWCRM
        Screen.MousePointer = vbDefault
    Case 2
    
        If CByte(RecuperaValor(lwCRM.Tag, 1)) = 4 Then
            If lwCRM.SelectedItem Is Nothing Then Exit Sub
            If MsgBox("�Desea eliminar las observaciones del departamento " & Me.lwCRM.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
            BuscaChekc = "DELETE from scrmobsclien  WHERE codclien = " & Me.Data1.Recordset!CodClien & " AND dpto=" & lwCRM.SelectedItem.SubItems(3)
            If ejecutar(BuscaChekc, False) Then CargaDatosLWCRM
            BuscaChekc = ""
        End If
    End Select
End Sub

Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
              If InsertarDesdeForm(Me, 1) Then
                 'Si pone en la cuenta contable, crear nueva cta contable
                 If Text2(35).Text = vbCrearNuevaCta Then
                    If Not InsertarCuentaCble(Text1(35).Text, Text1(0).Text) Then
                        MsgBox "Se ha producido un error insertando la cuenta: " & Text1(1).Text & ". Compruebelo", vbExclamation
                        Exit Sub
                    End If
                 End If
                
                 PosicionarData
                 CargaFrameDirec
              End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    Text2(35) = PonerNombreCuenta(Text1(35), Modo, Text1(0).Text)
                    
                    TerminaBloquear
                    PosicionarData
                End If
            End If
                
         Case 5 'InsertarModificar linea
            'Actualizar el registro en la tabla de lineas 'sdirec' (Direcciones/Departamentos)
            If SSTab1.Tab = 2 Then
                If InsertarModificarLinea Then
                    cad = "coddirec = " & Text3(0).Text & ""
                    CargaFrameDirec
                    If SituarData(Data2, cad, Indicador) Then
                        PonerCamposDirecciones
                        ModificaLineas = 0
    '                    lblIndicador.Caption = Indicador
                        PonerModoFrame 0
                    End If
                End If
                PonerFocoBtn Me.cmdRegresar
            Else
                '[Monica]24/09/2012
                'clientes de albaran
                If ModificaLineas = 1 Then
                    If InsertarLinea Then
                        CargaGrid DataGrid1, Adodc2
                        BotonAnyadirLinea2
                    End If
                Else
                    If ModificaLineas = 2 Then
                        If ModificarLinea Then
                            TerminaBloquear
                            CargaTxtAux False, False
                            CargaGrid DataGrid1, Adodc2
                            ModificaLineas = 0
                            PonerBotonCabecera True
                        End If
                    End If
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0
            Set frmCli = New frmFacClientes2
            frmCli.DatosADevolverBusqueda = "0"
            frmCli.Show vbModal
            Set frmCli = Nothing
            
    End Select

End Sub

Private Sub cmdCancelar_Click()
Dim cad As String
Dim Indicador As String

    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            If SSTab1.Tab = 2 Then
                PonerModoFrame 0
                If ModificaLineas = 1 Then '1 = Insertar
                    If Not Data2.Recordset.EOF Then
                        Data2.Recordset.MoveFirst
                        PonerCamposDirecciones
                    Else
                        LimpiarCamposDirecciones
                    End If
                ElseIf ModificaLineas = 2 Then 'Modificar
                     cad = "(coddirec=" & Text3(0).Text & ")"
                     If SituarData(Data2, cad, Indicador) Then
                        PonerCamposDirecciones
    '                    lblIndicador.Caption = Indicador
                     End If
                End If
                ModificaLineas = 0
                PonerModoOpcionesMenu
                PonerFoco Text3(1)
           Else
            TerminaBloquear
            CargaTxtAux False, False
            
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Adodc2.Recordset.EOF Then Adodc2.Recordset.MoveFirst
            Else
                ModificaLineas = 0
            End If
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
           End If
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos 'Vac�a los TextBox
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    MostrarSituacion False
    
    Text1(0).Text = SugerirCodigoSiguienteStr("scliente", "codclien")
    FormateaCampo Text1(0)
    Text1(13).Text = Format(Now, "dd/mm/yyyy")
    'Sugerir el tipo de IVA como NORMAL
    Me.cboTipoIVA.ListIndex = 0
    'Sugerir valorar albaran con: TODO
    Me.cboAlbaran.ListIndex = 0
    'Sugerir tipo facturacion a: Factura colectiva
    Me.cboFacturacion.ListIndex = 0
    
    Me.chkCorreo.Value = 1
    'Sugerimos periodo y repeticion , a 1
    Text1(38).Text = 1
    Text1(39).Text = 1
    
    'A cero los descuentos
    Text1(24).Text = "0,00"
    Text1(25).Text = "0,00"
    
    'Valores por defecto desde parametros
    If vParamAplic.PorDefecto_Activ > 0 Then
        Text1(9).Text = vParamAplic.PorDefecto_Activ
        Text1_LostFocus 9
    End If
    If vParamAplic.PorDefecto_Envio > 0 Then
        Text1(10).Text = vParamAplic.PorDefecto_Envio
        Text1_LostFocus 10
    End If
    If vParamAplic.PorDefecto_Situ >= 0 Then
        Text1(42).Text = vParamAplic.PorDefecto_Situ
        Text1_LostFocus 42
    End If
    If vParamAplic.PorDefecto_Tarifa > 0 Then
        Text1(37).Text = vParamAplic.PorDefecto_Tarifa
        Text1_LostFocus 37
    End If
    If vParamAplic.PorDefecto_Agente > 0 Then
        Text1(36).Text = vParamAplic.PorDefecto_Agente
        Text1_LostFocus 36
    End If
    Me.SSTab1.Tab = 0
    PonerFoco Text1(0)
    ConseguirFoco Text1(0), Modo
End Sub


Private Sub BotonAnyadirLinea()
Dim vWhere As String
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    Me.SSTab1.Tab = 2
    PonerModoFrame 3   '3: Insertar
    ModificaLineas = 1 'Insertar
    lblIndicador.Caption = "Insertar Linea"
    PonerModoOpcionesMenu

    'Obtenemos la siguiente numero de Direc./Dpto
    vWhere = "codclien=" & Text1(0).Text
    Text3(0).Text = SugerirCodigoSiguienteStr("sdirec", "coddirec", vWhere)
    PonerFoco Text3(0)
End Sub


Private Sub BotonAnyadirLinea2()
Dim vWhere As String
'
'    'Si no estaba modificando lineas salimos
'    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 2 Then Exit Sub
'
'    'A�adiremos el boton de aceptar y demas objetos para insertar
'    Me.SSTab1.Tab = 5
'    PonerModoFrame 3   '3: Insertar
'    ModificaLineas = 1 'Insertar
'    lblIndicador.Caption = "Insertar Linea"
'    PonerModoOpcionesMenu
'
'    'Obtenemos la siguiente numero de Direc./Dpto
'    vWhere = "codclien=" & Text1(0).Text
''    Text3(0).Text = SugerirCodigoSiguienteStr("sdirec", "coddirec", vWhere)
''    PonerFoco Text3(0)
'    PonerFoco txtAux1(0)

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo A�adir Linea
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR Subcliente"
    
    AnyadirLinea DataGrid1, Adodc2
    CargaTxtAux True, True
   
    PonerFoco txtAux1(0)
    Me.DataGrid1.Enabled = False

End Sub



Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(1)
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
'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & " where codclien < 1000000 " & Ordenacion
        PonerCadenaBusqueda
    End If
    
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data

    Select Case Modo
        Case 5 'Modo Mantenimiento de Direcc./Dptos (Lineas)
            If Data2.Recordset.EOF Then Exit Sub
            DesplazamientoData Data2, Index
            PonerCamposDirecciones
            
        Case Else 'Datos de Cabecera
            If Data1.Recordset.EOF Then Exit Sub
            DesplazamientoData Data1, Index
            PonerCampos
            MostrarSituacion True
            CargaFrameDirec
            
'            PonerModoOpcionesMenu
    End Select
End Sub



Private Sub DesplazamientoLineas(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data

'    Select Case Modo
'        Case 5 'Modo Mantenimiento de Direcc./Dptos (Lineas)
            If Data2.Recordset.EOF Then Exit Sub
            DesplazamientoData Data2, Index
            PonerCamposDirecciones
            If Modo = 5 And ModoFrame = 0 Then
                Me.lblIndicador.Caption = "Lineas Detalle"
                If Not Data2.Recordset.EOF Then Me.lblIndicador.Caption = Me.lblIndicador.Caption & ": " & Me.Data2.Recordset.AbsolutePosition & " de " & Me.Data2.Recordset.RecordCount
            End If
            
'        Case Else 'Datos de Cabecera
'            If Data1.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data1, Index
'            PonerCampos
'            MostrarSituacion True
'            CargaFrameDirec
'    End Select
End Sub


Private Sub BotonModificar()
    'A�adiremos el boton de aceptar y demas objetos para insertar
    If Me.SSTab1.Tab = 2 Then Me.SSTab1.Tab = 0
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea

    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    Me.SSTab1.Tab = 2
       
    'A�adiremos el boton de aceptar y demas objetos para insertar
    
    PonerModoFrame 4 'ModoFrame=4 -> Modificar
    Me.lblIndicador.Caption = "Modificar Linea"
    ModificaLineas = 2 'Modificar
    PonerModoOpcionesMenu
    
    'Como el campo 1 es clave primaria, NO se puede modificar
    BloquearTxt Text3(0), True
    
    PonerFoco Text3(1)
        
   
End Sub


Private Sub BotonEliminar()
Dim cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    '### a mano
    cad = "�Seguro que desea eliminar el Cliente?"
    cad = cad & vbCrLf & "Cod. : " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Data1.Recordset.Delete
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'solo habia un registro
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Cliente", Err.Description
    End If
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String, cad2 As String
Dim i As Integer

    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    'Dependiendo del parametro de la aplicacion trabajamos con Dpto o Direc.
    If vParamAplic.Departamento Then
        cad2 = " Dpto. "
        cad = " el Departamento?"
    Else
        cad2 = " Direc. "
        cad = " la Direcci�n?"
    End If
    
    cad = "�Seguro que desea eliminar " & cad & vbCrLf
    cad = cad & vbCrLf & "Cod." & cad2 & ": " & Format(Data2.Recordset.Fields(1), FormatoCampo(Text3(0)))
    cad = cad & vbCrLf & "Nombre" & cad2 & ": " & Data2.Recordset.Fields(2)
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data2.Recordset.AbsolutePosition
        Data2.Recordset.Delete
        
        
        'Para borrar en arimoeny
        If Text1(35).Text <> "" Then
            'SI NO tiene cta contable NO tiene dpto
            cad2 = " WHERE codmacta= '" & Text1(35).Text & "' AND Dpto = " & Text3(0).Text
            ConnConta.Execute "DELETE FROM departamentos " & cad2
        End If
        
        
        If SituarDataTrasEliminar(Data2, NumRegElim) Then
            PonerCamposDirecciones
        Else
             'Solo habia un registro
            LimpiarCamposDirecciones
'            PonerModoFrame 0
        End If
       
        ModificaLineas = 0
        PonerModoFrame 0
    End If
    
    Screen.MousePointer = vbDefault
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data2.Recordset.CancelUpdate
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub


Private Sub BotonDirecciones()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErrorDirec
    
    Me.SSTab1.Tab = 2
    
'    'Crear las lineas de Direcciones/Departamentos para el cliente
'    'ASignamos un SQL al DATA2
'    Me.Data2.ConnectionString = Conn
'    Data2.RecordSource = "Select * from sdirec where codclien = " & Val(Text1(0).Text) & ";"
'    Data2.Refresh
        
    'Poner el modo en el formulario
    PonerModo (5) 'Modo 5: Modificar lineas
    PonerModoFrame 0 'TextBox Bloqueados inicialmente
        
'    If Data2.Recordset.RecordCount > 0 Then
'        Data2.Recordset.MoveFirst
'        PonerCamposDirecciones
'    Else
'        LimpiarCamposDirecciones
'    End If
    
    PonerFocoBtn Me.cmdRegresar
    Screen.MousePointer = vbDefault

    Exit Sub
ErrorDirec:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonSubclientes()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErrorDirec
    
    Me.SSTab1.Tab = 5
    
    ModificaLineas = 0
    PonerModo 5
    PonerBotonCabecera True
    
    Screen.MousePointer = vbDefault

    Exit Sub
ErrorDirec:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub




Private Sub cmdRegresar_Click()
Dim cad As String
Dim Indicador As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Lineas Direcciones/Departamentos
        cad = "(codclien=" & Text1(0).Text & ")"
        If SituarData(Data1, cad, Indicador) Then
'            PonerLineaVisible False
            PonerModo 2
            lblIndicador.Caption = Indicador
        End If
    Else 'Regresar
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
        
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        cad = cad & Data1.Recordset!perclie1 & "|"
        cad = cad & Data1.Recordset!maiclie1 & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For i = 0 To Me.imgFecha.Count - 1
        imgFecha(i).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next

    
    
    'Icono de e-mail
    For kCampo = 0 To Me.ImgMail.Count - 1
        Me.ImgMail(kCampo).Picture = frmPpal.imgListComun.ListImages(20).Picture
    Next kCampo



    ' ICONITOS DE LA BARRA
    btnAnyadir = 6
    btnPrimero = 15
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Bot�n Buscar
        .Buttons(2).Image = 2   'Bot�n Todos
        .Buttons(6).Image = 3   'Insertar Nuevo
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        .Buttons(10).Image = 10 ' Direcciones/Departamentos
        .Buttons(11).Image = 45 ' Subclientes
        .Buttons(12).Image = 16 'imprimir
        .Buttons(13).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
    End With
    Me.SSTab1.Tab = 0
    
    
    'BARRA DE LAS LINEAS de DIRECCION/DPTOS
    With Me.ToolAux
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 6 'primero
        .Buttons(2).Image = 7 'Anterior
        .Buttons(3).Image = 8 'Siguiente
        .Buttons(4).Image = 9 '�ltimo
    End With
    
    
    'La nevegacion para albaranes, facturas....
    ImagenesNavegacion
    
    '[Monica]31/03/2014: los que la tienen marcada no imprimen la factura
    Me.chkTasaReciclado.Caption = "Facturaci�n Electr�nica" '"Tasa reciclado"
'    Me.FrameDirecciones.Top = 1860
'    Me.FrameDirecciones.Left = 360
'    Me.FrameDirecciones.Width = 10815
'    Me.FrameDirecciones.Height = 2600
    
    'Comprobar si es Departamento o Direccion (segun paramatro)
    If vParamAplic.Departamento Then
        Me.Toolbar1.Buttons(10).ToolTipText = "Departamentos"
        Me.FrameDirecciones.Caption = "Departamentos"
        Me.Label1(22).Caption = "Cod. Dpto"
        Me.SSTab1.TabCaption(2) = "Departamentos"
        Me.FrameCtaBanDpto.visible = True
    Else
        Me.Toolbar1.Buttons(10).ToolTipText = "Direcciones"
        Me.FrameDirecciones.Caption = "Direcciones"
        Me.Label1(22).Caption = "Cod. Direc."
        Me.SSTab1.TabCaption(2) = "Direcciones"
        Me.FrameCtaBanDpto.visible = False
    End If
    
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    ModificaLineas = 0
       
    'Si hay algun combo los cargamos
    CargarComboAlbaran
    CargarComboFacturacion
    CargarComboTipoIVA
    
    Me.lblSituacion.visible = False
    Me.Frame1(1).visible = False
    
    'Pone el Tag del primer bot�n de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sclien, BD: Aritaxi
    'Si tag>0 abre busqueda en la tabla: cuentas, BD: Conta.
    imgBuscar(0).Tag = "-1"
         
    '## A mano
    NombreTabla = "scliente"
    Ordenacion = " ORDER BY codclien"
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codclien=-1"
    Data1.Refresh
    
    'Asignamos un SQL al DATA2
    CargaFrameDirec
    
    'Ponemos los datos del listview
    imgFecha(3).Tag = vEmpresa.FechaIni
    CargaColumnas 1
    
    If vParamAplic.TieneCRM Then CargaColumnasCRM 0
    
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
    End If
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkClienteV.Value = 0
    Me.chkAbonos.Value = 0
    Me.chkPromociones.Value = 0
    Me.chkReferencia.Value = 0
    Me.chkTasaReciclado.Value = 0
    Me.cboAlbaran.ListIndex = -1
    Me.cboFacturacion.ListIndex = -1
    Me.cboTipoIVA.ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub LimpiarCamposDirecciones()
Dim i As Byte
    'Limpia los controles TextBox3
    For i = 0 To Text3.Count - 1
        Text3(i).Text = ""
    Next i
'    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Actividades
    Text1(9).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(9)
    Text2(9).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmAC_DatoSeleccionado(CadenaSeleccion As String)
'Agentes Comerciales
    Text1(36).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(36)
    Text2(36).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
  
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If Val(imgBuscar(0).Tag) >= 0 Then
            'Se llama desde el bot�n de busqueda del campo Tipos de IVA
            'Recuperar solo el campo c�digo y Descripci�n
'            Indice = Val(Me.imgBuscar(0).Tag)
            Text1(35).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(35).Text = RecuperaValor(CadenaDevuelta, 2)
        Else
            'Recupera todo el registro de Art�culos
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim indice As Byte
Dim devuelve As String

    If CByte(Me.imgBuscar(0).Tag) = 9 Then indice = 4
    If indice = 4 Then 'Form Principal de Clientes
        Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
        'Poblacion
        Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve)
        'provincia
        Text1(indice + 2).Text = devuelve

    Else 'Lineas de Direcciones/Dptos
        Text3(3).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
        Text3(4).Text = ObtenerPoblacion(Text3(3).Text, devuelve)  'Poblacion
        'provincia
        Text3(5).Text = devuelve
    End If
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux1(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmE_DatoSeleccionado(CadenaSeleccion As String)
    'Formas de Env�o
    Text1(10).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(10)
    Text2(10).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    Text1(49).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(49).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim indice As Byte
    Select Case Val(imgFecha(0).Tag)
        Case 0
            indice = 13
        Case 1
            indice = 40
        Case 2
            indice = 41
        Case 3
            indice = 46
    End Select
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Formas de Pago
    Text1(23).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(23)
    Text2(23).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
'Situaciones
    Text1(42).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(42)
    Text2(42).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Tarifas
    Text1(37).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(37)
    Text2(37).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    'Disitnto de Observaciones
    If Index <> 11 Or Index <> 12 Then
        If Modo = 2 Or Modo = 0 Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Codigo Actividad
            indice = 9
            Set frmA = New frmFacActividades
            frmA.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 1  'Cod. Envio
            indice = 10
            Set frmE = New frmFacFormasEnvio
            frmE.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
            frmE.Show vbModal
            Set frmE = Nothing
            
            
            
        Case 4  'Cod. Forma de Pago
            indice = 23
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 5  'Cuenta Contable
            imgBuscar(0).Tag = Index
            MandaBusquedaPrevia "apudirec= 'S'"
            imgBuscar(0).Tag = -1
            indice = 35
            
        Case 6 'C�digo de Agente
            indice = 36
            Set frmAC = New frmFacAgentesCom
            frmAC.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
            frmAC.Show vbModal
            Set frmAC = Nothing
            
        Case 7 'C�digo de Tarifa
            indice = 37
            Set frmT = New frmFacTarifas
            frmT.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 8 'C�digo de Situaci�n
            indice = 42
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
             If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
            frmS.Show vbModal
            Set frmS = Nothing
            
        Case 9, 10 'CPostal
            Me.imgBuscar(0).Tag = Index
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            If Index = 9 Then
                indice = 4
            Else
                PonerFoco Text3(3)
            End If
            Me.imgBuscar(0).Tag = -1
            VieneDeBuscar = True
        Case 11
            If Modo = 5 Or Modo = 0 Then
            Else
                
                If Modo = 3 Or Modo = 4 Then
                    CadenaDesdeOtroForm = Text1(22).Text
                Else
                    CadenaDesdeOtroForm = ""
                    If Not Data1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Data1.Recordset!observac, "T")
                End If
                frmFacClienteObser.Modificar = Modo >= 3
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 3 Or Modo = 4 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(22).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
            End If
            
        Case 12
            If Modo = 5 Or Modo = 0 Then
            Else
                
                If Modo = 3 Or Modo = 4 Then
                    CadenaDesdeOtroForm = Text1(47).Text
                Else
                    CadenaDesdeOtroForm = ""
                    If Not Data1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Data1.Recordset!observa1, "T")
                End If
                frmFacClienteObser.Modificar = Modo >= 3
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 3 Or Modo = 4 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(47).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
            End If
            
        Case 13
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
        
            
    End Select
    If Index <> 10 Then PonerFoco Text1(indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then
        If Index <> 3 Then Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
     Case 0
        indice = 13
     Case 1
        indice = 40
     Case 2
        indice = 41
     Case 3
        indice = 46
   End Select
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   
   'Para la fecha de la navegacion
   If Index = 3 And Text1(46).Text <> "" Then
        imgFecha(3).Tag = Text1(46).Text
        CargaDatosLWDoc
    End If
End Sub

Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0: dirMail = Text1(17).Text
        Case 1: dirMail = Text1(21).Text
        Case 2: dirMail = Text3(9).Text
    End Select

    If LanzaMailGnral(dirMail) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgWeb_Click()
'Abrimos el explorador de windows con la pagina Web del cliente
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If LanzaHome("websoporte") Then espera 2
    If LanzaHomeGnral(Text1(8).Text) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub





Private Sub lw1_DblClick()
Dim Seleccionado As Long
    If Modo <> 2 Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub


    If Me.DatosADevolverBusqueda <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un cliente. No puede ver los documentos.", vbExclamation
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

    Case 0
        'OFERTAS
    Case 3
        'FACTURAS del cliente scafaccli (facturas de publicidad FPC y de llamadas FAC)
        'Este no necesitamos crear instancias
        
        'Lo que ocurre que esta preparado para abrir la factura a partir de un albaran, con lo cual
        'En la funcion abrir factura, buscare un albaran de la factura para abrirlo
        AbrirFacturaLW
        
        
    Case 4
        'Precios especiales
        'No creamos instancias

        frmFacPreciosEspecial.CadenaSituarData = "'" & DevNombreSQL(lw1.SelectedItem.Text) & "'|" & Data1.Recordset!CodClien & "|"
        frmFacPreciosEspecial.Show vbModal
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
Dim It As ListItem
Dim i As Integer
    If Modo <> 2 Then Exit Sub
    If lwCRM.ListItems.Count = 0 Then Exit Sub
    If lwCRM.SelectedItem Is Nothing Then Exit Sub




     'Llegados aqui
    Select Case CByte(RecuperaValor(lwCRM.Tag, 1))
    Case 0
        'Aciones comerciales
        ' modificar o insertar acciones comerciales
        frmCRMMto.DesdeElCliente = Data1.Recordset!CodClien
        frmCRMMto.TipoPredefinido = 0  'sin tipo predefinido
        frmCRMMto.DatosADevolverBusqueda = "`fechora`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & _
            " AND scrmacciones.Tipo = " & lwCRM.SelectedItem.SubItems(4) & " And codClien = " & Data1.Recordset!CodClien
        frmCRMMto.Show vbModal
    Case 1
        'Llamadas
        If lwCRM.SelectedItem.SmallIcon = 27 Then
            'Lee de sllama
            
            CadenaDesdeOtroForm = "`feholla`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & " and `usuario`=" & DBSet(lwCRM.SelectedItem.SubItems(1), "T")
            frmLLamadasDatos2.SoloVer = True
            frmLLamadasDatos2.vModo = 4
            frmLLamadasDatos2.Show vbModal
        Else
            'Lee de acciones realizadas con tipo=1 .....
            
            frmCRMMto.DesdeElCliente = Data1.Recordset!CodClien
            frmCRMMto.TipoPredefinido = 1 'Llamadas realizadas
            frmCRMMto.DatosADevolverBusqueda = "`fechora`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & " AND scrmacciones.Tipo = 1 And codClien = " & Data1.Recordset!CodClien
            frmCRMMto.Show vbModal
            
        End If
    Case 2
        'MAIL
        frmMensajes.OpcionMensaje = 21
        If lwCRM.SelectedItem.SmallIcon = 28 Then
            frmMensajes.cadWHERE2 = "0"
        Else
            frmMensajes.cadWHERE2 = "1"
        End If
        frmMensajes.cadWHERE = "codclien = " & Text1(0).Text & " AND  entryID = '" & lwCRM.SelectedItem.SubItems(5) & "'"
        frmMensajes.Show vbModal
    Case 3
        'Cobros. NO HACEMOS NADA
        'Nos piramos
        Exit Sub
        
    Case 4
        frmCrmObsDpto.Nuevo = False
        BuscaChekc = "dpto = " & Me.lwCRM.SelectedItem.SubItems(3) & " AND codclien "
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "observa", "scrmobsclien", BuscaChekc, CStr(Data1.Recordset!CodClien))
        
        frmCrmObsDpto.Dpto = CByte(Me.lwCRM.SelectedItem.SubItems(3))
        frmCrmObsDpto.Label2.Caption = Data1.Recordset!nomclien
        frmCrmObsDpto.Tag = Data1.Recordset!CodClien
        frmCrmObsDpto.Show vbModal
    End Select
    Me.Refresh
    DoEvents
    Set It = lwCRM.SelectedItem
    
    CargaDatosLWCRM
    Set lwCRM.SelectedItem = Nothing
    For i = 1 To lwCRM.ListItems.Count
        If It.Text = lwCRM.ListItems(i).Text Then
            Set lwCRM.SelectedItem = lwCRM.ListItems(i)
        Else
            lwCRM.ListItems(i).Selected = False
        End If
    Next
    Set It = Nothing
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
     If Modo = 5 Then 'Eliminar lineas Art�culos x Almacen
        If SSTab1.Tab = 2 Then
            BotonEliminarLinea
        Else
            BotonEliminarLinea2
        End If
     Else   'Eliminar Art�culo
        BotonEliminar
     End If
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnModificar_Click()
     If Modo = 5 Then 'Modificar lineas Art�culos x Almacen
        'FALTA: bloquear la linea !!!!
        BotonModificarLinea
     Else   'Modificar Art�culos
        If BLOQUEADesdeFormulario(Me, 1) Then BotonModificar
     End If
End Sub

Private Sub mnNuevo_Click()
     If Modo = 5 Then          'A�adir lineas Art�culos x Almacen
        Select Case SSTab1.Tab
            Case 2
               BotonAnyadirLinea
            Case 5
               BotonAnyadirLinea2
        End Select
    Else 'A�adir Art�culos
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



Private Sub Text1_Change(Index As Integer)
    If Index = 4 Then HaCambiadoCP = True 'CPostal ha cambiado
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index = 4 Then HaCambiadoCP = False
    If Index <> 22 Then ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 22 And KeyCode = 40 Then 'Flecha abajo
        Me.SSTab1.Tab = 1
        PonerFoco Text1(23)
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 22 Then KEYpress KeyAscii
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
Dim campo As String
Dim Codigo As String
Dim Tabla As String
Dim Titulo As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Cod Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
            
        Case 4 'CPostal
             If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, campo)
                Text1(Index + 2).Text = campo
             End If
             VieneDeBuscar = False
        
        Case 7 'NIF
            If Text1(Index).Text <> "" And Me.chkClienteV.Value = False Then
                Text1(Index).Text = UCase(Text1(Index).Text)
                ValidarNIF Text1(Index).Text
                If Modo = 3 And Text1(45).Text = "" Then Text1(45).Text = Text1(Index).Text
            End If
        
        Case 9 'Codigo de Actividad
            campo = "nomactiv"
            Codigo = "codactiv"
            Tabla = "sactiv"
            Titulo = "Actividades"
            
        Case 10 'C�digo de Env�o
            campo = "nomenvio"
            Codigo = "codenvio"
            Tabla = "senvio"
            Titulo = "Formas de Env�o"
            
                       

        Case 22 'Observaciones
            If Modo = 3 Or Modo = 4 Then 'Insertando o modificando
                'si se pierde el foco con un TAB y pasaria al siguiente campo que
                'esta en la otra pesta�a. si movemos foco a otro campo de la
                'misma pesta�a no cambiamos
                If Screen.ActiveControl.Name = "Text1" Then
                    If Screen.ActiveControl.Index = 23 Then
                        Me.SSTab1.Tab = 1
                        PonerFoco Text1(23)
                    End If
                End If
            End If

        Case 47 'Observaciones 1
            If Modo = 3 Or Modo = 4 Then 'Insertando o modificando
                'si se pierde el foco con un TAB y pasaria al siguiente campo que
                'esta en la otra pesta�a. si movemos foco a otro campo de la
                'misma pesta�a no cambiamos
                If Screen.ActiveControl.Name = "Text1" Then
                    If Screen.ActiveControl.Index = 47 Then
                        Me.SSTab1.Tab = 1
                        PonerFoco Text1(47)
                    End If
                End If
            End If



         Case 23 'Codigo Formas de pago
            campo = "nomforpa"
            Tabla = "sforpa"
            Codigo = "codforpa"
            Titulo = "Forma de Pago"
            
        Case 24, 25 'Descuento Pronto Pago, Descuento General
                'Formato tipo 4: Decimal(4,2)
            If Text1(Index).Text <> "" And Modo <> 1 Then PonerFormatoDecimal Text1(Index), 4
            
        Case 31, 32 'codbanco, sucursal
            PonerFormatoEntero Text1(Index)
            
        Case 35 'Cuenta contable
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(0).Text)
            
        Case 36 'Codigo Agente Comercial
            campo = "nomagent"
            Tabla = "sagent"
            Codigo = "codagent"
            Titulo = "Agente Comercial"
            
        Case 37 'Codigo Tarifa
            campo = "nomlista"
            Codigo = "codlista"
            Tabla = "starif"
            Titulo = "Tarifa"
                                    
        Case 13, 40, 41 'Fecha alta, Fecha �ltimo mov.,fecha reclamaci�n
             If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
             
        Case 42 'C�digo Situaci�n
            campo = "nomsitua"
            Codigo = "codsitua"
            Tabla = "ssitua"
            Titulo = "Situaci�n"
            
        Case 43 'L�mite Cr�dito
            'Formato tipo 1: Decimal(12,2)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 1
        
        Case 44 'Distancia Km
'            PonerFormatoDecimal Text1(Index), 5
            PonerFormatoEntero Text1(Index)
        Case 11
            If Text1(Index).Text <> "" Then
                PonerFormatoDecimal Text1(Index), 6
            End If
        Case 48 ' codigo de iban
            Text1(Index).Text = UCase(Text1(Index).Text)
            
        Case 49 'Banco Propio
            campo = "nombanpr"
            Tabla = "sbanpr"
            Codigo = "codbanpr"
            Titulo = "Banco Propio"
    
    
    
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 31 Or Index = 32 Or Index = 33 Or Index = 34 Then
        Dim cta As String
        Dim CC As String
        If Text1(31).Text <> "" And Text1(32).Text <> "" And Text1(33).Text <> "" And Text1(34).Text <> "" Then
            
            cta = Format(Text1(31).Text, "0000") & Format(Text1(32).Text, "0000") & Format(Text1(33).Text, "00") & Format(Text1(34).Text, "0000000000")
            If Len(cta) = 20 Then
                If Text1(48).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then Text1(48).Text = "ES" & cta
                Else
                    CC = CStr(Mid(Text1(48).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(Text1(48).Text, 3) <> cta Then
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
            End If
        End If
    End If
            
    If (Index = 9 Or Index = 10 Or Index = 12) Or Index = 23 Or Index = 36 Or Index = 37 Or Index = 42 Or Index = 49 Then
        If PonerFormatoEntero(Text1(Index)) Then
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, Tabla, campo, Codigo, Titulo)
            If Text2(Index).Text = "" Then PonerFoco Text1(Index)
        Else
            Text2(Index).Text = ""
        End If
    End If
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    cadB1 = ObtenerBusqueda(Me, True)
    
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
Dim Conexion As Byte

    'Llamamos a al form
    '##A mano
    cad = ""
    Select Case Val(Me.imgBuscar(0).Tag)
        Case 5  'Cuenta Contable
            'Se llama a Busqueda desde el campo Cuenta contable
            '#A MANO: Porque busca en la tabla cuentas
            'de la base de datos de Contabilidad
            cad = cad & "C�digo|cuentas|codmacta|T||30�Denominacion|cuentas|nommacta|T||70�"
            Tabla = "cuentas"
            Titulo = "Cuentas Contables"
            Conexion = conConta    'Conexi�n a BD: Conta
        Case Else   'Registro de la tabla de cabeceras: sartic
            cad = cad & ParaGrid(Text1(0), 10, "C�digo")
            cad = cad & ParaGrid(Text1(1), 50, "Nombre")
            cad = cad & ParaGrid(Text1(2), 40, "Nombre Comercial")
            Tabla = "scliente"
            Titulo = "Clientes"
            Conexion = conAri    'Conexi�n a BD: Aritaxi
    End Select
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = Conexion
        frmB.vCargaFrame = (Conexion = 2)
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
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
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        
        PonerCampos
        CargaFrameDirec
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    Text2(9).Text = PonerNombreDeCod(Text1(9), conAri, "sactiv", "nomactiv")
    Text2(10).Text = PonerNombreDeCod(Text1(10), conAri, "senvio", "nomenvio")
    Text2(23).Text = PonerNombreDeCod(Text1(23), conAri, "sforpa", "nomforpa")
    Text2(35).Text = PonerNombreDeCod(Text1(35), conConta, "cuentas", "nommacta")
    Text2(36).Text = PonerNombreDeCod(Text1(36), conAri, "sagent", "nomagent")
    Text2(37).Text = PonerNombreDeCod(Text1(37), conAri, "starif", "nomlista", "codlista")
    Text2(42).Text = PonerNombreDeCod(Text1(42), conAri, "ssitua", "nomsitua")
    Text2(49).Text = PonerNombreDeCod(Text1(49), conAri, "sbanpr", "nombanpr")
    
    BloquearChecks Me, Modo
    
    
'[Monica]24/09/2012:solapa de clientes de albaran
    'data2 para el grid de las lineas clientes de albaran
    Adodc2.ConnectionString = conn
    '[Monica]29/08/2013: a�ado el order by 3
    Adodc2.RecordSource = "select scliente_albaran.codclien,scliente_albaran.numlinea,scliente_albaran.codclienalb,scliente.nomclien from scliente_albaran inner join scliente on scliente_albaran.codclien=" & Text1(0).Text & " and scliente_albaran.codclienalb=scliente.codclien" & " order by 3 "
    Adodc2.Refresh
    
    CargaGrid DataGrid1, Adodc2

    
    lblIndicador.Caption = "Datos navegacion"
    Me.Refresh
    DoEvents
    CargaDatosLWDoc
    If vParamAplic.TieneCRM Then CargaDatosLWCRM
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

End Sub


Private Sub PonerCamposDirecciones()
Dim X As Boolean

    If Data2.Recordset.EOF Then Exit Sub
    
    X = PonerCamposFormaFrame(Me, "Text3", Data2)
    
    '-- Esto permanece para saber donde estamos
'    lblIndicador.Caption = Data2.Recordset.AbsolutePosition & " de " & Data2.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diversos campos del
'   formulario en funcion del modo en k vayamos a trabajar
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
    PonerIndicador Me.lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
        Me.cmdRegresar.Caption = "Regresar"
    Else
        cmdRegresar.visible = False
    End If
    
     'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
         
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    'El campo 46 NUNCA se puede escribir en el
    Text1(46).Enabled = False
    Text1(46).Text = Me.imgFecha(3).Tag
    
    'Bloquear los Text3
    For i = 0 To Me.Text3.Count - 1
        BloquearTxt Me.Text3(i), Not (Modo = 5)
    Next i
        
        
        
    Select Case Kmodo
        Case 2    'Preparamos para que pueda Modificar
            MostrarSituacion True
    
'        Case 5 'Lineas Direcciones/Departamentos
'             BloquearTxt Text3(0), True
    End Select
    
'    Me.FrameDirecciones.visible = (Modo = 5)
        
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    cboAlbaran.Enabled = b
    cboFacturacion.Enabled = b
    cboTipoIVA.Enabled = b
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    'Bloquear los checkbox
    BloquearChecks Me, Modo
    
    For i = 0 To Me.imgFecha.Count - 1
        If i <> 3 Then Me.imgFecha(i).Enabled = b
    Next i
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    imgBuscar(11).Enabled = Modo >= 2 And Modo < 5
    
    'CRM
    cmdAccCRM(0).visible = vParamAplic.TieneCRM And Modo = 2
    cmdAccCRM(1).visible = vParamAplic.TieneCRM And Modo = 2
    
    
    '-----------------------------
'    If (Modo = 5) Then 'Lineas Direcciones/Departamentos
''        PonerLineaVisible True
'        Me.Toolbar1.Buttons(10).Enabled = False
'    End If
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opcines de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
                        
    'El listview
    If Modo <> 2 Then
        lw1.ListItems.Clear
        If vParamAplic.TieneCRM Then lwCRM.ListItems.Clear
    End If

                        
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
Dim b1 As Boolean

    b = (Modo = 2 Or Modo = 0 Or (Modo = 5 And ModificaLineas = 0))
    'Insertar
    Toolbar1.Buttons(6).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b1 = (Modo = 2 Or (Modo = 5 And ModificaLineas = 0 And SSTab1.Tab <> 5))
    b = (Modo = 2 Or (Modo = 5 And ModificaLineas = 0))
    'Modificar
    Toolbar1.Buttons(7).Enabled = b1 And Not DeConsulta
    Me.mnModificar.Enabled = b1 And Not DeConsulta
    'eliminar
    Toolbar1.Buttons(8).Enabled = b And Not DeConsulta
    Me.mnEliminar.Enabled = b And Not DeConsulta
    
    'Lineas Direcciones/Departamentos
    Toolbar1.Buttons(10).Enabled = (Modo = 2) And Not DeConsulta
    
    'Imprimir
    Toolbar1.Buttons(11).Enabled = (Modo = 2) And Not DeConsulta
    Me.mnImprimir.Enabled = (Modo = 2) And Not DeConsulta
    
    '-----------------------------
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    
    
    'BARRA DE DIRECCIONES
    Me.ToolAux.visible = (Modo <> 0)
    If Me.ToolAux.visible Then Me.ToolAux.visible = (Me.Data2.Recordset.RecordCount > 0)
    If Me.ToolAux.visible Then
        b = Not (Modo = 5 And (ModoFrame = 3 Or ModoFrame = 4))
        Me.ToolAux.Buttons(1).Enabled = b
        Me.ToolAux.Buttons(2).Enabled = b
        Me.ToolAux.Buttons(3).Enabled = b
        Me.ToolAux.Buttons(4).Enabled = b
    End If
    
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub



Private Sub PonerModoFrame(Kmodo As Byte)
Dim i As Byte
On Error GoTo EPonerModoFr

    ModoFrame = Kmodo
    PonerModo 5
    
    If ModoFrame = 0 Then
        DesplazamientoVisible Me.ToolAux, 1, True, Data2.Recordset.RecordCount
    Else
        DesplazamientoVisible Me.Toolbar1, btnPrimero, False, 1
    
    End If
    
    If SSTab1.Tab = 2 Then
    
        'Bloquear TextBox sino modo 3 o 4
        For i = 0 To Me.Text3.Count - 1
            If ModoFrame = 3 Then Text3(i).Text = ""
            BloquearTxt Text3(i), (ModoFrame = 0)
        Next i
        
        'Si modo es modificar bloquear Clave Primaria
        If ModoFrame = 4 Then BloquearTxt Text3(0), True
        
        Select Case ModoFrame
            Case 0  'MODO INICIAL
                Me.imgBuscar(10).Enabled = False
                PonerBotonCabecera True
            Case 3, 4 'Modo INSERTAR o MODIFICAR
                '3=Insertar,  4=Modificar
                Me.imgBuscar(10).Enabled = True
                If Modo = 3 Then PonerFoco Text3(0)
                PonerBotonCabecera False
        End Select
    
'    Else
'        'Bloquear TextBox sino modo 3 o 4
'        For i = 0 To Me.txtAux1.Count - 1
'            If ModoFrame = 3 Then txtAux1(i).Text = ""
'            BloquearTxt txtAux1(i), (ModoFrame = 0)
'        Next i
'
'        'Si modo es modificar bloquear Clave Primaria
'        If ModoFrame = 4 Then BloquearTxt txtAux1(0), True
'
'        Select Case ModoFrame
'            Case 0  'MODO INICIAL
'                Me.cmdAux(0).Enabled = False
'                PonerBotonCabecera True
'            Case 3, 4 'Modo INSERTAR o MODIFICAR
'                '3=Insertar,  4=Modificar
'                Me.cmdAux(0).Enabled = True
'                If Modo = 3 Then PonerFoco txtAux1(0)
'                PonerBotonCabecera False
'        End Select
'
    End If
EPonerModoFr:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLineaVisible(bol As Boolean)
'bol=true : Se pone visible el frame ArticulosxAlmacen
'bol=false : se pone visible Datos Articulos
'On Error Resume Next
'
'    Me.frameComercial.visible = Not bol
'
'    Me.Label1(37).visible = Not bol 'Web
'    Me.Text1(8).visible = Not bol
'
'    Me.Label1(5).visible = Not bol 'Cod Actividad
'    Me.imgBuscar(0).visible = Not bol
'    Me.Text1(9).visible = Not bol
'    Me.Text2(0).visible = Not bol
'
'    Me.Label1(6).visible = Not bol 'Cod. Env�o
'    Me.imgBuscar(1).visible = Not bol
'    Me.Text1(10).visible = Not bol
'    Me.Text2(1).visible = Not bol
'
'    Me.Label1(7).visible = Not bol 'Cod. Zona
'    Me.imgBuscar(2).visible = Not bol
'    Me.Text1(11).visible = Not bol
'    Me.Text2(2).visible = Not bol
'
'    Me.Label1(17).visible = Not bol 'Cod Ruta
'    Me.imgBuscar(3).visible = Not bol
'    Me.Text1(12).visible = Not bol
'    Me.Text2(3).visible = Not bol
'    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim fec As Date
Dim cta As String
Dim cadMen As String


    On Error GoTo EDatosOK

    DatosOk = False
    
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
       
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then b = False
    End If
    If Not b Then Exit Function
    
    
    '- Validar que la cuenta bancaria es correcta
'[Monica]22/11/2013: iban
'    Comprueba_CuentaBan (Text1(31).Text & Text1(32).Text & Text1(33).Text & Text1(34).Text)
    If b And (Modo = 3 Or Modo = 4) Then
        
        
        '[Monica]22/08/2013: a�adida la comprobacion de que la cuenta contable sea correcta
        If Text1(31).Text = "" Or Text1(32).Text = "" Or Text1(33).Text = "" Or Text1(34).Text = "" Then
            '[Monica]20/11/2013: a�adido el codigo de iban
            Text1(48).Text = ""
            Text1(31).Text = ""
            Text1(32).Text = ""
            Text1(33).Text = ""
            Text1(34).Text = ""
        Else
            cta = Format(Text1(31).Text, "0000") & Format(Text1(32).Text, "0000") & Format(Text1(33).Text, "00") & Format(Text1(34).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El cliente no tiene asignada cuenta bancaria."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria del cliente no es correcta. � Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(31)
                    b = False
                End If
            Else
                BuscaChekc = ""
                If Me.Text1(48).Text <> "" Then BuscaChekc = Mid(Text1(48).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.Text1(48).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(48).Text = BuscaChekc & cta
                    Else
                        If Mid(Text1(48).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.Text1(48).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco Text1(48)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If


    '- comprobar q dia de vto atrasado tiene valor solo si mes a no girar tiene valor
    If Trim(Text1(26).Text) = "" And Trim(Text1(27).Text) <> "" Then
        b = False
        MsgBox "El d�a de Vto. atrasado solo debe tener valor si hay mes a no girar.", vbInformation
    ElseIf Trim(Text1(26).Text) <> "" And Trim(Text1(27).Text) <> "" Then
        If Trim(Text1(28).Text) <> "" Or Trim(Text1(29).Text) <> "" Or Trim(Text1(30).Text) <> "" Then
            b = False
            MsgBox "Si hay dias de pago no puede haber d�a de vto. atrasado.", vbInformation
        Else
            'comprobar q el dia de vto atrasado introducido existe para
            'el mes siguiente al mes a no girar
              If CInt(Text1(26).Text) + 1 < 13 Then
                If Not IsDate(Text1(27).Text & "/" & CInt(Text1(26).Text) + 1 & "/" & Year(Now)) Then
                    b = False
                    MsgBox "La fecha del dia de vto atrasado para el mes " & CInt(Text1(26).Text) + 1 & " NO es valida.", vbInformation
                End If
              Else
                If Not IsDate(Text1(27).Text & "/1/" & Year(Now) + 1) Then
                    b = False
                    MsgBox "La fecha del dia de vto atrasado para el mes 1" & " NO es valida.", vbInformation
                End If
              End If
        End If
    End If

    'QUito esto   11 Enero 09
    'Text1(22).Text = QuitarCaracterEnter(Text1(22))
    
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim devuelve As String
On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True
    
    If Text3(1).Text = "" Then 'Campo Nombre Direc./Dpto
        MsgBox "El campo Nombre no puede ser nulo", vbExclamation
        b = False
    End If

    If Text3(2).Text = "" Then 'Campo Domicilio Direc./Dpto
        MsgBox "El campo Domicilio no puede ser nulo", vbExclamation
        b = False
        If Not b Then Exit Function
    End If

    If Text3(3).Text = "" Then 'Campo CPostal Direc./Dpto
        MsgBox "El campo C.Postal no puede ser nulo", vbExclamation
        b = False
    End If
    
    If Text3(4).Text = "" Then 'Campo Poblaci�n Direc./Dpto
        MsgBox "El campo Poblaci�n no puede ser nulo", vbExclamation
        b = False
    End If
    
    If Text3(5).Text = "" Then 'Campo Provincia Direc./Dpto
        MsgBox "El campo Provincia no puede ser nulo", vbExclamation
        b = False
    End If
    If Not b Then Exit Function
    
    'Comprobamos  si ya existe Si estamos insertando
    'conAri: conexion a BD Aritaxi
    devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "coddirec", "codclien", Text1(0).Text, "N", , "coddirec", Text3(0).Text, "N")
    'If ModificaLineas = 1 And DevuelveExisteEnBD(conAri, "sdirec", "codclien", Text1(0).Text, "N", "coddirec", Text3(0).Text, "N") Then
    If ModificaLineas = 1 And devuelve <> "" Then
        b = False
        If vParamAplic.Departamento Then
            devuelve = " el Departamento "
        Else
            devuelve = " la Direcci�n "
        End If
        devuelve = "Ya existe" & devuelve & " del Cliente: " & vbCrLf
        devuelve = devuelve & "Codigo: " & Text3(0).Text & vbCrLf
        MsgBox devuelve, vbExclamation
    End If
    
    
    'comprobar los datos de la cuenta bancaria si param. de departamentos
    If Me.FrameCtaBanDpto.visible Then
        'Validar que la cuenta bancaria es correcta
        Comprueba_CuentaBan (Text3(10).Text & Text3(11).Text & Text3(12).Text & Text3(13).Text)
    End If
    
    
    
    
    
    
    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text3_Change(Index As Integer)
    If Index = 3 Then HaCambiadoCP = True
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    If Index = 3 Then HaCambiadoCP = False
    ConseguirFoco Text3(Index), 3
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        If (Index = 9 And Me.FrameCtaBanDpto.visible = False) Or Index = 13 Then
            PonerFocoBtn Me.cmdAceptar
        Else
            SendKeys "{tab}"
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Text3_LostFocus(Index As Integer)
Dim devuelve As String

    On Error Resume Next
    
    If Not PerderFocoGnralLineas(Text3(Index), ModificaLineas) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'Codigo Direc/Dpto
            If Trim(Text3(Index).Text) = "" Then Exit Sub
            FormateaCampo Text3(Index)

        Case 3 'Cod. Postal
            If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                Text3(Index + 1).Text = ObtenerPoblacion(Text3(Index).Text, devuelve)
                Text3(Index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
            
        Case 10, 11 'codbanco, sucursal
            PonerFormatoEntero Text3(Index)
            
        Case 12, 13 'DC, cta banco
            FormateaCampo Text3(Index)
            If Index = 13 Then PonerFocoBtn Me.cmdAceptar
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub ToolAux_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 To 4 'Flechas Desplazamiento
            DesplazamientoLineas (Button.Index - 1)
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 6  'Nuevo
           mnNuevo_Click
        Case 7  'Modificar
           mnModificar_Click
        Case 8  'Borrar
           mnEliminar_Click
        Case 10 'Direcciones/Departamentos
            BotonDirecciones
        Case 11 '
            BotonSubclientes
        Case 12 'Imprimir
            mnImprimir_Click
        Case 13 'Salir
            mnSalir_Click
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


Private Sub CargarComboAlbaran()
'### Combo Valorar Albaran con
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Todo, 1-Cantidad y Precio, 2-Cantidad

    cboAlbaran.Clear
    cboAlbaran.AddItem "Todo"
    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 0

    cboAlbaran.AddItem "Cantidad y Precio"
    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 1

    cboAlbaran.AddItem "Cantidad"
    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 2

End Sub


Private Sub CargarComboFacturacion()
'### Combo Tipo Facturaci�n
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


Private Sub CargarComboTipoIVA()
'### Combo Tipo de IVA a Aplicar
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Normal, 1-Con Recargo de Equivalencia, 2-Exento de IVA

    Me.cboTipoIVA.Clear
    cboTipoIVA.AddItem "Normal"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 0

    cboTipoIVA.AddItem "Recargo Equivalencia"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 1

    cboTipoIVA.AddItem "Exento de IVA"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 2

    cboTipoIVA.AddItem "Intracomunitario"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 3


End Sub

 
    
Private Function InsertarModificarLinea() As Boolean
Dim i As Byte
Dim Sql As String

    On Error GoTo EInsertarModificarLinea
    
    InsertarModificarLinea = False
    Sql = ""
    Select Case ModificaLineas
    Case 1  'INSERTAR
        If DatosOkLinea Then
            Sql = "INSERT INTO sdirec (codclien,coddirec,nomdirec,domdirec,codpobla,pobdirec,prodirec,perdirec,teldirec,faxdirec,maidirec,codbanco,codsucur,digcontr,cuentaba) VALUES ("
            Sql = Sql & Text1(0).Text & ", "
            Sql = Sql & Text3(0).Text
            For i = 1 To 5
                Sql = Sql & ", "
                Sql = Sql & DBSet(Text3(i).Text, "T")
            Next i
                    
            For i = 6 To 13 'campos opcionales
                Sql = Sql & ", "
                Sql = Sql & DBSet(Text3(i).Text, "T", "S")
'                If i <> 13 Then SQL = SQL & ", "
            Next i
                        
            Sql = Sql & ")"
        End If
        
    Case 2  'MODIFICAR
        If DatosOkLinea Then
            Sql = "UPDATE sdirec Set nomdirec = " & DBSet(Text3(1).Text, "T")
            Sql = Sql & ", domdirec = " & DBSet(Text3(2).Text, "T")
            Sql = Sql & ", codpobla = " & DBSet(Text3(3).Text, "T")
            Sql = Sql & ", pobdirec = " & DBSet(Text3(4).Text, "T")
            Sql = Sql & ", prodirec = " & DBSet(Text3(5).Text, "T")
            Sql = Sql & ", perdirec = " & DBSet(Text3(6).Text, "T")
            'If Text3(7).Text <> "" Then SQL = SQL & ", fechainv = '" & Format(Text3(7).Text, "yyyy-mm-dd") & "'"
            'If Text3(8).Text <> "" Then SQL = SQL & ", horainve = '" & Format(Text3(8).Text, "hh:mm:ss") & "'"
            Sql = Sql & ", teldirec = " & DBSet(Text3(7).Text, "T")
            Sql = Sql & ", faxdirec = " & DBSet(Text3(8).Text, "T")
            Sql = Sql & ", maidirec = " & DBSet(Text3(9).Text, "T")
            'datos cuenta bancaria
            If Me.FrameCtaBanDpto.visible Then
                Sql = Sql & ", codbanco = " & DBSet(Text3(10).Text, "N", "S")
                Sql = Sql & ", codsucur = " & DBSet(Text3(11).Text, "N", "S")
                Sql = Sql & ", digcontr = " & DBSet(Text3(12).Text, "T")
                Sql = Sql & ", cuentaba = " & DBSet(Text3(13).Text, "T")
            End If
            
            Sql = Sql & " WHERE codclien =" & (Text1(0).Text) & " AND "
            Sql = Sql & " coddirec =" & (Text3(0).Text)
        End If
    End Select
        
    If Sql <> "" Then
        conn.Execute Sql
        InsertarModificarLinea = True
        TratarDptoEnTesoreria   'TESOERIA
    Else
        PonerFoco Text3(1)
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Direcciones/Departamentos" & vbCrLf & Err.Description
End Function
    

Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Lineas Detalle"
        If Not Data2.Recordset.EOF Then Me.lblIndicador.Caption = Me.lblIndicador.Caption & ": " & Me.Data2.Recordset.AbsolutePosition & " de " & Me.Data2.Recordset.RecordCount
    End If
End Sub


Private Sub MostrarSituacion(vMostrar As Boolean)
Dim Codigo As Integer
Dim Bloquea As String
Dim DescBloqueo As String

    On Error GoTo EMostrarSitu

    If Data1.Recordset.EOF Then Exit Sub
    If vMostrar Then
        Codigo = Data1.Recordset!codsitua
        If Not IsNull(Codigo) Then
            Me.lblSituacion.visible = (Codigo <> 0)
            Me.Frame1(1).visible = (Codigo <> 0)
            If Not (Codigo = 0) Then
            'Si situacion=0 (activo) no mostrar situacion
                Bloquea = DevuelveDesdeBDNew(conAri, "ssitua", "tipositu", "codsitua", CStr(Codigo), "N")
                DescBloqueo = DevuelveDesdeBDNew(conAri, "ssitua", "nomsitua", "codsitua", CStr(Codigo), "N")
                If Val(Bloquea) = 0 Then
                    'Cliente NO Bloqueado
                    Me.lblSituacion.Caption = UCase(DescBloqueo)
                    Me.lblSituacion.ForeColor = vbBlue
                Else
                    'Cliente Bloqueado
                    Me.lblSituacion.Caption = "BLOQUEADO POR: " & UCase(DescBloqueo)
                    Me.lblSituacion.ForeColor = vbRed
                End If
            End If
        End If
    Else
        Me.lblSituacion.visible = False
        Me.Frame1(1).visible = False
    End If
EMostrarSitu:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PosicionarData()
Dim Indicador As String, cad As String

    cad = "(codclien=" & Val(Text1(0).Text) & ")"
    If SituarData(Data1, cad, Indicador) Then
'       PonerModo 2
       lblIndicador.Caption = Indicador
'[Monica]02/02/2017: daba error al dar de alta un cliente nuevo
'       '[Monica]24/09/2012:a�ado la linea siguiente
'       CargaGrid DataGrid1, Adodc2
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP & Ordenacion
        PonerCadenaBusqueda
    End If
    PonerModo 2
End Sub


Private Function ObtenerWhereCP() As String
On Error Resume Next
    ObtenerWhereCP = " WHERE  codclien= " & Val(Text1(0).Text)
End Function


Private Sub CargaFrameDirec()
Dim cadCli As String

    'Crear las lineas de Direcciones/Departamentos para el cliente
    'ASignamos un SQL al DATA2
    Me.Data2.ConnectionString = conn
    If Text1(0).Text = "" Then
        cadCli = -1
    Else
        cadCli = Val(Text1(0).Text)
    End If
    Data2.RecordSource = "Select * from sdirec where codclien = " & cadCli & ";"
    Data2.Refresh
    
    
    If Data2.Recordset.RecordCount > 0 Then
        Data2.Recordset.MoveFirst
        PonerCamposDirecciones
    Else
        LimpiarCamposDirecciones
    End If
    PonerModoOpcionesMenu
    DesplazamientoVisible Me.ToolAux, 1, True, Data2.Recordset.RecordCount
End Sub




'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'
'
'       El listview tendra los datos de albaranes, facturas... que tenga el cliente
'       Con lo cual, a partir de un click tendremos que ser capaces de situarnos en
'       el formulario correspondiente
'
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------


Private Sub ImagenesNavegacion()
    With Me.Toolbar2
        .ImageList = frmPpal.ImgListPpal
        .Buttons(1).Image = 7
        .Buttons(3).Image = 8
        
    End With
    
    Set lw1.SmallIcons = frmPpal.ImgListPpal
    
    SSTab1.TabVisible(4) = vParamAplic.TieneCRM
    If vParamAplic.TieneCRM Then
    
        With Me.Toolbar3
            .ImageList = frmPpal.ImgListPpal
            .Buttons(1).Image = 3
            .Buttons(3).Image = 30
            .Buttons(5).Image = 25
            .Buttons(7).Image = 13
            .Buttons(9).Image = 31
'            .Buttons(11).Image = 12
        End With
        
        Set lwCRM.SmallIcons = frmPpal.ImgListPpal
        
    End If
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
        'LLAMADAS
        LabelDoc.Caption = "Llamadas"
        Columnas = "Fecha|Hora|N� V|Tipo|Domicilio|Importe|"
        Ancho = "1400|600|600|500|3400|1300|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|1|0|0|1|"
        'Formatos
        Formato = "dd/mm/yyyy|hh:mm|0000|0||" & FormatoImporte & "|"
        Ncol = 6
    
    Case 3
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
    Text1(46).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")
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
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & LabelDoc.Caption
    lblIndicador.Refresh
    CargaDatosLWDoc2
    Me.lblIndicador.Caption = C
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
    Text1(46).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")
    EsDTOFam = False
    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 2
        'LLAMADAS
        cad = "select fecha,hora,numeruve,tipservi,dirllama,impventa from shilla WHERE 1=1"
        GroupBy = "1,2,3"
        BuscaChekc = "fecha"
        
    Case 3
        'FACTURAS
        cad = "select codtipom,numfactu,fecfactu,totalfac from scafaccli WHERE 1=1"
        BuscaChekc = "fecfactu"
        GroupBy = "1,2,3"
    End Select
    
    
    'Para todos menos para Dtofamila marca
    
    If Not EsDTOFam Then
            'EL where del codclien
        If cad <> "" Then
            cad = cad & " and codclien=" & Data1.Recordset!CodClien
            
            'La fecha
            If BuscaChekc <> "" Then cad = cad & " and " & BuscaChekc & " >='" & Format(imgFecha(3).Tag, FormatoFecha) & "'"
            
            
            'El group by
            If GroupBy <> "" Then cad = cad & " GROUP BY " & GroupBy
            
            'El ORDER BY
            'BuscaChekc="" si es la opcion de precios especiales
            If BuscaChekc = "" Then BuscaChekc = " codartic "
            cad = cad & " ORDER BY " & BuscaChekc & " DESC"
        End If
    
    Else
        'Para familia marca
        cad = cad & " (codclien=" & Data1.Recordset!CodClien & " AND codactiv is null)"
        cad = cad & " OR (codactiv = " & Data1.Recordset!codactiv & " AND codclien is null)"
    End If
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
        
            'Para familia /dto
            If EsDTOFam Then
                'Si codclien es >0 then
                If DBLet(RS!CodClien, "N") > 0 Then It.Bold = True
            End If
        
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
            Case "FPC"
                With frmPubliHcoFacCli
                        .DesdeFichaCliente = True
                        .hcoCodMovim = RecuperaValor(s, 2)
                        .hcoCodTipoM = RecuperaValor(s, 1)
                        .hcoFechaMov = RecuperaValor(s, 3)
                        .Show vbModal
                End With
            Case "FAC", "FRN", "FVC"
                With frmFCliHcoFac
                        .DesdeFichaCliente = True
                        .hcoCodMovim = RecuperaValor(s, 2)
                        .hcoCodTipoM = RecuperaValor(s, 1)
                        .hcoFechaMov = RecuperaValor(s, 3)
                        .Show vbModal
                End With
            
            
        End Select
            
    
    End If
End Sub


Private Function TratarDptoEnTesoreria() As Boolean
Dim Existe As Boolean
Dim C As String

    If Text1(35).Text = "" Or Text2(35).Text = "" Then
        
        MsgBox "Cuenta contable erronea.", vbExclamation
        Exit Function
    End If


    Existe = False
    C = "codmacta = '" & Text1(35).Text & "' and Dpto "
    C = DevuelveDesdeBD(conConta, "descripcion", "departamentos", C, Text3(0).Text)
    If C <> "" Then Existe = True
    
    
    If Existe Then
        If ModificaLineas = 1 Then
            'Estamos insertando y ya existe. UPDATEAMOS
            
        End If
        'UPDATEAMOS
        C = "UPDATE  departamentos set Descripcion = " & DBSet(Text3(1).Text, "T")
        C = C & " WHERE codmacta= '" & Text1(35).Text & "' AND Dpto = " & Text3(0).Text
    Else
        'NO EXISTE... creamos
        C = "insert into `departamentos` (`codmacta`,`Dpto`,`Descripcion`) values ('"
        C = C & Text1(35).Text & "'," & Text3(0).Text & "," & DBSet(Text3(1).Text, "T") & ")"
        
    End If
    ConnConta.Execute C
    
End Function


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
        LabelCRM.Caption = "Acciones comerciales"
        
        Columnas = "Fecha|Usuario|Estado|Medio|Tipo|Descripcion|"   'fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion
        Ancho = "2100|1000|1200|1200|800|2300|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm:ss||||0000||"
        Ncol = 6
               
    Case 1
        'Llamadas
        LabelCRM.Caption = "Llamadas "
        
        Columnas = "Fecha|Usuario|Tipo/Trab|Observaciones|Orden|"   'fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion
        Ancho = "2100|1000|1400|4000|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm:ss||||"
        Ncol = 5
    
        Ordena = 5
        
    Case 2
        LabelCRM.Caption = "E-mail"
        Columnas = "Fecha|Enviado|Email|Asunto|Adj|entryID|"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1800|825|2565|3899|495|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm||||||"
        Ncol = 6
    
    Case 3
        'COBROS
        LabelCRM.Caption = "Cobros pendientes"
        Columnas = "Fecha Vto.|Factura|Fecha factura|Forma pago|Pendiente|"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1600|1500|1300|2400|1495|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|1|0|0|1|"
        'Formatos
        Formato = "dd/mm/yyyy||dd/mm/yyyy||" & FormatoImporte & "|"
        Ncol = 5
        
    Case 4
        'COBROS
        LabelCRM.Caption = "Observaciones departamento"
        Columnas = "Departamento|Fecha|Observaciones||"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1600|1500|6500|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|"
        'Formatos
        Formato = "|dd/mm/yyyy|||"
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
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & LabelCRM.Caption
    lblIndicador.Refresh
    CargaDatosLWcrm2
    Me.lblIndicador.Caption = C
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
    Text1(46).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")

    'EL where del codclien     lo lleva cada sql
    Kopc = CByte(RecuperaValor(lwCRM.Tag, 1))
    Select Case Kopc
    Case 0
        'Acciones comerciales
        cad = "select fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion from scrmacciones,scrmtipo WHERE scrmacciones.tipo= scrmtipo.codigo "
        cad = cad & " and codclien=" & Data1.Recordset!CodClien & " and tipo > 20"  'las 20 primerasprobablemebne no sepongan aqui
        GroupBy = ""
        BuscaChekc = "fechora"
    Case 1
        'Llamadas
        cad = "select feholla,usuario,nomllama1,observac,date_format(feholla,""%Y%m%d%H%i%s"") from sllama,sllama1  where"
        cad = cad & " sllama.codllama1 = sllama1.codllama1"
        cad = cad & " and codclien=" & Data1.Recordset!CodClien
        GroupBy = ""
        BuscaChekc = "feholla"
    
    Case 2
    
        'eMAIL
        cad = "select fechahora, if(enviado=1,""Enviado"",""Recibido""),email,asunto,"
        cad = cad & "if(adjuntos<>"""",""*"","""") ,entryID from scrmmail"
        cad = cad & " WHERE codclien=" & Data1.Recordset!CodClien
        GroupBy = ""
        BuscaChekc = "fechahora"
        
    Case 3
        'Cobros pendientes
        If vParamAplic.ContabilidadNueva Then
            cad = "SELECT fecvenci,concat(numserie,right(concat(""00000000"",codfaccl),7)),fecfactu,nomforpa,"
            cad = cad & "impvenci+if(gastos is null,0,gastos)-if(impcobro is null,0,impcobro)  tot"
            cad = cad & " FROM  cobros INNER JOIN formapago ON cobros.codforpa=formapago.codforpa "
            cad = cad & " WHERE cobros.codmacta = '" & Text1(35).Text & "' AND (formapago.tipforpa between 0 and 3) "
        Else
            cad = "SELECT fecvenci,concat(numserie,right(concat(""00000000"",codfaccl),7)),fecfaccl,nomforpa,"
            cad = cad & "impvenci+if(gastos is null,0,gastos)-if(impcobro is null,0,impcobro)  tot"
            cad = cad & " FROM  scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
            cad = cad & " WHERE scobro.codmacta = '" & Text1(35).Text & "' AND (sforpa.tipforpa between 0 and 3) "
        End If
        BuscaChekc = "fecvenci"
        
    Case 4
        'Observaciones departamento
        cad = "select if(dpto=1,""Administracion"",if(dpto=2,""Comercial"",""SAT"")),fecha,observa,dpto from scrmobsclien"
        cad = cad & " WHERE codclien=" & Data1.Recordset!CodClien
        BuscaChekc = "dpto"
        
    End Select
    
    
    
    
    'El group by
    If GroupBy <> "" Then cad = cad & " GROUP BY " & GroupBy
    
    'El ORDER BY
    cad = cad & " ORDER BY " & BuscaChekc
     If Kopc <> 4 Then cad = cad & " DESC"

    
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
                            If NumRegElim = 3 And Kopc = 0 Then DevuelveMedio cad
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
    
    
    If Kopc = 1 Then
        'Llamadas. Las efectuadas las hago desde este punto
        cad = "select fechora ,usuario,nomtraba ,observaciones,date_format(fechora,""%Y%m%d%H%i%s"") from"
        cad = cad & " scrmacciones left join straba on scrmacciones.codtraba=straba.codtraba "
        cad = cad & " WHERE scrmacciones.tipo=1  and codclien= " & Data1.Recordset!CodClien
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            '
            'Coje datos desde dos tablas
            Set It = lwCRM.ListItems.Add()
            It.Text = Format(RS.Fields(0), lwCRM.ColumnHeaders(1).Tag)
           
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
  
                        It.SubItems(NumRegElim - 1) = cad
                    
                        
                        
                    End If
                End If
            Next
            It.SmallIcon = 26
            RS.MoveNext
        Wend
        RS.Close
    End If
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


Private Sub LanzarProgramaEmails()
Dim TieneDatosDpto As Boolean

    On Error GoTo ELanzarProgramaEmails

    If Dir(App.Path & "\AriOutlook.exe", vbArchive) = "" Then
        MsgBox "No tienen el programa de asignacion de mails al CRM de Ariadna", vbExclamation
        Exit Sub
    End If
    
    TieneDatosDpto = False
    If Not Data2.Recordset Is Nothing Then
        If Not Data2.Recordset.EOF Then TieneDatosDpto = True
    End If
        
        
    'Como lanzamos el programa
    '*************************  dbaritaxi|codclien|nombre||||mails que se utlizan|
    If TieneDatosDpto Then
        BuscaChekc = "Select distinct(maidirec) from sdirec where codclien=" & Data1.Recordset!CodClien & " AND maidirec <>"""""
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open BuscaChekc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    
    BuscaChekc = ""
    If Text1(17).Text <> "" Then BuscaChekc = BuscaChekc & Text1(17).Text & "|"  'mail1
    If Text1(18).Text <> "" Then BuscaChekc = BuscaChekc & Text1(18).Text & "|"  'mail1
        
        
    If TieneDatosDpto Then
        While Not miRsAux.EOF
            If Not IsNull(miRsAux!maidirec) Then
                If miRsAux!maidirec <> "" Then BuscaChekc = BuscaChekc & miRsAux!maidirec & "|"
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    BuscaChekc = vUsu.CadenaConexion & "|" & Data1.Recordset!CodClien & "|" & Data1.Recordset!nomclien & "||||" & BuscaChekc
    
    Shell App.Path & "\AriOutlook.exe " & BuscaChekc, vbNormalFocus
    
    Espera 2
    
    
ELanzarProgramaEmails:
    If Err.Number <> 0 Then MuestraError Err.Number, "Lanzar Programa Email"
    Set miRsAux = Nothing
    BuscaChekc = ""
End Sub


Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "scliente"
        .Informe2 = "rManClientes.rpt"
        If cadB1 <> "" Then
            .cadRegSelec = cadB1 'SQL2SF(cadB1)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(Data1, Me)
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




'************************************
Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)   ', enlaza As Boolean)
Dim i As Integer
Dim Sql As String

On Error GoTo ECargaGrid

    vData.Refresh
    Set vDataGrid.DataSource = vData
    vDataGrid.Columns(0).visible = False 'codclien

    If vDataGrid.Name = "DataGrid1" Then
        vDataGrid.Columns(1).visible = False 'numlinea
        vDataGrid.Columns(2).Caption = "Cliente Albaran"
        vDataGrid.Columns(2).Width = 1800
        vDataGrid.Columns(2).NumberFormat = "000000"
        vDataGrid.Columns(3).Caption = "Nombre"
        vDataGrid.Columns(3).Width = 3800
        
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



Private Function InsertarLinea() As Boolean
Dim Sql As String
Dim vWhere As String
Dim numF As String
On Error GoTo EInsertarLinea

    conn.BeginTrans

    InsertarLinea = False
    Sql = ""
    If DatosOkLinea2 Then
        vWhere = "codclien=" & Val(Text1(0).Text)
        numF = SugerirCodigoSiguienteStr("scliente_albaran", "numlinea", vWhere)
        Sql = "INSERT INTO scliente_albaran "
        Sql = Sql & "(codclien, numlinea, codclienalb) "
        Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
        Sql = Sql & DBSet(txtAux1(0).Text, "N") & ")"
        
        conn.Execute Sql
     
        InsertarLinea = True
    End If
    conn.CommitTrans
    Exit Function
EInsertarLinea:
    conn.RollbackTrans
    MuestraError Err.Number, "Insertar Lineas Clientes Albaran" & vbCrLf & Err.Description
End Function

Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim Sql As String
On Error GoTo EModificarLinea

    ModificarLinea = False
    Sql = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        Sql = "UPDATE scliente_albaran Set codclienalb = " & txtAux1(0).Text
        Sql = Sql & " where codclien=" & Adodc2.Recordset!CodClien & " AND numlinea=" & Adodc2.Recordset!numlinea
        
        conn.Execute Sql
        
        ModificarLinea = True
    End If
    
    Exit Function

EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Clientes Albar�n" & vbCrLf & Err.Description
End Function



Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
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
        
        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 20)
        
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).Top = alto
            txtAux1(i).Height = DataGrid1.RowHeight
        Next i
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'cliente
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
        
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).visible = visible
        Next i
        Me.cmdAux(0).Height = Me.DataGrid1.RowHeight
        Me.cmdAux(0).Top = alto
        Me.cmdAux(0).visible = visible
'        cmdAux1.Top = alto
'        cmdAux1.visible = visible
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
        Case 0
            If txtAux1(Index).Text <> "" Then
                txtAux1(Index).Text = Format(txtAux1(Index).Text, "000000")
                encontrado = DevuelveDesdeBD(conAri, "nomclien", "scliente", "codclien", txtAux1(Index).Text, "N")
                If encontrado <> "" Then
                    txtAux1(1).Text = encontrado
                    
                    cmdAceptar.SetFocus
                Else
                    MsgBox "No existe el c�digo de cliente introducido.", vbExclamation
                    PonerFoco txtAux1(Index)
                End If
            End If
    End Select

End Sub



Private Function DatosOkLinea2() As Boolean
Dim b As Boolean
Dim devuelve As String
On Error GoTo EDatosOkLinea

    DatosOkLinea2 = False
    b = True
    
    If txtAux1(0).Text = "" Then 'subcliente
        MsgBox "El campo Subcliente no puede ser nulo", vbExclamation
        b = False
    End If
    If Not b Then Exit Function
    
    'Comprobamos  si ya existe Si estamos insertando
    'conAri: conexion a BD Aritaxi
    devuelve = DevuelveDesdeBDNew(conAri, "scliente_albaran", "codclien", "codclien", Text1(0).Text, "N", , "codclienalb", txtAux1(0).Text, "N")
    'If ModificaLineas = 1 And DevuelveExisteEnBD(conAri, "sdirec", "codclien", Text1(0).Text, "N", "coddirec", Text3(0).Text, "N") Then
    If devuelve <> "" Then
        b = False
        devuelve = "Ya existe como subcliente del Cliente: " & vbCrLf
        devuelve = devuelve & "Codigo: " & txtAux1(0).Text & vbCrLf
        MsgBox devuelve, vbExclamation
    End If
    If Not b Then Exit Function
    
    'Comprobamos que no existe como subcliente en otro cliente
    'conAri: conexion a BD Aritaxi
    devuelve = DevuelveDesdeBDNew(conAri, "scliente_albaran", "codclien", "codclienalb", txtAux1(0).Text, "N")
    'If ModificaLineas = 1 And DevuelveExisteEnBD(conAri, "sdirec", "codclien", Text1(0).Text, "N", "coddirec", Text3(0).Text, "N") Then
    If devuelve <> "" Then
        b = False
        devuelve = "Ya existe como subcliente de otro Cliente: " & vbCrLf
        devuelve = devuelve & "Codigo: " & txtAux1(0).Text & vbCrLf
        MsgBox devuelve, vbExclamation
    End If
    If Not b Then Exit Function
    
    DatosOkLinea2 = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub BotonEliminarLinea2()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String, cad2 As String
Dim i As Integer
Dim Sql As String




'    If Adodc2.Recordset.EOF Then Exit Sub
'    If Adodc2.Recordset.RecordCount < 1 Then Exit Sub
'
'    'Si no estaba modificando lineas salimos
'    'Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
'
'    ModificaLineas = 3 'Eliminar
    
    cad2 = " SubCliente. "
    cad = " el SubCliente?"
    
    cad = "�Seguro que desea eliminar " & cad & vbCrLf
    cad = cad & vbCrLf & "Cod." & cad2 & ": " & Format(Me.Adodc2.Recordset.Fields(2), FormatoCampo(txtAux1(0)))
    cad = cad & vbCrLf & "Nombre" & cad2 & ": " & Adodc2.Recordset.Fields(3)
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Adodc2.Recordset.AbsolutePosition
        
        Sql = "delete from scliente_albaran where codclien = " & Me.Adodc2.Recordset.Fields(0) & " and numlinea = " & Me.Adodc2.Recordset.Fields(1)
        conn.Execute Sql
        
        CargaGrid DataGrid1, Me.Adodc2
    End If
    
    Screen.MousePointer = vbDefault
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub

