VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAlmArticulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Art�culos"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   12510
   Icon            =   "frmAlmArticulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   9780
      TabIndex        =   157
      Top             =   270
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3900
      TabIndex        =   148
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   149
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
      Left            =   240
      TabIndex        =   146
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   147
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
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar denominaci�n"
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
      Left            =   5760
      TabIndex        =   105
      Top             =   7680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5970
      Left            =   240
      TabIndex        =   43
      Top             =   1560
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   10530
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
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
      TabCaption(0)   =   "Datos b�sicos   "
      TabPicture(0)   =   "frmAlmArticulos.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgCuentas(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(17)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgCuentas(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgCuentas(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "imgCuentas(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgCuentas(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgCuentas(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(9)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(20)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(19)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(4)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(16)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "imgFecha(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblSumaStocks"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(37)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "FrameLitrosUd"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "FrameDatosAlmacen2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chkSeries"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "chkConjunto"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text2(3)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(5)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(2)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(3)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(7)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(4)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text2(2)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text2(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text2(1)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text2(0)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text2(4)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(6)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(12)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(11)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(9)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cboStatus"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text1(10)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtSumaStock"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "chkCtrStock"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "chkMateriaPrima"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text1(8)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).ControlCount=   45
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmAlmArticulos.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "framePortes"
      Tab(1).Control(1)=   "Text1(22)"
      Tab(1).Control(2)=   "FrameServicios"
      Tab(1).Control(3)=   "Text1(19)"
      Tab(1).Control(4)=   "Text1(20)"
      Tab(1).Control(5)=   "Text1(21)"
      Tab(1).Control(6)=   "Label2(1)"
      Tab(1).Control(7)=   "Label2(11)"
      Tab(1).Control(8)=   "Label2(2)"
      Tab(1).Control(9)=   "Label2(3)"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Componentes"
      TabPicture(2)   =   "frmAlmArticulos.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameToolAux(2)"
      Tab(2).Control(1)=   "Data2"
      Tab(2).Control(2)=   "txtAux(6)"
      Tab(2).Control(3)=   "cmdActualizarImportes1(1)"
      Tab(2).Control(4)=   "cmdActualizarImportes1(0)"
      Tab(2).Control(5)=   "txtConjunto(5)"
      Tab(2).Control(6)=   "txtConjunto(4)"
      Tab(2).Control(7)=   "txtConjunto(3)"
      Tab(2).Control(8)=   "txtConjunto(2)"
      Tab(2).Control(9)=   "txtConjunto(1)"
      Tab(2).Control(10)=   "txtConjunto(0)"
      Tab(2).Control(11)=   "txtAux(5)"
      Tab(2).Control(12)=   "txtAux(4)"
      Tab(2).Control(13)=   "txtAux(3)"
      Tab(2).Control(14)=   "txtAux2"
      Tab(2).Control(15)=   "txtAux(1)"
      Tab(2).Control(16)=   "txtAux(0)"
      Tab(2).Control(17)=   "cmdAux"
      Tab(2).Control(18)=   "DataGrid1"
      Tab(2).Control(19)=   "Line5"
      Tab(2).Control(20)=   "Label5(5)"
      Tab(2).Control(21)=   "Label5(4)"
      Tab(2).Control(22)=   "Label5(3)"
      Tab(2).Control(23)=   "Label5(2)"
      Tab(2).Control(24)=   "Label5(1)"
      Tab(2).Control(25)=   "Label5(0)"
      Tab(2).Control(26)=   "Line4"
      Tab(2).ControlCount=   27
      TabCaption(3)   =   "Control instalaci�n / producci�n"
      TabPicture(3)   =   "frmAlmArticulos.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameToolAux(3)"
      Tab(3).Control(1)=   "txtAux(2)"
      Tab(3).Control(2)=   "DataGrid2"
      Tab(3).Control(3)=   "Data3"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Stocks"
      TabPicture(4)   =   "frmAlmArticulos.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameToolAux(0)"
      Tab(4).Control(1)=   "cmdAlma"
      Tab(4).Control(2)=   "Text3(0)"
      Tab(4).Control(3)=   "Text2(8)"
      Tab(4).Control(4)=   "Text3(1)"
      Tab(4).Control(5)=   "FrameArtxAlmac"
      Tab(4).Control(6)=   "DataGrid3"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "Cod. EAN"
      TabPicture(5)   =   "frmAlmArticulos.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FrameToolAux(1)"
      Tab(5).Control(1)=   "txtAux(7)"
      Tab(5).Control(2)=   "DataGrid4"
      Tab(5).Control(3)=   "Data5"
      Tab(5).Control(4)=   "Label2(4)"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Documentos"
      TabPicture(6)   =   "frmAlmArticulos.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame4"
      Tab(6).Control(1)=   "FrameDisponible"
      Tab(6).Control(2)=   "lw1"
      Tab(6).Control(3)=   "Label2(0)"
      Tab(6).ControlCount=   4
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   3
         Left            =   -74280
         TabIndex        =   155
         Top             =   390
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   2
            Left            =   210
            TabIndex        =   156
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   2
         Left            =   -74760
         TabIndex        =   154
         Top             =   360
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   1
            Left            =   270
            TabIndex        =   158
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   1
         Left            =   -74280
         TabIndex        =   152
         Top             =   420
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   3
            Left            =   210
            TabIndex        =   153
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   0
         Left            =   -74760
         TabIndex        =   150
         Top             =   390
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   210
            TabIndex        =   151
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
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
         Height          =   290
         Index           =   7
         Left            =   -74280
         MaxLength       =   60
         TabIndex        =   143
         Text            =   "Dat"
         Top             =   4440
         Visible         =   0   'False
         Width           =   2595
      End
      Begin MSAdodcLib.Adodc Data2 
         Height          =   330
         Left            =   -66240
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74880
         TabIndex        =   141
         Top             =   600
         Width           =   855
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   3030
            Left            =   120
            TabIndex        =   142
            Top             =   0
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   5345
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Appearance      =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   11
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Tarifas"
                  Object.Tag             =   "0"
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Precios especiales"
                  Object.Tag             =   "1"
                  Style           =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Promociones"
                  Object.Tag             =   "2"
                  Style           =   2
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Pedidos"
                  Object.Tag             =   "3"
                  Style           =   2
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Precios especiales"
                  Style           =   2
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Movimientos"
                  Object.Tag             =   "4"
                  Style           =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameDisponible 
         Height          =   2295
         Left            =   -66840
         TabIndex        =   131
         Top             =   3000
         Width           =   3195
         Begin VB.TextBox Text4 
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
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   135
            Text            =   "Text4"
            Top             =   240
            Width           =   1635
         End
         Begin VB.TextBox Text4 
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
            Index           =   1
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   134
            Text            =   "Text4"
            Top             =   720
            Width           =   1635
         End
         Begin VB.TextBox Text4 
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
            Index           =   2
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   133
            Text            =   "Text4"
            Top             =   1200
            Width           =   1635
         End
         Begin VB.TextBox Text4 
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
            Index           =   3
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   132
            Text            =   "Text4"
            Top             =   1800
            Width           =   1635
         End
         Begin VB.Label Label4 
            Caption         =   "Reservas"
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
            Left            =   240
            TabIndex        =   139
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Pedidos"
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
            Left            =   240
            TabIndex        =   138
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Stock"
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
            Left            =   240
            TabIndex        =   137
            Top             =   240
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   3000
            Y1              =   1650
            Y2              =   1650
         End
         Begin VB.Label Label4 
            Caption         =   "Disponible"
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
            Left            =   240
            TabIndex        =   136
            Top             =   1860
            Width           =   1005
         End
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
         Index           =   8
         Left            =   9390
         MaxLength       =   10
         TabIndex        =   28
         Tag             =   "Num. orden|N|S|||sartic|numorden|||"
         Text            =   "Text1"
         Top             =   1590
         Width           =   1335
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   290
         Index           =   6
         Left            =   -66480
         TabIndex        =   127
         Tag             =   "C|T|S|||||||"
         Text            =   "Dato2"
         ToolTipText     =   "Materia prima"
         Top             =   2880
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CheckBox chkMateriaPrima 
         Caption         =   "Materia prima"
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
         Left            =   9390
         TabIndex        =   126
         Tag             =   "Materia prima|N|N|0|1|sartic|mateprima||N|"
         Top             =   4200
         Width           =   1875
      End
      Begin VB.Frame framePortes 
         Caption         =   "Portes"
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
         Left            =   -68520
         TabIndex        =   124
         Top             =   4200
         Width           =   3975
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
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   34
            Tag             =   "Kilos|N|S|||sartic|pesoarti|#,##0.00||"
            Text            =   "Tex"
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Kilos"
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
            Index           =   36
            Left            =   1920
            TabIndex        =   125
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdActualizarImportes1 
         Height          =   375
         Index           =   1
         Left            =   -65640
         Picture         =   "frmAlmArticulos.frx":00D0
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Modificar componente"
         Top             =   5310
         Width           =   375
      End
      Begin VB.CommandButton cmdActualizarImportes1 
         Height          =   375
         Index           =   0
         Left            =   -65040
         Picture         =   "frmAlmArticulos.frx":0AD2
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Actualizar importes"
         Top             =   5310
         Width           =   375
      End
      Begin VB.TextBox txtConjunto 
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
         Index           =   5
         Left            =   -67680
         TabIndex        =   119
         Text            =   "Text5"
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
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
         Index           =   4
         Left            =   -69000
         TabIndex        =   117
         Text            =   "Text5"
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
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
         Index           =   3
         Left            =   -70320
         TabIndex        =   115
         Text            =   "Text5"
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
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
         Index           =   2
         Left            =   -71910
         TabIndex        =   113
         Text            =   "Text5"
         Top             =   5520
         Width           =   1335
      End
      Begin VB.TextBox txtConjunto 
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
         Index           =   1
         Left            =   -73230
         TabIndex        =   111
         Text            =   "Text5"
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
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
         Index           =   0
         Left            =   -74760
         TabIndex        =   109
         Text            =   "Text5"
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   290
         Index           =   5
         Left            =   -67200
         TabIndex        =   108
         Tag             =   "C|N|S|||||###,##0.0000||"
         Text            =   "Dato2"
         Top             =   3180
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   290
         Index           =   4
         Left            =   -65880
         TabIndex        =   107
         Tag             =   "C|N|S|||||###,##0.0000||"
         Text            =   "Dato2"
         Top             =   3180
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   290
         Index           =   3
         Left            =   -68640
         TabIndex        =   106
         Tag             =   "C|N|S|||||###,##0.0000||"
         Text            =   "Dato2"
         Top             =   3180
         Visible         =   0   'False
         Width           =   1395
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   30
         Tag             =   "Taux|T|S|||sartic|txtauxdocumento|||"
         Top             =   4680
         Width           =   6015
      End
      Begin VB.Frame FrameServicios 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2055
         Left            =   -68640
         TabIndex        =   100
         Top             =   360
         Width           =   4575
      End
      Begin VB.CommandButton cmdAlma 
         Caption         =   "+"
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
         Left            =   -74040
         TabIndex        =   99
         Top             =   3630
         Visible         =   0   'False
         Width           =   255
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
         Index           =   0
         Left            =   -74760
         MaxLength       =   8
         TabIndex        =   81
         Tag             =   "C�digo Almacen|N|N|||salmac|codalmac|0|S|"
         Text            =   "Text3"
         Top             =   3600
         Visible         =   0   'False
         Width           =   780
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
         Index           =   8
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   97
         Text            =   "Text2"
         Top             =   3600
         Visible         =   0   'False
         Width           =   3075
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
         Left            =   -70800
         MaxLength       =   16
         TabIndex        =   82
         Tag             =   "Cantidad Stock|N|N|||salmac|canstock|#,###,###,##0.00|N|"
         Text            =   "Text3"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Frame FrameArtxAlmac 
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
         Height          =   4545
         Left            =   -68640
         TabIndex        =   79
         Top             =   960
         Width           =   4455
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
            Height          =   315
            Index           =   2
            Left            =   2760
            MaxLength       =   16
            TabIndex        =   84
            Tag             =   "Stock M�nimo|N|S|||salmac|stockmin|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   960
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
            Height          =   315
            Index           =   3
            Left            =   2760
            MaxLength       =   16
            TabIndex        =   85
            Tag             =   "Punto de Pedido|N|S|||salmac|puntoped|#,###,###,##0.00|N|"
            Text            =   "Text3"
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
            Index           =   5
            Left            =   2760
            MaxLength       =   16
            TabIndex        =   83
            Tag             =   "Stock inventario|N|S|||salmac|stockinv|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   3480
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
            Index           =   6
            Left            =   240
            MaxLength       =   10
            TabIndex        =   87
            Tag             =   "Fecha inventario|F|S|||salmac|fechainv|dd/mm/yyyy|N|"
            Text            =   "Text3"
            Top             =   3480
            Width           =   1240
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
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   88
            Tag             =   "Hora Inventario|H|S|||salmac|horainve|hh:mm:ss|N|"
            Text            =   "Text3"
            Top             =   3480
            Width           =   1125
         End
         Begin VB.CheckBox chkInventario 
            Height          =   195
            Left            =   240
            TabIndex        =   90
            Top             =   4080
            Width           =   255
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
            Height          =   315
            Index           =   4
            Left            =   2760
            MaxLength       =   16
            TabIndex        =   86
            Tag             =   "Stock M�ximo|N|S|||salmac|stockmax|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "INVENTARIO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   145
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   4320
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label3 
            Caption         =   "Realizando Inventario"
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
            Left            =   600
            TabIndex        =   89
            Top             =   4080
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Stock M�nimo"
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
            Left            =   240
            TabIndex        =   96
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Punto de Pedido"
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
            Left            =   240
            TabIndex        =   95
            Top             =   1500
            Width           =   2025
         End
         Begin VB.Label Label1 
            Caption         =   "Stock "
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
            Left            =   3600
            TabIndex        =   94
            Top             =   3240
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Hora "
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
            Index           =   30
            Left            =   1560
            TabIndex        =   92
            Top             =   3240
            Width           =   495
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   1200
            ToolTipText     =   "Buscar fecha"
            Top             =   3240
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Stock M�ximo"
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
            TabIndex        =   91
            Top             =   1920
            Width           =   1785
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   6
            Left            =   4080
            ToolTipText     =   "Buscar almacen"
            Top             =   4080
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha "
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
            Left            =   240
            TabIndex        =   93
            Top             =   3240
            Width           =   735
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
         Height          =   975
         Index           =   19
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Tag             =   "Texto para Ventas|T|S|||sartic|textoven|||"
         Top             =   840
         Width           =   6015
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
         Height          =   855
         Index           =   20
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Tag             =   "Texto para compras|T|S|||sartic|textocom|||"
         Top             =   2160
         Width           =   6015
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
         Height          =   855
         Index           =   21
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Tag             =   "Control de instalaci�n|T|S|||sartic|controli|||"
         Top             =   3340
         Width           =   6015
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
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
         Height          =   360
         Index           =   2
         Left            =   -74040
         MaxLength       =   60
         TabIndex        =   65
         Text            =   "Dat"
         Top             =   2880
         Visible         =   0   'False
         Width           =   7035
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Height          =   290
         Left            =   -73200
         TabIndex        =   63
         Text            =   "Dato2"
         Top             =   3180
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   290
         Index           =   1
         Left            =   -70320
         TabIndex        =   62
         Tag             =   "C|N|N|||||###,##0.0000||"
         Text            =   "Dato2"
         Top             =   3180
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
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
         Height          =   290
         Index           =   0
         Left            =   -74280
         TabIndex        =   61
         Text            =   "Dat"
         Top             =   3180
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
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
         Left            =   -73440
         TabIndex        =   60
         Top             =   3180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CheckBox chkCtrStock 
         Caption         =   "�Control de stock?"
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
         Left            =   9390
         TabIndex        =   25
         Tag             =   "Control de stock|N|N|0|1|sartic|ctrstock||N|"
         Top             =   3840
         Width           =   2205
      End
      Begin VB.TextBox txtSumaStock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9390
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   5280
         Width           =   2325
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
         Left            =   9390
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Fecha de Alta|F|N|||sartic|fecaltas|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   1215
         Width           =   1335
      End
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9390
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Tag             =   "Situaci�n Art�culo|N|N|||sartic|codstatu||N|"
         Top             =   1965
         Width           =   1335
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
         Left            =   9390
         MaxLength       =   18
         TabIndex        =   18
         Tag             =   "C�digo Asociaci�n|T|S|||sartic|codtelem||N|"
         Text            =   "Text1"
         Top             =   795
         Width           =   1830
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
         Index           =   11
         Left            =   9390
         MaxLength       =   8
         TabIndex        =   21
         Tag             =   "D�as de garantia|N|N|0|99999|sartic|garantia||N|"
         Text            =   "Text1"
         Top             =   2325
         Width           =   990
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
         Left            =   9390
         MaxLength       =   8
         TabIndex        =   22
         Tag             =   "Unidades por caja|N|N|||sartic|unicajas||N|"
         Text            =   "Text1"
         Top             =   2685
         Width           =   990
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
         Index           =   6
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "Cod. Tipo Art�culo|T|N|||sartic|codtipar||N|"
         Text            =   "Te"
         Top             =   2223
         Width           =   945
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
         Index           =   4
         Left            =   3180
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   2223
         Width           =   3945
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
         Left            =   3180
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   495
         Width           =   3945
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
         Left            =   3180
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   927
         Width           =   3945
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
         Index           =   5
         Left            =   3180
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   2655
         Width           =   3945
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
         Left            =   3180
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   1359
         Width           =   3945
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
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "Cod. Marca|N|N|0|9999|sartic|codmarca|0000|N|"
         Text            =   "Text1"
         Top             =   1359
         Width           =   945
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "Tipo de IVA|N|N|0|99|sartic|codigiva|00|N|"
         Text            =   "Ti"
         Top             =   2655
         Width           =   945
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
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "Cod. Familia|N|N|0|9999|sartic|codfamia|0000|N|"
         Text            =   "Text1"
         Top             =   927
         Width           =   945
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Cod. Proveedor|N|N|0|999999|sartic|codprove|000000|N|"
         Text            =   "Text1"
         Top             =   495
         Width           =   945
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
         Index           =   5
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "Cod. Tipo Unidad|N|N|0|99|sartic|codunida|00|N|"
         Text            =   "Text1"
         Top             =   1791
         Width           =   945
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
         Left            =   3180
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   40
         Text            =   "Text2"
         Top             =   1791
         Width           =   3945
      End
      Begin VB.CheckBox chkConjunto 
         Caption         =   "Tiene componentes"
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
         Left            =   9390
         TabIndex        =   24
         Tag             =   "�Es conjunto?|N|N|0|1|sartic|conjunto||N|"
         Top             =   3480
         Width           =   2565
      End
      Begin VB.CheckBox chkSeries 
         Caption         =   "�Control N� Serie?"
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
         Left            =   9390
         TabIndex        =   23
         Tag             =   "�Control n� serie?|N|N|0|1|sartic|nseriesn||N|"
         Top             =   3120
         Width           =   2355
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4545
         Left            =   -74280
         TabIndex        =   66
         Top             =   1050
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8017
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
      Begin MSAdodcLib.Adodc Data3 
         Height          =   330
         Left            =   -66360
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4200
         Left            =   -74760
         TabIndex        =   64
         Top             =   990
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   7408
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
      Begin VB.Frame FrameDatosAlmacen2 
         Caption         =   "Datos Relacionados con Almacen"
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
         Height          =   2535
         Left            =   120
         TabIndex        =   71
         Top             =   3120
         Width           =   9195
         Begin VB.TextBox txtPVPIVA 
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
            Left            =   7110
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   102
            Text            =   "Text1"
            Top             =   2010
            Width           =   1425
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
            Left            =   7890
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Fecha �ltimo cambio P.V.P.|F|S|||sartic|ultfecpvp|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   840
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
            Index           =   25
            Left            =   7110
            MaxLength       =   12
            TabIndex        =   16
            Tag             =   "Precio anual matenimiento|N|S|0|999999.00|sartic|preanuman|###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1440
            Width           =   1425
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
            Left            =   3090
            MaxLength       =   6
            TabIndex        =   15
            Tag             =   "Margen comercial|N|S|0|999.00|sartic|margecom|##0.00|N|"
            Text            =   "Text1"
            Top             =   1440
            Width           =   1320
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
            Left            =   4680
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Fecha �ltima compra|F|S|||sartic|ultfecco|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1240
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
            Left            =   3090
            MaxLength       =   12
            TabIndex        =   17
            Tag             =   "Precio Venta al p�blico|N|N|0|999999.0000|sartic|preciove|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   2010
            Width           =   1305
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
            Left            =   4680
            MaxLength       =   12
            TabIndex        =   10
            Tag             =   "Precio Standard|N|S|0|999999.0000|sartic|preciost|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   1240
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
            Left            =   1920
            MaxLength       =   12
            TabIndex        =   12
            Tag             =   "Precio Ultima Compra|N|S|0|999999.0000|sartic|preciouc|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1095
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
            Left            =   7890
            MaxLength       =   12
            TabIndex        =   11
            Tag             =   "Precio Medio Acumulado|N|S|0|999999.0000|sartic|precioma|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   270
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
            Index           =   13
            Left            =   1920
            MaxLength       =   12
            TabIndex        =   9
            Tag             =   "Precio Medio Ponderado|N|S|0|999999.0000|sartic|preciomp|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   270
            Width           =   1095
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C0C0C0&
            X1              =   1200
            X2              =   8550
            Y1              =   1830
            Y2              =   1830
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   1200
            X2              =   8550
            Y1              =   1290
            Y2              =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "P.V.P. + IVA"
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
            Left            =   4680
            TabIndex        =   103
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Fec.Cambio P.V.P."
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
            Left            =   6000
            TabIndex        =   101
            Top             =   840
            Width           =   2355
         End
         Begin VB.Label Label1 
            Caption         =   "Pr.Anual Mantenimiento"
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
            Left            =   4680
            TabIndex        =   80
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Margen Comercial"
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
            Left            =   1200
            TabIndex        =   78
            Top             =   1440
            Width           =   1905
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   4410
            ToolTipText     =   "Buscar fecha"
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "�lt.F.Compra"
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
            Left            =   3120
            TabIndex        =   77
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "P.V.P."
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
            Left            =   1920
            TabIndex        =   76
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Pr.Standard"
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
            Left            =   3120
            TabIndex        =   75
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Pr.�ltima Compra"
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
            Left            =   120
            TabIndex        =   74
            Top             =   840
            Width           =   2205
         End
         Begin VB.Label Label1 
            Caption         =   "Pr.Med Acumulado"
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
            Left            =   6000
            TabIndex        =   73
            Top             =   300
            Width           =   2025
         End
         Begin VB.Label Label1 
            Caption         =   "Pr Med.Ponderado"
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
            Left            =   120
            TabIndex        =   72
            Top             =   300
            Width           =   1875
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   4455
         Left            =   -74760
         TabIndex        =   98
         Top             =   1050
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7858
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
      Begin VB.Frame FrameLitrosUd 
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
         ForeColor       =   &H00972E0B&
         Height          =   735
         Left            =   9000
         TabIndex        =   123
         Top             =   4320
         Width           =   1455
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   4815
         Left            =   -74040
         TabIndex        =   130
         Top             =   600
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   8493
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   4440
         Left            =   -74280
         TabIndex        =   140
         Top             =   1050
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   7832
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
      Begin MSAdodcLib.Adodc Data5 
         Height          =   330
         Left            =   -68160
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label Label2 
         Caption         =   "C�digos de Barras"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   300
         Index           =   4
         Left            =   -72330
         TabIndex        =   144
         Top             =   570
         Width           =   2865
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   -66960
         TabIndex        =   129
         Top             =   480
         Width           =   2865
      End
      Begin VB.Label Label1 
         Caption         =   "N� Orden"
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
         Left            =   7440
         TabIndex        =   128
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label lblSumaStocks 
         Caption         =   "Suma Stock Almacenes"
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
         Left            =   9390
         TabIndex        =   56
         Top             =   5040
         Width           =   2655
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   -70320
         X2              =   -66480
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Label Label5 
         Caption         =   "Diferencia"
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
         Left            =   -67680
         TabIndex        =   120
         Top             =   5280
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "PVP real"
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
         Left            =   -69000
         TabIndex        =   118
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "PVP articulo"
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
         Left            =   -70320
         TabIndex        =   116
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Diferencia"
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
         Left            =   -71910
         TabIndex        =   114
         Top             =   5280
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "Coste real"
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
         Left            =   -73230
         TabIndex        =   112
         Top             =   5280
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Coste art�culo"
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
         Left            =   -74760
         TabIndex        =   110
         Top             =   5280
         Width           =   1665
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   -74760
         X2              =   -70620
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Label Label2 
         Caption         =   "Texto auxiliar documentos"
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
         Index           =   1
         Left            =   -74760
         TabIndex        =   104
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Texto para Ventas"
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
         Index           =   11
         Left            =   -74760
         TabIndex        =   70
         Top             =   600
         Width           =   2265
      End
      Begin VB.Label Label2 
         Caption         =   "Texto para Compras"
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
         Index           =   2
         Left            =   -74760
         TabIndex        =   69
         Top             =   1920
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Control de Instalaci�n"
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
         Index           =   3
         Left            =   -74760
         TabIndex        =   68
         Top             =   3120
         Width           =   2505
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   9120
         ToolTipText     =   "Buscar fecha"
         Top             =   1245
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Alta"
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
         Left            =   7440
         TabIndex        =   54
         Top             =   1245
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Situaci�n Art�culo"
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
         Left            =   7440
         TabIndex        =   53
         Top             =   1995
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Asociaci�n"
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
         Left            =   7440
         TabIndex        =   52
         Top             =   855
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "D�as de Garant�a"
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
         Left            =   7440
         TabIndex        =   51
         Top             =   2340
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Unidades por Caja"
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
         Left            =   7440
         TabIndex        =   50
         Top             =   2700
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Art�culo"
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
         Left            =   270
         TabIndex        =   49
         Top             =   2265
         Width           =   1545
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   4
         Left            =   1860
         ToolTipText     =   "Buscar tipo art�culo"
         Top             =   2259
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   5
         Left            =   1860
         ToolTipText     =   "Buscar tipo IVA"
         Top             =   2700
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   1860
         ToolTipText     =   "Buscar familia"
         Top             =   936
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   1860
         ToolTipText     =   "Buscar marca"
         Top             =   1377
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   1860
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   495
         Width           =   240
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
         Index           =   5
         Left            =   270
         TabIndex        =   48
         Top             =   510
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Familia"
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
         Left            =   270
         TabIndex        =   47
         Top             =   945
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de I.V.A."
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
         TabIndex        =   46
         Top             =   2700
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Marca"
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
         Left            =   270
         TabIndex        =   45
         Top             =   1380
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Unidad"
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
         Left            =   270
         TabIndex        =   44
         Top             =   1830
         Width           =   1515
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   3
         Left            =   1860
         ToolTipText     =   "Buscar tipo unidad"
         Top             =   1818
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   240
      TabIndex        =   57
      Top             =   795
      Width           =   12055
      Begin VB.ComboBox cboArticuloVarios 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10530
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Art�culo de Varios|N|N|||sartic|artvario||N|"
         Top             =   210
         Width           =   1335
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
         Index           =   1
         Left            =   4395
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Denominaci�n Art�culo|T|N|||sartic|nomartic||N|"
         Text            =   "Text1"
         Top             =   210
         Width           =   4245
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
         Index           =   0
         Left            =   1040
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "C�digo Art�culo|T|N|||sartic|codartic||S|"
         Text            =   "Text1"
         Top             =   210
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "Art�culo Varios"
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
         Left            =   8940
         TabIndex        =   67
         Top             =   255
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Denominaci�n"
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
         Left            =   2955
         TabIndex        =   59
         Top             =   255
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo Art."
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
         Left            =   195
         TabIndex        =   58
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Index           =   0
      Left            =   240
      TabIndex        =   38
      Top             =   7560
      Width           =   3255
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
         TabIndex        =   39
         Top             =   180
         Width           =   2835
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
      Left            =   11160
      TabIndex        =   27
      Top             =   7680
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
      Left            =   9960
      TabIndex        =   26
      Top             =   7680
      Width           =   1135
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4320
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSAdodcLib.Adodc Data4 
      Height          =   330
      Left            =   3600
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   11160
      TabIndex        =   29
      Top             =   7680
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnMantenimientos 
      Caption         =   "&Mantenimientos"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnMtoStocksAlm 
         Caption         =   "&Stocks Almacenes"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnMtoConjuntos 
         Caption         =   "&Conjuntos"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnMtoInstalaciones 
         Caption         =   "&Instalaciones"
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "frmAlmArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
' ==== Modificaciones:  =====
' ---- [14/09/2009] (LAURA)  --> Modificar funcion "InsertarPreciosPorTarifa2" para crear en funci�n del par�metro
                                 '"creatarifart" solo tarifa generar o todas las tarifas para el articulo
                                 
'---- [23/09/2009] LAURA  --> A�adir lineas de Cod. EAN
'---- [02/11/2009] LAURA  --> abrir el form y situarse en solapa Documentos|Pedidos
' ===========================


Public DatosADevolverBusqueda2 As String    'Tendra el n� de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

'---- [02/11/2009] LAURA  --> abrir el form y situarse en solapa Documentos|Pedidos
Public parNumTAb As Byte 'n� de tab en el q queremos q se situe al abrir el form
'----


Public Event DatoSeleccionado(CadenaSeleccion As String)


Private WithEvents frmArt As frmBasico2 'Form para busquedas
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmP As frmComProveedores
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmM As frmAlmMarcas 'Marcas de Art�culos
Attribute frmM.VB_VarHelpID = -1
Private WithEvents frmTU As frmAlmTipoUnidad
Attribute frmTU.VB_VarHelpID = -1
Private WithEvents frmTA As frmAlmTipoArticulo
Attribute frmTA.VB_VarHelpID = -1
Private WithEvents frmFA As frmAlmFamiliaArticulo
Attribute frmFA.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmAlPropios 'Almacenes Propios
Attribute frmA.VB_VarHelpID = -1


Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar un registro
'   5.-  Mantenimiento Lineas de Articulos x Almacen
'   6.-  Mantenimiento Lineas de Componentes de Conjuntos
'   7.-  Mantenimiento Lineas de Control de Instalaciones
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
Private ModoAnterior As Byte

Private ModoFrame As Byte
'ModoFrame: 0.-Inicio, 3.-Insertar, 4.-Modificar

Private CadenaConsulta As String
'SQL de la tabla principal del formulario

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1

Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim PrimeraVez As Boolean

Private TagText3 As String

'NUEVO: JULIO 2007. PARA BUSCAR POR CHECKS TB
'------------------------------------------------
Private BuscaChekc As String

'NUevo: Nov 2008
' Hay un campo para pintar el PVP con el IVA
' Guardaremos el tipo de iva y el % (para no tener que recaluclarlo cada ve
Private mPorIva As String

Private PriVezForm As Boolean

'Cunado esta metiendo componentes, si es materia prima, y va por porcentajes
Private MateriaPrima As Boolean


Dim NumTabMto As Byte 'Indica que numero de Tab que esta en modo Mantenimiento




Private Sub cboArticuloVarios_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkConjunto_Click()
 If Modo = 1 Then CheckCadenaBusqueda chkConjunto, BuscaChekc
End Sub

Private Sub chkConjunto_GotFocus()
     ConseguirfocoChk Modo
End Sub

Private Sub chkConjunto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCtrStock_Click()
 If Modo = 1 Then CheckCadenaBusqueda chkctrstock, BuscaChekc

End Sub

Private Sub chkctrstock_GotFocus()
     ConseguirfocoChk Modo
End Sub

Private Sub chkctrstock_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkInventario_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkInventario, BuscaChekc
End Sub

Private Sub chkInventario_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkInventario_LostFocus()
    PonerFocoBtn Me.cmdAceptar
End Sub



Private Sub chkSeries_Click()
 If Modo = 1 Then CheckCadenaBusqueda chkSeries, BuscaChekc
 
End Sub

Private Sub chkSeries_GotFocus()
     ConseguirfocoChk Modo
End Sub

Private Sub chkSeries_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub cmdAceptar_Click()
Dim Cad As String, Indicador As String
Dim bol As Boolean

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
'              If InsertarDesdeForm(Me) Then
'                InsetarArticulosPorAlmacen
'                InsertarPreciosPorTarifa
                If InsertarArticulo Then
'                    MsgBox "Los precios del art�culo por tarifa se han introducido correctamente.", vbInformation
                    PosicionarData
                End If
'              End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    'comprobar si se ha modificado el precio ult. compra
                    'y preguntar si modificar PVP y Tarifas
                    If CCur(DBLet(Data1.Recordset!precioUC, "N")) <> ImporteFormateado(Text1(15).Text) Then
                        ActualizarPreciosVenta
                    ElseIf CCur(DBLet(Data1.Recordset!preciove, "N")) <> ImporteFormateado(Text1(17).Text) Then
                        'Comprobar si se ha modificado el precio de venta PVP y preguntar
                        'si se quieren actualizar las tarifas de precios
                        ActualizarPreciosPorTarifa
                    ElseIf CCur(DBLet(Data1.Recordset!margecom, "N")) <> ImporteFormateado(Text1(24).Text) Then
                        'comprobar si se ha modificado el margen comercial
                        'y preguntar si modificar PVP y Tarifas
                         ActualizarPreciosVenta
                    End If
                    
'                    DesBloqueaRegistroForm Text1(0)
                    PosicionarData
                End If
            End If
                
         Case 5 'InsertarModificar linea  '----------------
         
            'Actualizar el registro en la tabla de lineas 'salmac' (Art�culos x Almacen)
            If InsertarModificarLinea Then
'                DesBloqueaRegistroForm Text1(0)
      
                NumRegElim = Data4.Recordset.AbsolutePosition
                TerminaBloquear
                LLamaLineas2 0, 0, 4
                DataGrid3.AllowAddNew = False
                CargaGrid Me.DataGrid3, Me.Data4, True
                SituarDataPosicion Data4, NumRegElim, Indicador
                
                lblIndicador.Caption = Indicador
                PonerModoFrame 0
                PonerSumaStocks
                
               
                
            End If
            
          '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
          Case 6, 7, 8 '6: InsertarModificar Conjuntos
                    '7: InsertarModificar Instalaciones
                    '8: InsertarModificar cod. EAN
             If Modo = 6 Then bol = InsertarModificarConjunto
             If Modo = 7 Then bol = InsertarModificarInstalacion
             If Modo = 8 Then bol = InsertarModificarCodigosEAN
             
             If bol Then
                TerminaBloquear
                If Modo = 6 Then 'Conjunto
                  txtAux(0).visible = False
                  txtAux(1).visible = False
                  txtAux2.visible = False
                  cmdAux.visible = False
                  CargaGrid Me.DataGrid1, Me.Data2, True
                ElseIf Modo = 7 Then 'Instalacion
                    txtAux(2).visible = False
                    CargaGrid Me.DataGrid2, Me.Data3, True
                '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
                ElseIf Modo = 8 Then 'codigos EAN
                    txtAux(7).visible = False
                    CargaGrid Me.DataGrid4, Me.Data5, True
                '----
                End If
                
                If ModificaLineas = 2 Then 'Modificar
                    DesBloqueaRegistroForm Text1(0)
                    If Modo = 6 Then
                        Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                    ElseIf Modo = 7 Then
                        Data3.Recordset.Find (Data3.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
                    ElseIf Modo = 8 Then
                        Data5.Recordset.Find (Data5.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                    '----
                    End If
'--                    PonerBotonCabecera True
'                    Me.lblIndicador.Caption = ""
                    PonerFocoBtn Me.cmdAceptar
                    ModificaLineas = 0
                ElseIf ModificaLineas = 1 Then 'Insertar
                    BotonAnyadirConjunto2
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Sub cmdActualizarImportes1_Click(Index As Integer)
Dim frmAr As frmAlmArticulos

    If Modo <> 6 Then Exit Sub
    
    If ModificaLineas <> 0 Then
        MsgBox "Esta cambiando datos", vbExclamation
        Exit Sub
    End If
    
    If Index = 0 Then
        If txtConjunto(1).Text = "" Or txtConjunto(1).Text = "" Then
            MsgBox "Falta importes calculados", vbExclamation
            Exit Sub
        End If
        BuscaChekc = "�Desea cambiar los importes PVP y UPC del �rticulo principal?"
        If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    If Index = 0 Then
        'ACtualizar importes
    
        'Haremos lo siguiente
        If BLOQUEADesdeFormulario(Me) Then
            'Fijaremos los nuevos importes
             
             If ModificarImportesDesdeConjuntos Then
                    TerminaBloquear
                    Text1(15).Text = Me.txtConjunto(1).Text
                    Text1(17).Text = Me.txtConjunto(4).Text
                    'comprobar si se ha modificado el precio ult. compra
                    'y preguntar si modificar PVP y Tarifas
                    If CCur(DBLet(Data1.Recordset!precioUC, "N")) <> ImporteFormateado(Text1(15).Text) Then ActualizarPreciosVenta
                    'Comprobar si se ha modificado el precio de venta PVP y preguntar
                    'si se quieren actualizar las tarifas de precios
                    If CCur(DBLet(Data1.Recordset!preciove, "N")) <> ImporteFormateado(Text1(17).Text) Then ActualizarPreciosPorTarifa
                    

                    PosicionarData
            End If
        End If
    Else
        'VER ARTICULO LINEA
        Set frmAr = New frmAlmArticulos
        frmAr.DeConsulta = True
        frmAr.DatosADevolverBusqueda2 = "::" & DevNombreSQL(Data2.Recordset!codarti1)
        frmAr.Show vbModal
        Set frmAr = Nothing
        
        'Por si acaso ha cambiado
        'recargo el grid
        '--------------------------------------------------------------------------------------
        NumRegElim = Data2.Recordset.AbsolutePosition - 1
        
        CancelaADODC Me.Data2
        CargaGrid Me.DataGrid1, Me.Data2, True
        CancelaADODC Me.Data2
        ponerDatosConjuntos
        If NumRegElim > 0 Then Data2.Recordset.Move NumRegElim, 1
        
    End If
    BuscaChekc = ""
End Sub

Private Function ModificarImportesDesdeConjuntos() As Boolean
    On Error GoTo EM
    ModificarImportesDesdeConjuntos = False
    BuscaChekc = "UPDATE sartic set precioUC = " & TransformaComasPuntos(CStr(ImporteFormateado(Me.txtConjunto(1).Text)))
    BuscaChekc = BuscaChekc & " , preciove =" & TransformaComasPuntos(CStr(ImporteFormateado(Me.txtConjunto(4).Text)))
    BuscaChekc = BuscaChekc & " WHERE codartic = '" & DevNombreSQL(Data1.Recordset!codArtic) & "'"
    conn.Execute BuscaChekc
    ModificarImportesDesdeConjuntos = True
    Exit Function
EM:
    MuestraError Err.Number, "", Err.Description
End Function


Private Sub cmdAlma_Click()
    imgCuentas_Click 6
End Sub

Private Sub cmdAux_Click()
    MandaBusquedaPrevia " conjunto=0 "
    PonerFoco txtAux(1)
End Sub

Private Sub cmdCancelar_Click()
On Error Resume Next
    Select Case Modo
        Case 1 'Busqueda
            LimpiarCampos
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 3 'Insertar
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
                
        Case 4  'Modificar
            TerminaBloquear
'            DesBloqueaRegistroForm Text1(0)
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
        
        
        
        '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
        Case 5, 6, 7, 8 'Lineas Conjuntos, Lineas Instalaciones
            ModificaLineas = 0
'            DesBloqueoManual NombreTabla
            TerminaBloquear
            Select Case Modo
            Case 5
                DataGrid3.AllowAddNew = False
                DataGrid2.Enabled = False
                PonerModoFrame 0
                LLamaLineas2 0, 0, 4
                NumRegElim = Data4.Recordset.AbsolutePosition
                CargaGrid DataGrid3, Data4, True
                SituarDataPosicion Data4, NumRegElim, Me.lblIndicador.Caption
                If Not Data4.Recordset.EOF Then PonerCamposAlmacenes2
                
            Case 6
                txtAux(0).visible = False
                txtAux(1).visible = False
                txtAux2.visible = False
                cmdAux.visible = False
                DataGrid1.AllowAddNew = False
                If Not (ModificaLineas = 2) Then 'Modificar
                    If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                End If
                DataGrid1.Enabled = True
                
            Case 7
                txtAux(2).visible = False
                DataGrid2.AllowAddNew = False
                If Not (ModificaLineas = 2) Then 'Modificar
                    If Not Data3.Recordset.EOF Then Data3.Recordset.MoveFirst
                End If
                DataGrid2.Enabled = True
            
                 
            '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
            Case 8 'Lineas codigos EAN
                txtAux(7).visible = False
                DataGrid4.AllowAddNew = False
                If Not (ModificaLineas = 2) Then 'Modificar
                    If Not Data5.Recordset.EOF Then Data5.Recordset.MoveFirst
                End If
                DataGrid4.Enabled = True
            '----
            End Select
            
'--            PonerBotonCabecera True
            PonerModo 2
            PonerFocoBtn Me.cmdRegresar
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos 'Vac�a los TextBox
    ModoAnterior = Modo 'Para el bot�n Cancelar en Modo Insertar
    PonerModo 3
    Me.SSTab1.Tab = 0
    
    'Poner valores por defecto
    Me.chkctrstock.Value = 1 'por defecto hay control de stock
    Me.Text1(10).Text = Format(Now, "dd/mm/yyyy") 'fecha alta
    Me.cboArticuloVarios.ListIndex = 0
    Me.cboStatus.ListIndex = 0
    Me.Text1(11).Text = "0"
    Me.Text1(12).Text = "1"
    Me.chkMateriaPrima.Value = 0
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLinea()
Dim vWhere As String
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    Me.SSTab1.Tab = 4
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModoFrame 3    '3: Insertar
    ModificaLineas = 1 'Insertar

    'Obtenemos la siguiente numero de Art�culo
    vWhere = "codartic=" & DBSet(Text1(0).Text, "T")
    Text3(0).Text = SugerirCodigoSiguienteStr("salmac", "codalmac", vWhere)
    lblIndicador.Caption = "INSERTAR ALMACEN"
    PonerFoco Text3(0)
End Sub





Private Sub BotonAnyadirConjunto2()
Dim numF As String
Dim vWhere As String
Dim anc As Single
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    ModificaLineas = 1
'--    PonerBotonCabecera False
    
    vWhere = "codartic=" & DBSet(Text1(0).Text, "T")
'    ancIni = 200
    Select Case Modo
    Case 5 'Lineas STOCK
        Me.SSTab1.Tab = 4
        lblIndicador.Caption = "INSERTAR STOCK"
        numF = 1
        
    Case 6
        numF = SugerirCodigoSiguienteStr("sarti1", "numlinea", vWhere)
        Me.SSTab1.Tab = 2
        lblIndicador.Caption = "INSERTAR CONJUNTO"
        
    Case 7 'Lineas Instalaciones
        numF = SugerirCodigoSiguienteStr("sarti2", "numlinea", vWhere)
        Me.SSTab1.Tab = 3
        lblIndicador.Caption = "INSERTAR INSTALACI�N"
    
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
    Case 8 'Lineas cod. EAN
        numF = SugerirCodigoSiguienteStr("sarti3", "numlinea", vWhere)
        Me.SSTab1.Tab = 5
        lblIndicador.Caption = "INSERTAR COD. EAN"
    '----
    End Select
    cmdAceptar.Tag = numF
    
    PonerModo Modo
    
    Select Case Modo
    'If Modo = 6 Then 'Conjuntos
    Case 5 'Lineas STOCK
        PonerDatosForaGrid True
        PonerModoFrame 3
        AnyadirLinea DataGrid3, Data4
        anc = ObtenerAlto(DataGrid3, 20)
        LLamaLineas2 anc, 1, 4
        PonerFoco Text3(0)
        BloquearTxt Text3(0), False
        
    Case 6
        txtAux(0).Text = ""
        txtAux2.Text = ""
        txtAux(1).Text = ""
        'Situamos el grid al final
        AnyadirLinea DataGrid1, Data2

        anc = ObtenerAlto(DataGrid1, 20)
        LLamaLineas2 anc, 1, 2
        
        BloquearTxt txtAux(0), False
        Me.cmdAux.Enabled = True
        PonerFoco txtAux(0)
        
    Case 7 'Lineas INSTALACIONES
        Me.txtAux(2).Text = ""
        AnyadirLinea DataGrid2, Data3
        anc = ObtenerAlto(DataGrid2, 20)
        LLamaLineas2 anc, 1, 3
        PonerFoco txtAux(2)

    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
    Case 8 'Lineas cod. EAN
        Me.txtAux(7).Text = ""
        AnyadirLinea DataGrid4, Data5
        anc = ObtenerAlto(DataGrid4, 20)
        LLamaLineas2 anc, 1, 5
        PonerFoco txtAux(7)

    '----
    End Select
End Sub


Private Sub BotonBuscar()
'Buscar
    LimpiarCampos
    If Modo <> 1 Then 'Modo 1: Busqueda
        BuscaChekc = ""
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue 'vbYellow
        'Si es de buscqueda , buscamos solo activos
        If DeConsulta Then Me.cboStatus.ListIndex = 0
    Else
        If DeConsulta Then
            If cboStatus.ListIndex < 0 Then cboStatus.ListIndex = 0
        End If
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue 'vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
Dim C As String
  
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        C = ""
        If DeConsulta Then C = "codstatu = 0"
    
        MandaBusquedaPrevia C
    Else
        C = "Select * from " & NombreTabla
        If DeConsulta Then C = C & " WHERE codstatu = 0 "
        CadenaConsulta = C & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
''Botones de Desplazamiento de la Toolbar
'    Select Case Modo
'        Case 5 'Modo Mantenimiento de Almacenes (Lineas)
'            If Data4.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data4, Index
'            PonerCamposAlmacenes2
'        Case Else 'Datos de Cabecera
'            If Data1.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data1, Index
'            PonerCampos
'            PonerModoOpcionesMenu (Modo) 'Poner opciones de menu seg�n modo
'            PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
'                                'de permisos del usuario
'    End Select
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub BotonModificar()
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
End Sub



Private Sub BotonModificarConjunto(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim anc As Single
Dim i As Integer

    If vData.Recordset.EOF Then Exit Sub
    If vData.Recordset.RecordCount < 1 Then Exit Sub
   
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
'--    PonerBotonCabecera False
         
    If vDataGrid.Bookmark < vDataGrid.FirstRow Or vDataGrid.Bookmark > (vDataGrid.FirstRow + vDataGrid.VisibleRows - 1) Then
        i = vDataGrid.Bookmark - vDataGrid.FirstRow
        vDataGrid.Scroll 0, i
        vDataGrid.Refresh
    End If
    PonerFocoBtn Me.cmdAceptar
    vDataGrid.Enabled = False

    anc = ObtenerAlto(vDataGrid, 20)
    
    If Modo = 5 Then
        cmdAceptar.Tag = vData.Recordset!codAlmac
    Else
        cmdAceptar.Tag = vData.Recordset!numlinea
    End If
    
    Select Case Modo
    Case 5
        PonerModoFrame 4 'ModoFrame=4 -> Modificar
        Me.lblIndicador.Caption = "MODIFICAR ALMACEN"
        LLamaLineas2 anc, 2, 4
        BloquearTxt Text3(0), True
        Text3(0).Text = Data4.Recordset!codAlmac
        Text3(1).Text = Data4.Recordset!CanStock
        PonerFoco Text3(1)

    Case 6
        MateriaPrima = CStr(DBLet(vData.Recordset!MateriaPrima, "T")) = "*"
    ' If Modo = 6 Then 'Componentes de Conjunto
        Me.lblIndicador.Caption = "MODIFICAR CONJUNTO"
        Me.SSTab1.Tab = 2
         'Llamamos al form
        txtAux(0).Text = DataGrid1.Columns(2).Text
        BloquearTxt txtAux(0), True
        Me.txtAux2.Text = DataGrid1.Columns(3).Text
        txtAux(1).Text = DataGrid1.Columns(4).Text
        LLamaLineas2 anc, 2, 2
        PonerFoco txtAux(1)
        If ModificaLineas = 2 Then cmdAux.Enabled = False
    'Poner el foco
    'ElseIf Modo = 7 Then
    Case 7
        Me.lblIndicador.Caption = "MODIFICAR INSTALACI�N"
        Me.SSTab1.Tab = 3
        txtAux(2).Text = DataGrid2.Columns(2).Text
        LLamaLineas2 anc, 2, 3
        PonerFoco txtAux(2)
        
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
    Case 8 'Lineas cod. EAN
        Me.lblIndicador.Caption = "MODIFICAR COD. EAN"
        Me.SSTab1.Tab = 5
        txtAux(7).Text = DataGrid4.Columns(2).Text
        LLamaLineas2 anc, 2, 5
        PonerFoco txtAux(7)
    End Select
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonEliminar()
Dim Sql As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'No esta bloqueado
    If Val(Data1.Recordset!codstatu) = 1 Then
        MsgBox "Articulo bloqueado", vbExclamation
        Exit Sub
    End If
    
    
    'Tiene stock
    If ImporteFormateado(txtSumaStock.Text) <> 0 Then
        MsgBox "El articulo tiene stock", vbExclamation
        Exit Sub
    End If
    

    
    BuscaChekc = lblIndicador.Caption
    Sql = SePuedeEliminarArticulo(CStr(Data1.Recordset!codArtic), lblIndicador)
    lblIndicador.Caption = BuscaChekc
    BuscaChekc = ""
    If Sql <> "" Then
        Sql = "No se puede eliminar el articulo: " & Data1.Recordset!codArtic & vbCrLf & vbCrLf & Sql
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
    Sql = "Cabecera de Art�culos." & vbCrLf
    Sql = Sql & "---------------------------        " & vbCrLf & vbCrLf
    Sql = Sql & "Va a eliminar el Art�culo:"
    Sql = Sql & vbCrLf & "Cod. Artic. :   " & Data1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Descripci�n :   " & Data1.Recordset.Fields(1)
    Sql = Sql & vbCrLf & vbCrLf & " �Desea Eliminarlo? "
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        TerminaBloquear
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerModo 2
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Articulo", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De ArticulosxAlmacen
Dim Cad As String

     On Error GoTo Error2

    If Data4.Recordset.EOF Then Exit Sub
    If Data4.Recordset.RecordCount < 1 Then Exit Sub
    If vUsu.Nivel > 1 Then Exit Sub
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
    ModificaLineas = 3 'Eliminar
    
    '### a mano
    Cad = "Seguro que desea eliminar de la BD el registro:"
    Cad = Cad & vbCrLf & "Cod. Art�culo: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Cod. Almacen: " & Data4.Recordset.Fields(1)

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
       
        Screen.MousePointer = vbHourglass
        NumRegElim = Data4.Recordset.AbsolutePosition
        
        Cad = "DELETE FROM salmac where codartic = '" & DevNombreSQL(Data1.Recordset.Fields(0)) & "' AND codalmac = " & Data4.Recordset!codAlmac
        conn.Execute Cad
        
        CargaGrid Me.DataGrid3, Me.Data4, True
        If Data4.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCamposAlmacenes
            PonerModoFrame 0
        Else
            SituarDataPosicion Me.Data4, NumRegElim, Cad
            PonerCamposAlmacenes2
        End If
        ModificaLineas = 0
    End If
    Screen.MousePointer = vbDefault
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data4.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Linea de Articulo", Err.Description
    End If
End Sub


Private Sub BotonEliminarConjunto()
Dim Sql As String
    On Error GoTo Error2
    
    'Ciertas comprobaciones
    If Data2.Recordset.EOF Then Exit Sub
    
    Sql = "Seguro que desea eliminar el Componente de Conjunto:"
    Sql = Sql & vbCrLf & "C�digo: " & Data2.Recordset!codarti1
    Sql = Sql & vbCrLf & "Descripci�n: " & Data2.Recordset.Fields(3)
    
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        Sql = "Delete from sarti1 where codartic=" & DBSet(Data2.Recordset!codArtic, "T")
        Sql = Sql & " and numlinea=" & Data2.Recordset!numlinea
        Sql = Sql & " and codarti1=" & DBSet(Data2.Recordset!codarti1, "T")
        conn.Execute Sql
        CancelaADODC Me.Data2
        CargaGrid Me.DataGrid1, Me.Data2, True
        CancelaADODC Me.Data2
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Componente de Conjunto", Err.Description
End Sub


Private Sub BotonEliminarInstalacion()
Dim Sql As String
    On Error GoTo Error2

    'Ciertas comprobaciones
    If Data3.Recordset.EOF Then Exit Sub
    
    Sql = "Seguro que desea eliminar el control de instalaci�n:"
    Sql = Sql & vbCrLf & "Linea: " & Data3.Recordset!numlinea
    Sql = Sql & vbCrLf & "Descripci�n: " & Data3.Recordset!licontro
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Sql = "Delete from sarti2 where codartic=" & DBSet(Data3.Recordset!codArtic, "T")
        Sql = Sql & " and numlinea=" & Data3.Recordset!numlinea
        conn.Execute Sql
        CancelaADODC Me.Data3
        CargaGrid Me.DataGrid2, Me.Data3, True
        CancelaADODC Me.Data3
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Control de Instalaciones", Err.Description
End Sub


'---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
Private Sub BotonEliminarCodigosEAN()
Dim Sql As String
    On Error GoTo ErrElimEAN

    'Ciertas comprobaciones
    If Data5.Recordset.EOF Then Exit Sub
    
    Sql = "Seguro que desea eliminar el codigo EAN:"
    Sql = Sql & vbCrLf & "Linea: " & Data5.Recordset!numlinea
    Sql = Sql & vbCrLf & "Cod. EAN: " & Data5.Recordset!codigoea
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Sql = "Delete from sarti3 where codartic=" & DBSet(Data5.Recordset!codArtic, "T")
        Sql = Sql & " and numlinea=" & Data5.Recordset!numlinea
        conn.Execute Sql
        CancelaADODC Me.Data5
        CargaGrid Me.DataGrid4, Me.Data5, True
        CancelaADODC Me.Data5
    End If
    Exit Sub
    
ErrElimEAN:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Codigos EAN", Err.Description
End Sub
'----



Private Sub BotonArticulosxAlmac()
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrorArticAlmac
    
    Screen.MousePointer = vbHourglass
    'RESTAURO LOS tag's
    AccionesSobreTagText3_ False, False

    Me.SSTab1.Tab = 4
    PonerModo (5)
'--    PonerBotonCabecera True
    PonerFocoBtn Me.cmdRegresar
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ErrorArticAlmac:
    MuestraError Err.Number, "PonerCadenaBusqueda", Err.Description
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonConjuntos()
    On Error GoTo ErrorConjuntos
    
    Screen.MousePointer = vbHourglass
    Me.SSTab1.Tab = 2
    
    PonerModo (6)
'--    PonerBotonCabecera True
    
    DataGrid1.Enabled = True
    Me.DataGrid1.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorConjuntos:
    MuestraError Err.Number, "Conjuntos", Err.Description
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonInstalaciones()
    On Error GoTo ErrorInstala

    Screen.MousePointer = vbHourglass
    Me.SSTab1.Tab = 3
    PonerModo (7)
'--    PonerBotonCabecera True
    PonerFocoBtn Me.cmdRegresar
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorInstala:
    MuestraError Err.Number, "Instalaciones", Err.Description
    Screen.MousePointer = vbDefault
End Sub



Private Sub BotonCodigosEAN()
    On Error GoTo ErrEAN
    Screen.MousePointer = vbHourglass
    
    Me.SSTab1.Tab = 5
    PonerModo (8)
'--    PonerBotonCabecera True
    PonerFocoBtn Me.cmdRegresar
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrEAN:
    MuestraError Err.Number, "Codigos EAN", Err.Description
    Screen.MousePointer = vbDefault
End Sub



Private Sub cmdGenerar_Click()
    Dim Aux As String
    Aux = Text2(2) & " " & Text2(1) & " " & Text2(4) & " " & Text2(3)
    Text1(1).Text = Replace(Left(Aux, 40), "*", "")
    Text1(0).Text = SugerirCodAutomatico(Text1(4), Text1(3), Text1(6), Text1(5))
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
    If Modo = 5 Or Modo = 6 Or Modo = 7 Or Modo = 8 Then
    
        If Modo = 6 Then
            'Componentes
            If vParamAplic.ComponentePorcentaje Then
                'Son porcentajes. Compruebo que la suma es 100
                If Not ComprobarPorcentajesCorrectos Then Exit Sub
            End If
        End If
        'modo 5: Lineas Articulos x Almacen
        'modo 6: Lineas Conjuntos
        'modo 7: Lineas Instalaciones
        'modo 8: Lineas cod. EAN
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
        If DataGrid2.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid2
            DataGrid2.Bookmark = 1
        End If
        PonerModo 2

    Else 'Se llamo desde un bot�n de Prism�tico
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
        If DeConsulta Then
            If cboStatus.ListIndex > 0 Then
                MsgBox "Articulo " & cboStatus.Text, vbExclamation
                Exit Sub
            End If
        End If
            
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
        Cad = Cad & Data1.Recordset.Fields(8).Value & "|"
        Cad = Cad & Text2(4).Text & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub




Private Sub Data4_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Modo = 5 And ModificaLineas > 0 Then Exit Sub
    If Not Data4.Recordset.EOF Then
        If Not PrimeraVez Then PonerCamposAlmacenes2
    End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        If ModificaLineas = 0 Then
            PonerFocoBtn Me.cmdRegresar
        Else
            SendKeys "{tab}"
        End If
    End If
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(1)
    
    If PriVezForm Then
        PriVezForm = False
        
        'He abierto el form queriendo cargar un articulo
        If Mid(DatosADevolverBusqueda2, 1, 2) = "::" Then
            DatosADevolverBusqueda2 = Mid(DatosADevolverBusqueda2, 3)
            CadenaConsulta = "Select * from " & NombreTabla & " where codartic='" & DatosADevolverBusqueda2 & "'"
            PonerCadenaBusqueda
            
            If Me.chkConjunto.Value > 0 And vUsu.Nivel <= 1 Then
                Toolbar1.Buttons(11).Enabled = True
                Me.mnMtoConjuntos.Enabled = True
            End If
            
            If Me.parNumTAb = 6 Then
                Toolbar2.Buttons(7).Value = tbrPressed
                Toolbar2_ButtonClick Toolbar2.Buttons(7)
            End If
         End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer

    PriVezForm = True
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
    btnAnyadir = 6
    btnPrimero = 18 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Bot�n Buscar
'        .Buttons(2).Image = 2   'Bot�n Todos
'        .Buttons(6).Image = 3   'Insertar Nuevo
'        .Buttons(7).Image = 4   'Modificar
'        .Buttons(8).Image = 5   'Borrar
'        .Buttons(10).Image = 10 'Stocks Almacenes
'        .Buttons(11).Image = 11 'Conjuntos
'        .Buttons(12).Image = 36 'Instalaciones
'        .Buttons(13).Image = 23 'Cod. EAN '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
'
'        .Buttons(15).Image = 16  'Imprimir
'        .Buttons(16).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
'    End With
    
    
    With Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun1
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
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun1
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    'ICONITOS DE LAS BARRAS EN LOS TABS DE LINEA
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
            '.ImageList = frmPpal.imgListComun_VELL
            '  ### [Monica] 02/10/2006 acabo de comentarlo
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i
    
    
    
    
    
    For i = 0 To Me.imgCuentas.Count - 1
        imgCuentas(i).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next
    For i = 0 To Me.imgFecha.Count - 1
        imgFecha(i).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next
   
    
    If Me.parNumTAb > 0 Then
        Me.SSTab1.Tab = Me.parNumTAb
    Else
        Me.SSTab1.Tab = 0
    End If
    Me.SSTab1.TabVisible(2) = False
    Me.SSTab1.TabVisible(3) = False
    'Me.FrameDatosAlmacen.Left = 360
    'Me.FrameDatosAlmacen.Top = 2780
    'Me.FrameArtxAlmac2.visible = False
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    
        
    'Si no tiene servicios nmo muesotr el frame
    FrameServicios.visible = vParamAplic.Servicios
    FrameLitrosUd.visible = vParamAplic.Descriptores
    
    framePortes.visible = vParamAplic.ArtPortes <> ""
    
    'Si hay algun combo los cargamos
    CargarComboStatus
    CargarComboArticuloVarios
    
    
    'El tag de los stocks
    AccionesSobreTagText3_ True, True
    
    'Pone el Tag del primer bot�n de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sartic, BD: Aritaxi
    'Si tag>0 abre busqueda en la tabla asociada al indice.
    imgCuentas(0).Tag = "-1"
         
    '## A mano
    NombreTabla = "sartic"
    Ordenacion = " ORDER BY codartic"
  
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codartic is null "
    Data1.Refresh
    
    LimpiarDataGrids
    
    If DatosADevolverBusqueda2 = "" Then
        PonerModo 0
        PonerCamposLineas False
    Else
        If DatosADevolverBusqueda2 = "@1@" Then 'Poner Modo Busqueda
            BotonBuscar
        Else 'Poner Modo Insertar
            If Mid(DatosADevolverBusqueda2, 1, 2) = "::" Then
                'Abrimos el articulo poniendo un articulo especificado a continuacion
                
                'Lo haremos en el ACTIVATE
            Else
                PonerModo 3
                Text1(0).Text = DatosADevolverBusqueda2
            End If
        End If
    End If
    
    '-- Descriptores especiales y bot�n de composici�n (Rafa VRS 4.0.9)
    If vParamAplic.Descriptores Then
        'cmdGenerar.visible = True  estara en poner modo
        Label1(6) = "Cod. Categoria"
        Label1(9) = "Cod. Modelo"
        Label1(17) = "Cod. Formato"
        '-- Aqui cambiamos los tag para evitar lios.
        CambiaTagDescriptores Text1(3), "Cod. Categoria"
        CambiaTagDescriptores Text1(5), "Cod. Formato"
        CambiaTagDescriptores Text1(6), "Cod. Modelo"
    Else
        cmdGenerar.visible = False
    End If
    '--
    ImagenesNavegacion
    CargaColumnas 0
End Sub

Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ning�n registro
Dim Cad As String
On Error Resume Next

    Cad = "Select * from sarti1 where codartic is null"
    CargaGrid DataGrid1, Data2, False
    Cad = "Select * from sarti2 where codartic is null"
    CargaGrid DataGrid2, Data3, False
    Cad = "Select * from salmac where codartic is null"
    CargaGrid DataGrid3, Data4, False
    Cad = "Select * from sarti3 where codartic is null"
    CargaGrid DataGrid4, Data5, False
    
    PrimeraVez = False
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    CargaGrid Me.DataGrid3, Me.Data4, False  'Desenlazamos el GRID
    'Aqui va el especifico de cada form es
    Me.chkConjunto.Value = 0
    Me.chkSeries.Value = 0
    Me.chkctrstock.Value = 0
    Me.chkMateriaPrima.Value = 0
    Me.cboArticuloVarios.ListIndex = -1
    Me.cboStatus.ListIndex = -1
End Sub


Private Sub LimpiarCamposAlmacenes()
Dim i As Byte
    Text3(0).BackColor = vbRed
    For i = 0 To Text3.Count - 1
        Text3(i).Text = ""
    Next i
    Text2(8).Text = ""
    Me.chkInventario.Value = 0
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Me.parNumTAb = 0
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacenes Propios
    Text3(0).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text3(0)
    Text2(8).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
Dim indice As Integer
      
    If CadenaDevuelta <> "" Then
        If Val(imgCuentas(0).Tag) >= 0 Then
            'Se llama desde el bot�n de busqueda del campo Tipos de IVA
            'Recuperar solo el campo c�digo y Descripci�n
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            indice = Val(Me.imgCuentas(0).Tag)
            Text1(indice + 2).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(indice).Text = RecuperaValor(CadenaDevuelta, 2)
        Else
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            
            If Modo <> 6 Then
                'Recupera todo el registro de Art�culos
                'Sabemos que campos son los que nos devuelve
                'Creamos una cadena consulta y ponemos los datos
                CadB = ""
                Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
                CadB = Aux
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
                PonerCadenaBusqueda
            Else
                'Llamamos desde el boton auxiliar de Conjuntos
                txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
                txtAux2.Text = RecuperaValor(CadenaDevuelta, 2)
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    CadB = "codartic = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "T")
    
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
    PonerCadenaBusqueda
    Screen.MousePointer = vbDefault

End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas

    Select Case Val(imgFecha(0).Tag)
        Case 0
            Text1(10).Text = Format(vFecha, "dd/mm/yyyy")
        Case 1
            Text1(18).Text = Format(vFecha, "dd/mm/yyyy")
        Case 2
            Text3(6).Text = Format(vFecha, "dd/mm/yyyy")
            
        Case 3
            Text1(24).Text = Format(vFecha, "dd/mm/yyyy")
    End Select
End Sub


Private Sub frmFA_DatoSeleccionado(CadenaSeleccion As String)
'Familia de Articulo
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(3)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmM_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Marcas
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(4)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
'Proveedores
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(2)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTA_DatoSeleccionado(CadenaSeleccion As String)
'Tipo de Articulo
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTU_DatoSeleccionado(CadenaSeleccion As String)
'Tipo de Unidad
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(5)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgCuentas_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Codigo Proveedor
            Set frmP = New frmComProveedores
            frmP.DatosADevolverBusqueda = "0"
            frmP.Show vbModal
            Set frmP = Nothing
        Case 1  'Cod. Familia
            Set frmFA = New frmAlmFamiliaArticulo
            frmFA.DatosADevolverBusqueda = "0"
            frmFA.Show vbModal
            Set frmFA = Nothing
        Case 2  'Cod. Marca
            Set frmM = New frmAlmMarcas
            frmM.DatosADevolverBusqueda = "0"
            frmM.Show vbModal
            Set frmM = Nothing
        Case 3  'Cod. Tipo Unidad
            Set frmTU = New frmAlmTipoUnidad
            frmTU.DatosADevolverBusqueda = "0"
            frmTU.Show vbModal
            Set frmTU = Nothing
        Case 4  'Cod. Tipo Articulo
            Set frmTA = New frmAlmTipoArticulo
            frmTA.DatosADevolverBusqueda = "0"
            frmTA.Show vbModal
            Set frmTA = Nothing
            
        Case 5  'Tipos de IVA. Tabla de la BD Contabilidad
            imgCuentas(0).Tag = Index
            MandaBusquedaPrevia ""
            imgCuentas(0).Tag = -1
            
        Case 6 'C�digo de Almacen
            Set frmA = New frmAlmAlPropios
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 7 'cod. ubicaciones
            
    End Select
    
    If Index = 6 Then
        PonerFoco Text3(0)
    ElseIf Index = 7 Then
        PonerFoco Text3(1)
    ElseIf Index = 8 Then
        PonerFoco Text1(22)
    Else
        PonerFoco Text1(Index + 2)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
     Case 0, 1, 3
        If Index = 0 Then
            indice = 10
        ElseIf Index = 1 Then
            indice = 18
        Else
            indice = 24
        End If
        PonerFormatoFecha Text1(indice)
        If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

     Case 2
        PonerFormatoFecha Text3(6)
         If Text3(6).Text <> "" Then frmF.Fecha = CDate(Text3(6).Text)
   End Select
   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
End Sub



Private Sub lw1_DblClick()
Dim Seleccionado As Long
Dim Sql As String
    If Modo <> 2 Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub


    If Me.DatosADevolverBusqueda2 <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un cliente. No puede ver los documentos.", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 0, 1, 2
        
    Case 3
        If lw1.SelectedItem.SmallIcon = 6 Then
            'PEDIDO CLIENTE
            
            
'            frmFacEntPedidos.DatosADevolverBusqueda2 = lw1.SelectedItem.Text
'            frmFacEntPedidos.EsHistorico = False
'            frmFacEntPedidos.Show vbModal
            
        Else
            'PROVEEDOR
            frmComEntPedidos.MostrarDatos = lw1.SelectedItem.Text
            frmComEntPedidos.EsHistorico = False
            frmComEntPedidos.Show vbModal

        End If
    Case 4
        'Deberia ver el o lo k siese
        DataGrid1EnSMOVAL
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    lw1.SetFocus
    Seleccionado = lw1.SelectedItem.Index
    CargaDatosLW
    lw1.SelectedItem.Selected = False
    Set lw1.SelectedItem = Nothing
    If lw1.ListItems.Count >= Seleccionado Then
            lw1.ListItems(Seleccionado).Selected = True
            lw1.ListItems(Seleccionado).EnsureVisible
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    Select Case Modo
        Case 5  'Eliminar lineas Art�culos x Almacen
            BotonEliminarLinea
        Case 6 'Eliminar L�neas Conjuntos
            BotonEliminarConjunto
        Case 7 'Eliminar Lineas de Control de Instalacion
            BotonEliminarInstalacion
            
        '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
        Case 8 'Eliminar lineas de codigos EAN
            BotonEliminarCodigosEAN
        '----
        
        Case Else   'Eliminar Art�culo
            BotonEliminar
    End Select
End Sub


Private Sub mnModificar_Click()
Dim Cad As String
Dim Aux As String
Dim i As Integer

    Select Case Modo
        Case 5  'Modificar lineas Art�culos x Almacen
'                cad = Text1(0).Text
'                i = InStr(1, cad, """")
'                If i > 0 Then
'                    Aux = Mid(cad, 1, i)
'                    Aux = Aux & """"
'                    Aux = Aux & Mid(cad, i + 1, Len(cad))
'                End If
'                NombreSQL cad
'                If BloqueoManual(NombreTabla, "'" & cad & "|" & Text3(0).Text & "|'") Then BotonModificarLinea
                Aux = " codartic=" & DBSet(Text1(0).Text, "T")
                If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid3, Me.Data4
                
                
        Case 6 'Modificar L�neas Conjuntos
'                If BloqueoManual(NombreTabla, "|'" & Text1(0).Text & "'|" & txtAux(0).Text & "|") Then
'                    BotonModificarConjunto Me.DataGrid1, Me.Data2
'                End If
                Aux = " codartic=" & DBSet(Text1(0).Text, "T")
                If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid1, Me.Data2
                
                
        Case 7  'Modificar Linea de Control de Instalacion
'                If BloqueoManual(NombreTabla, "|'" & Text1(0).Text & "'|" & cmdAceptar.Tag & "|") Then
'                    BotonModificarConjunto Me.DataGrid2, Me.Data3
'                End If
                Aux = " codartic=" & DBSet(Text1(0).Text, "T")
                If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid2, Me.Data3
                
        '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
        Case 8 'Modificar linea codigos EAN
            Aux = " codartic=" & DBSet(Text1(0).Text, "T")
            If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid4, Me.Data5
            
        Case Else   'Modificar Art�culos
            If BLOQUEADesdeFormulario(Me) Then BotonModificar
'            If BloqueaRegistroForm(Me) Then BotonModificar
    End Select
End Sub


Private Sub mnMtoConjuntos_Click()
    BotonConjuntos
End Sub

Private Sub mnMtoInstalaciones_Click()
    BotonInstalaciones
End Sub

Private Sub mnMtoStocksAlm_Click()
    BotonArticulosxAlmac
End Sub


Private Sub mnMtoCodigosEAN_Click()
    BotonCodigosEAN
End Sub


Private Sub mnNuevo_Click()
     Select Case Modo
        'Case 5 'A�adir lineas Art�culos x Almacen
         '       BotonAnyadirLinea   'QUITAR EL PROCEDEIMIENTO
         
        '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
        Case 5, 6, 7, 8 'A�adir L�neas Conjuntos
                  'A�adir Linea de Control de Instalacion
                BotonAnyadirConjunto2
        Case Else 'A�adir Art�culos
                BotonAnyadir
    End Select
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If Modo = 5 Then
        '------------------------------------------------------
        'Si esta insertando lineas es una cosa, si no es otra
        cmdCancelar_Click
    Else
        If (Modo = 6) Or (Modo = 7) Then 'Modo 5: Mto Lineas
                        'Modo 6: Conjuntos, Modo 7: Instalaciones
                        
            cmdRegresar_Click
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If (Not Text1(Index).MultiLine) And (Text1(Index).ScrollBars) = 0 Then ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not Text1(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not Text1(Index).MultiLine Then KEYpress KeyAscii
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

    'Si modo=1 busqueda y pierde el foco el control del nombre articulo
    'entonces pongo el foco en aceptar, ya que el 99 % de las veces
    'buscare por nomartic
    If Modo = 1 And Index = 1 Then PonerFocoObjeto cmdAceptar



    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
        
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    

    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Codigo Art�culo
            'Comprobar si ya existe el cod de articulo en la tabla
            If Modo = 3 Then 'Insertar
                If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
            End If

        Case 2 'Codigo de Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 3 'C�digo de Familia
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "sfamia", "nomfamia")
                If Text2(Index - 2).Text = "" Then Text1(Index).Text = ""
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 4 'C�digo de Marca
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "smarca", "nommarca")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 5 'C�digo Tipo Unidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "sunida", "nomunida")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 6 'Codigo Tipo Art�culo
            Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "stipar", "nomtipar")
            If Text1(Index).Text <> "" And Text2(Index - 2).Text = "" Then PonerFoco Text1(Index)
            
        Case 7 'Tipo de IVA
            'conConta: BD Contabilidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conConta, "tiposiva", "nombriva")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 10, 18, 24 'Fecha alta, Fecha �ltima compra, FECHA VIGENCIA
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)

        '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN  (se borra campo de la cabecera index 31 pasa a ser el 8)
'        Case 11, 12, 31 'numericos
        Case 11, 12, 8 'numericos
            PonerFormatoEntero Text1(Index)

        Case 13, 14, 15, 16, 17 'Precios
            'Formato tipo 2: Decimal(10,4)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 2
        
        Case 21 'Texto Control de instalaci�n
            If (Modo <> 0) Then PonerFocoBtn Me.cmdAceptar
            
        Case 24 'Margen comercial
                       
            If PonerFormatoDecimal(Text1(Index), 7) Then
                ' ---- [06/11/2009] [LAURA] : calcular el PVP
                If Modo = 3 Then PonerPrecioPVP
            End If
        Case 25
             'Precio anual mantenimiento.  Lo que ponga en su tag
             PonerFormatoDecimal Text1(Index), 8
        
        
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
    
    CadB = ObtenerBusqueda(Me, False, BuscaChekc)

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String
Dim Conexion As Byte

'    'Llamamos a al form
'    '##A mano
'    Cad = ""
'    Select Case Val(Me.imgCuentas(0).Tag)
'        Case 5  'Tipo de IVA
'            'Se llama a Busqueda desde el campo Tipos IVA
'            '#A MANO: Porque busca en la tabla tiposiva
'            'de la base de datos de Contabilidad
'            Cad = Cad & "C�digo|tiposiva|codigiva|N||20�Denominacion|tiposiva|nombriva|T||70�"
'            Tabla = "tiposiva"
'            Titulo = "Tipos de IVA"
'            Conexion = conConta    'Conexi�n a BD: Conta
'        Case Else   'Registro de la tabla de cabeceras: sartic
'            Cad = Cad & ParaGrid(Text1(0), 23, "C�digo")
'            Cad = Cad & ParaGrid(Text1(1), 58, "Denominaci�n")
'            Cad = Cad & ParaGrid(Text1(9), 19, "Cod. asoc.")
'            Tabla = "sartic"
'            Titulo = "Art�culos"
'            Conexion = conAri    'Conexi�n a BD: Aritaxi
'    End Select
'
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vTabla = Tabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 1
'        frmB.vConexionGrid = Conexion
''        frmB.vBuscaPrevia = VPrevia
'        frmB.vCargaFrame = False
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda2 <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If
'    Screen.MousePointer = vbDefault


    Set frmArt = New frmBasico2
    
    AyudaArticulos frmArt, Text1(0).Text, CadB
    
    Set frmArt = Nothing
    

End Sub


Private Sub PonerCadenaBusqueda()

    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ning�n registro en la tabla " & NombreTabla & " para ese criterio de B�squeda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
        PonerCamposAlmacenes2
        'David 28 Nov 2008
        ' Si es conjunto mostrare sus solapa
        If Me.chkConjunto.Value = 1 Or Me.chkConjunto.Value = 2 Then PonerModoOpcionesMenu 2
    
        If DatosADevolverBusqueda2 <> "" Then PonerFocoBtn Me.cmdRegresar
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda", Err.Description
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim Impor As Currency

    If Data1.Recordset.EOF Then Exit Sub
    
    lblIndicador.Caption = "Datos articulo"
    lblIndicador.Refresh
    PonerCamposForma Me, Data1
    

    
    
    Text2(0).Text = PonerNombreDeCod(Text1(2), conAri, "sprove", "nomprove")
    Text2(1).Text = PonerNombreDeCod(Text1(3), conAri, "sfamia", "nomfamia")
    Text2(2).Text = PonerNombreDeCod(Text1(4), conAri, "smarca", "nommarca")
    Text2(3).Text = PonerNombreDeCod(Text1(5), conAri, "sunida", "nomunida")
    Text2(4).Text = PonerNombreDeCod(Text1(6), conAri, "stipar", "nomtipar")
    mPorIva = "porceiva"
    Text2(5).Text = DevuelveDesdeBD(conConta, "nombriva", "tiposiva", "codigiva", Text1(7).Text, "N", mPorIva)
    
    
    lblIndicador.Caption = "Importes"
    lblIndicador.Refresh
    PonerSumaStocks 'Poner la suma total de stocks de los almacenes donde esta el artic
    
    BloquearChecks Me, Modo

    PrimeraVez = False

    PonerCamposLineas True 'Pone los datos de las tablas de lineas de Componentes e Instalaciones
    
    'Lista campos
    CargaDatosLW
    
    'Pongo el PVP con IVA
    If mPorIva = "porceiva" Then mPorIva = 0
    Impor = CCur(mPorIva)
    Impor = Round2((Impor * Data1.Recordset!preciove) / 100, 4) + Data1.Recordset!preciove
    Me.txtPVPIVA.Text = Format(Impor, FormatoPrecio)
    
    
    
    
    
    'Si tiene conjuntos
    If Val(Data1.Recordset!Conjunto) = 1 Then ponerDatosConjuntos
    
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub PonerCamposLineas(enlaza As Boolean)
'Carga las Pesta�as con las tablas de lineas de Conjunto o Instalaciones
'segun la pesta�a de datos a mostrar
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    'Conjuntos
    CargaGrid DataGrid1, Data2, enlaza
    'Instalaciones
    CargaGrid DataGrid2, Data3, enlaza
    'Stocks
    CargaGrid DataGrid3, Data4, enlaza

    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
    'lineas codigos EAN
    CargaGrid DataGrid4, Data5, enlaza
    '----

    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas", Err.Description
'    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerSumaStocks()
Dim rst As ADODB.Recordset
Dim Sql As String
    
    Sql = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", Text1(0).Text, "T")
    If Sql <> "" Then
        Sql = "select sum(canstock) from salmac where codartic=" & DBSet(Text1(0).Text, "T")
        Set rst = New ADODB.Recordset
        rst.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not rst.EOF Then
            Me.txtSumaStock.Text = rst.Fields(0).Value
        End If
        rst.Close
        Set rst = Nothing
    Else
        Me.txtSumaStock.Text = 0
    End If
End Sub


Private Sub PonerPrecioPVP()
Dim cArt As CArticulo

    Set cArt = New CArticulo
    cArt.Codigo = Text1(0).Text
    cArt.PrecioUltCom = ComprobarCero(Text1(15).Text)
    cArt.MargenComercial = ComprobarCero(Text1(25).Text)
    cArt.PrecioVenta = ComprobarCero(Text1(17).Text)
    Text1(17).Text = cArt.AplicarMargenComercial 'obtiene el nuevo PVP
    FormateaCampo Text1(17)
    
    cArt.TipoIVA = Text1(7).Text
    
    
'    If cArt.LeerDatos(Me.parCodArtic) Then
'        Text1(2).Text = cArt.PrecioUltCom
'        Text1(2).Text = Format(Text1(2).Text, FormatoPrecio)
'
'        Text1(3).Text = cArt.PrecioVenta 'precio venta actual
'        Text1(3).Text = Format(Text1(3).Text, FormatoPrecio)
'        Text1(5).Text = cArt.MargenComercial
'        Text1(5).Text = Format(Text1(5).Text, FormatoPorcen)
'
'        Text1(4).Text = cArt.AplicarMargenComercial 'obtiene el nuevo PVP
'        Text1(4).Text = Format(Text1(4).Text, FormatoPrecio)
'    End If
    Set cArt = Nothing


End Sub



Private Sub PonerCamposAlmacenes2()
    If Data4.Recordset.EOF Then Exit Sub
    PonerCamposFormaFrame Me, "Text3", Data4
    
    'Rellenar el nombre correspondiente al c�digo de los TextBox de indice 8
    Text2(8).Text = PonerNombreDeCod(Text3(0), conAri, "salmpr", "nomalmac", "codalmac")
    
    'El check del inventario
    chkInventario.Value = DBLet(Data4.Recordset!statusin, "N")
    
    '-- Esto permanece para saber donde estamos
'    lblIndicador.Caption = Data4.Recordset.AbsolutePosition & " de " & Data4.Recordset.RecordCount
End Sub


'Private Function ComprobarEsInstalacion() As Boolean
'Dim devuelve As String
'Dim EsInstal As Boolean
'
'    EsInstal = False
'    If Not (vParamAplic.Frecuencias) Then Exit Function ' si no estan activadas las frecuencias no se muestra n�
'    If Text1(3).Text <> "" Then
'        devuelve = DevuelveDesdeBDNew(conAri, "sfamia", "instalac", "codfamia", Text1(3).Text, "N")
'        If devuelve = "1" Then
'            EsInstal = CBool(devuelve)
'        Else
'            EsInstal = False
'        End If
'    End If
'    ComprobarEsInstalacion = EsInstal
'End Function
'
'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
Dim NumReg As Byte

    'Actualiza Iconos Insertar,Modificar,Eliminar
'--    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    For i = 0 To txtAux.Count - 1
        Text1(i).BackColor = vbWhite
    Next i
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    
    'Modo 2. Hay datos y estamos visualizandolos
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
    b = (Kmodo = 2) Or (Modo = 5) Or (Modo = 6) Or (Modo = 7) Or (Modo = 8)
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda2 <> "" Then
        cmdRegresar.visible = b
        cmdRegresar.Caption = "&Regresar"
    Else
        cmdRegresar.visible = False
    End If
    
    b = Modo <> 0 And Modo <> 2 'And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    'Poner Flechas de Desplazamiento Visibles o no
    NumReg = 1
    If (Modo = 2) Then
        If Not Data1.Recordset.EOF Then
            If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
    ElseIf Modo = 5 Then
        If Not Data4.Recordset.EOF Then
            If Data4.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
    End If
    b = (Modo = 2) Or (Modo = 5)
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    'campos Precio medio y ponderado bloqueados, pq son calculados
    BloquearTxt Text1(13), True
    BloquearTxt Text1(14), True
    'fecha ultimo cambio PVP bloqueado pq se actualiza automaticamente
    BloquearTxt Text1(27), True
    
    
    Me.FrameArtxAlmac.Enabled = (Modo = 5)
    'Me.FrameArtxAlmac2.visible = (Modo = 5)
    If Me.FrameArtxAlmac.Enabled Then
        If Modo = 5 And ModificaLineas = 2 Then BloquearTxt Text3(0), True
         'Me.FrameArtxAlmac.Height = 2010
         'Me.FrameArtxAlmac.Top = 2260
         'Me.FrameArtxAlmac.Left = 360
    End If
    Me.FrameDatosAlmacen2.visible = (Modo <> 5)
        
    b = (Modo = 1 Or Modo = 3 Or Modo = 4) '1:Busqueda, 3:Insertar, 4:Modificar
    cboArticuloVarios.Enabled = b
    cboStatus.Enabled = b
    'Bloquear los checkbox
    BloquearChecks Me, Modo
'--
'    cmdCancelar.visible = b
'    cmdAceptar.visible = b
    For i = 0 To 1
        Me.imgFecha(i).Enabled = b
    Next i
    For i = 0 To 5
        Me.imgCuentas(i).Enabled = b
    Next i
    
    'Numero de orden
    'Busquedas o insertar modificar el supr usuario
    b = vUsu.Nivel = 0 And (Modo = 3 Or Modo = 4)
    b = b Or Modo = 1
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN  (se borra campo de la cabecera index 31 pasa a ser el 8)
'    BloquearTxt Text1(31), Not b
    BloquearTxt Text1(8), Not b
    '----
    
    chkVistaPrevia.Enabled = (Modo <= 2)

    'Bton generar denominacion solo en descriptores y en modo insertar
    Me.cmdGenerar.visible = vParamAplic.Descriptores And Modo = 3

    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Poner opciones de menu seg�n modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
                        
                        
                        
    'Los tag's de los campos de sctock NO estaran visibles si
    'inserto,modifico o busco en la PPAL
    If Modo = 1 Or Modo = 3 Or Modo = 4 Then
        AccionesSobreTagText3_ True, False
    Else
        'Los vuelvo a poner
        AccionesSobreTagText3_ False, False
    End If
    
    'El listview
    If Modo <> 2 Then lw1.ListItems.Clear


    'cmdACtualizar importes en conjuntos
    cmdActualizarImportes1(0).visible = Modo = 6 And (ModificaLineas <> 1)
    cmdActualizarImportes1(1).visible = Modo = 6 And (ModificaLineas <> 1)
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean
Dim EsInstal As Boolean
Dim i As Integer
Dim bAux As Boolean

    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
    b = (Modo = 2) Or (Modo = 5) Or (Modo = 6) Or (Modo = 7) Or (Modo = 8)
    'Insertar
    Toolbar1.Buttons(1).Enabled = (b Or Modo = 0 Or Modo = 1)
    Me.mnNuevo.Enabled = (b Or Modo = 0 Or Modo = 1)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    Toolbar1.Buttons(8).Enabled = Not DeConsulta

'--
'    b = (Modo = 2) And Not DeConsulta
'    'Lineas Articulos x Almacen
'    Toolbar1.Buttons(10).Enabled = b And vUsu.Nivel <= 1
'    Me.mnMtoStocksAlm.Enabled = b And vUsu.Nivel <= 1
'
'    'Lineas Conjuntos
'    '                       antes era B, ahora true
'    Toolbar1.Buttons(11).Enabled = (b And (Me.chkConjunto.Value = 1 Or Me.chkConjunto.Value = 2))
'    '                       antes era B, ahora true
'    Me.mnMtoConjuntos.Enabled = (True And (Me.chkConjunto.Value = 1 Or Me.chkConjunto.Value = 2))
'    Me.SSTab1.TabVisible(2) = (Me.chkConjunto.Value = 1 Or Me.chkConjunto.Value = 2)
'    If Me.SSTab1.TabVisible(2) Then
'        'Me.cmdActualizarImportes1(0).Enabled = Not DeConsulta And vUsu.Nivel <= 1
'        Me.cmdActualizarImportes1(0).Enabled = vUsu.Nivel <= 1
'        'Me.cmdActualizarImportes1(1).Enabled = Not DeConsulta And vUsu.Nivel <= 1
'        Me.cmdActualizarImportes1(1).Enabled = vUsu.Nivel <= 1
'    End If
'
'    'Lineas Instalaciones
'    'EsInstal = ComprobarEsInstalacion
'    EsInstal = True
'    b = b And EsInstal
'    Toolbar1.Buttons(12).Enabled = b
'    Me.mnMtoInstalaciones.Enabled = b
'    Me.SSTab1.TabVisible(3) = EsInstal
'
'
'    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
'    'Lineas cod. EAN
'    Toolbar1.Buttons(13).Enabled = b
''    Me.mnMtoInstalaciones.Enabled = B
'    ' -----

    b = (Modo = 0) Or (Modo = 2) Or (Modo = 1)
    'Buscar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    b = (Modo = 2 Or Modo = 3 Or Modo = 4)
    
    For i = 0 To ToolAux.Count - 1 '[Monica]30/09/2013: antes - 1
        If i = 0 Or i = 1 Then
            ToolAux(i).Buttons(1).Enabled = b And vUsu.Nivel <= 1
        Else
            ToolAux(i).Buttons(1).Enabled = b
        End If
        
        Select Case i
            Case 0 'stocks
                If b Then bAux = (b And Me.Data4.Recordset.RecordCount > 0) And vUsu.Nivel <= 1
            Case 1 'conjuntos
                If b Then bAux = (b And Me.Data2.Recordset.RecordCount > 0) And vUsu.Nivel <= 1
            Case 2 'instalaciones
                If b Then bAux = (b And Me.Data3.Recordset.RecordCount > 0)
            Case 3 'codean
                If b Then bAux = (b And Me.Data5.Recordset.RecordCount > 0)
        End Select
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
    
    
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoFrame(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
    ModoFrame = Kmodo
    
    Select Case ModoFrame
        Case 0  'MODO INICIAL
                For i = 0 To Me.Text3.Count - 1
                    BloquearTxt Text3(i), True
                Next i
                Me.imgFecha(2).Enabled = False
                Me.imgCuentas(6).Enabled = False
                Me.chkInventario.Enabled = False
'--                PonerBotonCabecera True
                
        Case 3  'Modo INSERTAR
                
                BloquearTxt Text3(0), False
                Text2(8).Text = ""
    End Select
    If ModoFrame = 3 Or ModoFrame = 4 Then
        '3=Insertar,  4=Modificar
        
        'Nuevo Marzo 2010
        ' Ni stock, ni los datos de inventario se pueden insertar
        BloquearTxt Text3(0), ModoFrame = 3
        
        For i = 1 To Me.Text3.Count - 1
        
            If i = 1 Or i >= 5 Then
                b = True
            Else
                b = False
            End If
            BloquearTxt Text3(i), b
            If ModoFrame = 3 Then
                If b And i = 1 Then
                    Text3(i).Text = "0"
                Else
                    Text3(i).Text = ""
                End If
            End If
        Next i
        chkInventario.Enabled = False
        Me.imgFecha(2).Enabled = False
        Me.imgCuentas(6).Enabled = (ModoFrame = 3)
        PonerFoco Text3(1)
'--        PonerBotonCabecera False
    End If
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean


    DatosOk = False
    
    'Comprobamos que el campo dias de garantia si no tiene valor lo
    'ponemos a 0 para q no de error que no puede ser nulo
    If Trim(Me.Text1(11).Text) = "" Then Text1(11).Text = "0"
    
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    'Comprobar si ya existe el cod de trabajador en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then
            b = False
        Else
            'No podemos crear este articulo ya que es una constante que utiliza
            If Text1(0).Text = "@1@" Then
                MsgBox "Imposible crear articulo @1@", vbExclamation
                b = False
            End If
            
            If Mid(Text1(0).Text, 1, 2) = "::" Then
                MsgBox "Imposible crear articulo ::", vbExclamation
                b = False
            End If
        
        
        End If
        
        
        If b Then
            'Comprobamos si ha puesto(insertando) el numero de orden
            'Si es asi, k tiene valor
            '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN  (se borra campo de la cabecera index 31 pasa a ser el 8)
'            If Text1(31).Text <> "" Then
            If Text1(8).Text <> "" Then
                BuscaChekc = DevuelveDesdeBD(conAri, "codartic", "sartic", "numorden", Text1(8).Text)
                If BuscaChekc <> "" Then
                    MsgBox "Ya existe el numero de orden", vbExclamation
                    b = False
                End If
            Else
                'No ha puesto ninguno. Le asigno
                'ASigno el max mas uno
                BuscaChekc = DevuelveDesdeBD(conAri, "max(numorden)", "sartic", "1", "1")
                If BuscaChekc = "" Then BuscaChekc = "0"
                BuscaChekc = Val(BuscaChekc) + 1
                Text1(8).Text = BuscaChekc
            End If
            '----
        End If
    End If
    
    
    'si se ha cambiado el precio venta PVP actualizamos la fecha de
    'ult. cambio PVP
    If Modo = 4 Then 'modo modificar
        'si se ha modificado el ult. precio compra la fecha ult. compra
        'debe tener valor
        If Text1(15).Text <> "" And Trim(Text1(18).Text) = "" Then
            b = False
            MsgBox "Si hay precio de ult. compra la fecha de ult. compra debe tener valor.", vbInformation
        End If
        
        
        'si se ha modificado el precio venta PVP actualizamos campos
        'para guardarlo correctamente
        If CCur(Me.Text1(17).Text) <> CCur(Me.Data1.Recordset!preciove) Then
            Me.Text1(26).Text = Format(Now, "dd/mm/yyyy")
        End If
        
        
        
        'Cuando modificamos, si pasamos un articulo a CADUCADO, entonces comproaremos
        'si tiene sctock. Si es asi NO dejammos continuar
        If Me.cboStatus.ListIndex = 2 And Val(Data1.Recordset!codstatu) < 2 Then
            If Me.chkctrstock.Value = 1 Then
                'Lleva stcok
                'Comprobamos k valor tiene
                BuscaChekc = TotalRegistros("select sum(canstock) from salmac where codartic='" & DevNombreSQL(Text1(0).Text) & "'")
                If Val(BuscaChekc) > 0 Then
                    MsgBox "No podemos pasar un �rticulo a caducado teniendo stock.", vbExclamation
                    Exit Function
                End If
            End If
        End If
        
        
        
    End If 'Modificando
    
    DatosOk = b
End Function


Private Function DatosOkConjunto() As Boolean
Dim b As Boolean
Dim devuelve As String

    DatosOkConjunto = False
    b = True
    If txtAux(1).Text = "" Then
         MsgBox "El campo Cantidad no puede ser nulo", vbExclamation, "Art�culos"
         b = False
    End If
        
    If Not IsNumeric(txtAux(1).Text) Then
        MsgBox "La cantidad de Art�culos tiene que ser num�rico", vbExclamation
        b = False
    End If
    If Not b Then Exit Function
    
    'Comprobamos  si existe, solo si estamos insertando (ModificaLineas=1)
    'conAri: conexion a BD Aritaxi
    devuelve = DevuelveDesdeBDNew(conAri, "sarti1", "codartic", "codartic", Text1(0).Text, "T", , "codarti1", txtAux(0).Text, "T")
    If ModificaLineas = 1 And devuelve <> "" Then
        b = False
        devuelve = "Ya existe el Art�culo de Conjunto: " & vbCrLf
        devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
        devuelve = devuelve & "Descripci�n: " & txtAux2.Text
        
        MsgBox devuelve, vbExclamation, "Art�culos"
    End If
    If Not b Then Exit Function
    
    'Comprobar que el articulo no tiene conjuntos, solo si estamos insertando (ModificaLineas=1)
    'Si tiene conjuntos no puede ser elemento de conjunto de otro articulo
    If ModificaLineas = 1 And DevuelveDesdeBDNew(conAri, "sartic", "conjunto", "codartic", txtAux(0).Text, "N") = "1" Then
        b = False
        devuelve = "No es un Art�culo de Conjunto: " & vbCrLf
        devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
        devuelve = devuelve & "Descripci�n: " & txtAux2.Text & vbCrLf & vbCrLf
        devuelve = devuelve & "�Continuar?"
        If MsgBox(devuelve, vbQuestion + vbYesNo) = vbYes Then b = True
    End If
    DatosOkConjunto = b
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim devuelve As String

    DatosOkLinea = False
    b = True
    
    'Campo de cantidad de Stock (Son decimales)
    If Trim(Text3(1).Text) = "" Or IsNull(Text3(1).Text) Then
        MsgBox "El campo Cantidad Stock no puede ser nulo", vbExclamation, "Art�culos"
        b = False
    End If
    If Not b Then Exit Function
    
    'Comprobamos  si existe
    devuelve = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", Text1(0).Text, "T", , "codalmac", Text3(0).Text, "N")
    If ModificaLineas = 1 And devuelve <> "" Then
        b = False
        devuelve = "Ya existe el Art�culo en el Almacen: " & vbCrLf
        devuelve = devuelve & "Codigo: " & Text3(0).Text & vbCrLf
        devuelve = devuelve & "Descripci�n: " & Text2(8).Text
        MsgBox devuelve, vbExclamation, "Art�culos"
    End If
    If b Then
        devuelve = ""
        If Text3(2).Text <> "" And Text3(4).Text <> "" Then
            If ImporteFormateado(Text3(2).Text) > ImporteFormateado(Text3(4).Text) Then devuelve = "Importe stock minimo mayor que el stock maximo"

        End If
        
        If devuelve = "" Then
            If Text3(3).Text <> "" Then
                'Veremos si esta entre maximo y minimo
                If Text3(2).Text <> "" Then
                    If ImporteFormateado(Text3(2).Text) > ImporteFormateado(Text3(3).Text) Then devuelve = "Importe stock minimo mayor que el punto pedido"
                End If
                
                If Text3(4).Text <> "" Then
                    If ImporteFormateado(Text3(3).Text) > ImporteFormateado(Text3(4).Text) Then devuelve = "Importe stock maximo menor que el punto pedido"
                End If
            End If
        End If
        
        If devuelve <> "" Then
            MsgBox devuelve, vbQuestion
            b = False
        End If
    End If
    DatosOkLinea = b
End Function


Private Sub Text3_GotFocus(Index As Integer)
    kCampo = Index
    If ModificaLineas <> 0 Then
        ConseguirFoco Text3(Index), 4
    Else
        ConseguirFoco Text3(Index), 2
    End If
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        If Index = 8 Then
            PonerFocoBtn Me.cmdAceptar
        Else
            KeyAscii = 0
            SendKeys "{tab}"
        End If
    End If
End Sub


Private Sub Text3_LostFocus(Index As Integer)
    
     If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'Codigo Almacen
             Text2(8).Text = PonerNombreDeCod(Text3(Index), conAri, "salmpr", "nomalmac")
             If Text2(8).Text = "" Then Text3(0).Text = ""
                
        Case 1, 2, 3, 4, 5 'Stocks, Punto Pedido
                'Formato tipo 1: Decimal(12,2)
                If Trim(Text3(Index)) <> "" Then PonerFormatoDecimal Text3(Index), 1
        
        Case 6  'Fecha Inventario
            If Text3(Index).Text <> "" Then PonerFormatoFecha Text3(Index)

        Case 7  'Hora Inventario
            If Trim(Text3(Index).Text) <> "" Then PonerFormatoHora Text3(Index)
    End Select
End Sub


Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    
    Modo = 5 + Index
    
    NumTabMto = Me.SSTab1.Tab
'    TituloLinea = Me.SSTab1.TabCaption(SSTab1.Tab)
    If Index <> 3 Then PonerModo 5 + Index
    
'    Select Case Index
'        Case 0
'            NomTablaLineas = "salmac"
'        Case 1
'            NomTablaLineas = "sarti1"
'        Case 2
'            NomTablaLineas = "sarti2"
'        Case 3
'            NomTablaLineas = "sarti3"
'    End Select

    Select Case Button.Index
        Case 1
'            BotonAnyadirLinea
            BotonAnyadirConjunto2
        Case 2
'            BotonModificarLinea
            Select Case Index
                Case 0
                    BotonModificarConjunto DataGrid3, Data4
                Case 1
                    BotonModificarConjunto DataGrid1, Data2
                Case 2
                    BotonModificarConjunto DataGrid2, Data3
                Case 3
                    BotonModificarConjunto DataGrid4, Data5
            End Select
        Case 3
            Select Case Index
                Case 0
                    BotonEliminarLinea
                Case 1
                    BotonEliminarConjunto
                Case 2
                    BotonEliminarInstalacion
                Case 3
                    BotonEliminarCodigosEAN
            End Select
            
            
            
            
        Case Else
    End Select


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            mnVerTodos_Click
        Case 1  'Nuevo
           mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
            
'        Case 10  'Stocks Almacenes
'            mnMtoStocksAlm_Click
'        Case 11 'Conjuntos
'            mnMtoConjuntos_Click
'        Case 12 'Instalaciones
'            mnMtoInstalaciones_Click
'
'        '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
'        Case 13 'Codigos EAN
'            mnMtoCodigosEAN_Click
'        '----
            
        Case 8 'Imprimir Listado de Articulos
            BotonImprimir
'        Case 16 'Salir
'            mnSalir_Click
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


Private Sub CargarComboStatus()
'### Combo Situaci�n Art�culo
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Normal, 1-Bloqueado, 2-Caducado

    cboStatus.Clear
    cboStatus.AddItem "Normal"
    cboStatus.ItemData(cboStatus.NewIndex) = 0
    
    cboStatus.AddItem "Bloqueado"
    cboStatus.ItemData(cboStatus.NewIndex) = 1
    
    cboStatus.AddItem "Caducado"
    cboStatus.ItemData(cboStatus.NewIndex) = 2
    
End Sub


Private Sub CargarComboArticuloVarios()
'### Combo Situaci�n Art�culo
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-No, 1-Si, 2-Rectificacion
 
    cboArticuloVarios.Clear
    cboArticuloVarios.AddItem "No"
    cboArticuloVarios.ItemData(cboArticuloVarios.NewIndex) = 0
    
    cboArticuloVarios.AddItem "Si"
    cboArticuloVarios.ItemData(cboArticuloVarios.NewIndex) = 1
    
    cboArticuloVarios.AddItem "Rectificaci�n"
    cboArticuloVarios.ItemData(cboArticuloVarios.NewIndex) = 2
    
End Sub


Private Function InsetarArticulosPorAlmacen(Optional cadErr As String) As Boolean
'Inserta en la tabla salmac una fila del art�culo que se esta insertando
'para cada uno de los almacenes que existen en la tabla salmpr
Dim vCodArtic As String, vcodalmac As Integer
Dim rsAlmPr As ADODB.Recordset
Dim Cad As String
    
    On Error GoTo EInsEnAlm

    vCodArtic = Text1(0).Text
    Set rsAlmPr = New ADODB.Recordset
    Cad = "Select codalmac from salmpr order by codalmac;"
    rsAlmPr.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not rsAlmPr.EOF
        vcodalmac = rsAlmPr.Fields(0).Value
        Cad = "INSERT INTO salmac (codartic,codalmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin)"
        Cad = Cad & " VALUES (" & DBSet(vCodArtic, "T") & "," & vcodalmac & ",0,0,0,0,0,NULL,NULL,0)"
        conn.Execute Cad
        rsAlmPr.MoveNext
    Wend
        
    rsAlmPr.Close
    Set rsAlmPr = Nothing
    InsetarArticulosPorAlmacen = True
    Exit Function
    
EInsEnAlm:
    InsetarArticulosPorAlmacen = False
    'MuestraError Err.Number, "Insertando Art�culo en Almacenes.", Err.Description
    cadErr = "Insertando Art�culo en Almacenes: " & vbCrLf & Err.Description
End Function
   
   
   
Private Function InsertarPreciosPorTarifa2(Optional cadErr As String) As Boolean
'Insertar en la lista de precios las tarifas para el articulo
Dim Sql As String
Dim RS As ADODB.Recordset
Dim cTar As CTarifaArt
Dim NoOK As Boolean
Dim Cad As String
'Dim codlista As Double

    On Error GoTo ErrInsPrecio
    
    'comprobamos que el PVP tenga valor y sea >0 para insertar lista de precios
    InsertarPreciosPorTarifa2 = True
    If Text1(17).Text = "" Then Exit Function
    If Not (CCur(Text1(17).Text) > 0) Then Exit Function
    
    

    InsertarPreciosPorTarifa2 = False
    
    
    
    'David. Enero 2009
    If vParamAplic.CreaTarifasArticulo = 0 Then
        'NO CREO NINGUNA.
        'Salgo dando OK
        InsertarPreciosPorTarifa2 = True
        Exit Function
    End If
    
    '---- [14/09/2009] LAURA
    If vParamAplic.CreaTarifasArticulo = 2 Then 'crear todas las tarifas
    '----
    
        Sql = "SELECT * FROM starif WHERE NOT ISNULL(margecom) "
        
    '---- [14/09/2009] LAURA
    Else 'crear solo la tarifa general
        Cad = DevuelveDesdeBD(conAri, "min(codlista)", "starif", "1", "1")
        If Cad = "" Then Cad = "0"
        Sql = "SELECT * FROM starif WHERE NOT ISNULL(margecom) and codlista = " & Val(Cad)

    
    End If
    '----
        
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada tarifa insertar un linea en la tabla de lista de precios
    'por cada codartic,codtarif
    
    '23 Abril 2008
    'Tb, en funcion sobre donde se aplica el margen se hara una cosa u otra
    ' Sobre PVP o sobre PUC
    'FALTA###
    NoOK = False
    While Not RS.EOF
        Set cTar = New CTarifaArt
        cTar.CodigoArticulo = Text1(0).Text
        cTar.CodigoTarifa = RS!codlista
        'Aqui dependera de una cosa u otra para lo del PVP / UPC
        ' 1.-  "    "  va sobre el UPC
        ' 0.- La tarifa va sobre el PVP
        If DBLet(RS!opcionINC, "N") = 0 Then
            'PVP
            cTar.PrecioActual = CCur(Text1(17).Text) 'precio venta al publico (pvp)
        Else
            If Text1(15).Text = "" Then
                cTar.PrecioActual = 0
            Else
                cTar.PrecioActual = CCur(Text1(15).Text) 'precio venta al publico (pUC)
            End If
        End If
        If cTar.InsertarPrecios = False Then NoOK = True
        Set cTar = Nothing
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
    If NoOK Then
        InsertarPreciosPorTarifa2 = False
        cadErr = "Los precios del art�culo por tarifa NO se han introducido correctamente."
    Else
        InsertarPreciosPorTarifa2 = True
    End If
        
    Exit Function
    
ErrInsPrecio:
    InsertarPreciosPorTarifa2 = False
    cadErr = "Insertar precios por tarifa: " & Err.Description
End Function
   
   
Private Function BloquearTarifas(codArtic As String) As Boolean
Dim cadWHERE As String
    cadWHERE = "codartic=" & DBSet(codArtic, "T")
    BloquearTarifas = BloqueaRegistro("slista", cadWHERE)
End Function
   
   
Private Function ActualizarPreciosVenta() As Boolean
'si se modifica el precio ult. compra a mano preguntar si quiere modificar
'el PVP y las tarifas de venta desde el formulario de actualizar precios
Dim precioUC As Currency 'precio ult. compra (valor actual)
Dim FechaUC As String
Dim newPrecioUC As Currency
Dim bActualizar As Boolean
Dim Cad As String

    'Comprobar si se ha modificado el precio desde la ultima compra
    'y preguntar quiere modificar el PVP del articulo aplicandole su margen
    'y el precio de las TArifas aplicandole el margen
    '-- Laura 19/12/2006: el precio de compra es el precio con los descuentos (importe/cantidad)
    precioUC = CCur(DBLet(Me.Data1.Recordset!precioUC, "N"))
    If Not IsNull(Me.Data1.Recordset!ultfecco) Then FechaUC = DBLet(Me.Data1.Recordset!ultfecco, "F")
    newPrecioUC = ImporteFormateado(Text1(15).Text)
    
    bActualizar = False
    Cad = ""
    If precioUC <> newPrecioUC Then
        If FechaUC = "" Then
            bActualizar = True
        ElseIf CDate(Text1(18).Text) >= CDate(FechaUC) Then
            bActualizar = True
        Else
            
        End If
        Cad = "precio de �ltima compra"
    End If
    
    
    '## LAURA 25/06/2008
    If Not bActualizar Then
        '-- comprobar si se ha modificado el margen comercial y
        '-- en este caso recalcular tambien el PVP y tarifas
        precioUC = CCur(DBLet(Me.Data1.Recordset!margecom, "N")) 'margen actual
        newPrecioUC = ImporteFormateado(Text1(25).Text) 'margen nuevo
        If precioUC <> newPrecioUC Then bActualizar = True
        Cad = "margen comercial"
    End If
    '##
    
    
     If bActualizar Then
            If MsgBox("Se ha modificado el " & Cad & "." & vbCrLf & "�Desea actualizar los precios de venta?", vbQuestion + vbYesNo) = vbYes Then
                'Comprobar que el art�culo tiene margen comercial
                If ArticuloTieneMargen(Text1(0).Text) Then
                    'Llamar al form de actualizar precios venta
                    frmComActPrecios.parCodArtic = Text1(0).Text
                    frmComActPrecios.parNomArtic = Text1(1).Text
                    frmComActPrecios.Show vbModal
                End If
            End If
        End If
    
    
    
'        If CDate(Text1(18).Text) >= CDate(FechaUC) Then
'            If MsgBox("Se ha modificado el precio �ltima compra." & vbCrLf & "�Desea actualizar los precios de venta?", vbQuestion + vbYesNo) = vbYes Then
'                'Comprobar que el art�culo tiene margen comercial
'                If ArticuloTieneMargen(txtAux(1).Text) Then
'                    'bloquear las tarifas del articulo para modificar
''                                If BloqueaRegistro("slista", "codartic=" & DBSet(txtAux(1).Text, "T")) Then
'                        'Aplicar margen comercial a los precios
'                        'Modificar precios de venta en articulo y tarifas
'                        frmComActPrecios.parCodArtic = txtAux(1).Text
'                        frmComActPrecios.parNomArtic = txtAux(2).Text
''                            frmcomactprecios.parPrecioUC =
'                        frmComActPrecios.Show vbModal
''                                End If
'                End If
'            End If
'        End If   'Fecha ultima compra
'    End If  'Precio ultima compra


End Function
  
  
  
  
Private Function ActualizarPreciosPorTarifa() As Boolean
Dim QueTipo As Byte
Dim Importe As Currency
Dim Aux As Currency
       'Reutilizo BuscaChekc
       QueTipo = 100
       BuscaChekc = ""
       
       '- ver si se ha modificado el precion venta PVP
       Importe = DBLet(Data1.Recordset!preciove, "N")
       If Importe <> CCur(Text1(17).Text) Then
            BuscaChekc = "-el precio de venta." & vbCrLf
            QueTipo = 0 'que mire tarifas PVP
       End If
        
       '- ver si se ha modificado el precio ultima compra
       Importe = DBLet(Data1.Recordset!precioUC, "N")
       Aux = 0
       If Text1(15).Text <> "" Then Aux = CCur(Text1(15).Text)
       If Importe <> Aux Then
            BuscaChekc = BuscaChekc & "-el precio de ultima compra." & vbCrLf
            If Aux = 0 Then BuscaChekc = BuscaChekc & "*****  Precio ultima compra=  CERO    ****** " & vbCrLf
            If QueTipo = 0 Then
                QueTipo = 2  'Que mire las dos
            Else
                QueTipo = 1  'que mire solo en tarifas U.P.C.
            End If
        End If
            
        '## LAURA 25/06/2008
        'si el tipo es 0 o 2 ya se va a modificar el PVP y no comprobamos margen
'        If QueTipo <> 0 And QueTipo <> 2 Then
'            'si se ha modificado el margen comercial tambien
'            'hay que actualizar el PVP
'            Importe = DBLet(Data1.Recordset!margecom, "N")
'            If Importe <> CCur(Text1(25).Text) Then
'                 BuscaChekc = "-el margen comercial." & vbCrLf
'                 If QueTipo = 1 Then
'                    QueTipo = 2  'Que mire las dos
'                 Else
'                    QueTipo = 0 'que mire tarifas PVP
'                 End If
'            End If
'        End If
        '##
        
            
        If QueTipo <> 100 Then
'[Monica]09/03/2011 Qu� tarifas ????
'            BuscaChekc = vbCrLf & BuscaChekc & vbCrLf
'            BuscaChekc = "Se han modificado: " & BuscaChekc & "�Desea actualizar las tarifas de precios?"
'            If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbYes Then
'
'                Screen.MousePointer = vbHourglass
'                ActualizarPreciosPorTarifaDOS QueTipo
'                Screen.MousePointer = vbDefault
'            End If
        End If
    
    
End Function
  
  
                    'QueTipoActualiza : 0. PVP
                    '                   1. UPC
                    '                   2. LOS DOS
Private Function ActualizarPreciosPorTarifaDOS(PVP As Byte, Optional cadErr As String) As Boolean
'Actualiza en la lista de precios las tarifas para el articulo
Dim Sql As String
Dim RS As ADODB.Recordset
Dim cTar As CTarifaArt
Dim NoOK As Boolean
Dim menErr As String
Dim newPrecio As Currency


    On Error GoTo ErrActPrecio
    
    '-- comprobamos que el PVP tenga valor y sea >0 para insertar lista de precios
    ActualizarPreciosPorTarifaDOS = True
    
    
    
    '-- comprobar que para ese articulo en la tabla de tarifas no haya ningun registros
    '   con valor en el campo precio_nuevo
    Sql = "SELECT COUNT(*) FROM slista WHERE codartic=" & DBSet(Text1(0).Text, "T")
    Sql = Sql & " AND not isnull(precionu) and precionu>0"
    If RegistrosAListar(Sql) > 0 Then
        MsgBox "No se pueden actualizar las tarifas del art�culo." & vbCrLf & "Tiene precios nuevos.", vbExclamation
        Exit Function
    End If
    
    
    ActualizarPreciosPorTarifaDOS = False
    
    If Not BloquearTarifas(Text1(0).Text) Then
        MsgBox "NO se han actualizado las tarifas de precios.", vbExclamation, "Actualizar precios"
        Exit Function
    End If
    
    
    '-- seleccionar todas las posibles tarifas
    Sql = "SELECT * FROM starif WHERE NOT ISNULL(margecom) "
    If PVP < 2 Then
        'Sera de uno de los tipos
        Sql = Sql & " AND opcionINC = " & CStr(PVP)
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada tarifa actualizar la linea en la tabla de lista de precios
    'por cada codartic,codtarif
    NoOK = False
    While Not RS.EOF
        If BloquearTarifas(Text1(0).Text) Then
            Set cTar = New CTarifaArt
            If cTar.LeerDatos(Text1(0).Text, RS!codlista) Then
                
                If cTar.TarifaSobre = 0 Then
                    'TARIFAS SOBRE PVP
                    newPrecio = Round2((CCur(Text1(17).Text) * cTar.MargenComercial) / 100, 4)
                    newPrecio = CCur(Text1(17).Text) + newPrecio
                    
                Else
                    'TARIFAS SOBRE UPC
                    newPrecio = Round2((CCur(Text1(15).Text) * cTar.MargenComercial) / 100, 4)
                    newPrecio = CCur(Text1(15).Text) + newPrecio
                End If
                
                If cTar.ActualizarPrecios(Format(Now, "dd/mm/yyyy"), newPrecio, 0, menErr, False) = False Then NoOK = True
            Else
                'si no existe el articulo con esa tarifa la damos de alta
                cTar.CodigoArticulo = Text1(0).Text
                cTar.CodigoTarifa = RS!codlista
                'Si la tarifa es sobre PVP, mando el PVP
                'Si es sobre el UPC mando el UPC
                If DBLet(RS!opcionINC, "N") = 0 Then
                    'PVP
                    cTar.PrecioActual = CCur(Text1(17).Text) 'precio venta al publico (pvp)
                Else
                    cTar.PrecioActual = CCur(Text1(15).Text) 'precio venta al publico (pUC)
                End If
                
                If Not cTar.InsertarPrecios Then NoOK = True
            End If
            Set cTar = Nothing
        Else
            NoOK = True
'            MsgBox "NO se han actualizado correctamente todas las tarifa del art�culo.", vbExclamation, "Actualizar precios"
            'Exit Function
        End If
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
    If NoOK Then
        ActualizarPreciosPorTarifaDOS = False
        cadErr = "NO se han actualizado correctamente todas las tarifa del art�culo."
        cadErr = cadErr & vbCrLf & menErr
        MsgBox cadErr, vbExclamation, "Actualizar Precios"
    Else
        ActualizarPreciosPorTarifaDOS = True
    End If
        
    Exit Function
    
ErrActPrecio:
    ActualizarPreciosPorTarifaDOS = False
    cadErr = "Actualizar precios por tarifa: " & Err.Description
    MsgBox cadErr, vbExclamation
End Function
   
    
Private Function InsertarModificarLinea() As Boolean
Dim i As Integer
Dim Sql As String

    On Error GoTo EInsertarModificarLinea

    InsertarModificarLinea = False
    Sql = ""
    Select Case ModificaLineas
    Case 1  'INSERTAR
        If DatosOkLinea Then 'INSERTAR
            Sql = "INSERT INTO salmac VALUES ("
            Sql = Sql & DBSet(Text1(0).Text, "T") & ", "
            Sql = Sql & Text3(0).Text & ", "
            
            'Campos Stocks (Son Decimales)
            Sql = Sql & DBSet(Text3(1).Text, "N", "N") & ", "
            For i = 2 To 5
                Sql = Sql & DBSet(Text3(i).Text, "N", "S") & ", "
            Next i
        
            'Campo Fecha
            Sql = Sql & DBSet(Text3(6).Text, "F", "S") & ", "
        
            If Trim(Text3(7).Text) <> "" Then     'Campo Hora
              Sql = Sql & Format(Text3(7).Text, "hh:mm:ss") & ", "
            Else
              Sql = Sql & "NULL, "
            End If
        
            Sql = Sql & chkInventario.Value & ")"
        End If
        
    Case 2  'MODIFICAR
        If DatosOkLinea Then
            Sql = "UPDATE salmac Set canstock = " & DBSet(Text3(1).Text, "N") & ", "
            Sql = Sql & " stockmin = " & DBSet(Text3(2).Text, "N", "S") & ", "
            Sql = Sql & " puntoped = " & DBSet(Text3(3).Text, "N", "S") & ", "
            Sql = Sql & " stockmax = " & DBSet(Text3(4).Text, "N", "S") & ", "
            Sql = Sql & " stockinv = " & DBSet(Text3(5).Text, "N", "S")
            If Trim(Text3(6).Text) <> "" Then _
            Sql = Sql & ", fechainv = " & DBSet(Text3(6).Text, "F", "S")
            If Trim(Text3(7).Text) <> "" Then
                Sql = Sql & ", horainve = '" & Format(Text3(7).Text, "hh:mm:ss") & "'"
            Else
                Sql = Sql & ", horainve = " & ValorNulo
            End If
            Sql = Sql & ", statusin = " & (chkInventario.Value)
            Sql = Sql & " WHERE codartic = " & DBSet(Text1(0).Text, "T") & " AND "
            Sql = Sql & " codalmac =" & Val(Text3(0).Text)
            
        End If
    End Select
        
    If Sql <> "" Then
        conn.Execute Sql
        InsertarModificarLinea = True
    Else
        PonerFoco Text3(1)
    End If
    Exit Function

EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Stocks Almacenes", Err.Description
End Function
    
    
Private Function InsertarArticulo() As Boolean
Dim b As Boolean
Dim menErr As String

    On Error GoTo ErrInsArt
    conn.BeginTrans
    
    b = InsertarDesdeForm(Me)
    If Not b Then menErr = "Insertando en tabla articulos"
    'insertar una linea en salmac para cada uno de los almacenes
    If b Then b = InsetarArticulosPorAlmacen(menErr)
    
    'insertar una linea de lista de precios para cada tarifa
    If b Then b = InsertarPreciosPorTarifa2(menErr)
                
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
        MsgBox menErr, vbExclamation
    End If
    InsertarArticulo = b
    Exit Function
                
ErrInsArt:
    conn.RollbackTrans
    InsertarArticulo = False
    MuestraError Err.Number, "Insertar art�culo.", Err.Description
End Function
    
    
    
    

Public Function InsertarModificarConjunto() As Boolean
Dim Sql As String
On Error GoTo EInsertarModificarLinea

    Sql = ""
    InsertarModificarConjunto = False
    
    If DatosOkConjunto Then
        Select Case ModificaLineas
        Case 1 'Insertar
                Sql = "INSERT INTO sarti1 VALUES ("
                Sql = Sql & DBSet(Text1(0).Text, "T") & ", "
                Sql = Sql & cmdAceptar.Tag & ", "
                Sql = Sql & DBSet(txtAux(0).Text, "T") & ", "
                Sql = Sql & DBSet(txtAux(1).Text, "N") & ") "
        Case 2 'Modificar
                Sql = "UPDATE sarti1 Set codarti1 = " & DBSet(txtAux(0).Text, "T")
                Sql = Sql & ", cantidad = " & DBSet(txtAux(1).Text, "N")
                Sql = Sql & " WHERE codartic =" & DBSet(Text1(0).Text, "T") & " AND "
                Sql = Sql & " numlinea =" & cmdAceptar.Tag
        End Select
    End If
    
    If Sql <> "" Then
        conn.Execute Sql
        InsertarModificarConjunto = True
    End If
    Exit Function
    
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Conjuntos", Err.Description
End Function


Public Function InsertarModificarInstalacion() As Boolean
Dim Sql As String
Dim Valor As String

On Error GoTo EInsertarModificarInstalacion
    InsertarModificarInstalacion = False
    Valor = Trim(txtAux(2).Text)
    If Valor = "" Then Valor = " "
    
    If ModificaLineas = 1 Then 'INSERTAR
        Sql = "INSERT INTO sarti2 VALUES ("
        Sql = Sql & DBSet(Text1(0).Text, "T") & ", "
        Sql = Sql & cmdAceptar.Tag & ", "
        Sql = Sql & DBSet(Valor, "T") & ") "
    ElseIf ModificaLineas = 2 Then 'MODIFICAR
        Sql = "UPDATE sarti2 Set licontro = " & DBSet(Valor, "T")
        Sql = Sql & " WHERE codartic =" & DBSet(Text1(0).Text, "T") & " AND "
        Sql = Sql & " numlinea =" & cmdAceptar.Tag
    End If
    
    conn.Execute Sql
    InsertarModificarInstalacion = True
    Exit Function

EInsertarModificarInstalacion:
    MuestraError Err.Number, "Insertar/Modificar Instalaci�n", Err.Description
End Function


'---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
Public Function InsertarModificarCodigosEAN() As Boolean
Dim Sql As String
Dim Valor As String

    On Error GoTo ErrInsModEAN
    InsertarModificarCodigosEAN = False
    
    Valor = Trim(txtAux(7).Text)
    If Valor = "" Then Valor = " "
    
    If ModificaLineas = 1 Then 'INSERTAR
        Sql = "INSERT INTO sarti3 VALUES ("
        Sql = Sql & DBSet(Text1(0).Text, "T") & ", "
        Sql = Sql & cmdAceptar.Tag & ", "
        Sql = Sql & DBSet(Valor, "T") & ") "
    ElseIf ModificaLineas = 2 Then 'MODIFICAR
        Sql = "UPDATE sarti3 Set codigoea = " & DBSet(Valor, "T")
        Sql = Sql & " WHERE codartic =" & DBSet(Text1(0).Text, "T") & " AND "
        Sql = Sql & " numlinea =" & cmdAceptar.Tag
    End If
    
    conn.Execute Sql
    InsertarModificarCodigosEAN = True
    Exit Function

ErrInsModEAN:
    MuestraError Err.Number, "Insertar/Modificar codigos EAN", Err.Description
End Function
'----




Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim tots As String
Dim Sql As String

    On Error GoTo ECargaGrid
    
      
    If vDataGrid.Name = "DataGrid1" Then
        Sql = MontaSQLCarga(enlaza, 2)
        CargaGridGnral DataGrid1, Me.Data2, Sql, PrimeraVez
        If vParamAplic.ComponentePorcentaje Then
            tots = "cant o %"
        Else
            tots = "cantidad"
        End If
        tots = "N||||0|;N||||0|;S|txtAux(0)|T|Cod. Art�culo|1750|;S|cmdAux|B||0|;S|txtAux2|T|Desc. Art�culo|4000|;S|txtAux(1)|T|" & tots & "|820|" & FormatoCantidad & "|;"
        tots = tots & "S|txtAux(3)|T|PVP|950|;S|txtAux(4)|T|UPC|950|;S|txtAux(5)|T|Pre.Tarif|950|;"
        'Materia prima
        tots = tots & "S|txtAux(6)|T|M.Pr.|600|;"
        arregla tots, DataGrid1, Me
        DataGrid1.Columns(4).Alignment = dbgCenter
        DataGrid1.ScrollBars = dbgAutomatic
        
    ElseIf vDataGrid.Name = "DataGrid2" Then
        Sql = MontaSQLCarga(enlaza, 3)
        CargaGridGnral DataGrid2, Me.Data3, Sql, PrimeraVez
        tots = "N||||0|;N||||0|;S|txtAux(2)|T|Control Instalaciones|7100|;"
        arregla tots, DataGrid2, Me
        DataGrid2.ScrollBars = dbgAutomatic
        
    ElseIf vDataGrid.Name = "DataGrid3" Then
        Sql = MontaSQLCarga(enlaza, 4)
        CargaGridGnral DataGrid3, Me.Data4, Sql, PrimeraVez
        tots = "S|Text3(0)|T|Cod.Alm|1200|;S|cmdAlma|B||0|;S|Text2(8)|T|Nombre Almacen|2400|;S|Text3(1)|T|Stock|1200|;"
        'Los campos que no se ven que van FUERA DEL GRID
        tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
        arregla tots, DataGrid3, Me
        DataGrid3.ScrollBars = dbgAutomatic
 
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
    ElseIf vDataGrid.Name = "DataGrid4" Then 'Lineas cod. EAN
        Sql = MontaSQLCarga(enlaza, 5)
        CargaGridGnral DataGrid4, Me.Data5, Sql, PrimeraVez
        tots = "N||||0|;N||||0|;S|txtAux(7)|T|Cod. EAN|2100|;"
        arregla tots, DataGrid4, Me
        DataGrid4.ScrollBars = dbgAutomatic
    '----
    End If
    
    
    
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el Data
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String

    If Opcion = 2 Then
        'cadena SQL para cargar los CONJUNTOS de la tabla sarti1
        'SQL = "SELECT sarti1.codartic,sarti1.numlinea,sarti1.codarti1,sartic.nomartic,sarti1.cantidad "
        'SQL = SQL & " FROM sarti1 INNER JOIN sartic ON sarti1.codarti1=sartic.codartic "
        
        
        Sql = "SELECT sarti1.codartic, numlinea, sarti1.codarti1,sartic.nomartic,"
        Sql = Sql & " sarti1.Cantidad , sartic.preciove, sartic.precioUC, slista.precioac,if (mateprima=1,""*"","" "") materiaprima"
        Sql = Sql & " FROM   sarti1 INNER JOIN sartic ON"
        Sql = Sql & " sarti1.codarti1 = sartic.codArtic"
        Sql = Sql & " LEFT OUTER JOIN slista ON sarti1.codarti1=slista.codartic AND slista.codlista = " & vParamAplic.CodTarifa
        Sql = Sql & " where sarti1.codartic="
        If enlaza Then
            Sql = Sql & DBSet(Text1(0).Text, "T")
        Else
            Sql = Sql & "'-1@#'"
        End If
        Sql = Sql & " ORDER BY sarti1.numlinea "
        
        
    ElseIf Opcion = 3 Then 'INSTALACIONES
        Sql = "SELECT sarti2.codartic, sarti2.numlinea, sarti2.licontro "
        Sql = Sql & " FROM sarti2"
        If enlaza Then
            Sql = Sql & " WHERE sarti2.codartic=" & DBSet(Text1(0), "T")
        Else
            Sql = Sql & " WHERE sarti2.codartic= '-1'"
        End If
        Sql = Sql & " ORDER BY sarti2.numlinea"
    
    ElseIf Opcion = 4 Then 'STOCK
        
        Sql = "select salmac.codalmac,nomalmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin  "
        Sql = Sql & " from salmac,salmpr where salmac.codalmac=salmpr.codalmac AND "
        If enlaza Then
            Sql = Sql & " codartic=" & DBSet(Text1(0), "T")
        Else
            Sql = Sql & " codartic= '-1'"
        End If
    
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
    ElseIf Opcion = 5 Then
        Sql = "SELECT *"
        Sql = Sql & " FROM sarti3"
        If enlaza Then
            Sql = Sql & " WHERE sarti3.codartic=" & DBSet(Text1(0), "T")
        Else
            Sql = Sql & " WHERE sarti3.codartic= '-1'"
        End If
        Sql = Sql & " ORDER BY sarti3.numlinea"
    '----
    End If
    
    MontaSQLCarga = Sql
End Function


Private Sub LLamaLineas2(alto As Single, xModo As Byte, Opcion As Byte)
Dim b As Boolean

    ModificaLineas = xModo
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN (se a�ade modo 8)
    b = (Modo >= 5 Or Modo <= 8) And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas

    Select Case Opcion
    Case 2 'CONJUNTOS
        DeseleccionaGrid Me.DataGrid1
        
        txtAux(0).Height = DataGrid1.RowHeight
        txtAux(0).visible = b
        txtAux(0).Top = alto
        txtAux(1).Height = DataGrid1.RowHeight
        txtAux(1).visible = b
        txtAux(1).Top = alto
        txtAux2.Height = DataGrid1.RowHeight
        txtAux2.visible = b
        txtAux2.Top = alto
        cmdAux.visible = b
        cmdAux.Top = alto
        cmdAux.Height = DataGrid1.RowHeight
         
    Case 3 'INSTALACIONES
        DeseleccionaGrid Me.DataGrid2
        txtAux(2).Height = DataGrid2.RowHeight
        txtAux(2).visible = True
        txtAux(2).Top = alto
        
        
    Case 4
        'STOCK
        DeseleccionaGrid Me.DataGrid3
        Text3(0).Height = DataGrid3.RowHeight
        Text3(0).visible = b
        Text3(0).Top = alto
        Text3(1).Height = DataGrid3.RowHeight
        Text3(1).visible = b
        Text3(1).Top = alto
        
        If b Then
            If ModificaLineas = 1 Then
                cmdAlma.visible = b And ModificaLineas = 1
                cmdAlma.Top = alto
                cmdAlma.Height = DataGrid1.RowHeight
            Else
                cmdAlma.visible = False
                Text3(0).Width = DataGrid3.Columns(0).Width
            End If
        Else
            cmdAlma.visible = False
        End If
        
    '---- [23/09/2009] LAURA : A�adir lineas de Cod. EAN
    Case 5 'Lineas Cod. EAN
        DeseleccionaGrid Me.DataGrid4
        txtAux(7).Height = DataGrid4.RowHeight
        txtAux(7).visible = True
        txtAux(7).Top = alto
    '----
    End Select
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
   Desplazamiento (Button.Index)
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        If (Index = 2) Or Index = 1 Then
            KeyAscii = 0
            PonerFocoBtn Me.cmdAceptar
            Exit Sub
        End If
    End If
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 0 'cod Articulo de conjunto
            TagText3 = "mateprima"
            txtAux2.Text = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", txtAux(0).Text, "T", TagText3)
            MateriaPrima = TagText3 = "1"
            TagText3 = ""
            
        Case 1
            'Si es materiaprima then
            If txtAux(1).Text <> "" Then
                If vParamAplic.ComponentePorcentaje And MateriaPrima Then
                    'Formato decimal
                    If Not PonerFormatoDecimal(txtAux(Index), 4) Then txtAux(1).Text = ""
                Else
                    If Not PonerFormatoDecimal(txtAux(Index), 2) Then txtAux(1).Text = ""
                End If
                If txtAux(1).Text = "" Then PonerFoco txtAux(1)
                
            End If
    End Select
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
    On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "&Cabecera"
    
    If b Then
        Me.lblIndicador.Caption = "Lineas Detalle"
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.lblIndicador.Caption = ""
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function Eliminar() As Boolean
    Set LOG = New cLOG
    
    

    conn.BeginTrans
    
    If EliminarArticulo(Data1.Recordset!codArtic, lblIndicador) Then
        LOG.Insertar 7, vUsu, Data1.Recordset!codArtic & " " & Data1.Recordset!NomArtic
        conn.CommitTrans
        Eliminar = True
    Else
        conn.RollbackTrans
        Eliminar = False
        
    End If
    Set LOG = Nothing
    lblIndicador.Caption = ""
End Function


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "codartic=" & DBSet(Text1(0).Text, "T")
    If SituarData(Data1, Cad, Indicador) Then
        PonerModo 2
        PonerCampos
        
        lblIndicador.Caption = Indicador
    ElseIf Not Data1.Recordset.EOF Then
'        Data1.Recordset.MoveFirst
        PonerCampos
        PonerModo 2
    ElseIf Modo = 3 Then
        'Acabamos de insertar un registro y lo seleccionamos en el recordset
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE codartic =" & DBSet(Text1(0).Text, "T")
        Data1.RecordSource = CadenaConsulta
        If SituarData(Data1, Cad, Indicador) Then
            PonerModo 2
            PonerCampos
            lblIndicador.Caption = Indicador
        End If
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub


Private Sub BotonImprimir()
    AbrirListado (6) '6: Informe de Articulos
End Sub


Private Sub AccionesSobreTagText3_(Guardar As Boolean, Cargando As Boolean)
Dim i As Integer

  
    If Guardar Then
        If Cargando Then TagText3 = ""
        For i = 0 To Text3.Count - 1
            If Cargando Then TagText3 = TagText3 & Replace(Text3(i).Tag, "|", ";") & "|"
            Text3(i).Tag = ""
        Next i
        
        'A�ADIMOS EL CHECK chkInventario.
        If Cargando Then TagText3 = TagText3 & Replace(chkInventario.Tag, "|", ";") & "|"
        chkInventario.Tag = ""
    Else
        For i = 0 To Text3.Count - 1
            Text3(i).Tag = Replace(RecuperaValor(TagText3, i + 1), ";", "|")
        Next i
        chkInventario.Tag = Replace(RecuperaValor(TagText3, i + 1), ";", "|")
    End If
End Sub


Private Sub PonerDatosForaGrid(ForzarLimpiar As Boolean)
Dim i As Integer
Dim Limp As Boolean

    Limp = True
    If Not ForzarLimpiar Then
        If Not (Data4.Recordset Is Nothing) Then
            If Not Data4.Recordset.EOF Then Limp = False
        End If
    End If
    
    
    If Limp Then

        'Limpiamos
        For i = 0 To Text3.Count - 1
            Text3(i).Text = ""
        Next i
'        Text2(6).Text = ""
        Text2(8).Text = ""
        chkInventario.Value = 0
        
    Else
        'EL
    End If
End Sub

'DAVID
'Para poner el foco en un objeto y si da error que no se arrastre
Private Sub PonerFocoObjeto(obj As Object)
    On Error Resume Next
    obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
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
        .Buttons(1).Image = 5
        .Buttons(3).Image = 6
        .Buttons(5).Image = 7
        .Buttons(7).Image = 1
        .Buttons(11).Image = 2
    End With
    
    Set lw1.SmallIcons = frmPpal.ImgListPpal
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Tag = "" Then Exit Sub
    Label2(0).Caption = ""
    'Levantamos todos los botones y dejamos pulsado el de ahora
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Tag <> "" Then
            If Toolbar2.Buttons(NumRegElim).Index <> Button.Index Then Toolbar2.Buttons(NumRegElim).Value = tbrUnpressed
        End If
    Next NumRegElim
    CargaColumnas CByte(Button.Tag)
    Me.Toolbar2.Refresh
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLW
End Sub





Private Sub CargaColumnas(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader

    Select Case OpcionList
    Case 0 'TARIFAS
        Label2(0).Caption = "Tarifas"
        Columnas = "Tarifa|Descripcion |Tipo|Importe|"
        Ancho = "800|2900|850|1500|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|2|"
        'Formatos
        Formato = "|||" & FormatoPrecio & "|"
        Ncol = 4
    
    Case 1 'PRECIOS ESPECIALES
        Label2(0).Caption = "Precios especiales"
        Columnas = "Cod. cli.|Nombre |Precio|"
        Ancho = "1200|3500|1300|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|1|"
        'Formatos
        Formato = "000||" & FormatoImporte & "|"
        Ncol = 3
        
    Case 2
        Label2(0).Caption = "Promociones"
        Columnas = "Tarifa|Descripcion|F. inicio|F. Fin| Precio|"
        Ancho = "900|2300|1100|1100|1150|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|"
        'Formatos
        Formato = "000||dd/mm/yyyy|dd/mm/yyyy|" & FormatoPrecio & "|"
        Ncol = 5
        
    Case 3 'PEDIDOS
        Label2(0).Caption = "PEDIDOS"
        Columnas = "N�Ped|Fecha|Cod.|Nombre|Candtidad|"
        Ancho = "1250|1100|800|2300|1000|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|"
        'Formatos
        Formato = "|dd/mm/yyyy|||" & FormatoImporte & "|"
        Ncol = 5
        
    Case 4
        'MOVIMIENTOS
        Label2(0).Caption = "MOVIMIENTOS ALMACEN"
        Columnas = "Alm|Fecha|Tipo|Entrada|Documento|Cantidad|C/P/T|"
        Ancho = "600|1100|900|900|1000|1000|900|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|1|1|"
        'Formatos
        Formato = "|dd/mm/yyyy||||" & FormatoCantidad & "||"
        Ncol = 7
    End Select
    
    Me.FrameDisponible.visible = OpcionList = 3

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


Private Sub CargaDatosLW()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & Label2(0).Caption
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLW2()
Dim Cad As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer



    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Value = tbrPressed Then
            ElIcono = Toolbar2.Buttons(NumRegElim).Image
            Exit For
        End If
    Next
    
    

    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 0
        'OFERTAS
        Cad = "select l.codlista,nomlista,if(opcionINC=0,""PVP"",""UPC""),precioac from slista l,starif c where c.codlista=l.codlista"

        BuscaChekc = ""
    Case 1
        'Precios especiales
        Cad = "select l.codclien,nomclien,precioac from sprees l,sclien s where s.codclien=l.codclien"
        BuscaChekc = ""

        
    Case 2
        'Promociones
        Cad = "select l.codlista,nomlista,fechaini,fechafin,precioac from spromo l, starif s where l.codlista=s.codlista"
        BuscaChekc = ""
   
    Case 3
        '*****************************
        'Es una funcion especial
        CargaDatosPedidos
        Exit Sub
        
    Case 4
        'Cargamos movimientos almacen
        Cad = "select codalmac,fechamov,detamovi,if(tipomovi=1,""*"","" ""),document,cantidad,codigope from smoval l WHERE 1=1 "
        BuscaChekc = "ORDER BY fechamov desc,horamovi desc"
    End Select
    
    
    'La fecha
    
    'EL where del codclien
    Cad = Cad & " and l.codartic='" & DevNombreSQL(Data1.Recordset!codArtic) & "'"
    
    
    

    
    'El ORDER BY
    If BuscaChekc <> "" Then Cad = Cad & " ORDER BY fechamov desc,horamovi desc"
    BuscaChekc = ""
    
    lw1.ListItems.Clear
    Set RS = New ADODB.Recordset
    
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number, "", Err.Description
    Set RS = Nothing
    
End Sub



Private Sub CargaDatosPedidos()
Dim C As String
Dim Importe As Currency
Dim T As Currency
    
    'Limpiamos
    lw1.ListItems.Clear
    For NumRegElim = 1 To 3
        Text4(NumRegElim).Text = ""
    Next
    
    'Cargamos el primer combo
    Text4(0).Text = txtSumaStock.Text
    T = 0
    If txtSumaStock.Text <> "" Then T = ImporteFormateado(txtSumaStock.Text)
        
        
    
    'Cargamos primero los de cliente
'    C = "select scaped.numpedcl,fecpedcl,codclien,nomclien,sum(cantidad) as cuantos"
'    C = C & " from scaped,sliped where scaped.numpedcl=sliped.numpedcl  and codartic='"
'    C = C & DevNombreSQL(Data1.Recordset!codArtic) & "' GROUP BY 1"
'    Importe = CargaListPedidos(6, C)
'    T = T - Importe
'    Text4(1).Text = Format(Importe, FormatoImporte)
    
    'Cargamos los comprados
    C = "select scappr.numpedpr,fecpedpr,codprove,nomprove,sum(cantidad) as cuantos"
    C = C & " from scappr,slippr where scappr.numpedpr=slippr.numpedpr  and codartic='"
    C = C & DevNombreSQL(Data1.Recordset!codArtic) & "' group by 1"
    Importe = CargaListPedidos(9, C)
    T = T + Importe
    Text4(2).Text = Format(Importe, FormatoImporte)
    'Disponible
    Text4(3).Text = Format(T, FormatoImporte)
End Sub


Private Function CargaListPedidos(ByRef ElIcono As Integer, Cad As String) As Currency
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim Cantidad As Currency

    Set RS = New ADODB.Recordset
    
    Cantidad = 0
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        Cantidad = Cantidad + DBLet(RS!Cuantos, "N")
        It.SmallIcon = ElIcono
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    CargaListPedidos = Cantidad
End Function




Private Sub ponerDatosConjuntos()
Dim Im1 As Currency
Dim Im2 As Currency
Dim Aux As Currency

    On Error GoTo EponerDatosConjuntos
    'Signo los valores del articulo del UPC y PVP
    txtConjunto(0).Text = Text1(15).Text
    txtConjunto(3).Text = Text1(17).Text
    
    If Data2.Recordset.RecordCount > 0 Then Me.Data2.Recordset.MoveFirst
        
    'Recorrer el RS buscando los importes reales
    While Not Data2.Recordset.EOF
    '
        'COSTE
        Aux = DBLet(Data2.Recordset!Cantidad, "N")
        Aux = Aux * DBLet(Data2.Recordset!precioUC, "N")
        Im1 = Im1 + Aux
        
        'PVP
        Aux = DBLet(Data2.Recordset!Cantidad, "N")
        Aux = Aux * Data2.Recordset!preciove
        Im2 = Im2 + Aux
            
        
        
        Data2.Recordset.MoveNext
    Wend
    If Data2.Recordset.RecordCount > 0 Then Me.Data2.Recordset.MoveFirst
    txtConjunto(1).Text = Format(Im1, FormatoPrecio)
    txtConjunto(4).Text = Format(Im2, FormatoPrecio)
    
    'Difernecias
    Im1 = ImporteFormateado(txtConjunto(0).Text) - Im1
    Im2 = ImporteFormateado(txtConjunto(3).Text) - Im2
    txtConjunto(2).Text = Format(Im1, FormatoPrecio)
    txtConjunto(5).Text = Format(Im2, FormatoPrecio)
    
    Exit Sub
EponerDatosConjuntos:
    MuestraError Err.Number, Err.Description
End Sub



Private Function ComprobarPorcentajesCorrectos() As Boolean
    On Error GoTo EComprobarPorcentajesCorrectos
    ComprobarPorcentajesCorrectos = True
    Set miRsAux = New ADODB.Recordset
    BuscaChekc = "SELECT  sum(sarti1.Cantidad) FROM   sarti1 INNER JOIN sartic ON sarti1.codarti1 = sartic.codArtic"
    BuscaChekc = BuscaChekc & " where mateprima=1 and sarti1.codartic=" & DBSet(Text1(0), "T")
    miRsAux.Open BuscaChekc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If Val(miRsAux.Fields(0)) > 0 And miRsAux.Fields(0) < 100 Then
                MsgBox "La suma de porcentajes de los componenetes no da 100(" & miRsAux.Fields(0) & ")", vbExclamation
                ComprobarPorcentajesCorrectos = False
            End If
        End If
    End If
    miRsAux.Close
EComprobarPorcentajesCorrectos:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "ComprobarPorcentajesCorrectos", Err.Description
        ComprobarPorcentajesCorrectos = False
    End If
    BuscaChekc = ""
    Set miRsAux = Nothing
End Function






Private Sub DataGrid1EnSMOVAL()
'Abrir el formulario del Mantenimiento del que viene el Movimiento
'Se busca en hist�rico o en Form
Dim Sql As String
    
    Select Case lw1.SelectedItem.SubItems(2)
            
        Case "REG" 'Movimientos de Almacen
                    'Movimientos de Almacen
            With frmAlmMovimientos
                .EsHistorico = True
                .hcoCodMovim = lw1.SelectedItem.SubItems(4)
                .hcoFechaMovim = lw1.SelectedItem.SubItems(1)
                .Show vbModal
            End With

        Case "ALV", "ART", "ARC", "ALM", "ALZ", "ALR", "ALS"
                                'ALV:Albaran de Venta (a clientes)
                                'ART: Albaran rectificativo
                                'ARC: Albaran rectificativo de cuotas
                                'ALM: ALbaran Mostrador
                                'ALZ: Albaranes "B"
            'comprobar si el Albaran esta facturado o no
            'si no esta facturado abrir el formulario de Entrada de Albaranes: frmFacEntAlbaranes
            'si esta ya facturado abrir el hist�rico de facturas: frmFacHcoFacturas
            
            'consultamos si existe el albaran en la tabla de albaranes: scaalb
            Sql = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", lw1.SelectedItem.SubItems(2), "T", , "numalbar", lw1.SelectedItem.SubItems(4), "N")
            If Sql <> "" Then 'existe el Albaran
                 With frmFacEntAlbaranes
                    If EsNumerico(lw1.SelectedItem.SubItems(4)) Then
                        .hcoCodMovim = Format(lw1.SelectedItem.SubItems(4), "0000000")
                    Else
                        .hcoCodMovim = lw1.SelectedItem.SubItems(4)
                    End If
                    .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                    .RecuperarFactu = False
                    .Show vbModal
                End With
            Else 'No existe en albaran, abrir Historico Factura
                With frmFacHcoFacturas2
                    .DesdeFichaCliente = False
                    If EsNumerico(lw1.SelectedItem.SubItems(4)) Then
                        .hcoCodMovim = Format(lw1.SelectedItem.SubItems(4), "0000000")
                    Else
                        .hcoCodMovim = lw1.SelectedItem.SubItems(4)
                    End If
                    .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                    .hcoFechaMov = lw1.SelectedItem.SubItems(1)
                    
                    .Show vbModal
                End With
            End If
            

'             With frmFacEntAlbaranes
'                If EsNumerico(lw1.SelectedItem.SubItems(4)) Then
'                    .hcoCodMovim = Format(lw1.SelectedItem.SubItems(4), "0000000")
'                Else
'                    .hcoCodMovim = lw1.SelectedItem.SubItems(4)
'                End If
'                .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
'                .RecuperarFactu = False
'                .Show vbModal
'            End With
            
        Case "ALC" 'Albaran de Compra (a Proveedores)
            'comprobar si el Albaran esta facturado o no
            'si no esta facturado abrir el formulario de Entrada de Albaranes: frmComEntAlbaranes
            'si esta ya facturado abrir el hist�rico de facturas: frmComHcoFacturas
            
            'consultamos si existe el albaran en la tabla de albaranes: scaalp
            Sql = DevuelveDesdeBDNew(conAri, "scaalp", "numalbar", "codprove", lw1.SelectedItem.SubItems(6), "N", , "numalbar", lw1.SelectedItem.SubItems(4), "T", "fechaalb", lw1.SelectedItem.SubItems(1), "F")
            If Sql <> "" Then 'existe el Albaran
                With frmComEntAlbaranes
                    .hcoCodMovim = Trim(lw1.SelectedItem.SubItems(4))
                    .hcoFechaMovim = lw1.SelectedItem.SubItems(1)
                    .hcoCodProve = lw1.SelectedItem.SubItems(6) 'aqui es el proveedor
                    .Show vbModal
                End With
            Else        'No existe en albaran, abrir Historico Factura
                With frmComHcoFacturas
                    .hcoCodMovim = Trim(lw1.SelectedItem.SubItems(4))
                    .hcoFechaMovim = lw1.SelectedItem.SubItems(1)
                    .hcoCodProve = lw1.SelectedItem.SubItems(6) 'aqui es el proveedor
                    .Show vbModal
                End With
            End If
            
            
        '**********************************
        'Laura: modificado 11/09/06
'        Case "FTI" 'Factura Ticket de venta
        Case "ATI" 'Albaran Ticket de venta
        '**********************************
            'Abrir el historico de facturas
             With frmFacHcoFacturas2
                .DesdeFichaCliente = False
                If EsNumerico(lw1.SelectedItem.SubItems(4)) Then
                    .hcoCodMovim = Format(lw1.SelectedItem.SubItems(4), "0000000")
                Else
                    .hcoCodMovim = lw1.SelectedItem.SubItems(4)
                End If
                .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                .hcoFechaMov = lw1.SelectedItem.SubItems(1)
                .Show vbModal
            End With
    Case "DFI"
        MsgBox "Diferencias de inventario.", vbInformation
    End Select
End Sub

