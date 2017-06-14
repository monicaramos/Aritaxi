VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdmTrabajadores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trabajadores"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   13035
   Icon            =   "frmAdmTrabajadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   13035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   210
      TabIndex        =   88
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   89
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
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3870
      TabIndex        =   86
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   87
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
      Left            =   9960
      TabIndex        =   85
      Top             =   240
      Width           =   1605
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
      Left            =   11640
      TabIndex        =   29
      Top             =   6480
      Width           =   1135
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   210
      TabIndex        =   78
      Top             =   855
      Width           =   12615
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
         Index           =   0
         Left            =   1065
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "Código Trabajador|N|N|0|9999|straba|codtraba|0000|S|"
         Text            =   "Text"
         Top             =   225
         Width           =   870
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
         Left            =   3075
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Nombre Trabajador|T|N|||straba|nomtraba||N|"
         Top             =   225
         Width           =   5805
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
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
         Left            =   255
         TabIndex        =   80
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
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
         Left            =   2160
         TabIndex        =   79
         Top             =   255
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
      Left            =   11640
      TabIndex        =   28
      Top             =   6480
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   6315
      Width           =   2895
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
         Left            =   120
         TabIndex        =   31
         Top             =   180
         Width           =   2715
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
      Left            =   10380
      TabIndex        =   27
      Top             =   6480
      Width           =   1135
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3120
      Top             =   5400
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
      Left            =   4560
      Top             =   5400
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
      Height          =   4605
      Left            =   240
      TabIndex        =   32
      Top             =   1590
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   8123
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
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
      TabPicture(0)   =   "frmAdmTrabajadores.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text1(10)"
      Tab(0).Control(1)=   "Text2(10)"
      Tab(0).Control(2)=   "Text2(24)"
      Tab(0).Control(3)=   "Text1(24)"
      Tab(0).Control(4)=   "Text1(2)"
      Tab(0).Control(5)=   "Text1(7)"
      Tab(0).Control(6)=   "Text1(23)"
      Tab(0).Control(7)=   "Text1(8)"
      Tab(0).Control(8)=   "frameDptoPersonal"
      Tab(0).Control(9)=   "frameBancos"
      Tab(0).Control(10)=   "Text1(9)"
      Tab(0).Control(11)=   "Text1(6)"
      Tab(0).Control(12)=   "Text1(5)"
      Tab(0).Control(13)=   "Text1(4)"
      Tab(0).Control(14)=   "Text1(3)"
      Tab(0).Control(15)=   "Label1(2)"
      Tab(0).Control(16)=   "imgBuscar(2)"
      Tab(0).Control(17)=   "ImgMail(0)"
      Tab(0).Control(18)=   "imgBuscar(1)"
      Tab(0).Control(19)=   "Label1(26)"
      Tab(0).Control(20)=   "Label1(25)"
      Tab(0).Control(21)=   "Label1(24)"
      Tab(0).Control(22)=   "Label1(12)"
      Tab(0).Control(23)=   "imgBuscar(0)"
      Tab(0).Control(24)=   "Label1(37)"
      Tab(0).Control(25)=   "Label1(36)"
      Tab(0).Control(26)=   "Label1(15)"
      Tab(0).Control(27)=   "Label1(34)"
      Tab(0).Control(28)=   "Label1(14)"
      Tab(0).Control(29)=   "Label1(13)"
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Estudios/Formación"
      TabPicture(1)   =   "frmAdmTrabajadores.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtAux1(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtAux1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "FrameToolAux(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Habilidades"
      TabPicture(2)   =   "frmAdmTrabajadores.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtAux2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "FrameToolAux(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Experiencia Laboral"
      TabPicture(3)   =   "frmAdmTrabajadores.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DataGrid3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "TxtAux3(0)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "TxtAux3(1)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "FrameToolAux(2)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Formación Realizada"
      TabPicture(4)   =   "frmAdmTrabajadores.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DataGrid4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "TxtAux4(0)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "TxtAux4(1)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "TxtAux4(2)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "TxtAux4(3)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "TxtAux4(4)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "FrameToolAux(3)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "Formación Empresa"
      TabPicture(5)   =   "frmAdmTrabajadores.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FrameToolAux(0)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "TxtAux5(2)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "TxtAux5(1)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "TxtAux5(0)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "DataGrid5"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).ControlCount=   5
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   4
         Left            =   480
         TabIndex        =   98
         Top             =   330
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   4
            Left            =   210
            TabIndex        =   99
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
         Index           =   3
         Left            =   -74520
         TabIndex        =   96
         Top             =   330
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   3
            Left            =   210
            TabIndex        =   97
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
         Left            =   -74520
         TabIndex        =   94
         Top             =   330
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   2
            Left            =   210
            TabIndex        =   95
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
         Index           =   1
         Left            =   -74520
         TabIndex        =   92
         Top             =   330
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   1
            Left            =   210
            TabIndex        =   93
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
         Left            =   -74520
         TabIndex        =   90
         Top             =   330
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   210
            TabIndex        =   91
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
         Left            =   -73695
         MaxLength       =   4
         TabIndex        =   12
         Tag             =   "Centro de coste|T|S|||straba|codccost||N|"
         Top             =   3750
         Width           =   630
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
         Index           =   10
         Left            =   -73035
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   81
         Text            =   "Text2"
         Top             =   3750
         Width           =   4275
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
         Index           =   24
         Left            =   -73035
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   77
         Text            =   "Text2"
         Top             =   3351
         Width           =   4275
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
         Index           =   24
         Left            =   -73695
         MaxLength       =   3
         TabIndex        =   11
         Tag             =   "Almacen por Defecto|N|N|0|999|straba|codalmac|000|N|"
         Text            =   "Text aldu dkdo sñsñs"
         Top             =   3351
         Width           =   630
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
         Index           =   2
         Left            =   -73695
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Domicilio|T|N|||straba|domtraba||N|"
         Text            =   "Text1"
         Top             =   988
         Width           =   4905
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
         Left            =   -73680
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "Teléfono|T|N|||straba|teltraba||N|"
         Text            =   "Text1"
         Top             =   2152
         Width           =   1830
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
         Left            =   -70515
         MaxLength       =   20
         TabIndex        =   3
         Tag             =   "Login Trabajador|T|S|||straba|login||N|"
         Text            =   "Text aldu dkdo sñsñs"
         Top             =   600
         Width           =   1710
      End
      Begin VB.TextBox TxtAux5 
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
         Height          =   315
         Index           =   2
         Left            =   -68160
         MaxLength       =   15
         TabIndex        =   74
         Tag             =   "Resultado|T|N|||strab5|resforma||N|"
         Text            =   "Resultado"
         Top             =   3720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox TxtAux5 
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
         Height          =   315
         Index           =   1
         Left            =   -72840
         MaxLength       =   50
         TabIndex        =   73
         Tag             =   "Formación|T|N|||strab5|formaci2||N|"
         Text            =   "Formacion"
         Top             =   3720
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.TextBox TxtAux5 
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
         Height          =   315
         Index           =   0
         Left            =   -74400
         MaxLength       =   12
         TabIndex        =   72
         Tag             =   "Fecha Formación|F|N|||strab5|fecforma||N|"
         Text            =   "F. Formac."
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtAux4 
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
         Height          =   315
         Index           =   4
         Left            =   -67200
         MaxLength       =   8
         TabIndex        =   71
         Tag             =   "Evaluación|T|N|||strab4|evaluaci||N|"
         Text            =   "Evaluaci"
         Top             =   3720
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox TxtAux4 
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
         Height          =   315
         Index           =   3
         Left            =   -69000
         MaxLength       =   15
         TabIndex        =   70
         Tag             =   "Centro|T|N|||strab4|codcentr||N|"
         Text            =   "Centro"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox TxtAux4 
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
         Height          =   315
         Index           =   2
         Left            =   -71760
         MaxLength       =   40
         TabIndex        =   69
         Tag             =   "Formacion|T|N|||strab4|formaci1||N|"
         Text            =   "Formacion"
         Top             =   3720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox TxtAux4 
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
         Height          =   315
         Index           =   1
         Left            =   -72960
         MaxLength       =   12
         TabIndex        =   68
         Tag             =   "Fecha Evaluación|F|N|||strab4|fecevalu|dd/mm/yyyy|N|"
         Text            =   "F.Evalua."
         Top             =   3720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtAux4 
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
         Height          =   315
         Index           =   0
         Left            =   -74280
         MaxLength       =   12
         TabIndex        =   67
         Tag             =   "Fecha Formación|F|N|||strab4|fecforma|dd/mm/yyyy|N|"
         Text            =   "F.Formac."
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtAux3 
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
         Height          =   315
         Index           =   1
         Left            =   -72360
         MaxLength       =   70
         TabIndex        =   66
         Tag             =   "Experiencia|T|N|||strab3|experien||N|"
         Text            =   "Experiencia"
         Top             =   3720
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.TextBox TxtAux3 
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
         Height          =   315
         Index           =   0
         Left            =   -74400
         MaxLength       =   15
         TabIndex        =   65
         Tag             =   "Periodo|T|N|||strab3|periodo1||N|"
         Text            =   "Periodo"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtAux2 
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
         Height          =   315
         Left            =   -74160
         MaxLength       =   70
         TabIndex        =   64
         Tag             =   "Habilidad|T|N|||strab2|habilida||N|"
         Text            =   "Habilidad"
         Top             =   3720
         Visible         =   0   'False
         Width           =   6735
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
         Left            =   -73695
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "Cargo en la empresa|T|S|||straba|cartraba||N|"
         Text            =   "Text1"
         Top             =   2540
         Width           =   4935
      End
      Begin VB.TextBox txtAux1 
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
         Height          =   315
         Index           =   1
         Left            =   2880
         MaxLength       =   70
         TabIndex        =   58
         Tag             =   "Formación|T|N|||strab1|formacio||N|"
         Text            =   "Formacion Formacion Formacion Formacion Formacion Formacion Formacion "
         Top             =   3660
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.TextBox txtAux1 
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
         Height          =   315
         Index           =   0
         Left            =   840
         MaxLength       =   15
         TabIndex        =   57
         Tag             =   "Periodo|T|N|||strab1|periodos||N|"
         Text            =   "Periodo"
         Top             =   3660
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame frameDptoPersonal 
         Caption         =   "Datos relacionados con Dpto Personal"
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
         Height          =   1635
         Left            =   -68640
         TabIndex        =   39
         Top             =   540
         Width           =   6105
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
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "Fecha de Baja|F|S|||straba|fechabaj|dd/mm/yyyy|N|"
            Top             =   840
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
            Index           =   12
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Fecha de Alta|F|N|||straba|fechaalt|dd/mm/yyyy|N|"
            Top             =   420
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
            Index           =   11
            Left            =   1740
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Fecha de Nacimiento|F|N|||straba|fechanac|dd/mm/yyyy|N|"
            Top             =   420
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
            Index           =   14
            Left            =   1545
            MaxLength       =   12
            TabIndex        =   14
            Tag             =   "Nº SSocial|T|S|||straba|nrosegur||N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Baja"
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
            Left            =   3150
            TabIndex        =   46
            Top             =   840
            Width           =   1095
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   4275
            ToolTipText     =   "Buscar fecha"
            Top             =   870
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Alta"
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
            Left            =   3150
            TabIndex        =   45
            Top             =   420
            Width           =   1125
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   4275
            ToolTipText     =   "Buscar fecha"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Nacimiento"
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
            TabIndex        =   44
            Top             =   420
            Width           =   1335
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1455
            ToolTipText     =   "Buscar fecha"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Seg.Social"
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
            Left            =   150
            TabIndex        =   40
            Top             =   840
            Width           =   1425
         End
      End
      Begin VB.Frame frameBancos 
         Caption         =   "Datos Bancarios"
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
         Height          =   1635
         Left            =   -68610
         TabIndex        =   41
         Top             =   2400
         Width           =   6075
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
            Index           =   26
            Left            =   960
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "IBAN Gastos1|T|S|||straba|iban1||N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   675
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
            Index           =   25
            Left            =   960
            MaxLength       =   4
            TabIndex        =   17
            Tag             =   "IBAN|T|S|||straba|iban||N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   675
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
            Left            =   1740
            MaxLength       =   4
            TabIndex        =   18
            Tag             =   "Código Banco Nómina|N|S|0|9999|straba|codbanco|0000|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   675
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
            Index           =   16
            Left            =   2505
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "Sucursal Nómina|N|S|0|9999|straba|codsucur|0000|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   675
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
            Index           =   17
            Left            =   3375
            MaxLength       =   2
            TabIndex        =   20
            Tag             =   "Dígito Control Nómina|T|S|||straba|digcontr|00||"
            Text            =   "Text1"
            Top             =   480
            Width           =   435
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
            Left            =   3870
            MaxLength       =   10
            TabIndex        =   21
            Tag             =   "Cuenta Bancaria Nómina|T|S|||straba|cuentaba|0000000000||"
            Text            =   "Text1"
            Top             =   480
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
            Index           =   19
            Left            =   1740
            MaxLength       =   4
            TabIndex        =   23
            Tag             =   "Código Banco Gastos|N|N|0|9999|straba|codbanc1|0000|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   675
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
            Left            =   2505
            MaxLength       =   4
            TabIndex        =   24
            Tag             =   "Sucursal Gastos|N|S|0|9999|straba|codsucu1|0000|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   675
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
            Left            =   3375
            MaxLength       =   2
            TabIndex        =   25
            Tag             =   "Dígito Control Gastos|T|S|||straba|digcont1|00||"
            Text            =   "Text1"
            Top             =   1200
            Width           =   435
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
            Left            =   3870
            MaxLength       =   10
            TabIndex        =   26
            Tag             =   "Cuenta Bancaria Gastos|T|S|||straba|cuentab1|0000000000||"
            Text            =   "Text1"
            Top             =   1200
            Width           =   1995
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
            Index           =   8
            Left            =   960
            TabIndex        =   84
            Top             =   960
            Width           =   495
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
            Index           =   4
            Left            =   960
            TabIndex        =   83
            Top             =   240
            Width           =   495
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
            Index           =   39
            Left            =   1740
            TabIndex        =   55
            Top             =   240
            Width           =   645
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
            Index           =   35
            Left            =   2505
            TabIndex        =   54
            Top             =   240
            Width           =   855
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
            Index           =   33
            Left            =   3450
            TabIndex        =   53
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Bancaria"
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
            Left            =   3870
            TabIndex        =   52
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Nómina"
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
            Index           =   20
            Left            =   180
            TabIndex        =   51
            Top             =   480
            Width           =   735
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
            Index           =   19
            Left            =   1740
            TabIndex        =   50
            Top             =   960
            Width           =   645
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
            Index           =   17
            Left            =   2505
            TabIndex        =   49
            Top             =   960
            Width           =   855
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
            Index           =   7
            Left            =   3450
            TabIndex        =   48
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Bancaria"
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
            Left            =   3870
            TabIndex        =   47
            Top             =   960
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "Gastos"
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
            Index           =   43
            Left            =   180
            TabIndex        =   42
            Top             =   1200
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
         Height          =   360
         Index           =   9
         Left            =   -73695
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "e-mail|T|S|||straba|maitraba||N|"
         Text            =   "Text1"
         Top             =   2930
         Width           =   4935
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
         Left            =   -73695
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "N.I.F.|T|N|||straba|niftraba||N|"
         Text            =   "Text1"
         Top             =   600
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
         Index           =   5
         Left            =   -73695
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Provincia|T|N|||straba|protraba||N|"
         Text            =   "Text1"
         Top             =   1764
         Width           =   4905
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
         Left            =   -71820
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Población|T|N|||straba|pobtraba||N|"
         Text            =   "Text1"
         Top             =   1376
         Width           =   3030
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
         Index           =   3
         Left            =   -73710
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "C.Postal|T|N|||straba|codpobla||N|"
         Text            =   "Text1"
         Top             =   1376
         Width           =   825
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmAdmTrabajadores.frx":00B4
         Height          =   3510
         Left            =   480
         TabIndex        =   56
         Top             =   915
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6191
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BorderStyle     =   0
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
         Bindings        =   "frmAdmTrabajadores.frx":00C9
         Height          =   3510
         Left            =   -74520
         TabIndex        =   60
         Top             =   915
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6191
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmAdmTrabajadores.frx":00DE
         Height          =   3510
         Left            =   -74520
         TabIndex        =   61
         Top             =   915
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6191
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
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
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmAdmTrabajadores.frx":00F3
         Height          =   3510
         Left            =   -74520
         TabIndex        =   62
         Top             =   915
         Width           =   10735
         _ExtentX        =   18944
         _ExtentY        =   6191
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
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
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "frmAdmTrabajadores.frx":0108
         Height          =   3510
         Left            =   -74520
         TabIndex        =   63
         Top             =   915
         Width           =   8855
         _ExtentX        =   15610
         _ExtentY        =   6191
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
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
      Begin VB.Label Label1 
         Caption         =   "CCoste"
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
         Left            =   -74880
         TabIndex        =   82
         Top             =   3750
         Width           =   795
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   -73995
         Tag             =   "-1"
         ToolTipText     =   "Buscar centro coste"
         Top             =   3750
         Width           =   240
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   0
         Left            =   -73995
         Tag             =   "-1"
         ToolTipText     =   "Enviar e-mail"
         Top             =   2925
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   -73995
         Tag             =   "-1"
         ToolTipText     =   "Buscar almacen"
         Top             =   3351
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
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
         Left            =   -74880
         TabIndex        =   76
         Top             =   3345
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Login"
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
         Left            =   -71280
         TabIndex        =   75
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cargo"
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
         Left            =   -74880
         TabIndex        =   59
         Top             =   2535
         Width           =   615
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
         Index           =   12
         Left            =   -74880
         TabIndex        =   43
         Top             =   2145
         Width           =   1275
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   -73965
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "E-mail"
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
         Left            =   -74880
         TabIndex        =   38
         Top             =   2925
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.F."
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
         Left            =   -74880
         TabIndex        =   37
         Top             =   600
         Width           =   675
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
         Index           =   15
         Left            =   -74880
         TabIndex        =   36
         Top             =   1770
         Width           =   1275
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
         Index           =   34
         Left            =   -72840
         TabIndex        =   35
         Top             =   1380
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "C. Postal"
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
         Left            =   -74880
         TabIndex        =   34
         Top             =   1380
         Width           =   1005
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
         Index           =   13
         Left            =   -74880
         TabIndex        =   33
         Top             =   990
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   5880
      Top             =   5400
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
   Begin MSAdodcLib.Adodc Data4 
      Height          =   330
      Left            =   7320
      Top             =   5400
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
   Begin MSAdodcLib.Adodc Data5 
      Height          =   330
      Left            =   3120
      Top             =   5640
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
   Begin MSAdodcLib.Adodc Data6 
      Height          =   330
      Left            =   4560
      Top             =   5640
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
   Begin VB.Menu mnMtoLineas 
      Caption         =   "&Mantenimiento Lineas"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnEstudios 
         Caption         =   "&Estudios/Formación"
         HelpContextID   =   2
      End
      Begin VB.Menu mnHabilidades 
         Caption         =   "&Habilidades"
         HelpContextID   =   2
      End
      Begin VB.Menu mnExperiencia 
         Caption         =   "Experiencia &Laboral"
         HelpContextID   =   2
      End
      Begin VB.Menu mnFormRealizada 
         Caption         =   "&Formación Realizada"
         HelpContextID   =   2
      End
      Begin VB.Menu mnFormEmpresa 
         Caption         =   "F&ormacion Empresa"
         HelpContextID   =   2
      End
   End
End
Attribute VB_Name = "frmAdmTrabajadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmTra As frmBasico2 'Form para busquedas trabajadores
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmCC As frmBasico2 'Form para busquedas trabajadores
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios  'Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1

Private Modo As Byte
'-------------------------------------------------------
'Se distinguen varios MODOS
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim NumTabMto As Byte 'Indica que numero de Tab que esta en modo Mantenimiento

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la tabla principal del formulario
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas del Mantenimiento en que estemos

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1

'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1
Dim btnPrimero As Byte

Dim BuscaChekc As String

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos


'===========================================================================
'       PROCEDIMIENTOS
'============================================================================

Private Sub cmdAceptar_Click()
Dim Cad As String, Indicador As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
              If InsertarDesdeForm(Me) Then PosicionarData
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
                
         Case 5 'INSERTAR MODIFICAR LINEA
            'Actualizar el registro en la tabla de lineas 'sdirec' (Direcciones/Departamentos)
            Cad = "Select * from " & NomTablaLineas & " where codtraba= " & Data1.Recordset!CodTraba
            Cad = Cad & " order by numlinea"
            
            If ModificaLineas = 1 Then 'INSERTAR lineas
                If InsertarLinea Then
                    Select Case Me.SSTab1.Tab
                        Case 1 'Estudios/Formacion - Datos de la tabla strab1
                            CargaGrid DataGrid1, Data2, Cad
                        Case 2 'Habilidades
                            CargaGrid DataGrid2, Data3, Cad
                        Case 3 'Experiencia Laboral
                            CargaGrid DataGrid3, Data4, Cad
                        Case 4 'Formacion Realizada
                            CargaGrid DataGrid4, Data5, Cad
                        Case 5 'Formacion Empresa
                            CargaGrid DataGrid5, Data6, Cad
                    End Select
                    BotonAnyadirLinea
                End If
                
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    PonerBotonCabecera True
                    ModificaLineas = 0
                    Select Case Me.SSTab1.Tab
                        Case 1 'Estudios/Formacion - Datos de la tabla strab1
                            NumRegElim = Data2.Recordset.AbsolutePosition
                            CargaTxtAux1 False, False
                            'CargaGrid DataGrid1, Data2, cad
                            CargaGrid2 DataGrid1, Data2
                            SituarDataPosicion Data2, NumRegElim, Indicador
                        Case 2 'Habilidades
                            NumRegElim = Data3.Recordset.AbsolutePosition
                            CargaTxtAux2 False, False
                            'CargaGrid DataGrid2, Data3, cad
                            CargaGrid2 DataGrid2, Data3
                            SituarDataPosicion Data3, NumRegElim, Indicador
                        Case 3 'Experiencia Laboral
                            NumRegElim = Data4.Recordset.AbsolutePosition
                            CargaTxtAux3 False, False
                            CargaGrid2 DataGrid3, Data4
                            SituarDataPosicion Data4, NumRegElim, Indicador
                        Case 4 'Formacion Realizada
                            NumRegElim = Data5.Recordset.AbsolutePosition
                            CargaTxtAux4 False, False
                            CargaGrid2 DataGrid4, Data5
                            SituarDataPosicion Data5, NumRegElim, Indicador
                        Case 5 'Formacion Empresa
                            NumRegElim = Data6.Recordset.AbsolutePosition
                            CargaTxtAux5 False, False
                            CargaGrid2 DataGrid5, Data6
                            SituarDataPosicion Data6, NumRegElim, Indicador
                    End Select
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
            Select Case Me.SSTab1.Tab
                Case 1 'Estudios/Formacion
                    CargaTxtAux1 False, False
                    DataGrid1.Enabled = True
                    If ModificaLineas = 1 Then 'Insertar
                        DataGrid1.AllowAddNew = False
                        If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                    End If
                Case 2 'Habilidades
                    CargaTxtAux2 False, False
                    DataGrid2.Enabled = True
                    If ModificaLineas = 1 Then 'INSERTAR
                        DataGrid2.AllowAddNew = False
                        If Not Data3.Recordset.EOF Then Data3.Recordset.MoveFirst
                    End If
                Case 3 'Experiencia Laboral
                    CargaTxtAux3 False, False
                    DataGrid3.Enabled = True
                    If ModificaLineas = 1 Then 'INSERTAR
                        DataGrid3.AllowAddNew = False
                        If Not Data4.Recordset.EOF Then Data4.Recordset.MoveFirst
                    End If
                Case 4 'Formacion Realizada
                    CargaTxtAux4 False, False
                    DataGrid4.Enabled = True
                    If ModificaLineas = 1 Then 'INSERTAR
                        DataGrid4.AllowAddNew = False
                        If Not Data5.Recordset.EOF Then Data5.Recordset.MoveFirst
                    End If
                Case 5 'Formacion Empresa
                    CargaTxtAux5 False, False
                    DataGrid5.Enabled = True
                    If ModificaLineas = 1 Then 'INSERTAR
                        DataGrid5.AllowAddNew = False
                        If Not Data6.Recordset.EOF Then Data6.Recordset.MoveFirst
                    End If
            End Select
'            PonerBotonCabecera True
            PonerModo 2
            ModificaLineas = 0
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de trabajadores: straba (Cabecera)

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    Text1(0).Text = SugerirCodigoSiguienteStr("straba", "codtraba")
    Text1(12).Text = Format(Now, "dd/mm/yyyy")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
        
    If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede Añadir. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If
    
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
'--    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    Select Case Me.SSTab1.Tab
        Case 1 'Estudios / Formacion
                'Situamos el grid al final
                AnyadirLinea DataGrid1, Data2
                CargaTxtAux1 True, True
                PonerFoco txtAux1(0)
        Case 2 'Habilidades
                AnyadirLinea DataGrid2, Data3
                CargaTxtAux2 True, True
                PonerFoco txtAux2
        Case 3 'Experiencia Laboral
                AnyadirLinea DataGrid3, Data4
                CargaTxtAux3 True, True
                PonerFoco txtAux3(0)
        Case 4 'Formacion Realizada
                AnyadirLinea DataGrid4, Data5
                CargaTxtAux4 True, True
                PonerFoco txtAux4(0)
        Case 5 'Formacion Empresa
                AnyadirLinea DataGrid5, Data6
                CargaTxtAux5 True, True
                PonerFoco TxtAux5(0)
    End Select
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
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
'            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    LimpiarCampos
    Me.SSTab1.Tab = 0
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
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
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4

    PonerFoco Text1(1)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede Modificar. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If
    
    vWhere = "codtraba=" & Val(Text1(0).Text) & " and numlinea="
    Select Case Me.SSTab1.Tab
        Case 1 'Estudios/Formacion
                If Data2.Recordset.EOF Then Exit Sub
                vWhere = vWhere & Data2.Recordset!numlinea
                If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
                CargaTxtAux1 True, False
                DataGrid1.Enabled = False
                PonerFoco txtAux1(0)
        Case 2 'Habilidades
                If Data3.Recordset.EOF Then Exit Sub
                vWhere = vWhere & Data3.Recordset!numlinea
                If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
                CargaTxtAux2 True, False
                DataGrid2.Enabled = False
                PonerFoco txtAux2
        Case 3 'Experiencia Laboral
                If Data4.Recordset.EOF Then Exit Sub
                vWhere = vWhere & Data4.Recordset!numlinea
                If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
                CargaTxtAux3 True, False
                DataGrid3.Enabled = False
                PonerFoco txtAux3(0)
        Case 4 'Formacion Realizada
                If Data5.Recordset.EOF Then Exit Sub
                vWhere = vWhere & Data5.Recordset!numlinea
                If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
                CargaTxtAux4 True, False
                DataGrid4.Enabled = False
                PonerFoco txtAux4(0)
        Case 5 'Formacion Empresa
                If Data6.Recordset.EOF Then Exit Sub
                vWhere = vWhere & Data6.Recordset!numlinea
                If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
                CargaTxtAux5 True, False
                DataGrid5.Enabled = False
                PonerFoco TxtAux5(0)
    End Select
    
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de trabajadores (straba)
Dim Cad As String
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    Cad = "Cabecera de Trabajadores." & vbCrLf
    Cad = Cad & "------------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Trabajador:"
    Cad = Cad & vbCrLf & "Código:   " & Format(Data1.Recordset.Fields(0), "000000")
    Cad = Cad & vbCrLf & "Descripción:   " & Data1.Recordset.Fields(1)
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        If Not Eliminar Then
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Trabajador", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Trabajador. Tablas: strab1, strab2, strab3, strab4, strab5
Dim Sql As String
Dim numlinea As Integer
On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

     If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede eliminar. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If

    Select Case Me.SSTab1.Tab
        Case 1 'EStudios/Formacion
            If Data2.Recordset.EOF Then Exit Sub
            numlinea = Data2.Recordset!numlinea
        Case 2 'Habilidades
            If Data3.Recordset.EOF Then Exit Sub
            numlinea = Data3.Recordset!numlinea
        Case 3 'Experiencia Laboral
            If Data4.Recordset.EOF Then Exit Sub
            numlinea = Data4.Recordset!numlinea
        Case 4 'Formacion Realizada
            If Data5.Recordset.EOF Then Exit Sub
            numlinea = Data5.Recordset!numlinea
        Case 5 'Formacion Empresa
            If Data6.Recordset.EOF Then Exit Sub
            numlinea = Data6.Recordset!numlinea
    End Select
    
    ModificaLineas = 3 'Eliminar
    Sql = "¿Seguro que desea eliminar la línea de " & TituloLinea & "?"
    Sql = Sql & vbCrLf & "Cod. Traba.: " & Format(Data1.Recordset!CodTraba, "000000")
    Sql = Sql & vbCrLf & "Nombre: " & Data1.Recordset!NomTraba
    Sql = Sql & vbCrLf & "Numlinea: " & numlinea
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Sql = "Delete from " & NomTablaLineas & " where codtraba=" & Data1.Recordset!CodTraba
        Sql = Sql & " and numlinea=" & numlinea
        conn.Execute Sql

        ModificaLineas = 0
        Select Case Me.SSTab1.Tab
            Case 1: 'Estudios/Formacion
'                CancelaADODC (Data2)
                CargaGrid2 DataGrid1, Data2
'                CancelaADODC (Data2)
            Case 2: 'Habilidades
                CargaGrid2 DataGrid2, Data3
            Case 3: 'Experiencia Laboral
                CargaGrid2 DataGrid3, Data4
            Case 4 'Formacion Realizada
                CargaGrid2 DataGrid4, Data5
            Case 5 'Formacion Empresa
                CargaGrid2 DataGrid5, Data6
        End Select
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Trabajador", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera tambien
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
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


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(0)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim I As Integer

    'Icono del form
    Me.Icon = frmPpal.Icon
    
    'Icono de imagen de e-mail
    Me.ImgMail(0).Picture = frmPpal.imgListComun.ListImages(20).Picture

    'ICONITOS DE LA BARRA
'    btnAnyadir = 5
'    btnPrimero = 19
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'        .Buttons(10).Image = 25 'Estudios/Formacion
'        .Buttons(11).Image = 27 'Habilidades
'        .Buttons(12).Image = 37 'Experiencia Laboral
'        .Buttons(13).Image = 28 'Formacion Realizada
'        .Buttons(14).Image = 29 'Formacion Empresa
'
'        .Buttons(16).Image = 16  'Salir
'        .Buttons(17).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
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
'        .Buttons(10).Image = 39 'Actualizar
        .Buttons(8).Image = 16 'Imprimir
'        .Buttons(13).Image = 15 'Salir
'        .Buttons(btnPrimero).Image = 6 'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
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
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            '.ImageList = frmPpal.imgListComun_VELL
            '  ### [Monica] 02/10/2006 acabo de comentarlo
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next I
    
    
    
    
    For I = 0 To Me.imgBuscar.Count - 1
        imgBuscar(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    
    For I = 0 To Me.imgFecha.Count - 1
        imgFecha(I).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next I
    
    Me.SSTab1.Tab = 0
      
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    PrimeraVez = True
         
    '## A mano
    NombreTabla = "straba"
    Ordenacion = " ORDER BY codtraba"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codtraba=-1"
    Data1.Refresh
    
    LimpiarDataGrids
    
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
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    Text1(24).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almac
    Text2(24).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Almac
End Sub




Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
      
    If CadenaSeleccion <> "" Then
        'Centro de coste
        Text1(10).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(10).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim indice As Byte
Dim devuelve As String
    
    indice = 3
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'Poblacion
    Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve)
    'provincia
    Text1(indice + 2).Text = devuelve
End Sub


Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim indice As Byte
    indice = Val(imgFecha(0).Tag) + 11
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String

    CadB = "codtraba = " & RecuperaValor(CadenaSeleccion, 1)
    
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & Ordenacion
    PonerCadenaBusqueda
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'CPostal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            indice = 4
            VieneDeBuscar = True
            
        Case 1 'Almacen por defecto del trabajador
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            indice = 24
            
        Case 2 'Centros de coste de la conta
            Me.imgFecha(0).Tag = 10
               
            Set frmCC = New frmBasico2
            
            AyudaCentroCoste frmCC
            
            Set frmCC = Nothing
        
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
   Me.imgFecha(0).Tag = Index
   indice = Index + 11
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub

Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If Index = 0 Then
        dirMail = Text1(9).Text
    End If
    If LanzaMailGnral(dirMail) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de trabajadores
         BotonEliminarLinea
    Else   'Eliminar Trabajador
         BotonEliminar
    End If
End Sub

Private Sub mnEstudios_Click()
'Abre Mantenimiento de lineas  Estudios/Formacion
    BotonMtoLineas 1, "Estudios/Formacion"
    NomTablaLineas = "strab1"
End Sub

Private Sub mnExperiencia_Click()
'Abre Mantenimiento de lineas Experiencia Laboral
    BotonMtoLineas 3, "Experiencia Laboral"
    NomTablaLineas = "strab3"
End Sub

Private Sub mnFormEmpresa_Click()
'Abre Mantenimiento de lineas Formacion Empresa
    BotonMtoLineas 5, "Formación Empresa"
    NomTablaLineas = "strab5"
End Sub

Private Sub mnFormRealizada_Click()
'Abre Mantenimiento de lineas Formacion Realizada
    BotonMtoLineas 4, "Formación Realizada"
    NomTablaLineas = "strab4"
End Sub

Private Sub mnHabilidades_Click()
'Abre Mantenimiento de lineas Habilidades
    BotonMtoLineas 2, "Habilidades"
    NomTablaLineas = "strab2"
End Sub

Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Trabajador
         Me.SSTab1.Tab = 0
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
    Else 'Añadir Trabajador
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub
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
Dim cta As String
Dim CC As String
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Cod. Trabajador
            If PonerFormatoEntero(Text1(Index)) Then
                'Comprobar si ya existe el cod de trabajador en la tabla
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
            
        Case 3 'CPostal
             If Not VieneDeBuscar Then
                Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                Text1(Index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
            
        Case 6 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
            
            
        Case 10
            ' ---- [19/10/2009] [LAURA] : Añadir funcion generica de ccoste
            Me.Text2(Index).Text = PonerNombreCCoste(Me.Text1(Index))
'            PonceCentroCoste
            
        Case 11, 12, 13 'Fecha Nacimiento, Fecha alta, Fecha baja
            'Si no es modo de Busqueda poner el formato
             If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
             
        Case 24 'Cod almacen
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "salmpr", "nomalmac", "codalmac")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 15, 16, 19, 20 'cod. banco, cod. sucursal
            PonerFormatoEntero Text1(Index)
            
        Case 25, 26 ' codigo de iban
            Text1(Index).Text = UCase(Text1(Index).Text)
            
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 15 Or Index = 16 Or Index = 17 Or Index = 18 Then
        If Text1(15).Text <> "" And Text1(16).Text <> "" And Text1(17).Text <> "" And Text1(18).Text <> "" Then
            
            cta = Format(Text1(15).Text, "0000") & Format(Text1(16).Text, "0000") & Format(Text1(17).Text, "00") & Format(Text1(18).Text, "0000000000")
            If Len(cta) = 20 Then
                If Text1(25).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then Text1(25).Text = "ES" & cta
                Else
                    CC = CStr(Mid(Text1(25).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(Text1(25).Text, 3) <> cta Then
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
            End If
        End If
    End If
            
    If Index = 19 Or Index = 20 Or Index = 21 Or Index = 22 Then
        If Text1(19).Text <> "" And Text1(20).Text <> "" And Text1(21).Text <> "" And Text1(22).Text <> "" Then
            cta = Format(Text1(19).Text, "0000") & Format(Text1(20).Text, "0000") & Format(Text1(21).Text, "00") & Format(Text1(22).Text, "0000000000")
            If Len(cta) = 20 Then
                If Text1(26).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then Text1(26).Text = "ES" & cta
                Else
                    CC = CStr(Mid(Text1(26).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(Text1(26).Text, 3) <> cta Then
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
            End If
        End If
    End If
            
End Sub

' ---- [19/10/2009] [LAURA] : Añadir funcion generica de ccoste
'Private Sub PonceCentroCoste()
'Dim C As String
'    text1(10).Text = Trim(text1(10).Text)
'    C = ""
'    If text1(10).Text <> "" Then
'        C = PonerNombreDeCod(text1(10), conConta, "cabccost", "nomccost", "codccost")
'        If C = "" Then
'            MsgBox "No existe centro de coste", vbExclamation
'            text1(10).Text = ""
'        End If
'    End If
'    Text2(10).Text = C
'
'End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        PonerFoco Text1(0)
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String

    Set frmTra = New frmBasico2
    
    AyudaTrabajadores frmTra, , CadB
    
    Set frmTra = Nothing

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
            PonerFoco Text1(0)
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
        PonerModo 2
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
Dim Sql As String
Dim vWhere As String
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
   
    vWhere = " WHERE codtraba= " & Data1.Recordset!CodTraba
    'Estudios/Formacion - Datos de la tabla strab1
    Sql = "Select * from strab1 " & vWhere
    Sql = Sql & " order by numlinea"
    CargaGrid DataGrid1, Data2, Sql
    
    'Habilidades
    Sql = "Select * from strab2 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
    Sql = Sql & " order by numlinea"
    CargaGrid DataGrid2, Data3, Sql

    'Experiencia Laboral
    Sql = "Select * from strab3 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
    Sql = Sql & " order by numlinea"
    CargaGrid DataGrid3, Data4, Sql

    'Formacion Realizada
    Sql = "Select * from strab4 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
    Sql = Sql & " order by numlinea"
    CargaGrid DataGrid4, Data5, Sql

    'Formacion Empresa
    Sql = "Select * from strab5 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
    Sql = Sql & " order by numlinea"
    CargaGrid DataGrid5, Data6, Sql

    PrimeraVez = False
    Screen.MousePointer = vbDefault
    Exit Sub
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error Resume Next
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    Text2(24).Text = PonerNombreDeCod(Text1(24), conAri, "salmpr", "nomalmac")
    
    ' ---- [19/10/2009] [LAURA] : Añadir funcion generica de ccoste
'    PonceCentroCoste
    Me.Text2(10).Text = PonerNombreCCoste(Me.Text1(10))
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas asociadas al trabajador
    
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

    'Visualizar el login solo si es administrador o root
    b = (vUsu.Nivel < 2)
    Me.Label1(25).visible = b
    Text1(23).visible = b

    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I



    'Actualiza Iconos Insertar,Modificar,Eliminar
'--    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
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
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 'And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b
    Next I
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    '-----------------------------
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub



Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim I As Byte

    b = (Modo = 2 Or Modo = 5 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 Or Modo = 5)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Mantenimiento lineas
    b = (Modo = 2)
'    For i = 10 To 14
'        Toolbar1.Buttons(i).Enabled = b
'    Next i
    Toolbar1.Buttons(8).Enabled = b Or Modo = 0
'    Me.mnEstudios.Enabled = b
'    Me.mnExperiencia.Enabled = b
'    Me.mnFormEmpresa.Enabled = b
'    Me.mnFormRealizada.Enabled = b
'    Me.mnHabilidades.Enabled = b
    
    '------------------------------------------
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    
    b = (Modo = 2 Or Modo = 3 Or Modo = 4)
    For I = 0 To ToolAux.Count - 1 '[Monica]30/09/2013: antes - 1
        ToolAux(I).Buttons(1).Enabled = b
        Select Case I
            Case 0
                If b Then bAux = (b And Me.Data6.Recordset.RecordCount > 0)
            Case 1
                If b Then bAux = (b And Me.Data3.Recordset.RecordCount > 0)
            Case 2
                If b Then bAux = (b And Me.Data4.Recordset.RecordCount > 0)
            Case 3
                If b Then bAux = (b And Me.Data5.Recordset.RecordCount > 0)
            Case 4
                If b Then bAux = (b And Me.Data2.Recordset.RecordCount > 0)
        End Select
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
        
    
    
    
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
  PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cta As String
Dim cadMen As String

On Error GoTo EDatosOK

    DatosOk = False
    b = True
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
'[Monica]22/11/2013: iban
'    Comprueba_CuentaBan (Text1(31).Text & Text1(32).Text & Text1(33).Text & Text1(34).Text)
    If b And (Modo = 3 Or Modo = 4) Then
        
        
        '[Monica]22/08/2013: añadida la comprobacion de que la cuenta contable sea correcta
        If Text1(15).Text = "" Or Text1(16).Text = "" Or Text1(17).Text = "" Or Text1(18).Text = "" Then
            '[Monica]20/11/2013: añadido el codigo de iban
            Text1(25).Text = ""
            Text1(15).Text = ""
            Text1(16).Text = ""
            Text1(17).Text = ""
            Text1(18).Text = ""
        Else
            cta = Format(Text1(15).Text, "0000") & Format(Text1(16).Text, "0000") & Format(Text1(17).Text, "00") & Format(Text1(18).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El trabajador no tiene asignada cuenta bancaria de nómina."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria de nómina del trabajador no es correcta. ¿ Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(15)
                    b = False
                End If
            Else
                BuscaChekc = ""
                If Me.Text1(25).Text <> "" Then BuscaChekc = Mid(Text1(25).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.Text1(25).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(25).Text = BuscaChekc & cta
                    Else
                        If Mid(Text1(25).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.Text1(25).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco Text1(25)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
        '[Monica]22/08/2013: añadida la comprobacion de que la cuenta contable sea correcta
        If Text1(19).Text = "" Or Text1(20).Text = "" Or Text1(21).Text = "" Or Text1(22).Text = "" Then
            '[Monica]20/11/2013: añadido el codigo de iban
            Text1(26).Text = ""
            Text1(19).Text = ""
            Text1(20).Text = ""
            Text1(21).Text = ""
            Text1(22).Text = ""
        Else
            cta = Format(Text1(19).Text, "0000") & Format(Text1(20).Text, "0000") & Format(Text1(21).Text, "00") & Format(Text1(22).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El trabajador no tiene asignada cuenta bancaria de gastos."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria de gastos del trabajador no es correcta. ¿ Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(19)
                    b = False
                End If
            Else
                BuscaChekc = ""
                If Me.Text1(26).Text <> "" Then BuscaChekc = Mid(Text1(26).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.Text1(26).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(26).Text = BuscaChekc & cta
                    Else
                        If Mid(Text1(26).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.Text1(26).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco Text1(19)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
    
    
    
    
    End If
          
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    Select Case Me.SSTab1.Tab
        Case 1 'Estudios/Formacion
            If (Not Trim(txtAux1(0).Text) <> "") And (Not Trim(txtAux1(1).Text) <> "") Then
                MsgBox "Los campos Periodo y Formación no pueden ser nulos", vbExclamation
                b = False
            End If
        Case 2 'Habilidades
            If Trim(txtAux2.Text) = "" Then
                MsgBox "El campo Habilidad no puede ser nulo", vbExclamation
                b = False
            End If
        Case 3 'Experiencial Laboral
            If (Not Trim(txtAux3(0).Text) <> "") And (Not Trim(txtAux3(1).Text) <> "") Then
                MsgBox "Los campos Periodo y Experiencia no pueden ser nulos", vbExclamation
                b = False
            End If
        Case 4 'Formacion Realizada
            If (Not Trim(txtAux4(0).Text) <> "") And (Not Trim(TxtAux5(1).Text) <> "") Then
                MsgBox "Los campos Fecha Formación y Fecha Evaluación no pueden ser nulos", vbExclamation
                b = False
            End If
        Case 5 'Formacion Empresa
            If Trim(TxtAux5(0).Text) = "" Then
                MsgBox "El campo Fecha Formación no puede ser nulo", vbExclamation
                b = False
            End If
    End Select
    
    DatosOkLinea = b
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Modo = 5
    
    NumTabMto = Me.SSTab1.Tab
    TituloLinea = Me.SSTab1.TabCaption(SSTab1.Tab)
    PonerModo 5
    
    Select Case Index
        Case 0
            NomTablaLineas = "strab5"
        Case 1
            NomTablaLineas = "strab2"
        Case 2
            NomTablaLineas = "strab3"
        Case 3
            NomTablaLineas = "strab4"
        Case 4
            NomTablaLineas = "strab1"
    End Select

    Select Case Button.Index
        Case 1
            BotonAnyadirLinea
        Case 2
            BotonModificarLinea
        Case 3
            BotonEliminarLinea
        Case Else
    End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 5  'Buscar
           mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
            
        Case 1  'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
            
'        Case 10  'Estudios/Formacion
'            mnEstudios_Click
'        Case 11  'Habilidades
'            mnHabilidades_Click
'        Case 12  'Experiencia Laboral
'            mnExperiencia_Click
'        Case 13 'Formacion Realizada
'            mnFormRealizada_Click
'        Case 14  'Formacion Empresa
'            mnFormEmpresa_Click
        Case 8
            frmListado2.Opcion = 17
            frmListado2.Show vbModal
'        Case 17    'Salir
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
   
    
Private Function InsertarLinea() As Boolean
Dim Sql As String
Dim vWhere As String
Dim numF As String
On Error GoTo EInsertarLinea

    InsertarLinea = False
    Sql = ""
    If DatosOkLinea Then
          vWhere = "codtraba=" & Val(Text1(0).Text)
          numF = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
          Select Case Me.SSTab1.Tab
             Case 1 'Estudios/Formacion
                 Sql = "INSERT INTO strab1 "
                 Sql = Sql & "(codtraba, numlinea, periodos, formacio) "
                 Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
                 Sql = Sql & DBSet(txtAux1(0).Text, "T") & "," & DBSet(txtAux1(1).Text, "T") & ")"
            Case 2 'Habilidades
                 Sql = "INSERT INTO strab2 "
                 Sql = Sql & "(codtraba, numlinea, habilida) "
                 Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
                 Sql = Sql & DBSet(txtAux2.Text, "T") & ")"
            Case 3 'Experiencia Laboral
                 Sql = "INSERT INTO strab3 "
                 Sql = Sql & "(codtraba, numlinea, periodo1, experien) "
                 Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
                 Sql = Sql & DBSet(txtAux3(0).Text, "T") & ", " & DBSet(txtAux3(1).Text, "T") & ")"
            Case 4 'Formacion Realizada
                 Sql = "INSERT INTO strab4 "
                 Sql = Sql & "(codtraba, numlinea, fecforma, fecevalu, formaci1, codcentr, evaluaci) "
                 Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ",'"
                 Sql = Sql & Format(txtAux4(0).Text, FormatoFecha) & "', '" & Format(txtAux4(1).Text, FormatoFecha) & "', "
                 Sql = Sql & DBSet(txtAux4(2).Text, "T") & ", " & DBSet(txtAux4(3).Text, "T") & ", " & DBSet(txtAux4(4).Text, "T") & ")"
            Case 5 'Formacion Empresa
                 Sql = "INSERT INTO strab5 "
                 Sql = Sql & "(codtraba, numlinea, fecforma, formaci2, resforma) "
                 Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ",'"
                 Sql = Sql & Format(TxtAux5(0).Text, FormatoFecha) & "', " & DBSet(TxtAux5(1).Text, "T") & ", " & DBSet(TxtAux5(2).Text, "T") & ")"
          End Select
     End If
    
    If Sql <> "" Then
        conn.Execute Sql
        InsertarLinea = True
    End If
    Exit Function
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Trabajador" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
Dim Sql As String
Dim vWhere As String
On Error GoTo EModificarLinea

    ModificarLinea = False
    Sql = ""
    If DatosOkLinea Then
         vWhere = "codtraba=" & Val(Text1(0).Text)
         Select Case Me.SSTab1.Tab
            Case 1 'Estudios/Formacion
                Sql = "UPDATE strab1 Set periodos = " & DBSet(txtAux1(0).Text, "T")
                Sql = Sql & ", formacio = " & DBSet(txtAux1(1).Text, "T")
                Sql = Sql & " WHERE " & vWhere & " AND numlinea=" & Data2.Recordset!numlinea
            Case 2 'Habilidades
                Sql = "UPDATE strab2 Set habilida = " & DBSet(txtAux2.Text, "T")
                Sql = Sql & " WHERE " & vWhere & " AND numlinea=" & Data3.Recordset!numlinea
            Case 3 'Experiencia Laboral
                Sql = "UPDATE strab3 Set periodo1 = " & DBSet(txtAux3(0).Text, "T") & ", "
                Sql = Sql & "experien=" & DBSet(txtAux3(1).Text, "T")
                Sql = Sql & " WHERE " & vWhere & " AND numlinea=" & Data4.Recordset!numlinea
            Case 4 'Formacion Realizada
                Sql = "UPDATE strab4 Set fecforma = " & DBSet(txtAux4(0).Text, "F") & ", fecevalu="
                Sql = Sql & DBSet(txtAux4(1).Text, "F") & ", formaci1=" & DBSet(txtAux4(2).Text, "T") & ", codcentr="
                Sql = Sql & DBSet(txtAux4(3).Text, "T") & ", evaluaci=" & DBSet(txtAux4(4).Text, "T")
                Sql = Sql & " WHERE " & vWhere & " AND numlinea=" & Data5.Recordset!numlinea
            Case 5 'Formacion Empresa
                Sql = "UPDATE strab5 Set fecforma = " & DBSet(TxtAux5(0).Text, "F") & ", formaci2="
                Sql = Sql & DBSet(TxtAux5(1).Text, "T") & ", resforma=" & DBSet(TxtAux5(2).Text, "T")
                Sql = Sql & " WHERE " & vWhere & " AND numlinea=" & Data6.Recordset!numlinea
        End Select
    End If

    If Sql <> "" Then
        conn.Execute Sql
        ModificarLinea = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Trabajador" & vbCrLf & Err.Description
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
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, Sql As String)
On Error GoTo ECargaGrid

    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    vDataGrid.RowHeight = 320
     
    CargaGrid2 vDataGrid, vData
    vDataGrid.Enabled = (Modo = 0 Or Modo = 2)
    vDataGrid.ScrollBars = dbgAutomatic
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Integer

On Error GoTo ECargaGrid

    vData.Refresh
    
    vDataGrid.Columns(0).visible = False 'codtraba
    vDataGrid.Columns(1).visible = False 'numlinea

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Estudios / Formacion
                vDataGrid.Columns(2).Caption = "Período"
                vDataGrid.Columns(2).Width = 2100
                vDataGrid.Columns(3).visible = True
                vDataGrid.Columns(3).Caption = "Formación"
                vDataGrid.Columns(3).Width = 6450
        Case "DataGrid2" 'Habilidades
                vDataGrid.Columns(2).Caption = "Habilidades"
                vDataGrid.Columns(2).Width = 7250
        Case "DataGrid3" 'Experiencia Laboral
                vDataGrid.Columns(2).Caption = "Periodo"
                vDataGrid.Columns(2).Width = 2100
                vDataGrid.Columns(3).visible = True
                vDataGrid.Columns(3).Caption = "Experiencia"
                vDataGrid.Columns(3).Width = 6450
        Case "DataGrid4" 'Formacion Realizada
                vDataGrid.Columns(2).Caption = "Fec.Formación"
                vDataGrid.Columns(2).Width = 1650
                vDataGrid.Columns(3).Caption = "Fec.Evaluación"
                vDataGrid.Columns(3).Width = 1650
                vDataGrid.Columns(4).Caption = "Formación"
                vDataGrid.Columns(4).Width = 4000
                vDataGrid.Columns(5).Caption = "Centro"
                vDataGrid.Columns(5).Width = 1670
                vDataGrid.Columns(6).Caption = "Evaluación"
                vDataGrid.Columns(6).Width = 1160
        Case "DataGrid5" 'Formacion Empresa
                vDataGrid.Columns(2).Caption = "Fec.Formación"
                vDataGrid.Columns(2).Width = 1700
                vDataGrid.Columns(3).Caption = "Formación"
                vDataGrid.Columns(3).Width = 4670
                vDataGrid.Columns(4).Caption = "Resultado"
                vDataGrid.Columns(4).Width = 1900
    End Select

    vDataGrid.Enabled = (Modo = 0) Or (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
    For I = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(I).Locked = True
        vDataGrid.Columns(I).AllowSizing = False
    Next I

    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux1(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux1.Count - 1 'TextBox
            txtAux1(I).Top = 290
            txtAux1(I).visible = visible
        Next I
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 0 To txtAux1.Count - 1
                txtAux1(I).Text = ""
                BloquearTxt txtAux1(I), False
            Next I
        Else
            For I = 0 To txtAux1.Count - 1
                txtAux1(I).Text = DataGrid1.Columns(I + 2).Text
                BloquearTxt txtAux1(I), False
            Next I
        End If


        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 8)
        
        For I = 0 To txtAux1.Count - 1
            txtAux1(I).Top = alto
            txtAux1(I).Height = DataGrid1.RowHeight
        Next I
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Periodo
        txtAux1(0).Left = DataGrid1.Left + 320
        txtAux1(0).Width = DataGrid1.Columns(2).Width - 20
        'Formacion
        txtAux1(1).Left = txtAux1(0).Left + txtAux1(0).Width + 20
        txtAux1(1).Width = DataGrid1.Columns(3).Width - 20
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux1.Count - 1
            txtAux1(I).visible = visible
        Next I
    End If
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux2(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
            txtAux2.Top = 290
            txtAux2.visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid2
            txtAux2.Text = ""
            BloquearTxt txtAux2, False
        Else
            txtAux2.Text = DataGrid2.Columns(2).Text
            BloquearTxt txtAux2, False
        End If


        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid2, 8)
        
        txtAux2.Top = alto
        txtAux2.Height = DataGrid2.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Habilidad
        txtAux2.Left = DataGrid2.Left + 320
        txtAux2.Width = DataGrid2.Columns(2).Width - 20
            
        'Los ponemos Visibles o No
        '--------------------------
        txtAux2.visible = visible
    End If
End Sub


Private Sub CargaTxtAux3(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte
    
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux3.Count - 1 'TextBox
            txtAux3(I).Top = 290
            txtAux3(I).visible = visible
        Next I
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid3
            For I = 0 To txtAux3.Count - 1
                txtAux3(I).Text = ""
                BloquearTxt txtAux3(I), False
            Next I
        Else
            For I = 0 To txtAux3.Count - 1
                txtAux3(I).Text = DataGrid3.Columns(I + 2).Text
                BloquearTxt txtAux3(I), False
            Next I
        End If


        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid3, 8)
        
        For I = 0 To txtAux3.Count - 1
            txtAux3(I).Top = alto
            txtAux3(I).Height = DataGrid3.RowHeight
        Next I
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Periodo
        txtAux3(0).Left = DataGrid3.Left + 320
        txtAux3(0).Width = DataGrid3.Columns(2).Width - 20
        'Experiencia
        txtAux3(1).Left = txtAux3(0).Left + txtAux3(0).Width + 20
        txtAux3(1).Width = DataGrid3.Columns(3).Width - 20
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux3.Count - 1
            txtAux3(I).visible = visible
        Next I
    End If
End Sub


Private Sub CargaTxtAux4(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux4.Count - 1 'TextBox
            txtAux4(I).Top = 290
            txtAux4(I).visible = visible
        Next I
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid4
            For I = 0 To txtAux4.Count - 1
                txtAux4(I).Text = ""
                BloquearTxt txtAux4(I), False
            Next I
        Else
            For I = 0 To txtAux4.Count - 1
                txtAux4(I).Text = DataGrid4.Columns(I + 2).Text
                BloquearTxt txtAux4(I), False
            Next I
        End If


        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid4, 8)
        
        For I = 0 To txtAux4.Count - 1
            txtAux4(I).Top = alto
            txtAux4(I).Height = DataGrid4.RowHeight
        Next I
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Fecha Formacion
        txtAux4(0).Left = DataGrid4.Left + 320
        txtAux4(0).Width = DataGrid4.Columns(2).Width - 20
        'Fecha Evaluacion
        For I = 1 To 4
            txtAux4(I).Left = txtAux4(I - 1).Left + txtAux4(I - 1).Width + 20
            txtAux4(I).Width = DataGrid4.Columns(I + 2).Width - 20
        Next I
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux4.Count - 1
            txtAux4(I).visible = visible
        Next I
    End If
End Sub


Private Sub CargaTxtAux5(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

'Formacion Empresa

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To TxtAux5.Count - 1 'TextBox
            TxtAux5(I).Top = 290
            TxtAux5(I).visible = visible
        Next I
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid5
            For I = 0 To TxtAux5.Count - 1
                TxtAux5(I).Text = ""
                BloquearTxt TxtAux5(I), False
            Next I
        Else
            For I = 0 To TxtAux5.Count - 1
                TxtAux5(I).Text = DataGrid5.Columns(I + 2).Text
                BloquearTxt TxtAux5(I), False
            Next I
        End If


        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid5, 8)
        
        For I = 0 To TxtAux5.Count - 1
            TxtAux5(I).Top = alto
            TxtAux5(I).Height = DataGrid5.RowHeight
        Next I
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Fecha Formacion
        TxtAux5(0).Left = DataGrid5.Left + 320
        TxtAux5(0).Width = DataGrid5.Columns(2).Width - 20
        'Formacion
        TxtAux5(1).Left = TxtAux5(0).Left + TxtAux5(0).Width + 20
        TxtAux5(1).Width = DataGrid5.Columns(3).Width - 20
        'Resultado
        TxtAux5(2).Left = TxtAux5(1).Left + TxtAux5(1).Width + 20
        TxtAux5(2).Width = DataGrid5.Columns(4).Width - 20
                
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To TxtAux5.Count - 1
            TxtAux5(I).visible = visible
        Next I
    End If
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
    ConseguirFoco txtAux1(Index), Modo
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
      If Not (Index = 0 And KeyCode = 38) Then
            KEYdown KeyCode
      End If
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
        Me.SSTab1.Tab = numTab
        NumTabMto = numTab
        TituloLinea = Cad
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Sub TxtAux1_LostFocus(Index As Integer)
    If Index = 1 And Me.SSTab1.Tab = 1 Then
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub txtAux2_GotFocus()
    ConseguirFoco txtAux2, Modo
End Sub

Private Sub txtAux2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then
           KEYdown KeyCode
    End If
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
     If Index = 1 And KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub TxtAux4_GotFocus(Index As Integer)
    ConseguirFoco txtAux4(Index), Modo
End Sub


Private Sub TxtAux4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then
           KEYdown KeyCode
    End If
End Sub

Private Sub TxtAux4_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 4 And KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub TxtAux4_LostFocus(Index As Integer)
    If Trim(txtAux4(Index).Text) = "" Then Exit Sub
    
    Select Case Index
        Case 0, 1
             PonerFormatoFecha txtAux4(Index)
    End Select
End Sub


Private Sub TxtAux5_GotFocus(Index As Integer)
    ConseguirFoco TxtAux5(Index), Modo
End Sub

Private Sub TxtAux5_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then
        KEYdown KeyCode
    End If
End Sub

Private Sub TxtAux5_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub TxtAux5_LostFocus(Index As Integer)
    If Trim(TxtAux5(Index).Text) = "" Then Exit Sub
    Select Case Index
        Case 0
            PonerFormatoFecha TxtAux5(Index)
        Case 2
'            If Me.ActiveControl.TabIndex <> 72 Then PonerFocoBtn Me.cmdAceptar
            PonerFocoBtn Me.cmdAceptar
    End Select
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String
On Error GoTo FinEliminar

        conn.BeginTrans
        Sql = " WHERE  codtraba=" & Data1.Recordset!CodTraba

        'Lineas Estudios/Formacion
        conn.Execute "Delete from strab1 " & Sql
        'Lineas Habilidades
        conn.Execute "Delete from strab2 " & Sql
        'Lineas Experiencia Laboral
        conn.Execute "Delete from strab3 " & Sql
        'Lineas Formacion Realizada
        conn.Execute "Delete from strab4 " & Sql
        'Lineas Experiencia Empresa
        conn.Execute "Delete from strab5 " & Sql
        'Cabeceras
        conn.Execute "Delete from straba " & Sql

FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function



Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
Dim Cad As String
On Error Resume Next

    Cad = "Select * from strab1 where codtraba is null"
    CargaGrid DataGrid1, Data2, Cad
    Cad = "Select * from strab2 where codtraba is null"
    CargaGrid DataGrid2, Data3, Cad
    Cad = "Select * from strab3 where codtraba is null"
    CargaGrid DataGrid3, Data4, Cad
    Cad = "Select * from strab4 where codtraba is null"
    CargaGrid DataGrid4, Data5, Cad
    Cad = "Select * from strab5 where codtraba is null"
    CargaGrid DataGrid5, Data6, Cad
    
    PrimeraVez = False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(codtraba=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
       PonerModo 2
       lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       'Poner los grid sin apuntar a nada
       LimpiarDataGrids
       PonerModo 0
    End If
End Sub



