VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmMovimientos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos Almacen"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12855
   Icon            =   "frmAlmMovimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   12855
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
      Left            =   8400
      TabIndex        =   38
      Top             =   180
      Width           =   1605
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3870
      TabIndex        =   36
      Top             =   30
      Width           =   915
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   37
         Top             =   180
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Actualizar"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4860
      TabIndex        =   34
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   35
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
      Left            =   210
      TabIndex        =   32
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   33
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
   Begin VB.Frame FrameAux0 
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   30
      TabIndex        =   22
      Top             =   2280
      Width           =   12735
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
         Height          =   330
         Index           =   0
         Left            =   240
         MaxLength       =   16
         TabIndex        =   26
         Text            =   "codartic"
         Top             =   2820
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
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
         Height          =   330
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   25
         Text            =   "nombre artic"
         Top             =   2820
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   2
         Left            =   3960
         MaxLength       =   16
         TabIndex        =   28
         Text            =   "cantidad"
         Top             =   2820
         Visible         =   0   'False
         Width           =   975
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
         Height          =   330
         Index           =   3
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   31
         Text            =   "observac"
         Top             =   2820
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Left            =   1200
         TabIndex        =   24
         ToolTipText     =   "Buscar art�culo"
         Top             =   2820
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.ComboBox cboAux 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Tag             =   "Situaci�n Impresi�n|N|N|||scamov|situacio||N|"
         Top             =   2820
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   120
         TabIndex        =   23
         Top             =   -60
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   210
            TabIndex        =   29
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmAlmMovimientos.frx":000C
         Height          =   3255
         Left            =   120
         TabIndex        =   27
         Top             =   540
         Width           =   12540
         _ExtentX        =   22119
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Left            =   6015
      MaxLength       =   8
      TabIndex        =   20
      Tag             =   "Hora|H|N|||scamov|hormovim|hh:mm:ss|N|"
      Text            =   "Text1"
      Top             =   825
      Width           =   855
   End
   Begin VB.CheckBox chkImpresion 
      Caption         =   "Impreso"
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
      Height          =   255
      Left            =   5700
      TabIndex        =   19
      Tag             =   "Situaci�n Impresi�n|N|N|||scamov|situacio||N|"
      Top             =   885
      Width           =   1275
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
      Left            =   10350
      TabIndex        =   5
      Top             =   6330
      Width           =   1135
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
      Left            =   11610
      TabIndex        =   6
      Top             =   6315
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
      Left            =   11610
      TabIndex        =   17
      Top             =   6330
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   60
      TabIndex        =   15
      Top             =   6240
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
         Height          =   240
         Left            =   240
         TabIndex        =   16
         Top             =   180
         Width           =   2115
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
      Index           =   0
      Left            =   2820
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   1380
      Width           =   4065
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
      Left            =   2820
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   1785
      Width           =   4065
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
      Height          =   1035
      Index           =   4
      Left            =   7110
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   4
      Tag             =   "Observaciones|T|S|||scamov|observa1||N|"
      Text            =   "frmAlmMovimientos.frx":0021
      Top             =   1110
      Width           =   5505
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
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "Cod. Trabajador|N|N|0|9999|scamov|codtraba|0000|N|"
      Text            =   "Text1"
      Top             =   1785
      Width           =   975
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "Cod. Almacen|N|N|0|999|scamov|codalmac|000|N|"
      Text            =   "Text1"
      Top             =   1380
      Width           =   975
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
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||scamov|fecmovim|dd/mm/yyyy|N|"
      Text            =   "Text1"
      Top             =   825
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8280
      Top             =   480
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FEF7E4&
      BeginProperty Font 
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
      Left            =   1770
      MaxLength       =   7
      TabIndex        =   0
      Tag             =   "N� Movimiento|N|S|0||scamov|codmovim|0000000|S|"
      Text            =   "Text1"
      Top             =   825
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   9720
      Top             =   480
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   5430
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Hora"
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
      Left            =   5370
      TabIndex        =   21
      Top             =   870
      Width           =   585
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1500
      ToolTipText     =   "Buscar trabajador"
      Top             =   1815
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1500
      ToolTipText     =   "Buscar almacen"
      Top             =   1425
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   3675
      Picture         =   "frmAlmMovimientos.frx":0027
      ToolTipText     =   "Buscar fecha"
      Top             =   825
      Width           =   240
   End
   Begin VB.Label Label6 
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
      Left            =   7140
      TabIndex        =   12
      Top             =   840
      Width           =   1725
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   11
      Top             =   1785
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   10
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
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
      Left            =   2955
      TabIndex        =   9
      Top             =   825
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "N� Movimiento"
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
      Left            =   240
      TabIndex        =   8
      Top             =   825
      Width           =   1485
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
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
End
Attribute VB_Name = "frmAlmMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 208


Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'historico schmov, y solo en modo de consulta

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del hist�rico de movimiento seleccionado (solo consulta)
Public hcoCodMovim As Long 'cod. movim del historico
Public hcoFechaMovim As Date 'Fecha del historico


'-----------------------------------------------------------------------

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmA As frmAlmAlPropios 'Almacen Origen/Destino
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores 'Mto de Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmArt As frmAlmArticulos   'Form Articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmAlmMov As frmAlmMovimientosPrev   'Form previsualizar
Attribute frmAlmMov.VB_VarHelpID = -1

Dim NombreTabla As String
Dim NomTablaLineas As String
Dim Ordenacion As String

Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1

Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en funci�n del valor en la tabla de par�metros: stipom

Dim CadenaConsulta As String
Dim cadSeleccion As String 'Cadena de seleccion para FormulaSelection del Informe


Private HaDevueltoDatos As Boolean



Private Sub cboAux_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboImpresion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    
    Select Case Modo
    Case 1 'BUSQUEDA
        cadSeleccion = ""
        HacerBusqueda
        
    Case 3 'INSERTAR
        If DatosOk Then InsertarCabecera

    Case 4 'MODIFICAR
        If DatosOk Then
            If ModificaDesdeFormulario(Me, 1) Then
                TerminaBloquear
                cad = "(" & ObtenerWhereCP(False) & ")"
                If SituarData(Data1, cad, Indicador) Then
                    PonerModo 2
                    lblIndicador.Caption = Indicador
                Else
                    PonerModo 0
                End If
            End If
        End If
            
    Case 5 'Lineas Movimientos Almacenes
        If InsertarModificarLinea Then
            'Reestablecemos los campos y ponemos el grid
            DataGrid1.AllowAddNew = False
'            CargaGrid True
            If ModificaLineas = 1 Then 'Insertar
                CargaGrid True
                ModificaLineas = 0
                BotonAnyadirLineas 0
            ElseIf ModificaLineas = 2 Then 'Modificar
                TerminaBloquear
                CargaGrid True
                Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                ModificaLineas = 0
'                PonerBotonCabecera True
                CargaTxtAux False, False
                Me.lblIndicador.Caption = ""
            End If
        End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdAux_Click()
    Set frmArt = New frmAlmArticulos
    frmArt.DatosADevolverBusqueda2 = "@1@" 'Abre en Modo busqueda
    frmArt.Show vbModal
    Set frmArt = Nothing
    PonerFoco txtAux(0)
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
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
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Mantenimiento Lineas traspasos
            CargaTxtAux False, False
            DataGrid1.AllowAddNew = False
            If Not ModificaLineas = 2 Then '2 = Modificar
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
'            PonerBotonCabecera True
            PonerModo 2
            PonerCampos
            DataGrid1.Refresh
            DataGrid1.Enabled = True
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then 'modo 5: Mantenimiento Lineas
'        PonerBotonCabecera False
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
    End If
End Sub


Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    If Modo = 5 And KeyAscii = 27 Then 'ESC 'Modo Lineas
        cmdRegresar_Click
    End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If Modo = 5 And KeyAscii = 27 Then 'ESC 'Modo Lineas
        cmdRegresar_Click
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmppal.Icon
    
    'ICONOS de La toolbar
    btnAnyadir = 5 'Posicion del boton A�adir en la toolbar1
    btnPrimero = 15 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
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
'        .Buttons(10).Image = 39 'Actualizar
        .Buttons(8).Image = 16 'Imprimir
'        .Buttons(13).Image = 15 'Salir
'        .Buttons(btnPrimero).Image = 6 'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    
    With Me.Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.imgListComun1
        .Buttons(1).Image = 39  'baja de socio
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
    
    'ICONITOS DE LAS BARRAS EN LOS TABS DE LINEA
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            '.ImageList = frmPpal.imgListComun_VELL
            '  ### [Monica] 02/10/2006 acabo de comentarlo
            .HotImageList = frmppal.imgListComun_OM16
            .DisabledImageList = frmppal.imgListComun_BN16
            .ImageList = frmppal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next I
    
    For I = 0 To Me.imgBuscar.Count - 1
        imgBuscar(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    CodTipoMov = "REG"
    
    'campo situacio solo en tabla scamov
    Me.chkImpresion.visible = Not EsHistorico
    'Campo Hora solo en el Historico
    Me.Label4.visible = EsHistorico
    Me.Text1(5).visible = EsHistorico
    
    cadSeleccion = ""
   
    If Not EsHistorico Then
        NombreTabla = "scamov"
        NomTablaLineas = "slimov" 'Tabla lineas de Movimientos
        Me.Caption = "Movimientos de Almacen"
    Else
        NombreTabla = "schmov"
        NomTablaLineas = "slhmov"
        CargarTagsHco Me, "scamov", NombreTabla
        Me.Caption = "Hist�rico Movimientos de Almacen"
    End If
    Ordenacion = " ORDER BY codmovim"
    
    CadenaConsulta = "Select * from " & NombreTabla
    If hcoCodMovim <> -1 Then
    'Se llama desde Dobleclick en frmAlmMovimArticulos
        CadenaConsulta = CadenaConsulta & " where codmovim=" & hcoCodMovim & " and fecmovim= """ & Format(hcoFechaMovim, "yyyy-mm-dd") & """"
    Else
         CadenaConsulta = CadenaConsulta & " WHERE codmovim = -1"
    End If
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Not Data1.Recordset.EOF Then 'Se llama desde DblClick frmAlmMovimArticulos
                                    'Se carga con el valor del registro del DblClick
        Data1.Recordset.MoveFirst
        Me.Text1(0).Text = Format(Data1.Recordset!codMovim, "0000000")
        Me.Text1(1).Text = Data1.Recordset!fecmovim
        Me.Text1(5).Text = Format(Data1.Recordset!hormovim, "hh:mm:ss")
        'Cod. Almacen
        Me.Text1(2).Text = Format(Data1.Recordset!codAlmac, "000")
        Me.Text2(0).Text = PonerNombreDeCod(Text1(2), conAri, "salmpr", "nomalmac", "codalmac")
        'Cod. Trabajador
        Me.Text1(3).Text = Format(Data1.Recordset!CodTraba, "0000")
        Me.Text2(1).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba", "codtraba")
        'Observaciones
        Text1(4).Text = DBLet(Data1.Recordset!observa1, "T")
        CargaGrid True
    Else
        CargaGrid False '(Modo = 2) 'False
    End If
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim Sql As String
On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, Sql, False
    
    DataGrid1.RowHeight = 350
    
    
    DataGrid1.Columns(0).visible = False 'Cod. Movim
    DataGrid1.Columns(1).visible = False 'Numlinea
    I = 2
    
    'Cod. Art�culo
    DataGrid1.Columns(I).Caption = "Cod. Articulo"
    DataGrid1.Columns(I).Width = 1700
    
    'Nombre Art�culo
    I = I + 1
    DataGrid1.Columns(I).Caption = "Nombre Articulo"
    DataGrid1.Columns(I).Width = 4100
    
    'Cantidad
    I = I + 1
    DataGrid1.Columns(I).Caption = "Cantidad"
    DataGrid1.Columns(I).Width = 1300
    DataGrid1.Columns(I).Alignment = dbgRight
    DataGrid1.Columns(I).NumberFormat = FormatoImporte
    
    'tipo Movimiento
    I = I + 1
    DataGrid1.Columns(I).Caption = "T.Mov."
    DataGrid1.Columns(I).Width = 700
    DataGrid1.Columns(I).Alignment = dbgCenter
    
    'Observaciones
    I = I + 1
    DataGrid1.Columns(I).Caption = "Observaciones"
    DataGrid1.Columns(I).Width = 4050
       
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
    Next I
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
    
    PonerModoOpcionesMenu
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim I As Byte
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1
            txtAux(I).top = 290
        Next I
        Me.cmdAux.top = 290
        Me.cboAux.top = 290
    Else
        DeseleccionaGrid Me.DataGrid1
        CargarComboAux
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
                If I <> 1 Then txtAux(I).Locked = False
            Next I
            cmdAux.Enabled = True
            cboAux.Enabled = True
            cboAux.ListIndex = -1
        Else  'Poner valor a los txtAux
            For I = 0 To txtAux.Count - 2
                txtAux(I).Text = DataGrid1.Columns(I + 2).Text
            Next I
            Select Case DataGrid1.Columns(5).Value
                Case "S"
                    Me.cboAux.ListIndex = 0
                Case "E"
                    Me.cboAux.ListIndex = 1
            End Select
            txtAux(3).Text = DataGrid1.Columns(6).Text
            txtAux(0).Locked = True
            cmdAux.Enabled = False
            cboAux.Enabled = True
            txtAux(2).Locked = False
            txtAux(3).Locked = False
        End If
        
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.top + 240 '220
        Else
            alto = DataGrid1.top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If
        
        'Fijamos altura y posici�n Top
        For I = 0 To txtAux.Count - 1
            txtAux(I).top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        Me.cmdAux.top = alto
        Me.cmdAux.Height = DataGrid1.RowHeight
        cboAux.top = alto - 5
        
        'Fijamos anchura y posicion Left
        txtAux(0).Left = DataGrid1.Left + 340 'codartic
        txtAux(0).Width = DataGrid1.Columns(2).Width - 200
        cmdAux.Left = txtAux(0).Left + txtAux(0).Width
        txtAux(1).Left = cmdAux.Left + cmdAux.Width  'Nombre Artic
        txtAux(1).Width = DataGrid1.Columns(3).Width - 35
        I = 2 'Cantidad
        txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 25
        txtAux(I).Width = DataGrid1.Columns(I + 2).Width - 20
        'Tipo Movimiento
        cboAux.Left = txtAux(2).Left + txtAux(2).Width + 20
        cboAux.Width = DataGrid1.Columns(5).Width + 10
        I = 3 'Observac
        txtAux(I).Left = cboAux.Left + cboAux.Width + 30
        txtAux(I).Width = DataGrid1.Columns(6).Width - 60
    End If

    'Los ponemos Visibles o No
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = visible
    Next I
    cmdAux.visible = visible
    cboAux.visible = visible
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacen Propios
Dim indice As Byte
    indice = CByte(Me.imgBuscar(0).Tag)
    Text1(indice + 2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmAlmMov_DatoSeleccionado(CadenaSeleccion As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
    
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        'Recupera todo el registro de Traspaso Almacenes
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
        CadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If Modo <> 5 Then 'Estamos en Cabecera
            'Recupera todo el registro de Traspaso Almacenes
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else 'Estamos en Lineas
            'Llamamos desde el boton auxiliar de Art�culos
            txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
            txtAux(1).Text = RecuperaValor(CadenaDevuelta, 2)
            PonerFoco txtAux(2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
Dim indice As Byte
    indice = 1
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Trabajadores
Dim indice As Byte
    indice = 3
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(indice - 2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    imgBuscar(0).Tag = Index
    
    Select Case Index
        Case 0 'Codigo Almacen
            Set frmA = New frmAlmAlPropios
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
        Case 1  'Cod. Trabajador
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
    End Select
    PonerFoco Text1(Index + 2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte

   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   indice = 1
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(1)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then   'Eliminar lineas Movimiento Almacenes
        BotonEliminarLinea 0
    Else 'Eliminar Cabecera Movimiento Almacenes
        BotonEliminar
    End If
End Sub

Private Sub mnModificar_Click()
Dim vWhere As String

    If Modo = 5 Then  'Modificar LINEAS
        vWhere = ObtenerWhereCP(False) & " and numlinea=" & Me.Data2.Recordset.Fields(1)
        If BloqueaRegistro(NomTablaLineas, vWhere) Then BotonModificarLinea 0
    Else 'Modificar Cabecera
       If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then  'A�adir lineas Movimiento Almacenes
        BotonAnyadirLineas 0
    Else 'A�adir Cabecera Movimiento Almacenes
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

Private Sub Text1_GotFocus(Index As Integer)
    If Index <> 4 Then ConseguirFoco Text1(Index), Modo
End Sub



Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 And Index = 3 And Modo = 1 Then
        PonerFocoBtn cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub Text1_LostFocus(Index As Integer)

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Bloquear el contador si no estamos en busquedas
    If (Modo <> 1) And (Index = 0) Then BloquearTxt Text1(0), True, True

    Select Case Index
        Case 0 'Codigo Movimiento Almacen
            Text1(Index).Text = Format(Text1(Index).Text, "0000000")
        Case 1 'Fecha
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
        Case 2 'Codigo Almacen
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "salmpr", "nomalmac", "codalmac")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 3  'Codigo Trabajador
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 4 'Observaciones
            If Text1(Index).Text <> "" Then Text1(Index).Text = QuitarCaracterEnter(Text1(Index).Text)
    End Select
End Sub

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            BotonAnyadirLineas Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
        Case Else
    End Select
    'End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Actualizar
           BotonActualizar
    End Select

End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Despla�ament; per a despla�ar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub




Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 3 And KeyCode = 40 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 And KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub

    Select Case Index
        Case 0 'Cod ARTICULO
            If txtAux(Index).Text = "" Then
                txtAux(Index + 1).Text = ""
            Else
                 PonerArticulo txtAux(0), txtAux(1), Text1(2).Text, CodTipoMov, ModificaLineas
            End If
            
        Case 2 'CANTIDAD (Comprobamos formato como si fuera un Importe)
            'Formato tipo 1: Decimal(12,2)
            PonerFormatoDecimal txtAux(Index), 1
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
'        Case 1 'Busqueda
'           mnBuscar_Click
'        Case 2 'Ver Todos
'           mnVerTodos_Click
'        Case 5 'Nuevo
'           mnNuevo_Click
'        Case 6  'Modificar
'           mnModificar_Click
'        Case 7 'Eliminar
'           mnEliminar_Click
'
'        Case 9 'Mantenimiento Lineas
'           BotonLineas
'        Case 10 'Actualizar
'           BotonActualizar
'        Case 12 'Imprimir
'           BotonImprimir
'        Case 13  'Salir
'           mnSalir_Click
'        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
'           Desplazamiento (Button.Index - btnPrimero)
    
        Case 5 'Busqueda
           mnBuscar_Click
        Case 6 'Ver Todos
           mnVerTodos_Click
        Case 1 'Nuevo
           mnNuevo_Click
        Case 2  'Modificar
           mnModificar_Click
        Case 3 'Eliminar
           mnEliminar_Click
        Case 8 'Imprimir
           BotonImprimir
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean
    
    'Actualiza Iconos Insertar,Modificar,Eliminar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    For I = 0 To txtAux.Count - 1
        Text1(I).BackColor = vbWhite
    Next I
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModificaLineas
    
    '--------------------------------------------
    b = (Kmodo = 2)
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    'DesplazamientoVisible  Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
    'Como el campo 0 es clave primaria, NO se puede modificar, es contador
    BloquearTxt Text1(0), (Modo <> 1), True
    
'    Me.cmdRegresar.visible = (Not b) And Not EsHistorico
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.visible = b
'    Else
'        cmdRegresar.visible = False
'    End If
    
    '=================================================
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b
    Next I
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I

    Me.chkVistaPrevia.Enabled = (Modo <= 2)

    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
    PonerModoUsuarioGnral Modo, "aritaxi"
                        
End Sub




Private Sub PonerModoUsuarioGnral(Modo As Byte, Aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(Aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF And Not EsHistorico Then
        Toolbar1.Buttons(1).Enabled = Toolbar1.Buttons(1).Enabled And DBLet(Rs!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = Toolbar1.Buttons(2).Enabled And DBLet(Rs!Modificar, "N")
        Toolbar1.Buttons(3).Enabled = Toolbar1.Buttons(3).Enabled And DBLet(Rs!creareliminar, "N")
        
        Toolbar1.Buttons(5).Enabled = Toolbar1.Buttons(5).Enabled And DBLet(Rs!Ver, "N")
        Toolbar1.Buttons(6).Enabled = Toolbar1.Buttons(6).Enabled And DBLet(Rs!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = Toolbar1.Buttons(8).Enabled And DBLet(Rs!Imprimir, "N")
         
        'Actualizar
        Me.Toolbar2.Buttons(1).Enabled = Me.Toolbar2.Buttons(1).Enabled And DBLet(Rs!especial, "N")
        
        'lineas
        ToolAux(0).Buttons(1).Enabled = ToolAux(0).Buttons(1).Enabled And DBLet(Rs!creareliminar, "N")
        ToolAux(0).Buttons(2).Enabled = ToolAux(0).Buttons(2).Enabled And DBLet(Rs!Modificar, "N")
        ToolAux(0).Buttons(3).Enabled = ToolAux(0).Buttons(3).Enabled And DBLet(Rs!creareliminar, "N")
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

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
Dim bAux As Boolean
Dim I As Byte

    'Si visualizamos el historico no mostrar botones de Mantenimiento, solo es consulta
'    For i = 5 To 10
'        Toolbar1.Buttons(i).visible = Not EsHistorico
'    Next i
    Me.mnNuevo.visible = Not EsHistorico
    Me.mnModificar.visible = Not EsHistorico
    Me.mnEliminar.visible = Not EsHistorico
    Me.mnBarra2.visible = Not EsHistorico
    
    If Not EsHistorico Then
        'Modo 2. Hay datos y estamos visualizandolos
        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(1).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(2).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(3).Enabled = b
        Me.mnEliminar.Enabled = b
        
        '--------------------------------
        b = (Modo = 2)
        'Lineas Movimientos Almacenes
'        Toolbar1.Buttons(9).Enabled = b
        'Actualizar
        Toolbar2.Buttons(1).Enabled = b
        
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
        
        
'
        b = (Modo = 2 Or Modo = 3 Or Modo = 4)
        For I = 0 To ToolAux.Count - 1 '[Monica]30/09/2013: antes - 1
            ToolAux(I).Buttons(1).Enabled = b
            If b Then bAux = (b And Me.Data2.Recordset.RecordCount > 0)
            ToolAux(I).Buttons(2).Enabled = bAux
            ToolAux(I).Buttons(3).Enabled = bAux
        Next I
        
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkImpresion.Value = 0
End Sub


'Private Sub Desplazamiento(Index As Integer)
''Botones de Desplazamiento de la Toolbar
'
'    Select Case Modo
'        Case 5 'Modo Mantenimiento de Almacenes (Lineas)
'            If Data2.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data2, Index
'        Case Else 'Datos de Cabecera
'            If Data1.Recordset.EOF Then Exit Sub
'            DesplazamientoData Data1, Index
'            PonerCampos
'    End Select
'End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
Dim Tabla As String
    
    Tabla = NomTablaLineas

    Sql = "SELECT " & Tabla & ".codmovim, "
    Sql = Sql & Tabla & ".numlinea, " & Tabla & ".codartic, Articulos.nomartic, "
    Sql = Sql & Tabla & ".cantidad, if(" & Tabla & ".tipomovi=0,""S"",""E"") as tipomovi, "
    Sql = Sql & Tabla & ".motimovi "
    Sql = Sql & " FROM ((" & Tabla & " LEFT JOIN sartic AS Articulos ON " & Tabla & ".codartic ="
    Sql = Sql & " Articulos.codartic))"
    If enlaza Then
        Sql = Sql & " WHERE codmovim = " & Data1.Recordset!codMovim
    Else
        Sql = Sql & " WHERE codmovim = -1"
    End If
    Sql = Sql & " ORDER BY " & Tabla & ".numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False

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
'Ver todos
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonLineas()
On Error GoTo ErrorLineas

    Screen.MousePointer = vbHourglass
    PonerModo (5)
    ModificaLineas = 0
'    PonerBotonCabecera True
    CargaGrid True
    DataGrid1.Enabled = True
    Me.DataGrid1.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Lineas"
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonAnyadir()
Dim NomTraba As String

    LimpiarCampos 'Vac�a los TextBox
    ModoAnterior = Modo 'Para el bot�n Cancelar en Modo Insertar
    PonerModo 3
           
    'Ponemos el grid lineas Traspaso enlazando a ningun sitio
    CargaGrid False
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    'Poner Trabajador por defecto el trabajador conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(1).Text = NomTraba
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLineas(Index As Integer)
Dim vWhere As String
    
    PonerModo 5
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    
    ModificaLineas = 1
    
    vWhere = ObtenerWhereCP(False)
    cmdAceptar.Tag = SugerirCodigoSiguienteStr("slimov", "numlinea", vWhere)
    
'    PonerBotonCabecera False
'    lblIndicador.Caption = "INSERTAR"
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data2

    DataGrid1.Enabled = False
    CargaTxtAux True, True
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo 1 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), True, True
    PonerFoco Text1(1)
End Sub


Private Sub BotonModificarLinea(indice As Integer)
Dim I As Integer

    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.RecordCount < 1 Then Exit Sub


    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub
    
    ModificaLineas = 2 'Modificar
    PonerModo 5

    Screen.MousePointer = vbHourglass
    
'    PonerBotonCabecera False
'    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    cmdAceptar.Tag = Data2.Recordset!numlinea
    
    CargaTxtAux True, False
    DataGrid1.Enabled = False
    PonerFoco txtAux(2) 'Poner el foco
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonEliminar()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    Sql = "Cabecera de Movimiento Almacen." & vbCrLf
    Sql = Sql & "----------------------------------------" & vbCrLf & vbCrLf
    
    Sql = Sql & "Va a eliminar el Movimiento:"
    Sql = Sql & vbCrLf & " N� Movim. : " & Text1(0).Text
    Sql = Sql & vbCrLf & " Fecha Mov.: " & CStr(Data1.Recordset.Fields(1))
    Sql = Sql & vbCrLf & " Almacen   : " & Text1(2).Text
    Sql = Sql & vbCrLf & vbCrLf & " �Desea continuar ? "
    
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub
    
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        NumRegElim = Data1.Recordset.Fields(0)
        vTipoMov.DevolverContador CodTipoMov, NumRegElim
        Set vTipoMov = Nothing
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        DataGrid1.Enabled = False
        
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
        End If
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Movimiento", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String
On Error GoTo FinEliminar
        
        conn.BeginTrans
        Sql = " WHERE  codmovim=" & Data1.Recordset!codMovim
        
        'Lineas
        conn.Execute "Delete  from slimov " & Sql
        
        'Cabeceras
        conn.Execute "Delete  from scamov " & Sql
                      
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


Private Sub BotonEliminarLinea(indice As Integer)
Dim Sql As String
On Error GoTo Error2
    
    'Ciertas comprobaciones
    If Data2.Recordset.EOF Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    '### a mano
    Sql = "Seguro que desea eliminar la l�nea del Art�culo:"
    Sql = Sql & vbCrLf & "C�digo: " & Data2.Recordset!codArtic
    Sql = Sql & vbCrLf & "Descripci�n: " & Data2.Recordset.Fields(3)
    
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        Sql = "Delete from slimov where codmovim=" & Data2.Recordset!codMovim
        Sql = Sql & " and numlinea=" & Data2.Recordset!numlinea
        Sql = Sql & " and codartic=" & DBSet(Data2.Recordset!codArtic, "T")
        conn.Execute Sql
        CancelaADODC Me.Data2
        CargaGrid True
        CancelaADODC Me.Data2
    End If
    ModificaLineas = 0
    
Error2:
    Screen.MousePointer = vbDefault
    ModificaLineas = 0
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar L�nea de Art�culo de Movimiento Almacen", Err.Description
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim vStock As String
'Dim vstockOrig As Single  'Stock en el almacen Origen
'Dim SQL As String, devuelve As String

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    

    'Comprobar que todos los Art�culos estan en el nuevo almacen
    If Modo = 4 Then 'Modificando
        b = ComprobarStocksLineas
    End If

    DatosOk = True
End Function



Private Function ComprobarStocksLineas() As Boolean
'Comprobar para todas las lineas del traspaso que:
' - todos los Art�culos entan en el almacen origen
' - Comprobar que hay suficiente stock en el Almacen Origen de ese Articulo
Dim b As Boolean

    If Not Data2.Recordset.EOF Then  'Si hay lineas
        Data2.Recordset.MoveFirst
        b = True
        
        While Not Data2.Recordset.EOF And b
            If Data2.Recordset!tipomovi = "S" Then 'Mov. de salida
                b = ComprobarStock(Data2.Recordset!codArtic, Text1(2).Text, Data2.Recordset!Cantidad, CodTipoMov)
            End If
            Data2.Recordset.MoveNext
        Wend
        Data2.Recordset.MoveFirst
    End If
    ComprobarStocksLineas = b
End Function




Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim devuelve As String

    DatosOkLinea = False
    b = True
        
    If txtAux(0).Text = "" Then
        MsgBox "El campo Cod. Art�culo no puede ser nulo", vbExclamation
        b = False
        Exit Function
    End If
        
    'Comprobamos el campo Cantidad
    If txtAux(2).Text = "" Then
         MsgBox "El campo Cantidad no puede ser nulo", vbExclamation, "Art�culos"
         b = False
    ElseIf Not IsNumeric(txtAux(2).Text) Then
        MsgBox "El campo Cantidad debe ser num�rico", vbExclamation
        b = False
    End If
    If Not b Then
        PonerFoco txtAux(2)
        Exit Function
    End If
     
    'Comprobamos si ya existe una linea con el art�culo, solo si estamos insertando (ModificaLineas=1)
    'BD 1: conexion a BD Aritaxi
    If ModificaLineas = 1 Then
        devuelve = DevuelveDesdeBDNew(conAri, "slimov", "codmovim", "codmovim", Text1(0).Text, "N", , "codartic", txtAux(0).Text, "T")
        If devuelve <> "" Then
            b = False
            devuelve = "Ya hay una l�nea con ese Art�culo: " & vbCrLf
            devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
            devuelve = devuelve & "Descripci�n: " & txtAux(1).Text
            MsgBox devuelve, vbExclamation
        End If
        
        'Comprobamos si existe el art�culo, solo si estamos insertando (ModificaLineas=1)
        If Trim(txtAux(1).Text) = "" Then
            b = False
            devuelve = "No existe el Art�culo " & vbCrLf
            devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
            devuelve = devuelve & "Descripci�n: " & txtAux(1).Text
            MsgBox devuelve, vbExclamation
        End If
    End If
    If Not b Then Exit Function
    
    'Comprobar que hay suficiente stock en el Almacen
    'Si es movimiento de Salida
    If Me.cboAux.ListIndex = 0 Then
        b = ComprobarStock(txtAux(0).Text, Text1(2).Text, txtAux(2).Text, CodTipoMov)
    End If
    DatosOkLinea = b
End Function


Private Sub PonerBotonCabecera(b As Boolean)
On Error Resume Next
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Lineas Detalle"
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.lblIndicador.Caption = ""
    End If
    'Habilitar las opciones correctas del menu seg�n Modo
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu seg�n Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function InsertarModificarLinea() As Boolean
Dim Sql As String, cad As String
On Error GoTo EInsertarModificarLinea
    
    Sql = ""
    InsertarModificarLinea = False
    
    Select Case ModificaLineas
    Case 1 'Insertar
        If DatosOkLinea Then 'INSERTAR
            Sql = "INSERT INTO slimov (codmovim,numlinea,codartic,cantidad,tipomovi,motimovi) "
            Sql = Sql & " VALUES (" & Val(Text1(0).Text) & ", "
            Sql = Sql & cmdAceptar.Tag & ", "
            Sql = Sql & DBSet(txtAux(0).Text, "T") & ", "
            Sql = Sql & DBSet(txtAux(2).Text, "N") & ", "
            If cboAux.ListIndex = -1 Then
                cad = ValorNulo
            Else
                 cad = cboAux.ItemData(cboAux.ListIndex)
            End If
            Sql = Sql & CSng(cad) & ","
            Sql = Sql & DBSet(txtAux(3).Text, "T") & ") "
        End If
    Case 2 'Modificar
        If DatosOkLinea Then
            Sql = "UPDATE slimov Set cantidad = " & DBSet(txtAux(2).Text, "N")
            Sql = Sql & ", tipomovi = " & cboAux.ItemData(cboAux.ListIndex)
            Sql = Sql & ", motimovi = " & DBSet(txtAux(3).Text, "T")
            Sql = Sql & " WHERE codmovim =" & Val(Text1(0).Text) & " AND "
            Sql = Sql & " numlinea =" & Val(cmdAceptar.Tag)
        End If
    End Select
            
    If Sql <> "" Then
        conn.Execute Sql
        InsertarModificarLinea = True
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Lineas Traspaso Almacenes" & vbCrLf & Err.Description
End Function


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

'    'Llamamos a al form
'    cad = ""
'    'Registro de la tabla de cabeceras: scamov
'    cad = cad & ParaGrid(Text1(0), 15, "N� Mov.")
'    cad = cad & ParaGrid(Text1(1), 20, "Fecha")
'    cad = cad & ParaGrid(Text1(2), 10, "Alm.")
'    cad = cad & "Desc. Alm. Orig|salmpr|nomalmac|T||40�"
'    Tabla = "(" & NombreTabla & " LEFT JOIN salmpr ON " & NombreTabla & ".codalmac=salmpr.codalmac" & ") "
'    Titulo = Me.Caption
'
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = Tabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 0
'        frmB.vConexionGrid = conAri 'Conexion a BD Aritaxi
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
''        Else   'de ha devuelto datos, es decir NO ha devuelto datos
''            If Modo = 5 Then
''                PonerFoco txtAux(0)
''            Else
'                PonerFoco Text1(kCampo)
''            End If
'        End If
'    End If
'    Screen.MousePointer = vbDefault

    
    Set frmAlmMov = New frmAlmMovimientosPrev

    frmAlmMov.EsHistorico = EsHistorico
    frmAlmMov.DatosADevolverBusqueda = "0|"
    frmAlmMov.Show vbModal
    
    Set frmAlmMov = Nothing

End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    cadSeleccion = ObtenerBusqueda(Me, True) 'Para la consulta de report

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        MsgBox "Introducir criterios de b�squeda", vbExclamation
        PonerFoco Text1(0)
    End If
    
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    Text2(0).Text = PonerNombreDeCod(Text1(2), conAri, "salmpr", "nomalmac")
    Text2(1).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Function ActualizarStocks() As Boolean
Dim Sql As String, EnAlmDest As String
Dim Cantidad As Single, vStock As Single
Dim devuelve As String
Dim vCantidad As String
    On Error GoTo EActualizarStock

    ActualizarStocks = False
    While Not Data2.Recordset.EOF
        'Actualizar el stock si el articulo tiene control de stock
        devuelve = DevuelveDesdeBDNew(conAri, "sartic", "ctrstock", "codartic", Data2.Recordset!codArtic, "T")
        If Val(devuelve) = 1 Then 'Hay control de stock

            Cantidad = Data2.Recordset!Cantidad 'Cant a traspasar
            vCantidad = TransformaComasPuntos(CStr(CCur(Cantidad)))
            If Data2.Recordset!tipomovi = "E" Then 'Mov. de Entrada
                '==== Aumentar el stock en el Almacen
                'Comprobar que existe el articulo en Almacen Destino
                EnAlmDest = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", Data2.Recordset!codArtic, "T", , "codalmac", Text1(2).Text, "N")
                If EnAlmDest = "" Then 'No hay de ese art�culo en Almacen
                    Sql = "INSERT INTO salmac (codartic,codalmac,ubialmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin)"
                    Sql = Sql & " VALUES (" & DBSet(Data2.Recordset!codArtic, "T") & "," & Val(Text1(2).Text) & ",''," & DBSet(Cantidad, "N") & ",0,0,0,0,NULL,NULL,0)"
                Else 'Existe el artic en almac. Dest -> Aumentar stock
                    Sql = "UPDATE salmac Set canstock = canstock + " & vCantidad
                    Sql = Sql & " WHERE codartic =" & DBSet(Data2.Recordset!codArtic, "T") & " AND "
                    Sql = Sql & " codalmac =" & Data1.Recordset!codAlmac
                End If
                
            Else 'Mov. de Salida
                '==== Disminuir Stock en Almacen Origen
                EnAlmDest = DevuelveDesdeBDNew(conAri, "salmac", "canstock", "codartic", Data2.Recordset!codArtic, "T", , "codalmac", Text1(2).Text, "N")
                If EnAlmDest = "" Then 'No hay de ese art�culo en Almacen
                    devuelve = "No existe en el Almacen: " & Data1.Recordset!codAlmac & vbCrLf
                    devuelve = devuelve & "El Art�culo: " & Data2.Recordset!codArtic
                    MsgBox devuelve, vbExclamation
                Else 'Existe el artic en almac. Dest -> Disminuir stock
                    vStock = CSng(EnAlmDest)
                    If ComprobarHayStock(vStock, Cantidad, Data2.Recordset!codArtic, Data2.Recordset!NomArtic, CodTipoMov) Then
                        Sql = "UPDATE salmac Set canstock = canstock - " & vCantidad
                        Sql = Sql & " WHERE codartic =" & DBSet(Data2.Recordset!codArtic, "T") & " AND "
                        Sql = Sql & " codalmac =" & Data1.Recordset!codAlmac
                    End If
                End If
            End If
            
            conn.Execute Sql
        End If
        Data2.Recordset.MoveNext
    Wend
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        ActualizarStocks = False
    Else
        ActualizarStocks = True
    End If
EActualizarStock:
End Function


Private Sub BotonActualizar()
'Actualizar Traspaso Almacen
Dim Sql As String

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n Movimiento para actualizar.", vbExclamation
        Exit Sub
    End If
    
    If Data2 Is Nothing Then Exit Sub
    If Data2.Recordset.EOF Then
        MsgBox "No hay lineas insertadas para este N� de Movimiento", vbExclamation
        Exit Sub
    End If
    
    Sql = "Actualizaci�n Movimientos Almacen." & vbCrLf
    Sql = Sql & "-------------------------------------------" & vbCrLf & vbCrLf

    If Not CBool(Data1.Recordset.Fields(5).Value) Then 'Informe No Impreso
        Sql = Sql & "NO ESTA IMPRESO EL MOVIMIENTO:" & vbCrLf
    End If
    Sql = Sql & vbCrLf & "N� Movim. : " & Format(Data1.Recordset.Fields(0), "0000000")
    Sql = Sql & vbCrLf & "Fecha        : " & CStr(Data1.Recordset.Fields(2))
    Sql = Sql & vbCrLf & "Almacen    : " & Format(Data1.Recordset.Fields(1), "000") & " - " & Text2(0).Text
    Sql = Sql & vbCrLf & vbCrLf & " �Desea continuar ? "
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
'    Else 'Informe Impreso
'        SQL = "Actualizaci�n Movimientos Almacen." & vbCrLf
'        SQL = SQL & "--------------------------------------------" & vbCrLf & vbCrLf
'
'        SQL = SQL & "Va a Actualizar el Movimiento:"
'        SQL = SQL & vbCrLf & " N� Movim.  :  " & Format(Data1.Recordset.Fields(0), "0000000")
'        SQL = SQL & vbCrLf & " Fecha Mov.:  " & CStr(Data1.Recordset.Fields(2))
'        SQL = SQL & vbCrLf & " Almacen     :  " & CStr(Format(Data1.Recordset.Fields(1), "000"))
'        SQL = SQL & vbCrLf & vbCrLf & " �Desea continuar ? "
'        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then
'            Exit Sub
'        End If
'    End If
    
    Me.ProgressBar1.visible = True
    Me.ProgressBar1.Value = 0
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    If ActualizarTraspaso Then
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
            PonerModo 2
        Else 'Solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
            Espera 0.3
            Me.Refresh
        End If
    
    End If
    Me.ProgressBar1.visible = False
End Sub


Private Function ActualizarTraspaso() As Boolean
Dim Donde As String
Dim devuelve As String
Dim bol As Boolean
On Error GoTo EActualizarTraspaso
    
    'Comprobamos que no existe en historico
    devuelve = DevuelveDesdeBDNew(conAri, "schmov", "codmovim", "codmovim", Data1.Recordset!codMovim, "N", , "fecmovim", Data1.Recordset!fecmovim, "F")
    If Trim(devuelve) <> "" Then
        devuelve = "Ya existe en el hist�rico el movimiento:" & vbCrLf
        devuelve = devuelve & " N�: " & Data1.Recordset!codMovim & vbCrLf
        devuelve = devuelve & " Fecha: " & Data1.Recordset!fecmovim
        MsgBox devuelve, vbExclamation
        Exit Function
    End If
    
    If Not ComprobarStocksLineas Then Exit Function
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    Donde = ""
    bol = ActualizarElTraspaso(Donde)

EActualizarTraspaso:
    If Err.Number <> 0 Or Donde <> "" Then
        devuelve = "Actualizar Movimiento." & vbCrLf & "----------------------------" & vbCrLf
        devuelve = devuelve & Donde
        MuestraError Err.Number, devuelve, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ActualizarTraspaso = True
    Else
        conn.RollbackTrans
        MuestraError Err.Number, devuelve, Err.Description
    End If
End Function


Private Function ActualizarElTraspaso(ByRef ADonde As String) As Boolean

    ActualizarElTraspaso = False
    
    'Insertamos en cabeceras Historico
    ADonde = "Insertando datos en historico cabeceras movimientos almacen"
    If Not InsertarCabeceraHistorico Then Exit Function
    IncrementarProgres 2
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Historico lineas Movimientos Almacen"
    If Not InsertarLineasHistorico Then Exit Function
    IncrementarProgres 2
    
    
     'Modificar stock
    ADonde = "Actualizando Stocks Almacenes"
    If Not ActualizarStocks() Then Exit Function
    IncrementarProgres 2
    
    
    'Insertamos en Movimientos Art�culos
    ADonde = "Insertando datos en Movimientos de Articulos"
    If Not InsertarMovimArticulos Then Exit Function
    IncrementarProgres 2
   
    
    'Borramos cabeceras y lineas del asiento
    ADonde = "Borrar cabeceras y lineas en Movimientos Almacen"
    If Not BorrarTraspaso(False) Then Exit Function
    IncrementarProgres 2
    
    ActualizarElTraspaso = True
    ADonde = ""
End Function


Private Function InsertarCabeceraHistorico() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
On Error GoTo EInsertarCab

    Sql = "SELECT codmovim,codalmac,fecmovim,codtraba,observa1 from scamov where "
    Sql = Sql & " codmovim =" & Data1.Recordset!codMovim
    Sql = Sql & " AND fecmovim='" & Format(Data1.Recordset!fecmovim, "yyyy-mm-dd") & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        Sql = "INSERT INTO schmov (codmovim, fecmovim,hormovim,codalmac,codtraba,observa1) "
        Sql = Sql & " VALUES (" & Rs.Fields(0).Value & ", '" & Format(Rs.Fields(2).Value, "yyyy-mm-dd") & "','"
        Sql = Sql & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & Rs.Fields(1).Value & ", " & Rs.Fields(3).Value
        Sql = Sql & ", " & DBSet(Rs.Fields(4).Value, "T") & ")"
    End If
    Rs.Close
    Set Rs = Nothing
    conn.Execute Sql
   
EInsertarCab:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarCabeceraHistorico = False
    Else
        InsertarCabeceraHistorico = True
    End If
End Function


Private Function InsertarLineasHistorico() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
On Error GoTo EInsertarLineas

    Sql = "SELECT codmovim, numlinea, codartic, cantidad, tipomovi, motimovi from slimov where "
    Sql = Sql & " codmovim =" & Data1.Recordset!codMovim
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Rs.MoveFirst
    While Not Rs.EOF
        Sql = "INSERT INTO slhmov (codmovim, fecmovim, numlinea, codartic, cantidad, tipomovi, motimovi)"
        Sql = Sql & " VALUES (" & Rs.Fields(0).Value & ", '" & Format(Data1.Recordset!fecmovim, "yyyy-mm-dd") & "', "
        Sql = Sql & Rs.Fields(1).Value & ", " & DBSet(Rs.Fields(2).Value, "T") & ", "
        Sql = Sql & DBSet(Rs.Fields(3).Value, "N") & ", " & Rs.Fields(4).Value
        Sql = Sql & ", '" & Rs.Fields(5).Value & "')"
        conn.Execute Sql
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
EInsertarLineas:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        Rs.Close
        Set Rs = Nothing
        InsertarLineasHistorico = False
    Else
        InsertarLineasHistorico = True
    End If
End Function


Private Function InsertarMovimArticulos() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim vImporte As Single, vPrecioVenta As String
Dim vTipoMov As CTiposMov
Dim bol As Boolean
Dim cad As String
On Error GoTo EInsertar

    bol = True
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        'Se han cargado correctamente los valores de la clase
        Sql = "SELECT scamov.codmovim, codalmac, fecmovim, codtraba, numlinea, codartic, cantidad, tipomovi "
        Sql = Sql & " from scamov LEFT JOIN slimov on scamov.codmovim=slimov.codmovim "
        Sql = Sql & " WHERE scamov.codmovim =" & Data1.Recordset!codMovim
        Sql = Sql & " AND fecmovim='" & Format(Data1.Recordset!fecmovim, "yyyy-mm-dd") & "'"
    
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not Rs.EOF
            'Obtener el precio de venta del articulo, si tiene control de stock
            cad = "ctrstock"
            vPrecioVenta = DevuelveDesdeBDNew(conAri, "sartic", "preciomp", "codartic", Rs.Fields!codArtic, "T", cad)
            If vPrecioVenta <> "" Then
                vImporte = Rs.Fields!Cantidad * CSng(vPrecioVenta)
            Else
                vImporte = 0
            End If
            If Val(cad) = 1 Then
                Sql = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
                Sql = Sql & " VALUES (" & DBSet(Rs.Fields!codArtic, "T") & ", " & Rs.Fields!codAlmac & ", '" & Format(Rs.Fields!fecmovim, "yyyy-mm-dd") & "', '"
                Sql = Sql & Format(Rs.Fields!fecmovim & " " & Time, "yyyy-mm-dd hh:mm:ss") & "', " & Rs.Fields!tipomovi & ", '" & vTipoMov.TipoMovimiento & "', " & DBSet(Rs.Fields!Cantidad, "N") & ", " & DBSet(vImporte, "N") & ", " & Rs.Fields!CodTraba & ", '"
                Sql = Sql & vTipoMov.LetraSerie & "', " & Rs.Fields!codMovim & ", " & Rs.Fields!numlinea & ")"
                conn.Execute Sql
            End If
            Rs.MoveNext
        Wend
    Else
        bol = False
    End If
    Set vTipoMov = Nothing
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        Set vTipoMov = Nothing
        Rs.Close
        Set Rs = Nothing
    End If
    If Err.Number <> 0 Or Not bol Then
         'Hay error , almacenamos y salimos
        InsertarMovimArticulos = False
    Else
        InsertarMovimArticulos = True
    End If
End Function



Private Sub IncrementarProgres(Veces As Integer)
On Error Resume Next
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + (Veces * 10)
    If Err.Number <> 0 Then Err.Clear
    Me.Refresh
End Sub


Private Function BorrarTraspaso(EnHistorico As Boolean) As Boolean
'Si EnHistorico=true borra de las tablas de historico: "schtra" y "slhtra"
'Si EnHistorico=false borra de las tablas de traspaso: "scatra" y "slitra"
Dim Sql As String

    BorrarTraspaso = False
    
    'Borramos las lineas
    Sql = "Delete from "
    If EnHistorico Then
        Sql = Sql & "slhmov"
        Sql = Sql & " WHERE codmovim = " & Data1.Recordset!codMovim
        Sql = Sql & " AND fecmovim = '" & Data1.Recordset!fecmovim & "'"
    Else
        Sql = Sql & "slimov"
        Sql = Sql & " WHERE codmovim = " & Data1.Recordset!codMovim
    End If
    conn.Execute Sql
    
    'La cabecera
    Sql = "Delete from "
    If EnHistorico Then
        Sql = Sql & "schmov"
        Sql = Sql & " WHERE codmovim =" & Data1.Recordset!codMovim
        Sql = Sql & " AND fecmovim='" & Data1.Recordset!fecmovim & "'"
    Else
        Sql = Sql & "scamov"
        Sql = Sql & " WHERE codmovim =" & Data1.Recordset!codMovim
    End If
    conn.Execute Sql
    
    If Err.Number <> 0 Then
        BorrarTraspaso = False
    Else
        BorrarTraspaso = True
    End If
End Function


Private Sub CargarComboAux()
'### Combo Tipo Movimiento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Entrada, 1-Salida

    cboAux.Clear
    cboAux.AddItem "S"
    cboAux.ItemData(cboAux.NewIndex) = 0
    
    cboAux.AddItem "E"
    cboAux.ItemData(cboAux.NewIndex) = 1
        
End Sub


Public Sub ActualizarSituacionImpresion()
Dim cad As String, Indicador As String
On Error GoTo EImpresion
   
    cad = "(" & ObtenerWhereCP(False) & ")"
    If SituarData(Data1, cad, Indicador) Then
        If Modo <> 5 Then
            PonerModo 2
        Else
            PonerModo 5
        End If
        PonerCampos
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
EImpresion:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonImprimir()
        If Text1(0).Text = "" Then Exit Sub
        frmListado.NumCod = Text1(0).Text
        If Not EsHistorico Then
            AbrirListado (8) '8: Informe Movimientos Almacen
            ActualizarSituacionImpresion
        Else
            BotonImprimirHco
        End If
End Sub


Private Sub BotonImprimirHco()
Dim indRPT As Byte
Dim cadParam As String
Dim cad As String
Dim numParam As Byte
Dim nomDocu As String


    cadParam = "|"
    numParam = 0
    If Not PonerParamEmpresa(cadParam, numParam) Then Exit Sub

    indRPT = 4 '4: Historico Movimientos de Almacen
    If PonerParamRPT(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt) Then
        With frmImprimir
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .NombreRPT = nomDocu
            .NombrePDF = pPdfRpt
            .EnvioEMail = False
            .Opcion = 8
            .Titulo = "Hist. Movimientos Alm."
            If cadSeleccion <> "" Then
                .FormulaSeleccion = cadSeleccion
            Else
                'Se Llama desde dobleclick en frmAlmMovimArticulos
                cad = "{schmov.codmovim}= " & Data1.Recordset!codMovim
                cad = cad & " and {schmov.fecmovim}= Date(" & Year(Data1.Recordset!fecmovim) & "," & Month(Data1.Recordset!fecmovim) & "," & Day(Data1.Recordset!fecmovim) & ")" & ""
                .FormulaSeleccion = cad
            End If
            .Show vbModal
        End With
    End If
End Sub



Private Function InsertarMovimiento(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean
On Error GoTo EInsertarMovim
    
    bol = True
    
    'Aqui empieza transaccion
    conn.BeginTrans
    
    MenError = "Error al insertar en la tabla de Movimientos(smovim)."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador del recibo."
    bol = vTipoMov.IncrementarContador(CodTipoMov)

EInsertarMovim:
        If Err.Number <> 0 Then
            MenError = "Insertando Movimiento." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarMovimiento = True
        Else
            conn.RollbackTrans
            InsertarMovimiento = False
        End If
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
'Obtiene la sentencia WHERE para seleccionar registros de la tabla por Clave Primaria
On Error Resume Next
    If conWhere Then
        ObtenerWhereCP = " WHERE codmovim= " & Val(Text1(0).Text)
    Else
        ObtenerWhereCP = " codmovim= " & Val(Text1(0).Text)
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    Set vTipoMov = New CTiposMov
    
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Text1(0).Text = Format(Text1(0).Text, "0000000")
        cmdCancelar.Caption = "Cancelar"
        Sql = CadenaInsertarDesdeForm(Me)
        
        If Sql <> "" Then
            If InsertarMovimiento(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                 'Ponerse en Modo Insertar Lineas
                BotonLineas
                BotonAnyadirLineas 0
            End If
        End If
    End If
    Set vTipoMov = Nothing
End Sub



