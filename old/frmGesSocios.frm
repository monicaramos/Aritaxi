VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGesSocios 
   Caption         =   "Socios."
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   4200
      Top             =   7680
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   5520
      Top             =   7800
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   240
      TabIndex        =   48
      Top             =   4140
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   -2147483624
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "frmGesSocios.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Contadores"
      TabPicture(1)   =   "frmGesSocios.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Choferes"
      TabPicture(2)   =   "frmGesSocios.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtAux1(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtAux1(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtAux1(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtAux1(3)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtAux1(4)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdAux(0)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Publicidad"
      TabPicture(3)   =   "frmGesSocios.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "DataGrid2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "txtAux2(0)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "txtAux2(1)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "txtAux2(2)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "txtAux2(3)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdAux1"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtAux2(4)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cmdAux(1)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      Begin VB.Frame Frame4 
         Caption         =   "Datos Bancarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -68190
         TabIndex        =   29
         Top             =   510
         Width           =   4065
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   22
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   25
            Tag             =   "Cuenta Banco|T|N|||sclien|cuentaba|0000000000||"
            Text            =   "9999999999"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   21
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   24
            Tag             =   "Digito Control|T|N|||sclien|digcontr|00||"
            Text            =   "99"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   20
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   23
            Tag             =   "Codigo Sucursal|N|N|||sclien|codsucur|0000||"
            Text            =   "9999"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   19
            Left            =   840
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "Codigo Banco|N|N|||sclien|codbanco|0000||"
            Text            =   "9999"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "C.C.C."
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   70
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos del Vehículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74760
         TabIndex        =   63
         Top             =   510
         Width           =   6495
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   24
            Left            =   4800
            MaxLength       =   9
            TabIndex        =   18
            Tag             =   "Poliza|T|S|||sclien|numpoliza|||"
            Text            =   "Text"
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2280
            TabIndex        =   73
            Text            =   "Text2"
            Top             =   360
            Width           =   3975
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   11
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   21
            Tag             =   "Codigo Situación|N|N|||sclien|codsitua|00||"
            Text            =   "Tex"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   1320
            MaxLength       =   9
            TabIndex        =   19
            Tag             =   "Licencia|N|S|||sclien|licencia|00000000||"
            Text            =   "Text"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   67
            Text            =   "Text2"
            Top             =   1800
            Width           =   4215
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   18
            Left            =   4800
            MaxLength       =   10
            TabIndex        =   20
            Tag             =   "Fecha Situación|F|S|||sclien|fechasit||dd/mm/yyyy|"
            Text            =   "99/99/9999"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   15
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   17
            Tag             =   "Matricula|T|S|||sclien|matricul|||"
            Text            =   "Text"
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   16
            Tag             =   "Codigo Coche|N|N|||sclien|codcoche|0000||"
            Text            =   "Text"
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label1 
            Caption         =   "Número de Poliza:"
            Height          =   255
            Index           =   23
            Left            =   3120
            TabIndex        =   74
            Top             =   840
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   960
            Picture         =   "frmGesSocios.frx":0070
            Tag             =   "-1"
            ToolTipText     =   "Buscar vehiculo"
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   960
            Picture         =   "frmGesSocios.frx":0172
            Tag             =   "-1"
            ToolTipText     =   "Buscar situación"
            Top             =   1800
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Situación:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   69
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Licencia:"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   68
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   4440
            Picture         =   "frmGesSocios.frx":0274
            ToolTipText     =   "Buscar fecha"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Situación:"
            Height          =   255
            Index           =   18
            Left            =   3120
            TabIndex        =   66
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Matricula:"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   65
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Vehiculo"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdAux 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   55
         ToolTipText     =   "Buscar cliente"
         Top             =   2520
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   4440
         TabIndex        =   59
         Text            =   "hasta"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cmdAux1 
         Height          =   315
         ItemData        =   "frmGesSocios.frx":02FF
         Left            =   7680
         List            =   "frmGesSocios.frx":0309
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2520
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   58
         Text            =   "desde"
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   57
         Text            =   "importe"
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   62
         Text            =   "nomcliente"
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   360
         MaxLength       =   6
         TabIndex        =   56
         Text            =   "Codclien"
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdAux 
         Caption         =   "+"
         Height          =   240
         Index           =   0
         Left            =   -73500
         TabIndex        =   50
         ToolTipText     =   "Buscar chofer"
         Top             =   2460
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   -69240
         MaxLength       =   40
         TabIndex        =   54
         Tag             =   "Observaciones|T|S|||ssocio_chofer|observac|||"
         Text            =   "Observaciones"
         Top             =   2460
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   -70320
         MaxLength       =   10
         TabIndex        =   53
         Tag             =   "Fecha Baja|F|S|||ssocio_chofer|fechabaj|||"
         Text            =   "FEcBaja"
         Top             =   2460
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -71640
         MaxLength       =   10
         TabIndex        =   52
         Tag             =   "Fecha Alta|F|S|||ssocio_chofer|fechaalt|||"
         Text            =   "FecAlta"
         Top             =   2460
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73680
         TabIndex        =   49
         Text            =   "nomchofe"
         Top             =   2460
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74760
         MaxLength       =   5
         TabIndex        =   51
         Tag             =   "Chofer|N|N|||ssocio_chofer|codchofe|||"
         Text            =   "chofe"
         Top             =   2460
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   71
         Top             =   720
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3201
         _Version        =   393216
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2010
         Left            =   240
         TabIndex        =   61
         Top             =   720
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   3545
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmGesSocios.frx":0320
         Height          =   2010
         Left            =   -74760
         TabIndex        =   75
         Top             =   720
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   3545
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
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6360
      Top             =   7440
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5280
      Top             =   7440
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
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8940
      TabIndex        =   26
      Top             =   7560
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10260
      TabIndex        =   27
      Top             =   7560
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10260
      TabIndex        =   28
      Top             =   7560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   240
      TabIndex        =   31
      Top             =   720
      Width           =   11055
      Begin VB.CheckBox Check1 
         Caption         =   "Es Socio"
         Height          =   375
         Left            =   4650
         TabIndex        =   14
         Tag             =   "Facturado|N|N|0|1|sclien|essocio|||"
         Top             =   2700
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   555
         Index           =   17
         Left            =   6480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Tag             =   "Observaciones|T|S|||sclien|observac|||"
         Text            =   "frmGesSocios.frx":0335
         Top             =   2520
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   13
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Fecha Baja|F|S|||sclien|fechabaj||dd/mm/yyyy|"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   12
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Fecha Alta|F|N|||sclien|fechaalt||dd/mm/yyyy|"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   10
         Left            =   7320
         MaxLength       =   40
         TabIndex        =   11
         Tag             =   "Mail|T|S|||sclien|maiclie1|||"
         Text            =   "Text"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Movil|T|S|||sclien|movclien|||"
         Text            =   "Text"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Telefono|T|S|||sclien|telclie1|||"
         Text            =   "963577679"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   13
         Tag             =   "CIF|T|N|||sclien|nifclien|||"
         Text            =   "Text"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   12
         Tag             =   "Provincia|T|N|||sclien|proclien|||"
         Text            =   "Text"
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   3480
         MaxLength       =   35
         TabIndex        =   10
         Tag             =   "Población|T|N|||sclien|pobclien|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   9
         Tag             =   "CP|T|N|||sclien|codpobla|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   7
         Tag             =   "Domicilio|T|N|||sclien|domclien|||"
         Text            =   "Text"
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   2
         Tag             =   "Nombre|T|N|||sclien|nomclien|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   1
         Left            =   9840
         MaxLength       =   5
         TabIndex        =   3
         Tag             =   "Num.Vehiculo|N|S|||sclien|numeruve|0000||"
         Text            =   "Text"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   1
         Tag             =   "Codigo Socio|N|N|||sclien|codclien|0000|S|"
         Text            =   "Text"
         Top             =   360
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   23
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   72
         Tag             =   "Codigo Socio|N|N|||sclien|codtarif|||"
         Text            =   "Text"
         Top             =   840
         Width           =   870
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4170
         Picture         =   "frmGesSocios.frx":033A
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Baja:"
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   47
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmGesSocios.frx":03C5
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Alta:"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   46
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   45
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   1
         Left            =   6960
         Tag             =   "-1"
         ToolTipText     =   "Enviar e-mail"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   7680
         Picture         =   "frmGesSocios.frx":0450
         Tag             =   "-1"
         ToolTipText     =   "Ver observaciones"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "E-mail"
         Height          =   255
         Index           =   11
         Left            =   6480
         TabIndex        =   44
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Movil:"
         Height          =   255
         Index           =   10
         Left            =   6480
         TabIndex        =   43
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono:"
         Height          =   255
         Index           =   9
         Left            =   6480
         TabIndex        =   42
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "CIF:"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   41
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   40
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Población:"
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   39
         Top             =   1800
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1200
         Picture         =   "frmGesSocios.frx":0552
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "CP:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   38
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones:"
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   34
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Número de Vehiculo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   7320
         TabIndex        =   33
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
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
            Object.ToolTipText     =   "Choferes"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Publicidad"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   30
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   240
      TabIndex        =   32
      Top             =   7440
      Width           =   3975
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnbuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnvertodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnbarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmGesSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Public WithEvents frmCP As frmCPostal
Attribute frmCP.VB_VarHelpID = -1
Public WithEvents frmS As frmFacSituaciones
Attribute frmS.VB_VarHelpID = -1
Public WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Public WithEvents frmCond As frmGesConduc
Attribute frmCond.VB_VarHelpID = -1
Public WithEvents frmBanco As frmFacBancosPropios
Attribute frmBanco.VB_VarHelpID = -1
Public WithEvents frmCal As frmCal
Attribute frmCal.VB_VarHelpID = -1
Public WithEvents frmV As frmGesVehic
Attribute frmV.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim btnAnyadir As Byte
Dim btnPrimero As Byte
Dim NombreTabla As String
Dim Ordenacion As String
Dim CadenaConsulta As String
Dim HaDevueltoDatos As Boolean
Private Modo As Byte
Dim kCampo As Byte
Dim ModificaLineas As Byte
Dim Fecha As Date
Dim Situacion As Boolean
'-------------------------------------------------------
'Se distinguen varios MODOS
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------
Private VieneDeBuscar As Boolean

Dim cadB1 As String

Dim BuscaChekc As String


Private Sub Check1_Click()
    If Modo = 1 Then CheckCadenaBusqueda Check1, BuscaChekc
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cadB As String
Dim cad As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
              If InsertarDesdeForm(Me) Then
                CrearContadores
                PosicionarData
              End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
         Case 5 'INSERTAR MODIFICAR LINEA
            
            If ModificaLineas = 1 Then 'INSERTAR lineas
                If SSTab1.Tab = 2 Then
                    'choferes
                    If InsertarLinea Then
                        CargaGrid DataGrid1, Adodc2
                        BotonAnyadirLinea
                    End If
                Else
                    'publicidad
                    If InsertarLinea2 Then
                        CargaGrid DataGrid2, Adodc3
                        BotonAnyadirLinea2
                    End If
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If SSTab1.Tab = 2 Then 'chofer
                    If ModificarLinea Then
                        TerminaBloquear
                        CargaTxtAux False, False
                        CargaGrid DataGrid1, Adodc2
                        ModificaLineas = 0
                        PonerBotonCabecera True
                    End If
                Else 'publicidad
                    If ModificarLinea2 Then
                        TerminaBloquear
                        CargaTxtAux2 False, False
                        CargaGrid DataGrid2, Adodc3
                        PonerBotonCabecera True
                    End If
                End If
                Me.DataGrid1.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault
Error1:
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub CrearContadores()
'creará los contadores del socio nuevo con contadores igual a 0 con movimientos
'que en stipom tengan tipodocu=2
Dim SQL As String

On Error GoTo EContadores

Set miRsAux = New ADODB.Recordset
SQL = "select * from stipom where tipodocu=2"
miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

While Not miRsAux.EOF
    SQL = "INSERT INTO sclien_contadores (codsocio,codtipom,contador) values ("
    SQL = SQL & Text1(0).Text & "," & DBSet(miRsAux!codtipom, "T") & ",0)"
    conn.Execute SQL
    miRsAux.MoveNext
Wend

miRsAux.Close
Set miRsAux = Nothing

EContadores:
If Err.Number <> 0 Then MsgBox "Error contadores: " & Err.Description

End Sub


Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a Mantenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    
    If b Then
        Me.lblIndicador(0).Caption = "Líneas "
        PonerFocoBtn Me.cmdRegresar
    End If
    
    'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim SQL As String
On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE sclien_chofer Set codchofe = " & txtAux1(0).Text & ", fechaalt='" & Format(txtAux1(2).Text, FormatoFecha) & "', "
        SQL = SQL & "fechabaj='" & Format(txtAux1(3).Text, FormatoFecha) & "', obsevac=" & DBSet(txtAux1(4).Text, "T")
        
        SQL = SQL & " where codsocio=" & Adodc2.Recordset!codSocio & " AND numlinea=" & Adodc2.Recordset!numlinea
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Chofer" & vbCrLf & Err.Description
End Function

Private Function ModificarLinea2() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim SQL As String
On Error GoTo EModificarLinea

    ModificarLinea2 = False
    SQL = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE sclien_publicidad Set codclien = " & txtAux2(0).Text & ", importes=" & TransformaComasPuntos(txtAux2(2).Text) & ", "
        SQL = SQL & "desdefec='" & Format(txtAux2(3).Text, FormatoFecha) & "',hastafec='" & Format(txtAux2(4).Text, FormatoFecha) & "', situacio=" & cmdAux1.ItemData(cmdAux1.ListIndex)
        SQL = SQL & " where codsocio=" & Adodc3.Recordset!codSocio & " AND numlinea=" & Adodc3.Recordset!numlinea
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea2 = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Publicidad" & vbCrLf & Err.Description
End Function

Private Function DatosOkLinea() As Boolean

DatosOkLinea = False
If SSTab1.Tab = 2 Then
'chofer
    If txtAux1(0).Text = "" Then
        MsgBox "Es necesario introducir el código de chofer.", vbExclamation
        DatosOkLinea = False
        Exit Function
    End If
Else
'publicidad
    If txtAux2(0).Text = "" Then
        MsgBox "Es necesario introducir el código de cliente.", vbExclamation
        DatosOkLinea = False
        Exit Function
    End If
    If txtAux2(2).Text = "" Then
        MsgBox "Es necesario introducir el importe.", vbExclamation
        DatosOkLinea = False
        Exit Function
    End If
    If txtAux2(3).Text = "" Then
        MsgBox "Es necesario introducir la fecha desde.", vbExclamation
        DatosOkLinea = False
        Exit Function
    End If
    If txtAux2(4).Text = "" Then
        MsgBox "Es necesario introducir la fecha hasta.", vbExclamation
        DatosOkLinea = False
        Exit Function
    End If
    If cmdAux1.Text = "" Then
        MsgBox "Es necesario introducir el código de situación.", vbExclamation
        DatosOkLinea = False
        Exit Function
    End If
End If
DatosOkLinea = True

End Function
Private Function InsertarLinea() As Boolean
Dim SQL As String
Dim vWhere As String
Dim numF As String
On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""
    If DatosOkLinea Then
        vWhere = "codsocio=" & Val(Text1(0).Text)
        numF = SugerirCodigoSiguienteStr("sclien_chofer", "numlinea", vWhere)
        SQL = "INSERT INTO sclien_chofer "
        SQL = SQL & "(codsocio, numlinea, codchofe, fechaalt,fechabaj,obsevac) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
        SQL = SQL & DBSet(txtAux1(0).Text, "T") & ",'" & Format(txtAux1(2).Text, FormatoFecha) & "','"
        SQL = SQL & Format(txtAux1(3).Text, FormatoFecha) & "'," & DBSet(txtAux1(4).Text, "T") & ")"
     End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea = True
    End If
    Exit Function
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas chofer" & vbCrLf & Err.Description
End Function
Private Function InsertarLinea2() As Boolean
Dim SQL As String
Dim vWhere As String
Dim numF As String
Dim Importe As Currency
On Error GoTo EInsertarLinea2

    InsertarLinea2 = False
    SQL = ""
    If DatosOkLinea Then
        Importe = ImporteFormateado(txtAux2(2).Text)
        vWhere = "codsocio=" & Val(Text1(0).Text)
        numF = SugerirCodigoSiguienteStr("sclien_publicidad", "numlinea", vWhere)
        SQL = "INSERT INTO sclien_publicidad "
        SQL = SQL & "(codsocio, numlinea, codclien, importes,desdefec,hastafec,situacio) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
        SQL = SQL & DBSet(txtAux2(0).Text, "N") & "," & TransformaComasPuntos(CStr(Importe)) & ",'"
        SQL = SQL & Format(txtAux2(3).Text, FormatoFecha) & "','" & Format(txtAux2(4).Text, FormatoFecha) & "'," & cmdAux1.ItemData(cmdAux1.ListIndex) & ")"
     End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea2 = True
    End If
    Exit Function
EInsertarLinea2:
    MuestraError Err.Number, "Insertar Lineas publicidad" & vbCrLf & Err.Description
End Function
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
    Else
        PonerFoco Text1(0)
    End If
End Sub
Private Function DatosOk() As Boolean

DatosOk = False

If Text1(0).Text = "" Then
    MsgBox "Es necesario introducir el codigo de socio.", vbExclamation
    DatosOk = False
    Exit Function
End If
If Text1(2).Text = "" Then
    MsgBox "Es necesario introducir el nombre del socio.", vbExclamation
    DatosOk = False
    Exit Function
End If
If Text1(3).Text = "" Then
    MsgBox "Es necesario introducir el domicilio del socio.", vbExclamation
    DatosOk = False
    Exit Function
End If
If Text1(4).Text = "" Then
    MsgBox "Es necesario introducir el código de población del socio.", vbExclamation
    DatosOk = False
    Exit Function
End If
If Text1(5).Text = "" Then
    MsgBox "Es necesario introducir la población del socio.", vbExclamation
    DatosOk = False
    Exit Function
End If
If Text1(6).Text = "" Then
    MsgBox "Es necesario introducir la provincia del socio.", vbExclamation
    DatosOk = False
    Exit Function
End If
If Text1(7).Text = "" Then
    MsgBox "Es necesario introducir el CIF del socio.", vbExclamation
    DatosOk = False
    Exit Function
End If
If Text1(12).Text = "" Then
    MsgBox "Es necesario introducir la fecha de alta del socio.", vbExclamation
    DatosOk = False
    Exit Function
End If
If Text1(14).Text = "" Then
    MsgBox "Es necesario introducir el còdigo de coche.", vbExclamation
    SSTab1.Tab = 0
    PonerFoco Text1(14)
    DatosOk = False
    Exit Function
End If
If Text1(11).Text = "" Then
    MsgBox "Es necesario introducir el código de situación.", vbExclamation
    SSTab1.Tab = 0
    PonerFoco Text1(11)
    DatosOk = False
    Exit Function
End If
If Text1(19).Text = "" Then
    MsgBox "Es necesario introducir el código del banco.", vbExclamation
    SSTab1.Tab = 0
    PonerFoco Text1(19)
    DatosOk = False
    Exit Function
End If
If Text1(20).Text = "" Then
    MsgBox "Es necesario introducir el código de sucursal.", vbExclamation
    SSTab1.Tab = 0
    PonerFoco Text1(19)
    DatosOk = False
    Exit Function
End If
If Text1(21).Text = "" Then
    MsgBox "Es necesario introducir el dígito de control.", vbExclamation
    SSTab1.Tab = 0
    PonerFoco Text1(21)
    DatosOk = False
    Exit Function
End If
If Text1(22).Text = "" Then
    MsgBox "Es necesario introducir la cuenta bancaria.", vbExclamation
    SSTab1.Tab = 0
    PonerFoco Text1(22)
    DatosOk = False
    Exit Function
End If

'- Validar que la cuenta bancaria es correcta
If Not Comprueba_CuentaBan(Text1(19).Text & Text1(20).Text & Text1(21).Text & Text1(22).Text) Then
    MsgBox "La cuenta bancaria no es correcta.", vbExclamation
    DatosOk = False
    Exit Function
End If
DatosOk = True
    
End Function

Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 1
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0"
            frmCli.Show vbModal
            Set frmCli = Nothing
        Case 2
            Situacion = True
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
            frmS.Show vbModal
            Set frmS = Nothing
        Case 0
            Set frmCond = New frmGesConduc
            frmCond.DatosADevolverBusqueda = "0"
            frmCond.Show vbModal
            Set frmCond = Nothing
    End Select
    
End Sub


Private Sub cmdAux1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAux1_LostFocus()
    cmdAceptar.SetFocus
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        Me.lblIndicador(0).Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
            DeseleccionaGrid DataGrid2
        End If
        cmdRegresar.Caption = "Regresar"
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If adodc1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        
        cad = adodc1.Recordset.Fields(0) & "|"
        cad = cad & adodc1.Recordset.Fields(2) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(2)
End Sub

Private Sub Form_Load()
    'Icono del form
    Me.Icon = frmPpal.Icon
    

    'ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 14
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(9).Image = 45 'Lineas chofer
        .Buttons(10).Image = 10 'Lineas publicidad
        .Buttons(11).Image = 16 'imprimir
        .Buttons(12).Image = 15 'salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
      
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    PrimeraVez = True
    
    ImgMail(1).Picture = frmPpal.imgListComun.ListImages(20).Picture
    
    '## A mano
    NombreTabla = "sclien"
    Ordenacion = " ORDER BY codclien"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    adodc1.ConnectionString = conn
    adodc1.RecordSource = "Select * from " & NombreTabla & " where codclien=-1"
    adodc1.Refresh
    
    LimpiarDataGrids
    
    Me.SSTab1.Tab = 0
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
    End If
    
End Sub
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo


    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    
    Modo = Kmodo
    PonerIndicador lblIndicador(0), Modo
    
    
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
    If Not adodc1.Recordset.EOF Then
        If adodc1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    Me.imgBuscar(2).Enabled = (Modo <= 4 And Modo > 1)
    chkVistaPrevia.Enabled = (Modo <= 2)
    Me.Check1.Enabled = Not (Modo = 0 Or Modo = 2)
    '-----------------------------
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub
Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub
Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean
Dim i As Byte

    b = (Modo = 2 Or Modo = 5 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    
    b = (Modo = 2 Or Modo = 5)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    b = (Modo = 0 Or Modo = 2)
    'lineas de chofer
    Toolbar1.Buttons(9).Enabled = b
    'lineas de publicidad
    Toolbar1.Buttons(10).Enabled = b
    'imprimir
    Toolbar1.Buttons(11).Enabled = b
    
    '------------------------------------------
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
End Sub
Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador(0).Caption = ""
    Check1.Value = 0
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        'Estamos en Cabecera
        'Recupera todo el registro de Tarifas de Precios
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmCal_Selec(vFecha As Date)
    Fecha = vFecha
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCond_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux1(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String
    
    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'Poblacion
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)
    'provincia
    Text1(Indice + 2).Text = devuelve

End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
'Situaciones
    If Situacion Then
        txtAux2(4).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        Text1(11).Text = RecuperaValor(CadenaSeleccion, 1)
        FormateaCampo Text1(10)
        Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmV_DatoSeleccionado(CadenaSeleccion As String)
    Text1(14).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

Select Case Index
    Case 2
            If Modo = 5 Or Modo = 0 Then
            Else
                
                If Modo = 3 Or Modo = 4 Then
                    CadenaDesdeOtroForm = Text1(17).Text
                Else
                    CadenaDesdeOtroForm = ""
                    If Not adodc1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(adodc1.Recordset!observac, "T")
                End If
                frmFacClienteObser.Modificar = Modo >= 3
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 3 Or Modo = 4 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(17).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
            End If
    Case 0 'codigo postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            If Index = 9 Then
                Indice = 4
            Else
                PonerFoco Text1(4)
            End If
    Case 1  'situaciones
            Situacion = False
            Indice = 11
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
             If Not IsNumeric(Text1(Indice)) Then Text1(Indice).Text = ""
            frmS.Show vbModal
            Set frmS = Nothing
    Case 3 'coches
            Set frmV = New frmGesVehic
            frmV.DatosADevolverBusqueda = "0|1|"
            frmV.Show vbModal
            Set frmV = Nothing
    End Select
End Sub

Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
   
    Set frmCal = New frmCal
    frmCal.Fecha = Now
       
    Select Case Index
        Case 0
            Indice = 12
            PonerFormatoFecha Text1(Indice)
            If Text1(Indice).Text <> "" Then frmCal.Fecha = CDate(Text1(Indice).Text)
        Case 1
            Indice = 13
            PonerFormatoFecha Text1(Indice)
            If Text1(Indice).Text <> "" Then frmCal.Fecha = CDate(Text1(Indice).Text)
        Case 2
            Indice = 18
            PonerFormatoFecha Text1(Indice)
            If Text1(Indice).Text <> "" Then frmCal.Fecha = CDate(Text1(Indice).Text)
    End Select
    Screen.MousePointer = vbDefault
    frmCal.Show vbModal
    If IsDate(Fecha) Then
        If Fecha <> "0:00:00" Then Text1(Indice) = Fecha
    End If
    Set frmCal = Nothing
    PonerFoco Text1(Indice)
End Sub




Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub

Private Sub Text1_GotFocus(Index As Integer)
kCampo = Index
ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim encontrado As String
Dim devuelve As String

If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
Text1(Index) = UCase(Text1(Index).Text)

Select Case Index
    Case 11 'cod situacion
        If Text1(Index).Text <> "" Then
            If IsNumeric(Text1(Index).Text) Then
               Text1(Index).Text = Format(Text1(Index).Text, "00")
                encontrado = DevuelveDesdeBD(conAri, "nomsitua", "ssitua", "codsitua", Text1(Index).Text, "T")
                If encontrado <> "" Then
                    Text2(0).Text = encontrado
                Else
                    MsgBox "El código de situación introducido no existe.", vbExclamation
                    PonerFoco Text1(Index)
                End If
            Else
                MsgBox "El código de situación debe ser numérico.", vbExclamation
                PonerFoco Text1(Index)
            End If
        End If
    Case 19 'banco
'        If Text1(Index).Text <> "" Then
'            If IsNumeric(Text1(Index).Text) Then
'                encontrado = DevuelveDesdeBD(conAri, "nombanpr", "sbanpr", "codbanpr", Text1(Index).Text, "T")
'                If encontrado <> "" Then
'                    'Text2(1).Text = encontrado
'                Else
'                    MsgBox "El código de banco introducido no existe.", vbExclamation
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                MsgBox "El código de banco debe ser numérico.", vbExclamation
'                PonerFoco Text1(Index)
'            End If
'        End If
    Case 0 'codsocio
        If Text1(Index).Text <> "" Then
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "El código de socio debe ser numérico.", vbExclamation
                PonerFoco Text1(Index)
            End If
        End If
    Case 1 'numeruve
        If Text1(Index).Text <> "" Then
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "El código de vehiculo debe ser numérico.", vbExclamation
                PonerFoco Text1(Index)
            ElseIf Text1(Index).Text <= 0 Then
                MsgBox "El código de vehiculo tiene que tener un valor mayor que 0.", vbExclamation
                PonerFoco Text1(Index)
            Else
                VerificarVehiculo
            End If
        End If
    Case 12, 13, 18 'fecha alta,baja y situación
        If Modo = 1 Then Exit Sub
    
        If Text1(Index).Text <> "" Then
            PonerFormatoFecha Text1(Index)
        End If
    Case 4
        If Text1(Index).Text <> "" Then
            'Poblacion
            Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
            'provincia
            Text1(Index + 2).Text = devuelve
        End If
    Case 14
        If Text1(Index).Text <> "" Then
            encontrado = DevuelveDesdeBD(conAri, "nomchofe", "scoche", "codcoche", Text1(Index).Text, "T")
            If encontrado = "" Then
                MsgBox "El código de coche introducido no existe.", vbExclamation
                PonerFoco Text1(Index)
            Else
                Text1(Index).Text = Format(Text1(Index).Text, "0000")
                Text2(1).Text = encontrado
            End If
        End If
    Case 16
        Text1(Index).Text = Format(Text1(Index).Text, "00000000")
    Case 7 'nif
        If Text1(Index).Text <> "" Then
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
        End If
End Select
End Sub
Private Sub VerificarVehiculo()
Dim encontrado As String
Dim Cliente As String

If Text1(13).Text <> "" Then 'si esta dado de baja no hace ninguna comprobación
    Cliente = "codclien"
    encontrado = DevuelveDesdeBD(conAri, "numeruve", "sclien", "numeruve", Text1(1).Text, "T", Cliente)
    Cliente = Format(Cliente, "0000")
    If encontrado <> "" Then
        If Not Cliente = Text1(0).Text Then
            MsgBox "El código de vehiculo ingresado esta asociado a otro Socio.", vbExclamation
        End If
    End If
End If

End Sub
Private Sub LimpiarDataGrids()
Dim SQL As String
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    'SQL = "select * from sclien_chofer where codsocio=-1"
    SQL = "select sclien_chofer.codsocio,sclien_chofer.numlinea,sclien_chofer.codchofe,schofe.nomchofe,sclien_chofer.fechaalt,sclien_chofer.fechabaj,sclien_chofer.obsevac from sclien_chofer inner join schofe on sclien_chofer.codsocio= -1 and sclien_chofer.codchofe=schofe.codchofe"
    CargaGridGnral DataGrid1, Adodc2, SQL, PrimeraVez
    CargaGrid DataGrid1, Adodc2

    SQL = "select sclien_publicidad.codsocio,sclien_publicidad.numlinea,sclien_publicidad.codclien,scliente.nomclien,sclien_publicidad.importes,sclien_publicidad.desdefec,sclien_publicidad.hastafec,if (sclien_publicidad.situacio=0, ""Activo"",""No Activo"") from sclien_publicidad inner join scliente on sclien_publicidad.codsocio= -1 and sclien_publicidad.codclien=scliente.codclien"
    CargaGridGnral DataGrid2, Adodc3, SQL, PrimeraVez
    CargaGrid DataGrid2, Adodc3

    SQL = "select sclien_contadores.codsocio,sclien_contadores.codtipom,stipom.nomtipom,sclien_contadores.contador from sclien_contadores inner join stipom on sclien_contadores.codsocio=-1 and sclien_contadores.codtipom=stipom.codtipom"
    CargaGridGnral DataGrid3, Adodc4, SQL, PrimeraVez
    CargaGrid DataGrid3, Adodc4
    
'    CargaGrid DataGrid1, Adodc2, False
    
    
    If Err.Number <> 0 Then Err.Clear
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
            CargaTxtAux2 False, False
            If ModificaLineas = 1 Then 'INSERTAR
                If SSTab1.Tab = 2 Then
                    ModificaLineas = 0
                    DataGrid1.AllowAddNew = False
                    If Not Adodc2.Recordset.EOF Then Adodc2.Recordset.MoveFirst
                Else
                    ModificaLineas = 0
                    DataGrid2.AllowAddNew = False
                    If Not Adodc3.Recordset.EOF Then Adodc3.Recordset.MoveFirst
                End If
            Else
                ModificaLineas = 0
            End If
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            Me.DataGrid2.Enabled = True
            
    End Select
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub
Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
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
        Case 9 'chofer
            mnLineas_Click
        Case 10 'publicidad
            mnLineas2_Click
        Case 11  'imprimir
            printNou
        Case 12  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub
Private Sub mnLineas_Click()
    BotonMtoLineas "Choferes"
End Sub
Private Sub mnLineas2_Click()
    BotonMtoLineas "Publicidad"
End Sub
Private Sub BotonMtoLineas(cad As String)
        If cad = "Choferes" Then
            SSTab1.Tab = 2
        Else
            SSTab1.Tab = 3
        End If
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
        
End Sub
Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
        If SSTab1.Tab = 2 Then
            BotonAnyadirLinea
        Else
            BotonAnyadirLinea2
        End If
    Else 'Añadir Cabecera de Pedidos
         BotonAnyadir
    End If
End Sub
Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador(0).Caption = "INSERTAR Chofer"
    
    AnyadirLinea DataGrid1, Adodc2
    CargaTxtAux True, True
   
    PonerFoco txtAux1(0)
    Me.DataGrid1.Enabled = False
End Sub
Private Sub BotonAnyadirLinea2()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador(0).Caption = "INSERTAR Publicidad"
    
    AnyadirLinea DataGrid2, Adodc3
    CargaTxtAux2 True, True
   
    PonerFoco txtAux2(0)
    Me.DataGrid2.Enabled = False
End Sub

Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux1.Count - 1 'TextBox
            txtAux1(i).Top = 290
            txtAux1(i).visible = visible
        Next i
        Me.cmdaux(0).visible = visible
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
                txtAux1(i).Locked = False
            Next i
        End If
        
        

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 20)
        
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).Top = alto
            txtAux1(i).Height = DataGrid1.RowHeight
        Next i
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'chofer
        txtAux1(0).Left = DataGrid1.Left + 330
        txtAux1(0).Width = DataGrid1.Columns(2).Width - 100
        'nombre
        txtAux1(1).Width = DataGrid1.Columns(3).Width - 100
        txtAux1(1).Left = txtAux1(0).Left + (txtAux1(0).Width + 100)
        'fecha alta
        txtAux1(2).Width = DataGrid1.Columns(4).Width - 100
        txtAux1(2).Left = txtAux1(1).Left + (txtAux1(1).Width + 100)
        'fecha baja
        txtAux1(3).Width = DataGrid1.Columns(5).Width - 100
        txtAux1(3).Left = txtAux1(2).Left + (txtAux1(2).Width + 100)
        'observaciones
        txtAux1(4).Width = DataGrid1.Columns(5).Width - 100
        txtAux1(4).Left = txtAux1(3).Left + (txtAux1(3).Width + 100)
        
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).visible = visible
        Next i
    Me.cmdaux(0).Height = Me.DataGrid1.RowHeight
    Me.cmdaux(0).Top = alto
    Me.cmdaux(0).visible = visible
    cmdAux1.Top = alto
    cmdAux1.visible = visible
    End If
End Sub
Private Sub CargaTxtAux2(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte


    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux2.Count - 1 'TextBox
            txtAux2(i).Top = 290
            txtAux2(i).visible = visible
        Next i
'        cmdAux1.Top = 290
'        cmdAux1.visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid2
            For i = 0 To txtAux2.Count - 1
                txtAux2(i).Text = ""
                BloquearTxt txtAux2(i), False
            Next i
            cmdAux1.ListIndex = 0
        Else 'Vamos a modificar
            For i = 0 To txtAux2.Count - 1
                txtAux2(i).Text = DataGrid2.Columns(i + 2).Text
                txtAux2(i).Locked = False
            Next i
        End If
            If Not Adodc3.Recordset.EOF Then
                If Adodc3.Recordset.Fields(7) = "Activo" Then
                    cmdAux1.ListIndex = 0
                Else
                    cmdAux1.ListIndex = 1
                End If
            End If

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid2, 20)
        
        For i = 0 To txtAux2.Count - 1
            txtAux2(i).Top = alto
            txtAux2(i).Height = DataGrid2.RowHeight
        Next i
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'cliente
        txtAux2(0).Left = DataGrid2.Left + 330
        txtAux2(0).Width = DataGrid2.Columns(2).Width - 100
        'nombre
        txtAux2(1).Width = DataGrid2.Columns(3).Width - 100
        txtAux2(1).Left = txtAux2(0).Left + (txtAux2(0).Width + 100)
        'importe
        txtAux2(2).Width = DataGrid2.Columns(4).Width - 100
        txtAux2(2).Left = txtAux2(1).Left + (txtAux2(1).Width + 100)
        'desde
        txtAux2(3).Width = DataGrid2.Columns(5).Width - 100
        txtAux2(3).Left = txtAux2(2).Left + (txtAux2(2).Width + 100)
        'hasta
        txtAux2(4).Width = DataGrid2.Columns(6).Width - 100
        txtAux2(4).Left = txtAux2(3).Left + (txtAux2(3).Width + 100)
        
        cmdAux1.Width = DataGrid2.Columns(7).Width - 100
        cmdAux1.Left = txtAux2(4).Left + (txtAux2(4).Width + 100)
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux2.Count - 1
            txtAux2(i).visible = visible
        Next i
    End If
    Me.cmdaux(1).Height = Me.DataGrid2.RowHeight
    Me.cmdaux(1).Top = alto
    Me.cmdaux(1).visible = visible
    cmdAux1.Top = alto
    cmdAux1.visible = visible
End Sub

Private Sub BotonAnyadir()
Dim SQL As String
Dim Codigo As String
'Añadir registro en tabla de trabajadores: straba (Cabecera)

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(1).BackColor = &HFFFFC0
    PonerFoco Text1(0)
    'busco el codtarif correspondiente al menor codlista que tenga como valor
    'en bonifica=0
    SQL = "select min(codlista) from starif where bonifica=0"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'no puede ser eof
    Text1(23).Text = miRsAux.Fields(0)
    miRsAux.Close
    Set miRsAux = Nothing
    Text1(12).Text = Date
    Codigo = SugerirCodigoSiguienteStr("sclien", "codclien")
    Text1(0).Text = Codigo
End Sub
Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Pedido
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub
Private Sub BotonModificarLinea()
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String
On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If SSTab1.Tab = 2 Then
        'chofer
        If Adodc2.Recordset.EOF Then Exit Sub
        vWhere = "codsocio=" & Adodc2.Recordset!codSocio & " and numlinea=" & Adodc2.Recordset!numlinea
    
        If Not BloqueaRegistro("sclien_chofer", vWhere) Then Exit Sub
    
        CargaTxtAux True, False
        Me.lblIndicador(0).Caption = "MODIFICAR CHOFERES"
        PonerFoco txtAux1(0)
        Me.DataGrid1.Enabled = False
    Else
        'publicidad
        If Adodc3.Recordset.EOF Then Exit Sub
        vWhere = "codsocio=" & Adodc3.Recordset!codSocio & " and numlinea=" & Adodc3.Recordset!numlinea
    
        If Not BloqueaRegistro("sclien_publicidad", vWhere) Then Exit Sub
    
        CargaTxtAux2 True, False
        Me.lblIndicador(0).Caption = "MODIFICAR PUBLICIDAD"
        PonerFoco txtAux2(0)
        Me.DataGrid2.Enabled = False
    End If
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
        
    
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub
Private Sub BotonModificar()
'Prepara el Form para Modificar la cabecera de Pedidos (tabla: scaped)
Dim DeVarios As Boolean
Dim SQL As String
On Error GoTo EModificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(1)
    Text1(1).BackColor = &HFFFFC0
    Text1(2).Enabled = False
    Text1(2).BackColor = &H80000018
    
   
EModificar:
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub mnEliminar_Click()
If Modo = 5 Then
    BotonEliminarFila
Else
    BotonEliminar
End If

End Sub
Private Sub BotonEliminar()
Dim msg As String
Dim SQL As String
Dim encontrado As String

On Error GoTo EEliminar

msg = "Esta seguro que desea eliminar el socio:" & Text1(0).Text & "?"
If MsgBox(msg, vbYesNo) = vbYes Then
    NumRegElim = adodc1.Recordset.AbsolutePosition
    encontrado = DevuelveDesdeBD(conAri, "codclien", "scafac", "codclien", Text1(0).Text, "T")
    If encontrado <> "" Then
        MsgBox "No es posible eliminar este socio, ya que tiene facturas asociadas.", vbExclamation
        Exit Sub
    End If
    SQL = "Delete from sclien where codclien=" & Text1(0).Text
    conn.Execute SQL
    'Ahora borramos las lineas
    SQL = "Delete from sclien_chofer where codsocio=" & Text1(0).Text
    conn.Execute SQL
    SQL = "Delete from sclien_publicidad where codsocio=" & Text1(0).Text
    conn.Execute SQL
End If

If SituarDataTrasEliminar(adodc1, NumRegElim) Then
    PonerCampos
End If

EEliminar:
If Err.Number <> 0 Then
    MsgBox "Error al eliminar Socio." & Err.Description
End If
End Sub

Private Sub BotonEliminarFila()
Dim msg As String
Dim SQL As String

On Error GoTo EEliminarLineas

msg = "Esta seguro que desea eliminar la linea?"
If MsgBox(msg, vbQuestion + vbYesNo) = vbYes Then
    If SSTab1.Tab = 2 Then
        SQL = "Delete from sclien_chofer where codsocio=" & Text1(0).Text & " and numlinea = " & Adodc2.Recordset!numlinea
        conn.Execute SQL
        
        CargaGrid DataGrid1, Me.Adodc2
     Else
        SQL = "Delete from sclien_publicidad where codsocio=" & Text1(0).Text & " and numlinea = " & Adodc3.Recordset!numlinea
        conn.Execute SQL
        
        CargaGrid DataGrid2, Me.Adodc3

    End If
End If

EEliminarLineas:
If Err.Number <> 0 Then
    MsgBox "Error al eliminar Lineas." & Err.Description
End If

End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub
Private Sub BotonVerTodos()
'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub
Private Sub mnBuscar_Click()
   BotonBuscar
End Sub
Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        BuscaChekc = ""
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If adodc1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If adodc1.Recordset.EOF Then Exit Sub
    DesplazamientoData adodc1, Index
    PonerCampos
End Sub
Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    adodc1.RecordSource = CadenaConsulta
    adodc1.Refresh
    If adodc1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(0)
            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        adodc1.Recordset.MoveFirst
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
Private Sub PonerCampos()
Dim encontrado As String

On Error Resume Next

    
    If adodc1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, adodc1
    
    
    'data2 para el grid de las lineas chofer
    Adodc2.ConnectionString = conn
    Adodc2.RecordSource = "select sclien_chofer.codsocio,sclien_chofer.numlinea,sclien_chofer.codchofe,schofe.nomchofe,sclien_chofer.fechaalt,sclien_chofer.fechabaj,sclien_chofer.obsevac from sclien_chofer inner join schofe on sclien_chofer.codsocio=" & Text1(0).Text & " and sclien_chofer.codchofe=schofe.codchofe"
    Adodc2.Refresh
    
    CargaGrid DataGrid1, Adodc2
    'data3 para el grid de las lineas publicidad
    Adodc3.ConnectionString = conn
    Adodc3.RecordSource = "select sclien_publicidad.codsocio,sclien_publicidad.numlinea,sclien_publicidad.codclien,scliente.nomclien,sclien_publicidad.importes,sclien_publicidad.desdefec,sclien_publicidad.hastafec,if (sclien_publicidad.situacio=0, ""Activo"",""No Activo"") from sclien_publicidad inner join scliente on sclien_publicidad.codsocio= " & Text1(0).Text & " and sclien_publicidad.codclien=scliente.codclien"
    Adodc3.Refresh
    
    CargaGrid DataGrid2, Adodc3
    'data4 para el grid contadores
    Adodc4.ConnectionString = conn
    Adodc4.RecordSource = "select sclien_contadores.codsocio,sclien_contadores.codtipom,stipom.nomtipom,sclien_contadores.contador from sclien_contadores inner join stipom on sclien_contadores.codsocio=" & Text1(0).Text & " and sclien_contadores.codtipom=stipom.codtipom"
    Adodc4.Refresh
    CargaGrid DataGrid3, Adodc4
    
    If Text1(11).Text <> "" Then
        encontrado = DevuelveDesdeBD(conAri, "nomsitua", "ssitua", "codsitua", Text1(11).Text, "T")
        Text2(0).Text = encontrado
    End If
'    If Text1(19).Text <> "" Then
'        encontrado = DevuelveDesdeBD(conAri, "nombanpr", "sbanpr", "codbanpr", Text1(19).Text, "T")
'        Text2(1).Text = encontrado
'    End If
    If Text1(14).Text <> "" Then
        encontrado = DevuelveDesdeBD(conAri, "nomchofe", "scoche", "codcoche", Text1(14).Text, "T")
        Text2(1).Text = encontrado
    End If
    
    '-- Esto permanece para saber donde estamos
    lblIndicador(0).Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)   ', enlaza As Boolean)
Dim i As Integer

On Error GoTo ECargaGrid

    vData.Refresh
    Set vDataGrid.DataSource = vData
    vDataGrid.Columns(0).visible = False 'codcoche

    If vDataGrid.Name = "DataGrid1" Then
        vDataGrid.Columns(1).visible = False 'numlinea
        vDataGrid.Columns(2).Caption = "Chofer"
        vDataGrid.Columns(2).Width = 1000
        vDataGrid.Columns(2).NumberFormat = "0000"
        vDataGrid.Columns(3).Caption = "Nombre"
        vDataGrid.Columns(3).Width = 2800
        vDataGrid.Columns(4).Caption = "Fecha Alta"
        vDataGrid.Columns(4).Width = 1100
        vDataGrid.Columns(5).Caption = "Fecha Baja"
        vDataGrid.Columns(5).Width = 1100
        vDataGrid.Columns(6).Caption = "Observaciones"
        vDataGrid.Columns(6).Width = 3600
    ElseIf vDataGrid.Name = "DataGrid2" Then
        vDataGrid.Columns(1).visible = False 'numlinea
        vDataGrid.Columns(2).Caption = "Cliente"
        vDataGrid.Columns(2).Width = 1100
        vDataGrid.Columns(2).NumberFormat = "000000"
        vDataGrid.Columns(3).Caption = "Nombre"
        vDataGrid.Columns(3).Width = 3300
        vDataGrid.Columns(4).Caption = "Importe"
        vDataGrid.Columns(4).Width = 1500
        vDataGrid.Columns(4).NumberFormat = "#,###,###,##0.00"
        vDataGrid.Columns(4).Alignment = dbgRight
        vDataGrid.Columns(5).Caption = "Desde"
        vDataGrid.Columns(5).Width = 1200
        vDataGrid.Columns(6).Caption = "Hasta"
        vDataGrid.Columns(6).Width = 1200
        vDataGrid.Columns(7).Caption = "Situación"
        vDataGrid.Columns(7).Width = 1300
    Else
        vDataGrid.Columns(1).Caption = "Tipo Movimiento"
        vDataGrid.Columns(1).Width = 1500
        vDataGrid.Columns(2).Caption = "Nombre"
        vDataGrid.Columns(2).Width = 4000
        vDataGrid.Columns(3).Caption = "Contador"
        vDataGrid.Columns(3).Width = 2100
        vDataGrid.Columns(3).Alignment = dbgRight
        
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





Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    '##A mano
    cad = ""
    cad = cad & ParaGrid(Text1(0), 14, "Código")
    cad = cad & ParaGrid(Text1(2), 55, "Nombre")
    cad = cad & ParaGrid(Text1(1), 14, "Uve")
    cad = cad & ParaGrid(Text1(8), 14, "Teléfono")

    Tabla = "sclien"
    Titulo = "Socios"
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|2|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 1: dirMail = Text1(10).Text
    End Select

    If LanzaMailGnral(dirMail) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub
Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(codclien=" & Text1(0).Text & ")"
    If SituarData(adodc1, cad, Indicador) Then
       PonerModo 2
       lblIndicador(0).Caption = Indicador
        'data4 para el grid contadores
        Adodc4.ConnectionString = conn
        Adodc4.RecordSource = "select sclien_contadores.codsocio,sclien_contadores.codtipom,stipom.nomtipom,sclien_contadores.contador from sclien_contadores left join stipom on sclien_contadores.codsocio=" & Text1(0).Text & " and sclien_contadores.codtipom=stipom.codtipom"
        Adodc4.Refresh
        CargaGrid DataGrid3, Adodc4
    Else
       LimpiarCampos
       PonerModo 0
    End If
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
        Case 2, 3
            PonerFormatoFecha txtAux1(Index)
        Case 0
            If txtAux1(Index).Text <> "" Then
                txtAux1(Index).Text = Format(txtAux1(Index).Text, "0000")
                encontrado = DevuelveDesdeBD(conAri, "nomchofe", "schofe", "codchofe", txtAux1(Index).Text, "T")
                If encontrado <> "" Then
                    txtAux1(1).Text = encontrado
                Else
                    MsgBox "No existe el código de chofer introducido.", vbExclamation
                    PonerFoco txtAux1(Index)
                End If
            End If
            
        Case 4
            cmdAceptar.SetFocus
    End Select
End Sub

Private Sub txtAux2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub txtAux2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux2_LostFocus(Index As Integer)
Dim encontrado As String

txtAux2(Index).Text = UCase(txtAux2(Index).Text)

    Select Case Index
        Case 0
            If txtAux2(Index).Text <> "" Then
                txtAux2(Index).Text = Format(txtAux2(Index).Text, "000000")
                encontrado = DevuelveDesdeBD(conAri, "nomclien", "scliente", "codclien", txtAux2(Index).Text, "T")
                If encontrado <> "" Then
                    txtAux2(1).Text = encontrado
                Else
                    MsgBox "No existe el código de cliente introducido.", vbExclamation
                    PonerFoco txtAux2(Index)
                End If
            End If
        Case 3, 4
            PonerFormatoFecha txtAux2(Index)
        Case 2
            PonerFormatoDecimal txtAux2(Index), 1
    End Select
End Sub


Private Sub printNou()


    With frmImprimir2
        .cadTabla2 = "sclien"
        .Informe2 = "rGesSocios.rpt"
        If cadB1 <> "" Then
            .cadRegSelec = cadB1 'SQL2SF(cadB1)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(adodc1, Me)
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


